import math
from enum import Enum

from drop import Drop
from itemSpawn import ItemSpawn
from xlDeal import XlDeal


class E_LoveType(Enum):
    GetCard = 1
    CardLevelUp = 2
    CardStarUp = 3
    CardAwakeUp = 4
    CardPhaseUp = 5
    LoveLevUp = 6
    PlayerLevUp = 7
    DayChange = 8
    SpecialDate = 9
    SoulTrial = 10
    GetItem = 11
    Doll = 12
    MiaoGacha = 13
    CardIdAndLoveLev = 14


class LovePoint:
    LevExpMap: dict[int, int] = {}  # lev:exp
    # man*1000+lev:{'Drop','Reward'}
    LevRewardMap: dict[int, dict[str, any]] = {}
    # id:{'ManType','DateType','UnlockItem','LoveLevelCondition'}
    SpecialDateMap: dict[int, dict[str, any]] = {}
    # id:{'Contact','Type'}
    PhoneMsgMap: dict[int, dict[str, any]] = {}
    # id:{'Contact','Type'}
    PhoneCallMap: dict[int, dict[str, any]] = {}
    PhoneMomentMap: dict[int, int] = {}
    TitleMap: dict[int, int] = {}
    # id:{'RoleID','Reward'}
    ASMRMap: dict[int, dict[str, any]] = {}
    PhotoActionMap: dict[int, int] = {}
    BGMMap: dict[int, int] = {}
    # id:{'RoleList','PartEnum'}
    FashionMap: dict[int, dict[str, any]] = {}
    # id:{'man','LoveLevelCondition','Reward'}
    RadioMap: dict[int, dict[str, any]] = {}
    loveExpMap = {}  # 存储不同行为可获得的经验
    loveLimitExpMap = {}  # 存储不同行为可获得的经验的次数

    def __init__(self, _man, _player) -> None:  # 事项经验表
        if len(LovePoint.LevExpMap) == 0:
            LovePoint.LovePointInit()
        self.man = _man
        self.player = _player
        self.loveLev = 1
        self.exp = 0
        self.loveTypeExpMap = {}  # 存储不同行为获得的牵绊度经验
        self.loveDetailTimesMap = {}  # 存储详细行为的触发次数
        self.limitMap = {}
        # id:{'ManType','DateType','UnlockItem','LoveLevelCondition','Reward'}
        self.manSpecialDateMap = {}  # 保存约会和广播剧，解锁后从map中删除
        self.manRadioMap = {}
        for key, items in LovePoint.SpecialDateMap.items():
            if LovePoint.SpecialDateMap[key]["ManType"] == _man:
                self.manSpecialDateMap[key] = {}
                for item in items:
                    self.manSpecialDateMap[key][item] = items[item]
        for key, items in LovePoint.RadioMap.items():
            if LovePoint.RadioMap[key]["man"] == _man:
                self.manRadioMap[key] = {}
                for item in items:
                    self.manRadioMap[key][item] = items[item]

    def LevUp(self):
        """牵绊度升级"""
        returnMap = {}
        upFlag = False
        while True:
            if self.exp >= LovePoint.LevExpMap[self.loveLev]:
                self.loveLev += 1
                upFlag = True
                self.player.loveLevList[self.man].AddLoveExp(E_LoveType.LoveLevUp, self.loveLev)
                key = self.man * 1000 + self.loveLev
                reward = LovePoint.LevRewardMap[key]["Reward"]
                if reward is not None:
                    rewardMap = ItemSpawn.GetItemNumMap(reward)
                    for item in rewardMap:
                        if item not in returnMap:
                            returnMap[item] = rewardMap[item]
                        else:
                            returnMap[item] += rewardMap[item]
                drop = LovePoint.LevRewardMap[key]["Drop"]
                if drop is not None:
                    reward = Drop.GetDropResult(drop)
                    rewardMap = ItemSpawn.GetItemNumMap(reward)
                    for item in rewardMap:
                        if item not in returnMap:
                            returnMap[item] = rewardMap[item]
                        else:
                            returnMap[item] += rewardMap[item]
            else:
                break
        self.player.GetNewItems(returnMap, "牵绊度")
        if upFlag is True:
            self.player.loveLevList[self.man].AddLoveExp(E_LoveType.SpecialDate)
            self.player.loveLevList[self.man].AddLoveExp(E_LoveType.CardIdAndLoveLev)

    def AddLoveExp(self, addType: E_LoveType, para1=None, para2=None):
        """获得牵绊度经验

        Args:
            addType (E_LoveType):
            para1 (int): 相关参数
            para2 (int): card品质
        """
        repeatTag = False
        keyInfo = ""
        match addType:
            case E_LoveType.GetCard:
                key = str(addType.value) + "=" + str(para1)
                info = f"获得品质{para1}的思念"
            case E_LoveType.CardLevelUp:
                key = str(addType.value) + "=" + str(para1) + "=" + str(para2)
                info = f"品质{para1}的思念升至{para2}级"
            case E_LoveType.CardStarUp:
                key = str(addType.value) + "=" + str(para1) + "=" + str(para2)
                info = f"品质{para1}的思念突破{para2*10}级"
            case E_LoveType.CardAwakeUp:
                key = str(addType.value) + "=" + str(para1)
                info = f"觉醒品质{para1}的思念"
            case E_LoveType.CardPhaseUp:
                key = str(addType.value) + "=" + str(para1)
                info = f"进阶品质{para1}的思念"
            case E_LoveType.LoveLevUp:
                key = str(addType.value) + "=" + str(para1)
                info = "首次任务次数"
            case E_LoveType.PlayerLevUp:
                key = str(addType.value) + "=" + str(para1)
                info = "首次任务次数"
            case E_LoveType.DayChange:
                key = str(addType.value) + "=" + str(para1)
                info = "首次任务次数"
            case E_LoveType.SpecialDate:
                key = 0
                for id in self.manSpecialDateMap:
                    needLev = self.manSpecialDateMap[id]["LoveLevelCondition"]
                    if needLev <= self.loveLev and self.manSpecialDateMap[id]["UnlockItem"] is not None:
                        cost = self.manSpecialDateMap[id]["UnlockItem"]
                        costKey, costNum = ItemSpawn.GetItemKeyAndValue(cost)
                        if costKey in self.player.itemMap and self.player.itemMap[costKey] >= costNum:
                            self.player.itemMap[costKey] -= costNum
                            key = str(addType.value) + "=16"
                            reward = self.manSpecialDateMap[id]["Reward"]
                            if reward is not None:
                                self.player.GetNewItems(ItemSpawn.GetItemNumMap(reward), "特约")
                            del self.manSpecialDateMap[id]
                            break
                info = "特约个数"
            case E_LoveType.SoulTrial:
                level = math.ceil((para1 / 20))
                key = str(addType.value) + "=" + str(level)
                info = "定向轨道层数"
            case E_LoveType.GetItem:
                key = str(addType.value) + "=" + str(para1)
                if para2 is not None:
                    key = key + "=" + str(para2)
                if key in LovePoint.loveLimitExpMap:
                    if key not in self.limitMap:
                        self.limitMap[key] = 1
                    elif self.limitMap[key] < LovePoint.loveLimitExpMap[key]:
                        self.limitMap[key] += 1
                    else:
                        key = 0
                match para1:
                    case "20":
                        if para2 == 6:
                            info = "首次任务次数"
                            keyInfo = E_LoveType.DayChange.label
                        else:
                            info = "短信"
                    case "21":
                        if para2 == 1:
                            info = "语音电话"
                        else:
                            info = "视频电话"
                    case "22":
                        info = "朋友圈"
                    case "31":
                        info = "首次任务次数"
                        keyInfo = E_LoveType.DayChange.label
                    case "43":
                        info = "ASMR"
                    case "54":
                        info = "拍照动作"
                    case "72":
                        info = "BGM"
                    case "101":
                        info = "服饰"
            case E_LoveType.Doll:
                key = str(addType.value)
                info = "娃娃数"
            case E_LoveType.MiaoGacha:
                key = str(addType.value)
                info = "喵呜徽章数"
            case E_LoveType.CardIdAndLoveLev:
                key = 0
                for id in self.manSpecialDateMap:
                    needLev = self.manSpecialDateMap[id]["LoveLevelCondition"]
                    if needLev <= self.loveLev and id in self.player.cardMap:
                        key = str(addType.value) + "=2"
                        info = "AVG"
                        reward = self.manSpecialDateMap[id]["Reward"]
                        if reward is not None:
                            self.player.GetNewItems(ItemSpawn.GetItemNumMap(reward), "AVG")
                        del self.manSpecialDateMap[id]
                        repeatTag = True
                        break
                for id in self.manRadioMap:
                    needLev = self.manRadioMap[id]["LoveLevelCondition"]
                    if needLev <= self.loveLev and id in self.player.cardMap:
                        key = str(addType.value) + "=1"
                        info = "广播剧"
                        reward = self.manRadioMap[id]["Reward"]
                        if reward is not None:
                            self.player.GetNewItems(ItemSpawn.GetItemNumMap(reward), "广播剧")
                        del self.manRadioMap[id]
                        repeatTag = True
                        break
        if key in LovePoint.loveExpMap:
            self.exp += LovePoint.loveExpMap[key]
            if keyInfo == "":
                keyInfo = addType.label
            if keyInfo not in self.loveTypeExpMap:
                self.loveTypeExpMap[keyInfo] = LovePoint.loveExpMap[key]
            else:
                self.loveTypeExpMap[keyInfo] += LovePoint.loveExpMap[key]
            if info not in self.loveDetailTimesMap:
                self.loveDetailTimesMap[info] = 1
            else:
                self.loveDetailTimesMap[info] += 1
            if repeatTag is True:
                self.AddLoveExp(E_LoveType.CardIdAndLoveLev)
            else:
                self.LevUp()

    @classmethod
    def LovePointInit(cls):
        cls.__InitLovePointLevel()
        cls.__InitLovePointReward()
        cls.__InitSpecialDate()
        cls.__InitPhoneMsg()
        cls.__InitPhoneCall()
        cls.__InitPhoneMoment()
        cls.__InitASMR()
        cls.__InitPhotoAction()
        cls.__InitBGM()
        cls.__InitFashion()
        cls.__InitRadio()
        cls.__InitTitle()
        cls.__InitEnum()

    @classmethod
    def __InitLovePointLevel(cls):
        LovePointLevel = XlDeal("LovePointLevel.xlsx", "LovePointLevel")
        levCol = LovePointLevel.GetColIndex("Level", 2)
        expCol = LovePointLevel.GetColIndex("NextAddLove", 2)
        for row in LovePointLevel.data[3:]:
            cls.LevExpMap[int(row[levCol])] = int(row[expCol])
        del LovePointLevel

    @classmethod
    def __InitLovePointReward(cls):
        LovePointReward = XlDeal("LovePointLevel.xlsx", "LovePointReward")
        manCol = LovePointReward.GetColIndex("RoleID", 2)
        LevCol = LovePointReward.GetColIndex("LevelID", 2)
        dropCol = LovePointReward.GetColIndex("Drop", 2)
        rewardCol = LovePointReward.GetColIndex("Reward", 2)
        for row in LovePointReward.data[3:]:
            id = int(row[manCol] * 1000 + row[LevCol])
            cls.LevRewardMap[id] = {}
            cls.LevRewardMap[id]["Drop"] = row[dropCol]
            cls.LevRewardMap[id]["Reward"] = row[rewardCol]
        LovePointReward.CloseBook()
        del LovePointReward

    @classmethod
    def __InitSpecialDate(cls):
        SpecialDate = XlDeal("SpecialDate.xlsx", "SpecialDateEntry")
        SpecialDateReward = XlDeal("SpecialDate.xlsx", "SpecialDateStoryTreeProcess")
        idCol = SpecialDate.GetColIndex("ID", 2)
        skipCol = SpecialDate.GetColIndex("SkipExport", 2)
        manCol = SpecialDate.GetColIndex("ManType", 2)
        typeCol = SpecialDate.GetColIndex("DateType", 2)
        unlockCol = SpecialDate.GetColIndex("UnlockItem", 2)
        loveLevCol = SpecialDate.GetColIndex("LoveLevelCondition", 2)

        rewardIdCol = SpecialDateReward.GetColIndex("DateID", 2)
        rewardItemCol = SpecialDateReward.GetColIndex("Reward", 2)
        rewardMap = {}
        for row in SpecialDateReward.data[3:]:
            id = int(row[rewardIdCol])
            if id not in rewardMap:
                rewardMap[id] = row[rewardItemCol]
            else:
                rewardMap[id] = rewardMap[id] + "|" + row[rewardItemCol]

        for row in SpecialDate.data[3:]:
            if row[skipCol] != 1:
                id = int(row[idCol])
                cls.SpecialDateMap[id] = {}
                cls.SpecialDateMap[id]["ManType"] = int(row[manCol])
                cls.SpecialDateMap[id]["DateType"] = int(row[typeCol])
                cls.SpecialDateMap[id]["UnlockItem"] = row[unlockCol]
                cls.SpecialDateMap[id]["LoveLevelCondition"] = int(row[loveLevCol])
                if id in rewardMap:
                    cls.SpecialDateMap[id]["Reward"] = rewardMap[id]
                else:
                    cls.SpecialDateMap[id]["Reward"] = None
        SpecialDate.CloseBook()
        del SpecialDate
        del SpecialDateReward
        del rewardMap

    @classmethod
    def __InitPhoneMsg(cls):
        PhoneMsg = XlDeal("PhoneMsg.xlsx", "PhoneMsg")
        idCol = PhoneMsg.GetColIndex("ID", 2)
        manCol = PhoneMsg.GetColIndex("Contact", 2)
        skipCol = PhoneMsg.GetColIndex("SkipExport", 2)
        typeCol = PhoneMsg.GetColIndex("Type", 2)
        for row in PhoneMsg.data[3:]:
            if row[skipCol] != 1:
                id = int(row[idCol])
                LovePoint.PhoneMsgMap[id] = {}
                LovePoint.PhoneMsgMap[id]["Contact"] = int(row[manCol])
                LovePoint.PhoneMsgMap[id]["Type"] = int(row[typeCol])

        PhoneMsg.CloseBook()
        del PhoneMsg

    @classmethod
    def __InitPhoneCall(cls):
        PhoneCall = XlDeal("PhoneCall.xlsx", "PhoneCall")
        idCol = PhoneCall.GetColIndex("ID", 2)
        manCol = PhoneCall.GetColIndex("Contact", 2)
        skipCol = PhoneCall.GetColIndex("SkipExport", 2)
        typeCol = PhoneCall.GetColIndex("Type", 2)
        for row in PhoneCall.data[3:]:
            if row[skipCol] != 1:
                id = int(row[idCol])
                LovePoint.PhoneCallMap[id] = {}
                LovePoint.PhoneCallMap[id]["Contact"] = int(row[manCol])
                LovePoint.PhoneCallMap[id]["Type"] = int(row[typeCol])
        PhoneCall.CloseBook()
        del PhoneCall

    @classmethod
    def __InitPhoneMoment(cls):
        PhoneMoment = XlDeal("PhoneMoment.xlsx", "PhoneMoment")
        idCol = PhoneMoment.GetColIndex("ID", 2)
        manCol = PhoneMoment.GetColIndex("Contact", 2)
        skipCol = PhoneMoment.GetColIndex("SkipExport", 2)
        for row in PhoneMoment.data[3:]:
            if row[skipCol] != 1:
                LovePoint.PhoneMomentMap[int(row[idCol])] = int(row[manCol])
        PhoneMoment.CloseBook()
        del PhoneMoment

    @classmethod
    def __InitASMR(cls):
        ASMR = XlDeal("ASMR.xlsx", "ASMRInfo")
        idCol = ASMR.GetColIndex("ASMRID", 2)
        manCol = ASMR.GetColIndex("RoleID", 2)
        rewardCol = ASMR.GetColIndex("Reward", 2)
        for row in ASMR.data[3:]:
            id = int(row[idCol])
            cls.ASMRMap[id] = {}
            cls.ASMRMap[id]["RoleID"] = int(row[manCol])
            cls.ASMRMap[id]["Reward"] = row[rewardCol]
        ASMR.CloseBook()
        del ASMR

    @classmethod
    def __InitPhotoAction(cls):
        PhotoAction = XlDeal("Photo.xlsx", "PhotoAction")
        idCol = PhotoAction.GetColIndex("ID", 2)
        manCol = PhotoAction.GetColIndex("Role", 2)
        skipCol = PhotoAction.GetColIndex("SkipExport", 2)
        for row in PhotoAction.data[3:]:
            if row[skipCol] != 1:
                cls.PhotoActionMap[int(row[idCol])] = int(row[manCol])
        PhotoAction.CloseBook()
        del PhotoAction

    @classmethod
    def __InitBGM(cls):
        BGM = XlDeal("MainUI.xlsx", "MainUIBGM")
        idCol = BGM.GetColIndex("ID", 2)
        manCol = BGM.GetColIndex("Male", 2)
        skipCol = BGM.GetColIndex("SkipExport", 2)
        for row in BGM.data[3:]:
            if row[skipCol] != 1 and row[idCol] is not None:
                cls.BGMMap[int(row[idCol])] = int(row[manCol])
        BGM.CloseBook()
        del BGM

    @classmethod
    def __InitFashion(cls):
        Fashion = XlDeal("Fashion.xlsx", "FashionData")
        idCol = Fashion.GetColIndex("ActivateItemID", 2)
        manCol = Fashion.GetColIndex("RoleList", 2)
        skipCol = Fashion.GetColIndex("SkipExport", 2)
        partCol = Fashion.GetColIndex("PartEnum", 2)
        for row in Fashion.data[3:]:
            if row[skipCol] != 1:
                id = int(row[idCol])
                cls.FashionMap[id] = {}
                cls.FashionMap[id]["RoleList"] = row[manCol]
                cls.FashionMap[id]["PartEnum"] = row[partCol]
        Fashion.CloseBook()
        del Fashion

    @classmethod
    def __InitRadio(cls):
        shtNames = [None, "RadioInfo", "RadioInfo!!!YS", None, None, "RadioInfo!!!RY"]
        for i in range(len(shtNames)):
            if shtNames[i] is not None:
                Radio = XlDeal("Radio.xlsx", shtNames[i])
                idCol = Radio.GetColIndex("CardID", 2)
                loveLevCol = Radio.GetColIndex("LoveLevelCondition", 2)
                rewardCol = Radio.GetColIndex("Reward", 2)
                skipCol = Radio.GetColIndex("SkipExport", 2)
                for row in Radio.data[3:]:
                    if row[skipCol] != 1:
                        id = int(row[idCol])
                        cls.RadioMap[id] = {}
                        cls.RadioMap[id]["man"] = i
                        cls.RadioMap[id]["LoveLevelCondition"] = int(row[loveLevCol])
                        cls.RadioMap[id]["Reward"] = row[rewardCol]
                if i == len(shtNames) - 1:
                    Radio.CloseBook()
                del Radio

    @classmethod
    def __InitTitle(cls):
        Title = XlDeal("Title.xlsx", "Title")
        idCol = Title.GetColIndex("ID", 2)
        manCol = Title.GetColIndex("Role", 2)
        for row in Title.data[3:]:
            if row[idCol] is not None:
                LovePoint.TitleMap[int(row[idCol])] = int(row[manCol])
        Title.CloseBook()
        del Title

    @classmethod
    def __InitEnum(cls):
        E_LoveType.GetCard.label = "获得思念"
        E_LoveType.CardLevelUp.label = "思念养成"
        E_LoveType.CardStarUp.label = "思念养成"
        E_LoveType.CardAwakeUp.label = "思念养成"
        E_LoveType.CardPhaseUp.label = "获得思念"
        E_LoveType.LoveLevUp.label = "首次任务"
        E_LoveType.PlayerLevUp.label = "首次任务"
        E_LoveType.DayChange.label = "首次任务"
        E_LoveType.SpecialDate.label = "情感类"
        E_LoveType.SoulTrial.label = "战斗"
        E_LoveType.GetItem.label = "情感类"
        E_LoveType.Doll.label = "小小快乐"
        E_LoveType.MiaoGacha.label = "小小快乐"
        E_LoveType.CardIdAndLoveLev.label = "获得思念"
