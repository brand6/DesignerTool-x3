from card import Card
from cardGacha import Gacha
from hangUp import HangUp
from itemSpawn import ItemSpawn
from lovePoint import E_LoveType, LovePoint
from miaoGacha import MiaoGacha
from soulTrial import SoulTrial
from stage import Stage
from xlDeal import XlDeal


class Player:
    # lev:{'NextAddExp','LevelPresent'}
    LevMap: dict[int, dict[str, any]] = {}
    TurnParaMap = {}

    def __init__(self, _name, _man="1|2|5", _dealCardRare=3) -> None:
        """_summary_

        Args:
            _name (_type_): 玩家名\n
            _man (_type_, optional): 培养的男主. Defaults to 1|2|5.\n
            _dealCardRare (int, optional): 最低培养的卡牌品质. Defaults to 3.
        """
        if len(Player.LevMap) == 0:
            Player.PlayerInit()
        self.name = _name
        self.playerLev = 1
        self.developMan = str(_man)
        self.developRare = int(_dealCardRare)  # 预期培养的最低品质
        self.developLevs = [2] * 6  # 预期培养等级[同tag第一张，第二张...]
        self.developNumPerTag = 1  # 预期培养数量
        self.developCardLevLimit = 20  # 当前卡牌养成等级上限

        self.itemSpawnList: list[ItemSpawn] = []  # 可获得奖励的系统
        self.itemMap: dict[str, int] = {}  # 拥有的道具
        self.cardMap: dict[int, Card] = {}  # 拥有的card
        self.miaoGachaMap: dict[int, MiaoGacha] = {}  # 喵喵徽章情况
        self.dollMap: dict[int, int] = {}  # 娃娃情况

        self.countMap = {}  # 保底计数可跨卡池，记在玩家身上
        self.gachaMap: dict[int, Gacha] = {}  # 抽过的卡池
        self.hangUpMap: dict[int, HangUp] = {}  # 抽过的挂机池
        self.totalCardGachaTimes = 0  # 总抽卡数
        self.newGacha = 0
        self.isNewFinish = False  # 新手卡池是否抽完

        self.soulTrialList = [None] * 6  # 定向轨道情况
        self.loveLevList: list[LovePoint] = [None] * 6  # 牵绊度情况
        for m in range(6):
            if m != 3 and m != 4:
                self.soulTrialList[m] = SoulTrial(m, self)
                self.loveLevList[m] = LovePoint(m, self)

        self.stageMap: dict[str, Stage] = {}  # 资源本情况
        self.stageMap["1=1"] = Stage(40101)
        self.stageMap["201=0"] = Stage(40201)
        self.stageMap["1|4"] = Stage(40301)
        self.stageMap["2|5"] = Stage(40401)
        self.stageMap["3|6"] = Stage(40501)

        self.itemChangeSystem = {}  # 用于存储道具变化相关的系统

        # 拥有的不同品质的card列表
        self.rareCardList: list[list[int]] = [[] for _ in range(5)]
        # 拥有的不同男主的card列表
        self.manCardList: list[list[int]] = [[] for _ in range(6)]
        # 当前培养的card列表，按tag分类
        self.dealTagCardList: list[list[int]] = [[] for _ in range(7)]
        # 指定的培养列表
        self.targetCardList = []

    def DevelopCard(self) -> None:
        """培养卡牌"""
        fullCardTag = True
        for i in range(self.developNumPerTag):
            lastCardId = 0
            while True:
                for rare in range(4, self.developRare - 1, -1):
                    for tag in range(1, 7):
                        if len(self.dealTagCardList[tag]) > i:
                            cardId = self.dealTagCardList[tag][i]
                            lastCardId = cardId
                            if self.cardMap[cardId].rare == rare:
                                while self.cardMap[cardId].cardLev < self.developLevs[i]:
                                    if self.__TryDevelopCard(cardId) is False:
                                        break
                        else:
                            fullCardTag = False
                if (
                    lastCardId > 0
                    and self.cardMap[lastCardId].cardLev == self.developLevs[i]
                    and self.developLevs[i] < self.developCardLevLimit
                ):
                    self.developLevs[i] += 1
                else:
                    break
        if (
            lastCardId > 0
            and fullCardTag is True
            and self.cardMap[lastCardId].cardLev == self.developLevs[self.developNumPerTag - 1]
            and self.cardMap[lastCardId].cardLev == self.developCardLevLimit
        ):
            self.developNumPerTag += 1
            self.UpdateDevelopCardList()

    def __TryDevelopCard(self, cardId) -> bool:
        _, info = self.cardMap[cardId].LevelUp(self.itemMap)
        match info:
            case "reachMaxLev":
                _, info2 = self.cardMap[cardId].StarUp(self.playerLev, self.itemMap)
                match info2:
                    case "reachMaxStar":
                        return False
                    case "unMatchPlayerLev":
                        return False
                    case "lackGold":
                        if self.itemMap["3=3"] > 8:
                            self.GetStageReward("1=1")
                        else:
                            return False
                    case "lackItem":
                        cardTag = self.cardMap[cardId].tag
                        if cardTag > 3:
                            stageKey = str(cardTag - 3) + "|" + str(cardTag)
                        else:
                            stageKey = str(cardTag) + "|" + str(cardTag + 3)
                        if self.itemMap["3=3"] > 8:
                            self.GetStageReward(stageKey)
                        else:
                            return False
            case "lackExp":
                if self.itemMap["3=3"] > 8:
                    self.GetStageReward("201=0")
                else:
                    return False
        return True

    def DevelopTargetCard(self) -> None:
        """培养卡牌"""
        while True:
            developCard = self.targetCardList[0]
            for card in self.targetCardList[1:]:
                if self.cardMap[card].cardLev < self.cardMap[developCard].cardLev:
                    developCard = card
            if self.__TryDevelopCard(developCard) is False:
                break

    def UpdateDevelopCardList(self) -> None:
        """更新养成列表"""
        for rare in range(4, self.developRare - 1, -1):
            for cardId in self.rareCardList[rare]:
                cardTag = self.cardMap[cardId].tag
                cardMan = self.cardMap[cardId].man
                cardLev = self.cardMap[cardId].cardLev
                if str(cardMan) in self.developMan and cardId not in self.dealTagCardList[cardTag]:
                    if len(self.dealTagCardList[cardTag]) < self.developNumPerTag:
                        self.dealTagCardList[cardTag].append(cardId)
                    else:
                        for i in range(len(self.dealTagCardList[cardTag])):
                            _cardId = self.dealTagCardList[cardTag][i]
                            if self.cardMap[_cardId].rare < rare:
                                self.dealTagCardList[cardTag][i] = cardId
                                break
                            elif self.cardMap[_cardId].rare == rare and self.cardMap[_cardId].cardLev < cardLev:
                                self.dealTagCardList[cardTag][i] = cardId
                                break
        self.SortDevelopCardList()
        self.DevelopCard()

    def SortDevelopCardList(self) -> None:
        """对养成列表进行排序，品质高的在前面，同品质等级高的在前面"""
        for cardList in self.dealTagCardList:
            cardList.sort(key=lambda x: self.cardMap[x].rare * 100 + self.cardMap[x].cardLev, reverse=True)

    def levUp(self):
        """玩家升级"""
        changeTag = False
        while True:
            if self.itemMap["4=4"] >= Player.LevMap[self.playerLev]["NextAddExp"]:
                self.itemMap["4=4"] -= Player.LevMap[self.playerLev]["NextAddExp"]
                self.playerLev += 1
                for i in range(1, 6):
                    if self.loveLevList[i] is not None:
                        self.loveLevList[i].AddLoveExp(E_LoveType.PlayerLevUp, self.playerLev)
                self.GetNewItems(ItemSpawn.GetItemNumMap(Player.LevMap[self.playerLev]["LevelPresent"]), "玩家升级")
                changeTag = True
            else:
                break
        # 更新玩家的关卡数据
        if changeTag is True:
            for key in list(self.stageMap.keys()):
                checkId = self.stageMap[key].stageId
                while True:
                    checkId += 1
                    if checkId in Stage.StageMap and Stage.StageMap[checkId]["NeedLevel"] <= self.playerLev:
                        pass
                    else:
                        checkId -= 1
                        break
                if checkId != self.stageMap[key].stageId:
                    del self.stageMap[key]
                    self.stageMap[key] = Stage(checkId)
            for i in range(len(Card.StarMap[201]) - 1):
                if self.playerLev >= Card.StarMap[201][i]["PlayerLevel"]:
                    self.developCardLevLimit = Card.StarMap[201][i + 1]["LevelLimit"]
                else:
                    break

    def GetStageReward(self, stageKey, times: int = 1):
        rewardMap = self.stageMap[stageKey].GetReward(times)
        self.GetNewItems(rewardMap, "关卡")

    def DrawCard(self, gachaId: int, drawTimes: int, is10=False) -> None:
        """单次抽卡

        Args:
            gachaId (int): 卡池id
            drawTimes (int): 抽卡次数
            is10：是否进行10连抽
        """
        if drawTimes > 0:
            if gachaId not in self.gachaMap:
                self.gachaMap[gachaId] = Gacha(gachaId, self.countMap)
            if is10 is True:
                self.totalCardGachaTimes += 10
                rewardMap = self.gachaMap[gachaId].Draw10(drawTimes, self.countMap)
            else:
                self.totalCardGachaTimes += 1
                rewardMap = self.gachaMap[gachaId].Draw1(drawTimes, self.countMap)
            self.GetNewItems(rewardMap, "卡池" + str(gachaId))

    def HangUpExplore(self, exploreId: int, drawTimes: int) -> None:
        """挂机抽奖

        Args:
            exploreId (int): 挂机池id
            drawTimes (int): 抽卡次数
        """
        if drawTimes > 0:
            if exploreId not in self.hangUpMap:
                self.hangUpMap[exploreId] = HangUp(exploreId)
            rewardMap = self.hangUpMap[exploreId].Explore(drawTimes)
            self.GetNewItems(rewardMap, "挂机")

    def AddItemSpawn(self, itemSpawn: ItemSpawn) -> None:
        """添加可获得奖励的道具生成器

        Args:
            itemSpawn (ItemSpawn): 系统
        """
        self.itemSpawnList.append(itemSpawn)

    def UpdateItemsByDay(self, day, lastDay) -> None:
        """根据天数更新拥有的奖励

        Args:
            day (int): 开服天数
            lastDay(int)：上次统计的天数
        """
        for itemSpawn in self.itemSpawnList:
            rewardList, costList = itemSpawn.GetItemsByDay(day, lastDay)
            if costList is not None:
                self.CostItems(ItemSpawn.GetItemNumMap(costList), itemSpawn.system, isForce=True)
            if rewardList is not None:
                self.GetNewItems(ItemSpawn.GetItemNumMap(rewardList), itemSpawn.system)

    def CostItems(self, costItems: dict, system: str, isForce=False) -> bool:
        """玩家消耗道具

        Args:
            costItems (dict): 需要消耗的道具
            system (str): 消耗道具的系统
            isForce(bool)：是否强制扣除道具
        Returns:
            bool: 是否扣除道具成功
        """
        # 检查是否拥有足够的资源
        if isForce is False:
            for item in costItems.keys():
                if item not in self.itemMap:
                    return False
                elif self.itemMap[item] < costItems[item]:
                    return False
        # 实际扣除消耗
        if system not in self.itemChangeSystem:
            self.itemChangeSystem[system] = {}
        for item in costItems.keys():
            if item not in self.itemMap:
                self.itemMap[item] = -costItems[item]
            else:
                self.itemMap[item] -= costItems[item]
            if item not in self.itemChangeSystem[system]:
                self.itemChangeSystem[system][item] = -costItems[item]
            else:
                self.itemChangeSystem[system][item] -= costItems[item]
        return True

    def GetNewItems(self, newItems: dict, system) -> None:
        """玩家获得新的道具

        Args:
            newItems (dict): 新道具
            system(str):道具来源的系统
        """
        if system not in self.itemChangeSystem:
            self.itemChangeSystem[system] = {}
        for item in newItems.keys():
            if item not in self.itemChangeSystem[system]:
                self.itemChangeSystem[system][item] = newItems[item]
            else:
                self.itemChangeSystem[system][item] += newItems[item]
            itemType, itemId = item.split("=")
            itemId = int(itemId)
            if itemType == "51":
                self.GetNewCard(itemId, newItems[item], system)
            else:
                # 资源转换
                if item in Player.TurnParaMap:
                    after = Player.TurnParaMap[item][0]
                    if after not in self.itemMap:
                        self.itemMap[after] = Player.TurnParaMap[item][1] * newItems[item]
                    else:
                        self.itemMap[after] += Player.TurnParaMap[item][1] * newItems[item]
                else:
                    if item not in self.itemMap:
                        self.itemMap[item] = newItems[item]
                    else:
                        self.itemMap[item] += newItems[item]
                    if item == "4=4":
                        self.levUp()
                # 牵绊度
                match itemType:
                    case "16":
                        for key in LovePoint.SpecialDateMap:
                            if (
                                LovePoint.SpecialDateMap[key]["UnlockItem"] is not None
                                and item in LovePoint.SpecialDateMap[key]["UnlockItem"]
                            ):
                                man = LovePoint.SpecialDateMap[key]["ManType"]
                                self.loveLevList[man].AddLoveExp(E_LoveType.SpecialDate)
                    case "20":
                        if itemId in LovePoint.PhoneMsgMap:
                            man = LovePoint.PhoneMsgMap[itemId]["Contact"]
                            if man > 0 and man < 6 and self.loveLevList[man] is not None:
                                msgType = LovePoint.PhoneMsgMap[itemId]["Type"]
                                self.loveLevList[man].AddLoveExp(E_LoveType.GetItem, itemType, msgType)
                    case "21":
                        man = LovePoint.PhoneCallMap[itemId]["Contact"]
                        if man > 0 and man < 6 and self.loveLevList[man] is not None:
                            callType = LovePoint.PhoneCallMap[itemId]["Type"]
                            self.loveLevList[man].AddLoveExp(E_LoveType.GetItem, itemType, callType)
                    case "22":
                        man = LovePoint.PhoneMomentMap[itemId]
                        if man > 0 and man < 6 and self.loveLevList[man] is not None:
                            self.loveLevList[man].AddLoveExp(E_LoveType.GetItem, itemType)
                    case "31":
                        man = LovePoint.TitleMap[itemId]
                        if man > 0 and man < 6 and self.loveLevList[man] is not None:
                            self.loveLevList[man].AddLoveExp(E_LoveType.GetItem, itemType)
                    case "43":
                        reward = LovePoint.ASMRMap[itemId]["Reward"]
                        if reward is not None:
                            self.GetNewItems(ItemSpawn.GetItemNumMap(reward))
                        man = int(LovePoint.ASMRMap[itemId]["RoleID"])
                        if man > 0 and man < 6 and self.loveLevList[man] is not None:
                            self.loveLevList[man].AddLoveExp(E_LoveType.GetItem, itemType)
                    case "54":
                        man = LovePoint.PhotoActionMap[itemId]
                        if man > 0 and man < 6 and self.loveLevList[man] is not None:
                            self.loveLevList[man].AddLoveExp(E_LoveType.GetItem, itemType)
                    case "72":
                        man = LovePoint.BGMMap[itemId]
                        if man > 0 and man < 6 and self.loveLevList[man] is not None:
                            self.loveLevList[man].AddLoveExp(E_LoveType.GetItem, itemType)
                    case "101":
                        man = LovePoint.FashionMap[itemId]["RoleList"]
                        if not isinstance(man, str):
                            man = int(man)
                            part = LovePoint.FashionMap[itemId]["PartEnum"]
                            if part == 1 and man > 0 and man < 6 and self.loveLevList[man] is not None:
                                self.loveLevList[man].AddLoveExp(E_LoveType.GetItem, itemType, part)

    def GetNewCard(self, cardId: int, cardNum: int, system="") -> None:
        """处理获取的card

        Args:
            itemKey (int): 51 = cardId
            itemNum (int): cardNum
        """
        if cardId not in self.cardMap:
            self.cardMap[cardId] = Card(cardId, self)
            man: int = self.cardMap[cardId].man
            rare: int = self.cardMap[cardId].rare
            self.rareCardList[rare].append(cardId)
            self.manCardList[man].append(cardId)
            self.loveLevList[man].AddLoveExp(E_LoveType.GetCard, rare)
            self.loveLevList[man].AddLoveExp(E_LoveType.CardIdAndLoveLev)
            if rare == 4 and system == "卡池" + str(self.newGacha):
                self.isNewFinish = True
            cardNum -= 1
        for _ in range(cardNum):
            self.cardMap[cardId].SplitCard(self.itemMap)
        # print(f"getNewCard{cardId},rare={self.cardMap[cardId].rare}")

    def GetTotalCardLevel(self) -> int:
        """获得总升级次数

        Returns:
            int: _description_
        """
        totalLevel = 0
        for card in self.cardMap:
            totalLevel += self.cardMap[card].cardLev - 1
        return totalLevel

    def GetTotalCardPhase(self, rare=2) -> int:
        """获得指定品质及以上的card总进阶次数

        Returns:
            _type_: _description_
        """
        totalPhase = 0
        for card in self.cardMap:
            if self.cardMap[card].rare >= rare:
                totalPhase += self.cardMap[card].cardPhase
        return totalPhase

    def GetAllCard(self, manList=[]):
        """获得所有的卡"""
        for cardId in Card.CardMap:
            if len(manList) == 0:
                self.GetNewCard(cardId, 1, "GM")
            else:
                for _man in manList:
                    if Card.CardMap[cardId]["ManType"] == _man:
                        self.GetNewCard(cardId, 1, "GM")
                        break

    def AddTargetCard(self, rare):
        order = len(self.targetCardList) % 6
        if order % 2 == 0:
            tag = order / 2 + 1
        else:
            tag = (order - 1) / 2 + 4
        for card in self.rareCardList[rare]:
            if card not in self.targetCardList and self.cardMap[card].tag == tag:
                self.targetCardList.append(card)
                break

    @classmethod
    def PlayerInit(cls) -> None:
        PlayerLevel = XlDeal("PlayerLevel.xlsx", "PlayerLevel")
        idCol = PlayerLevel.GetColIndex("Level", 2)
        expCol = PlayerLevel.GetColIndex("NextAddExp", 2)
        presentCol = PlayerLevel.GetColIndex("LevelPresent", 2)
        for row in PlayerLevel.data[3:]:
            levId = int(row[idCol])
            cls.LevMap[levId] = {}
            cls.LevMap[levId]["NextAddExp"] = row[expCol]
            cls.LevMap[levId]["LevelPresent"] = row[presentCol]
        PlayerLevel.CloseBook()
        del PlayerLevel
