import random

from card import Card
from itemSpawn import ItemSpawn
from xlDeal import XlDeal


class Gacha:
    # gachaId:{'CostTicket','Cost1','Cost10','Reward','Rule'}
    AllMap: dict[int, dict[str, any]] = {}
    # ruleId:{'CountID','RuleType','Param1','Param2','Param3',\
    #         'Priority','TotalWeight','WeightList','DropList'}
    RuleMap: dict[int, dict[str, any]] = {}
    # countId:{'CountType','Param1'}
    GuaranteeMap: dict[int, dict[str, any]] = {}
    # dropId:{'TotalWeight','WeightList','ItemIDList'}
    DropMap: dict[int, dict[str, any]] = {}

    def __init__(self, _gachaId: int, countMap: dict) -> None:
        if len(Gacha.AllMap) == 0:
            Gacha.GachaInit()
        self.gachaId = _gachaId
        self.drawTimes = 0
        _ticket = Gacha.AllMap[_gachaId]["CostTicket"]
        _ticietType = ItemSpawn.ItemMap[_ticket]
        _ticketKey = str(_ticietType) + "=" + str(_ticket)
        self.draw10Cost: dict = {_ticketKey: Gacha.AllMap[_gachaId]["Cost10"]}
        self.draw1Cost: dict = {_ticketKey: Gacha.AllMap[_gachaId]["Cost1"]}
        self.drawReward: dict = ItemSpawn.GetItemNumMap(Gacha.AllMap[_gachaId]["Reward"])
        self.ruleList = []
        self.countList = []
        for rule in Gacha.AllMap[_gachaId]["Rule"]:
            countId = Gacha.RuleMap[rule]["CountID"]
            if countId is not None:
                if countId not in self.countList:
                    self.countList.append(countId)
                if countId not in countMap:
                    countMap[countId] = 0
            if len(self.ruleList) == 0:
                self.ruleList.append(rule)
            else:
                ruleOrder = Gacha.RuleMap[rule]["Priority"]
                for i in range(len(self.ruleList)):
                    if ruleOrder < Gacha.RuleMap[self.ruleList[i]]["Priority"]:
                        self.ruleList.insert(i, rule)
                        break
                else:
                    self.ruleList.append(rule)

    def Draw1(self, drawTimes: int, countMap: dict) -> dict[str, int]:
        """抽一次卡

        Args:
            drawTimes (int): 抽卡次数

        Returns:
            dict: 抽卡行为对应的资源变化
        """
        rewardMap = {}
        for _ in range(drawTimes):
            # 获得抽到的卡
            itemKey, itemValue = self.__GetDrawResult(countMap)
            if itemKey not in rewardMap:
                rewardMap[itemKey] = itemValue
            else:
                rewardMap[itemKey] += itemValue
            # 获得每次抽卡的奖励
            for itemKey in self.drawReward:
                if itemKey not in rewardMap:
                    rewardMap[itemKey] = self.drawReward[itemKey]
                else:
                    rewardMap[itemKey] += self.drawReward[itemKey]
            # 扣除抽卡的消耗
            for itemKey in self.draw1Cost:
                if itemKey not in rewardMap:
                    rewardMap[itemKey] = -self.draw1Cost[itemKey]
                else:
                    rewardMap[itemKey] -= self.draw1Cost[itemKey]
        return rewardMap

    def Draw10(self, drawTimes: int, countMap: dict) -> dict[str, int]:
        """十连抽

        Args:
            drawTimes (int): 抽卡次数

        Returns:
            dict: 抽卡行为对应的资源变化
        """
        rewardMap = {}
        for _ in range(drawTimes):
            for _ in range(10):
                # 获得抽到的卡
                itemKey, itemValue = self.__GetDrawResult(countMap)
                if itemKey not in rewardMap:
                    rewardMap[itemKey] = itemValue
                else:
                    rewardMap[itemKey] += itemValue
                # 获得每次抽卡的奖励
                for itemKey in self.drawReward:
                    if itemKey not in rewardMap:
                        rewardMap[itemKey] = self.drawReward[itemKey]
                    else:
                        rewardMap[itemKey] += self.drawReward[itemKey]
            # 扣除抽卡的消耗
            for itemKey in self.draw10Cost:
                if itemKey not in rewardMap:
                    rewardMap[itemKey] = -self.draw10Cost[itemKey]
                else:
                    rewardMap[itemKey] -= self.draw10Cost[itemKey]
        return rewardMap

    @classmethod
    def GachaInit(cls) -> None:
        """从Gacha表读取数据"""
        cls.__InitAllMap()
        cls.__InitRuleMap()
        cls.__InitGuaranteeMap()
        cls.__InitDropMap()

    def __GetDrawResult(self, countMap) -> tuple[str, int]:
        self.drawTimes += 1
        for rule in self.ruleList:
            countId = Gacha.RuleMap[rule]["CountID"]
            if countId is not None:
                countNum = countMap[countId]
            param1 = Gacha.RuleMap[rule]["Param1"]
            param2 = Gacha.RuleMap[rule]["Param2"]
            param3 = Gacha.RuleMap[rule]["Param3"]
            totalWeight = Gacha.RuleMap[rule]["TotalWeight"]
            weightList = Gacha.RuleMap[rule]["WeightList"]
            dropList = Gacha.RuleMap[rule]["DropList"]
            match Gacha.RuleMap[rule]["RuleType"]:
                case 1:
                    if self.drawTimes == param1 and countNum < param2:
                        dropId = self.__GetRandomResult(totalWeight, weightList, dropList)
                        break
                case 2:
                    if countNum >= param1:
                        dropId = self.__GetRandomResult(totalWeight, weightList, dropList)
                        break
                case 4:
                    if countNum >= param1:
                        changeWeight = (countNum + 1 - param1) * param2
                        self.__adjustWeight(totalWeight, weightList, dropList, changeWeight, param3)
                        dropId = self.__GetRandomResult(totalWeight, weightList, dropList)
                        break
                case 0:
                    dropId = self.__GetRandomResult(totalWeight, weightList, dropList)
                    break
                case _:
                    print(rule, "规则类型无效，因为类型不是0/1/2/4！")
        totalWeight = Gacha.DropMap[dropId]["TotalWeight"]
        weightList = Gacha.DropMap[dropId]["WeightList"]
        itemList = Gacha.DropMap[dropId]["ItemIDList"]
        itemStr: str = self.__GetRandomResult(totalWeight, weightList, itemList)
        itemType, itemId, itemValue = itemStr.split("=")
        cardQuality = str(Card.CardMap[int(itemId)]["Quality"])
        for countId in self.countList:
            if Gacha.GuaranteeMap[countId]["CountType"] == 1:
                if cardQuality in Gacha.GuaranteeMap[countId]["Param1"]:
                    countMap[countId] += 1
            elif Gacha.GuaranteeMap[countId]["CountType"] == 2:
                if cardQuality not in Gacha.GuaranteeMap[countId]["Param1"]:
                    countMap[countId] += 1
                else:
                    countMap[countId] = 0
            else:
                print(countId, "保底计数无效，因为类型不是1或2！")
        # print(f"match rule {rule},dropID{dropId},cardid{itemId},rare{cardQuality}")
        return itemType + "=" + itemId, int(itemValue)

    @classmethod
    def __GetRandomResult(cls, totalWeight: int, weightList: list, dropIdList: list):
        rndValue = random.randint(1, totalWeight)
        for i in range(len(weightList)):
            if rndValue <= weightList[i]:
                return dropIdList[i]
            else:
                rndValue -= weightList[i]

    @classmethod
    def __adjustWeight(
        cls,
        totalWeight: int,
        weightList: list[int],
        dropIdList: list[int],
        changeWeight: int,
        upDrop: int,
    ) -> None:
        index = -1
        extraWeight = 0
        for i in range(len(dropIdList)):
            if dropIdList[i] == upDrop:
                index = i
                extraWeight = weightList[i]
                weightList[i] += changeWeight
                break
        for i in range(len(weightList)):
            if i != index:
                weightList[i] = weightList[i] / (totalWeight - extraWeight) * (totalWeight - weightList[index])

    @classmethod
    def __InitAllMap(cls) -> None:
        GachaAll = XlDeal("Gacha.xlsx", "GachaAll")
        idCol = GachaAll.GetColIndex("ID", 2)
        ticketCol = GachaAll.GetColIndex("CostTicket", 2)
        cost10Col = GachaAll.GetColIndex("Cost10", 2)
        cost1Col = GachaAll.GetColIndex("Cost1", 2)
        ruleCol = GachaAll.GetColIndex("Rule", 2)
        rewardCol = GachaAll.GetColIndex("Reward", 2)
        for row in GachaAll.data[3:]:
            allId = int(row[idCol])
            cls.AllMap[allId] = {}
            cls.AllMap[allId]["CostTicket"] = int(row[ticketCol])
            cls.AllMap[allId]["Cost10"] = int(row[cost10Col])
            cls.AllMap[allId]["Cost1"] = int(row[cost1Col])
            cls.AllMap[allId]["Reward"] = row[rewardCol]
            cls.AllMap[allId]["Rule"] = []
            rules = str.split(row[ruleCol], "|")
            for rule in rules:
                cls.AllMap[allId]["Rule"].append(int(rule))
        del GachaAll

    @classmethod
    def __InitRuleMap(cls) -> None:
        GachaRule = XlDeal("Gacha.xlsx", "GachaRule")
        idCol = GachaRule.GetColIndex("ID", 2)
        countCol = GachaRule.GetColIndex("CountID", 2)
        typeCol = GachaRule.GetColIndex("RuleType", 2)
        para1Col = GachaRule.GetColIndex("Param1", 2)
        para2Col = GachaRule.GetColIndex("Param2", 2)
        para3Col = GachaRule.GetColIndex("Param3", 2)
        groupCol = GachaRule.GetColIndex("RuleTypeGroup", 2)
        priorityCol = GachaRule.GetColIndex("Priority", 2)
        dropCol = GachaRule.GetColIndex("Drop", 2)
        for row in GachaRule.data[3:]:
            ruleId = int(row[idCol])
            cls.RuleMap[ruleId] = {}
            cls.RuleMap[ruleId]["CountID"] = int(row[countCol]) if row[countCol] is not None else None
            cls.RuleMap[ruleId]["RuleType"] = int(row[typeCol])
            cls.RuleMap[ruleId]["Param1"] = row[para1Col]
            cls.RuleMap[ruleId]["Param2"] = row[para2Col]
            cls.RuleMap[ruleId]["Param3"] = row[para3Col]
            ruleOrder = int(row[groupCol] * 100 - row[priorityCol])  # 数字小的优先
            cls.RuleMap[ruleId]["Priority"] = ruleOrder
            cls.RuleMap[ruleId]["TotalWeight"] = 0
            cls.RuleMap[ruleId]["WeightList"] = []
            cls.RuleMap[ruleId]["DropList"] = []
            drops = str.split(row[dropCol], "|")
            for drop in drops:
                dropId, dropWeight = str.split(drop, "=")
                dropWeight = int(dropWeight)
                cls.RuleMap[ruleId]["TotalWeight"] += dropWeight
                cls.RuleMap[ruleId]["WeightList"].append(dropWeight)
                cls.RuleMap[ruleId]["DropList"].append(int(dropId))
            drops.clear()
        del GachaRule

    @classmethod
    def __InitGuaranteeMap(cls) -> None:
        GachaGuarantee = XlDeal("Gacha.xlsx", "GuaranteeCount")
        idCol = GachaGuarantee.GetColIndex("ID", 2)
        typeCol = GachaGuarantee.GetColIndex("CountType", 2)
        para1Col = GachaGuarantee.GetColIndex("Param1", 2)
        for row in GachaGuarantee.data[3:]:
            guaranteeId = int(row[idCol])
            cls.GuaranteeMap[guaranteeId] = {}
            cls.GuaranteeMap[guaranteeId]["CountType"] = int(row[typeCol])
            cls.GuaranteeMap[guaranteeId]["Param1"] = str(row[para1Col])
        del GachaGuarantee

    @classmethod
    def __InitDropMap(cls) -> None:
        GachaDrop = XlDeal("Gacha.xlsx", "GachaDrop")
        skipCol = GachaDrop.GetColIndex("SkipExport", 2)
        groupCol = GachaDrop.GetColIndex("GroupID", 2)
        weightCol = GachaDrop.GetColIndex("Weight", 2)
        itemCol = GachaDrop.GetColIndex("ItemID", 2)
        for row in GachaDrop.data[3:]:
            if row[skipCol] is None:
                dropId = int(row[groupCol])
                if dropId not in cls.DropMap:
                    cls.DropMap[dropId] = {}
                    cls.DropMap[dropId]["TotalWeight"] = row[weightCol]
                    cls.DropMap[dropId]["WeightList"] = [row[weightCol]]
                    cls.DropMap[dropId]["ItemIDList"] = [row[itemCol]]
                else:
                    cls.DropMap[dropId]["TotalWeight"] += row[weightCol]
                    cls.DropMap[dropId]["WeightList"].append(row[weightCol])
                    cls.DropMap[dropId]["ItemIDList"].append(row[itemCol])
        GachaDrop.CloseBook()
        del GachaDrop
