import random

from drop import Drop
from itemSpawn import ItemSpawn
from xlDeal import XlDeal


class HangUp:
    # id:{'Cost','SpeedUpCost','ExRewardProb','GuaranteeTimes','ExReward',\
    #     'TotalWeight','WeightList','DropList'}
    HangUpMap: dict[int, dict[str, any]] = {}

    def __init__(self, _exploreId) -> None:
        if len(HangUp.HangUpMap) == 0:
            HangUp.HangUpInit()
        self.exploreId = _exploreId
        self.exCount = 0
        self.cost = ItemSpawn.GetItemNumMap(HangUp.HangUpMap[_exploreId]["Cost"])
        self.speedUpCost = ItemSpawn.GetItemNumMap(HangUp.HangUpMap[_exploreId]["SpeedUpCost"])
        self.exRewardProb = HangUp.HangUpMap[_exploreId]["ExRewardProb"]
        self.exReward = HangUp.HangUpMap[_exploreId]["ExReward"]
        self.guaranteeTimes = HangUp.HangUpMap[_exploreId]["GuaranteeTimes"]
        self.totalWeight = HangUp.HangUpMap[_exploreId]["TotalWeight"]
        self.weightList = HangUp.HangUpMap[_exploreId]["WeightList"]
        self.DropList = HangUp.HangUpMap[_exploreId]["DropList"]

    def Explore(self, exploreTimes: int, isSpeedUp=False) -> dict:
        rewardMap = {}
        for _ in range(exploreTimes):
            # 扣除消耗
            for itemKey in self.cost:
                if itemKey not in rewardMap:
                    rewardMap[itemKey] = -self.cost[itemKey]
                else:
                    rewardMap[itemKey] -= self.cost[itemKey]
            if isSpeedUp:
                for itemKey in self.speedUpCost:
                    if itemKey not in rewardMap:
                        rewardMap[itemKey] = -self.cost[itemKey]
                    else:
                        rewardMap[itemKey] -= self.cost[itemKey]
            # 抽奖
            items = ItemSpawn.GetItemNumMap(self.__GetExploreResult())
            for itemKey in items:
                if itemKey not in rewardMap:
                    rewardMap[itemKey] = items[itemKey]
                else:
                    rewardMap[itemKey] += items[itemKey]
            # 处理额外奖励
            if self.exRewardProb is not None:
                self.exCount += 1
                exTag = False
                if self.exCount >= self.guaranteeTimes:
                    exTag = True
                elif random.randint(1, 1000) < self.exRewardProb:
                    exTag = True
                if exTag is True:
                    self.exCount = 0
                    drop = self.exReward
                    itemStr = self.__GetRandomResult(
                        Drop.DropMap[drop]["TotalWeight"],
                        Drop.DropMap[drop]["WeightList"],
                        Drop.DropMap[drop]["ItemList"],
                    )
                    items = ItemSpawn.GetItemNumMap(itemStr)
                    for itemKey in items:
                        if itemKey not in rewardMap:
                            rewardMap[itemKey] = items[itemKey]
                        else:
                            rewardMap[itemKey] += items[itemKey]
                else:
                    self.exCount += 1
        return rewardMap

    def __GetExploreResult(self) -> str:
        drop = self.__GetRandomResult(self.totalWeight, self.weightList, self.DropList)
        return self.__GetRandomResult(
            Drop.DropMap[drop]["TotalWeight"],
            Drop.DropMap[drop]["WeightList"],
            Drop.DropMap[drop]["ItemList"],
        )

    @classmethod
    def __GetRandomResult(cls, totalWeight: int, weightList: list, dropIdList: list):
        rndValue = random.randint(1, totalWeight)
        for i in range(len(weightList)):
            if rndValue <= weightList[i]:
                return dropIdList[i]
            else:
                rndValue -= weightList[i]

    @classmethod
    def HangUpInit(cls):
        HangUpReward = XlDeal("HangUpReward.xlsx", "HangUpReward")
        idCol = HangUpReward.GetColIndex("ExploreID", 2)
        costCol = HangUpReward.GetColIndex("Cost", 2)
        speedUpCol = HangUpReward.GetColIndex("SpeedUpCost", 2)
        rewardCol = HangUpReward.GetColIndex("Reward", 2)
        exProbCol = HangUpReward.GetColIndex("ExRewardProb", 2)
        exRewardCol = HangUpReward.GetColIndex("ExReward", 2)
        exGuaranteeCol = HangUpReward.GetColIndex("GuaranteeTimes", 2)
        for row in HangUpReward.data[3:]:
            id = int(row[idCol])
            cls.HangUpMap[id] = {}
            cls.HangUpMap[id]["Cost"] = row[costCol]
            cls.HangUpMap[id]["SpeedUpCost"] = row[speedUpCol]
            cls.HangUpMap[id]["ExRewardProb"] = row[exProbCol]
            cls.HangUpMap[id]["GuaranteeTimes"] = row[exGuaranteeCol]
            cls.HangUpMap[id]["ExReward"] = int(row[exRewardCol]) if row[exRewardCol] is not None else None
            cls.HangUpMap[id]["TotalWeight"] = 0
            cls.HangUpMap[id]["WeightList"] = []
            cls.HangUpMap[id]["DropList"] = []
            drops = str.split(row[rewardCol], "|")
            for drop in drops:
                dropId, dropWeight = str.split(drop, "=")
                dropWeight = int(dropWeight)
                cls.HangUpMap[id]["TotalWeight"] += dropWeight
                cls.HangUpMap[id]["WeightList"].append(dropWeight)
                cls.HangUpMap[id]["DropList"].append(int(dropId))
            drops.clear()
        HangUpReward.CloseBook()
        del HangUpReward
