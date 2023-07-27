import random

from drop import Drop
from itemSpawn import ItemSpawn
from xlDeal import XlDeal


class Stage:
    # stageId:{'NeedLevel','EnterCost','LevelExp','CommonReward','LootID',\
    #          'FixNumRare','RandNum','RandRare'}
    StageMap: dict[int, dict[str, any]] = {}
    # lootId:{'ItemIDYieldNum','FixItemIDYieldNum'}
    LootMap: dict[int, dict[str, any]] = {}

    def __init__(self, _stageId: int) -> None:
        if len(Stage.StageMap) == 0:
            Stage.StageInit()
        self.stageId = _stageId
        self.costMap = ItemSpawn.GetItemNumMap(Stage.StageMap[_stageId]["EnterCost"])
        if Stage.StageMap[_stageId]["LevelExp"] is not None:
            self.expMap = ItemSpawn.GetItemNumMap(Stage.StageMap[_stageId]["LevelExp"])
        else:
            self.expMap = {"4=4": 0}
        self.unlockLev = Stage.StageMap[_stageId]["NeedLevel"]
        self.drops = []  # [[id1,times],[id2,times]...]
        commonReward = Stage.StageMap[_stageId]["CommonReward"]
        if commonReward is not None:
            drops = commonReward.split("|")
            for drop in drops:
                dropId, dropTimes = drop.split("=")
            self.drops.append([int(dropId), int(dropTimes)])

        self.loots = []
        _loot = Stage.StageMap[_stageId]["LootID"]
        if isinstance(_loot, str):
            loots = _loot.split("|")
            for loot in loots:
                self.loots.append(int(loot))
        else:
            self.loots = [int(_loot)]
        self.fixNumRare = Stage.StageMap[_stageId]["FixNumRare"]
        self.randNum = Stage.StageMap[_stageId]["RandNum"]
        self.randRare = Stage.StageMap[_stageId]["RandRare"]

    def GetReward(self, times: int = 1):
        resMap = {}
        # 扣除消耗
        for cost in self.costMap:
            if cost not in resMap:
                resMap[cost] = 0
            resMap[cost] -= self.costMap[cost] * times
        # 获得经验
        for exp in self.expMap:
            if exp not in resMap:
                resMap[exp] = self.expMap[exp] * times
            else:
                resMap[exp] += self.expMap[exp] * times
        for _ in range(times):
            for drop in self.drops:
                for _ in self.drops[drop]:
                    items = ItemSpawn.GetItemNumMap(
                        self.__GetRandomResult(
                            Drop.DropMap[drop]["TotalWeight"],
                            Drop.DropMap[drop]["WeightList"],
                            Drop.DropMap[drop]["ItemList"],
                        )
                    )
                    for item in items:
                        if item not in resMap:
                            resMap[item] = items[item]
                        else:
                            resMap[item] += items[item]
            for loot in self.loots:
                items = [
                    Stage.LootMap[loot]["ItemIDYieldNum"],
                    Stage.LootMap[loot]["FixItemIDYieldNum"],
                ]
                for itemStr in items:
                    if itemStr is not None:
                        itemList = itemStr.split("|")
                        for item in itemList:
                            itemId, itemNum = item.split("=")
                            itemType = ItemSpawn.ItemMap[int(itemId)]
                            itemKey = str(itemType) + "=" + itemId
                            if float(itemNum) - int(itemNum) > 0:
                                rnd = random.random()
                                if rnd < (float(itemNum) - int(itemNum)):
                                    itemNum = int(itemNum) + 1
                                else:
                                    itemNum = int(itemNum)
                            else:
                                itemNum = int(itemNum)
                            if itemKey not in resMap:
                                resMap[itemKey] = itemNum
                            else:
                                resMap[itemKey] += itemNum
        return resMap

    @classmethod
    def __GetRandomResult(cls, totalWeight: int, weightList: list, dropIdList: list):
        rndValue = random.randint(1, totalWeight)
        for i in range(len(weightList)):
            if rndValue <= weightList[i]:
                return dropIdList[i]
            else:
                rndValue -= weightList[i]

    @classmethod
    def StageInit(cls):
        cls.__CommonStageInit()
        cls.__LootInit()

    @classmethod
    def __CommonStageInit(cls):
        CommonStage = XlDeal("CommonStageEntry.xlsx", "CommonStageEntry")
        idCol = CommonStage.GetColIndex("ID", 2)
        skipCol = CommonStage.GetColIndex("SkipExport", 2)
        levCol = CommonStage.GetColIndex("NeedLevel", 2)
        costCol = CommonStage.GetColIndex("EnterCost", 2)
        expCol = CommonStage.GetColIndex("LevelExp", 2)
        rewardCol = CommonStage.GetColIndex("CommonReward", 2)
        lootCol = CommonStage.GetColIndex("LootID", 2)
        randFixCol = CommonStage.GetColIndex("FixNumRare", 2)
        randNumCol = CommonStage.GetColIndex("RandNum", 2)
        randRareCol = CommonStage.GetColIndex("RandRare", 2)
        for row in CommonStage.data[3:]:
            if row[skipCol] is None:
                stageId = int(row[idCol])
                cls.StageMap[stageId] = {}
                cls.StageMap[stageId]["NeedLevel"] = int(row[levCol]) if row[levCol] is not None else None
                cls.StageMap[stageId]["EnterCost"] = row[costCol]
                cls.StageMap[stageId]["LevelExp"] = row[expCol]
                cls.StageMap[stageId]["CommonReward"] = row[rewardCol]
                cls.StageMap[stageId]["LootID"] = row[lootCol]
                cls.StageMap[stageId]["FixNumRare"] = row[randFixCol]
                cls.StageMap[stageId]["RandNum"] = row[randNumCol]
                cls.StageMap[stageId]["RandRare"] = row[randRareCol]
        CommonStage.CloseBook()
        del CommonStage

    @classmethod
    def __LootInit(cls):
        Loot = XlDeal("Loot.xlsx", "Loot")
        idCol = Loot.GetColIndex("LootID", 2)
        numCol = Loot.GetColIndex("ItemIDYieldNum", 2)
        fixNumCol = Loot.GetColIndex("FixItemIDYieldNum", 2)
        for row in Loot.data[3:]:
            lootId = int(row[idCol])
            cls.LootMap[lootId] = {}
            cls.LootMap[lootId]["ItemIDYieldNum"] = row[numCol]
            cls.LootMap[lootId]["FixItemIDYieldNum"] = row[fixNumCol]
        Loot.CloseBook()
        del Loot
