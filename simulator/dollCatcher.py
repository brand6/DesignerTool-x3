import random

from itemSpawn import ItemSpawn
from lovePoint import E_LoveType
from player import Player
from xlDeal import XlDeal


class DollCatcher:
    # id:{'DollTotalNum','DollTypeMaxNum','DollPool'}
    DifficultyMap: dict[int, dict[str, any]] = {}
    # id:{'UnattainedNormalDollPR','ChangeColorPR'}
    TotalPoolMap: dict[int, dict[str, any]] = {}
    # dropId:{'TotalWeight','WeightList','DollIDList'}
    DollDropMap = {}
    # dollId:{'TotalWeight','WeightList','ColorDollIDList'}
    ColorMap = {}
    CollectionMap = {}  # 存储娃娃收集奖励
    CatcherAvgNum = 1.5
    ChangeColorNum = 5
    # week：{}
    DollCatcherMap: dict[int, list] = {}  # 存储不同周的娃娃池

    def __init__(self, _catcherId: int, _dropId: int) -> None:
        if len(DollCatcher.DifficultyMap) == 0:
            DollCatcher.DollCatcherInit()
        self.catcherId = _catcherId
        self.colorCondition = DollCatcher.ChangeColorNum
        self.dropId = _dropId
        self.avgGetNum = DollCatcher.CatcherAvgNum
        self.dollMaxNum = DollCatcher.DifficultyMap[_catcherId]["DollTotalNum"]
        self.typeMaxNum = DollCatcher.DifficultyMap[_catcherId]["DollTypeMaxNum"]
        self.dollPool = DollCatcher.DifficultyMap[_catcherId]["DollPool"]
        self.unGetRate = DollCatcher.TotalPoolMap[self.dollPool]["UnattainedNormalDollPR"]
        self.colorRate = DollCatcher.TotalPoolMap[self.dollPool]["ChangeColorPR"]

    def GetDoll(self, player: Player) -> list:
        returnList = []
        if self.dollMaxNum == 1:
            for _ in range(5):
                dollList = self.GetDollPool(player.dollMap)
                rndValue = random.random()
                if rndValue < self.avgGetNum / 5:
                    returnList.append(dollList[0])
        else:
            dollList = self.GetDollPool(player.dollMap)
            for _ in range(5):
                rndValue = random.random()
                if rndValue < self.avgGetNum / 5:
                    returnList.append(dollList[0])
                    dollList.remove(dollList[0])

        for doll in dollList:
            if doll not in player.dollMap:
                player.dollMap[doll] = 1
                for i in range(1, 6):
                    if player.loveLevList[i] is not None:
                        player.loveLevList[i].AddLoveExp(E_LoveType.Doll)
                        dollNum = len(player.dollMap)
                        if dollNum in DollCatcher.CollectionMap:
                            player.GetNewItems(ItemSpawn.GetItemNumMap(DollCatcher.CollectionMap[dollNum]), "娃娃收集")
                        elif len(DollCatcher.CollectionMap) == 0:
                            print("DollCatcher.CollectionMap 初始化失败")
            else:
                player.dollMap[doll] = +1

    def GetDollPool(self, dollMap: dict):
        dollList = DollCatcher.DollDropMap[self.dropId]["DollIDList"]
        returnList = []
        appearList = []
        for _ in range(self.dollMaxNum):
            getTag = False
            if len(appearList) < self.typeMaxNum:
                for doll in dollList:
                    if doll not in dollMap and doll not in returnList:
                        rndValue = random.randint(1, 1000)
                        if rndValue <= self.unGetRate:
                            returnList.insert(0, doll)
                            if doll not in appearList:
                                appearList.append(doll)
                            getTag = True
                            break
                if getTag is False:
                    TotalWeight = DollCatcher.DollDropMap[self.dropId]["TotalWeight"]
                    WeightList = DollCatcher.DollDropMap[self.dropId]["WeightList"]
                    doll = DollCatcher.__GetRandomResult(TotalWeight, WeightList, dollList)
                    if doll not in appearList:
                        appearList.append(doll)
                    colorDoll = self.__tryColor(doll, dollMap)
                    if colorDoll != 0:
                        returnList.insert(0, colorDoll)
                    else:
                        returnList.append(doll)
            else:
                rndValue = random.randrange(len(appearList))
                doll = appearList[rndValue]
                colorDoll = self.__tryColor(doll, dollMap)
                if colorDoll != 0:
                    returnList.insert(0, colorDoll)
                else:
                    returnList.append(doll)
        return returnList

    def __tryColor(self, doll, dollMap):
        if doll in dollMap and dollMap[doll] > self.colorCondition:
            rndValue = random.randint(1, 1000)
            if rndValue <= self.colorRate:
                _TotalWeight = DollCatcher.ColorMap[doll]["TotalWeight"]
                _WeightList = DollCatcher.ColorMap[doll]["WeightList"]
                _ColorDollIDList = DollCatcher.ColorMap[doll]["ColorDollIDList"]
                colorDoll = DollCatcher.__GetRandomResult(_TotalWeight, _WeightList, _ColorDollIDList)
                return colorDoll
        return 0

    @classmethod
    def __GetRandomResult(cls, totalWeight: int, weightList: list, dropIdList: list):
        rndValue = random.randint(1, totalWeight)
        for i in range(len(weightList)):
            if rndValue <= weightList[i]:
                return dropIdList[i]
            else:
                rndValue -= weightList[i]

    @classmethod
    def DollCatcherInit(cls):
        cls.__InitDifficulty()
        cls.__InitTotalPool()
        cls.__InitDollDrop()
        cls.__InitDollColor()
        cls.__InitDollCollection()

    @classmethod
    def __InitDifficulty(cls):
        Difficulty = XlDeal("UFOCatcher.xlsx", "UFOCatcherDifficulty")
        idCol = Difficulty.GetColIndex("ID", 2)
        dollNumCol = Difficulty.GetColIndex("DollTotalNum", 2)
        dollMaxTypeCol = Difficulty.GetColIndex("DollTypeMaxNum", 2)
        dollPoolCol = Difficulty.GetColIndex("DollPool", 2)
        for row in Difficulty.data[3:]:
            id = int(row[idCol])
            cls.DifficultyMap[id] = {}
            cls.DifficultyMap[id]["DollTotalNum"] = int(row[dollNumCol])
            cls.DifficultyMap[id]["DollTypeMaxNum"] = int(row[dollMaxTypeCol])
            cls.DifficultyMap[id]["DollPool"] = int(row[dollPoolCol])
        Difficulty.CloseBook()
        del Difficulty

    @classmethod
    def __InitTotalPool(cls):
        TotalPool = XlDeal("UFOCatcherDollPool.xlsx", "UFOCatcherTotalPool")
        idCol = TotalPool.GetColIndex("PoolID", 2)
        rateCol = TotalPool.GetColIndex("UnattainedNormalDollPR", 2)
        colorCol = TotalPool.GetColIndex("ChangeColorPR", 2)
        for row in TotalPool.data[3:]:
            id = int(row[idCol])
            cls.TotalPoolMap[id] = {}
            cls.TotalPoolMap[id]["UnattainedNormalDollPR"] = int(row[rateCol])
            cls.TotalPoolMap[id]["ChangeColorPR"] = int(row[colorCol])
        del TotalPool

    @classmethod
    def __InitDollDrop(cls):
        DollDrop = XlDeal("UFOCatcherDollPool.xlsx", "UFOCatcherDollDrop")
        idCol = DollDrop.GetColIndex("GroupID", 2)
        dollIdCol = DollDrop.GetColIndex("DollID", 2)
        weightCol = DollDrop.GetColIndex("Weight", 2)
        for row in DollDrop.data[3:]:
            if row[idCol] is not None:
                dropId = int(row[idCol])
                if dropId not in cls.DollDropMap:
                    cls.DollDropMap[dropId] = {}
                    cls.DollDropMap[dropId]["TotalWeight"] = row[weightCol]
                    cls.DollDropMap[dropId]["WeightList"] = [row[weightCol]]
                    cls.DollDropMap[dropId]["DollIDList"] = [int(row[dollIdCol])]
                else:
                    cls.DollDropMap[dropId]["TotalWeight"] += row[weightCol]
                    cls.DollDropMap[dropId]["WeightList"].append(row[weightCol])
                    cls.DollDropMap[dropId]["DollIDList"].append(int(row[dollIdCol]))
        del DollDrop

    @classmethod
    def __InitDollColor(cls):
        DollColor = XlDeal("UFOCatcherDollPool.xlsx", "UFOCatcherDollColor")
        idCol = DollColor.GetColIndex("OriginDollID", 2)
        colorIdCol = DollColor.GetColIndex("ColorDollID", 2)
        weightCol = DollColor.GetColIndex("Weight", 2)
        for row in DollColor.data[3:]:
            dropId = int(row[idCol])
            if dropId not in cls.ColorMap:
                cls.ColorMap[dropId] = {}
                cls.ColorMap[dropId]["TotalWeight"] = row[weightCol]
                cls.ColorMap[dropId]["WeightList"] = [row[weightCol]]
                cls.ColorMap[dropId]["ColorDollIDList"] = [int(row[colorIdCol])]
            else:
                cls.ColorMap[dropId]["TotalWeight"] += row[weightCol]
                cls.ColorMap[dropId]["WeightList"].append(row[weightCol])
                cls.ColorMap[dropId]["ColorDollIDList"].append(int(row[colorIdCol]))
        DollColor.CloseBook()
        del DollColor

    @classmethod
    def __InitDollCollection(cls):
        DollCollection = XlDeal("GalleryManual.xlsx", "GalleryDollCollection")
        numCol = DollCollection.GetColIndex("Num", 2)
        rewardCol = DollCollection.GetColIndex("Reward", 2)
        for row in DollCollection.data[3:]:
            if row[numCol] is not None:
                cls.CollectionMap[int(row[numCol])] = row[rewardCol]
        DollCollection.CloseBook()
        del DollCollection
