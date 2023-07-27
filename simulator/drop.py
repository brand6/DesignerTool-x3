import random

from xlDeal import XlDeal


class Drop:
    # dropId:{'TotalWeight','WeightList','ItemIDList'}
    DropMap: dict[int, dict[str, any]] = {}

    @classmethod
    def DropInit(cls):
        Drop = XlDeal("Drop.xlsx", "Drop")
        skipCol = Drop.GetColIndex("SkipExport", 2)
        groupCol = Drop.GetColIndex("GroupID", 2)
        weightCol = Drop.GetColIndex("Weight", 2)
        itemCol = Drop.GetColIndex("Item", 2)
        for row in Drop.data[3:]:
            if row[skipCol] is None:
                dropId = int(row[groupCol])
                if dropId not in cls.DropMap:
                    cls.DropMap[dropId] = {}
                    cls.DropMap[dropId]["TotalWeight"] = row[weightCol]
                    cls.DropMap[dropId]["WeightList"] = [row[weightCol]]
                    cls.DropMap[dropId]["ItemList"] = [row[itemCol]]
                else:
                    cls.DropMap[dropId]["TotalWeight"] += row[weightCol]
                    cls.DropMap[dropId]["WeightList"].append(row[weightCol])
                    cls.DropMap[dropId]["ItemList"].append(row[itemCol])
        Drop.CloseBook()
        del Drop

    @classmethod
    def GetDropResult(cls, dropID):
        drop = int(dropID)
        totalWeight = Drop.DropMap[drop]["TotalWeight"]
        WeightList = Drop.DropMap[drop]["WeightList"]
        ItemList = Drop.DropMap[drop]["ItemList"]
        return cls.GetRandomResult(totalWeight, WeightList, ItemList)

    @classmethod
    def GetRandomResult(cls, totalWeight: int, weightList: list, dropIdList: list):
        rndValue = random.randint(1, totalWeight)
        for i in range(len(weightList)):
            if rndValue <= weightList[i]:
                return dropIdList[i]
            else:
                rndValue -= weightList[i]
