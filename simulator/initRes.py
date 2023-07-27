from itemSpawn import ItemSpawn
from xlDeal import XlDeal


class InitRes:
    InitResMap = {}
    shareResMap = {}

    @classmethod
    def GetInitRes(cls):
        if len(cls.InitResMap) == 0:
            cls.__InitResMap()
        return cls.InitResMap

    @classmethod
    def GetShareRes(cls):
        if len(cls.shareResMap) == 0:
            cls.__InitShareMap()
        return cls.shareResMap

    @classmethod
    def InitResInit(cls):
        cls.__InitResMap()
        cls.__InitShareMap()

    @classmethod
    def __InitResMap(cls):
        InitItem = XlDeal("CreateRoleInitItem.xlsx", "CreateRoleInitItem")
        itemCol = InitItem.GetColIndex("InitItem", 2)
        skipCol = InitItem.GetColIndex("SkipExport", 2)
        for row in InitItem.data[3:]:
            if row[itemCol] is not None and row[skipCol] != 1:
                items = ItemSpawn.GetItemNumMap(row[itemCol])
                for key in items:
                    if key not in cls.InitResMap:
                        cls.InitResMap[key] = items[key]
                    else:
                        cls.InitResMap[key] += items[key]
        InitItem.CloseBook()
        del InitItem

    @classmethod
    def __InitShareMap(cls):
        Share = XlDeal("Share.xlsx", "ShareReward")
        itemCol = Share.GetColIndex("Reward", 2)
        for row in Share.data[3:]:
            if row[itemCol] is not None:
                items = ItemSpawn.GetItemNumMap(row[itemCol])
                for key in items:
                    if key not in cls.shareResMap:
                        cls.shareResMap[key] = items[key]
                    else:
                        cls.shareResMap[key] += items[key]
        Share.CloseBook()
        del Share
