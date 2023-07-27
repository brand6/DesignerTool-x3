import random

from itemSpawn import ItemSpawn
from lovePoint import E_LoveType
from xlDeal import XlDeal


class MiaoGacha:
    # id:{'DropGroupID','Cost1'}
    PackMap = {}
    DropMap: dict[int, list] = {}
    # id:{'ModuleList','Reward'}
    LibraryMap = {}
    ActivityPoolMap = {}  # 活动池数据

    def __init__(self, _packId, _player) -> None:
        if len(MiaoGacha.PackMap) == 0:
            MiaoGacha.MiaoGachaInit()
        self.packId = _packId
        self.player = _player
        _dropId = MiaoGacha.PackMap[_packId]["DropGroupID"]
        self.itemList = []
        for item in MiaoGacha.DropMap[_dropId]:  # 卡池剩余的列表
            self.itemList.append(item)
        self.getList = []  # 拥有的列表

    def Draw(self):
        rndValue = random.randrange(len(self.itemList))
        itemStr = self.itemList[rndValue]
        self.itemList.remove(itemStr)
        item = itemStr.split("=")[1]
        if item not in self.getList:
            self.getList.append(item)
            for i in range(1, 6):
                if self.player.loveLevList[i] is not None:
                    self.player.loveLevList[i].AddLoveExp(E_LoveType.MiaoGacha)
                    self.__GetLibraryReward(item)

    def __GetLibraryReward(self, card):
        for id in MiaoGacha.LibraryMap:
            if str(card) in MiaoGacha.LibraryMap[id]["ModuleList"]:
                for module in MiaoGacha.LibraryMap[id]["ModuleList"].split("|"):
                    if int(module) not in self.getList:
                        break
                else:
                    self.player.GetNewItems(ItemSpawn.GetItemNumMap(MiaoGacha.LibraryMap[id]["Reward"]), "喵呜集卡")

    @classmethod
    def MiaoGachaInit(cls):
        cls.__InitGachaPack()
        cls.__InitGachaDrop()
        cls.__InitGachaLibrary()

    @classmethod
    def __InitGachaPack(cls):
        GachaPack = XlDeal("MiaoGacha.xlsx", "MiaoGachaPack")
        idCol = GachaPack.GetColIndex("ID", 2)
        dropCol = GachaPack.GetColIndex("DropGroupID", 2)
        costCol = GachaPack.GetColIndex("Cost1", 2)
        for row in GachaPack.data[3:]:
            id = int(row[idCol])
            MiaoGacha.PackMap[id] = {}
            MiaoGacha.PackMap[id]["DropGroupID"] = int(row[dropCol])
            MiaoGacha.PackMap[id]["Cost1"] = row[costCol]
        del GachaPack

    @classmethod
    def __InitGachaDrop(cls):
        GachaDrop = XlDeal("MiaoGacha.xlsx", "MiaoGachaDropALL")
        idCol = GachaDrop.GetColIndex("DropGroupID", 2)
        itemCol = GachaDrop.GetColIndex("ItemID", 2)
        numCol = GachaDrop.GetColIndex("DropNum", 2)
        for row in GachaDrop.data[3:]:
            id = int(row[idCol])
            if id not in MiaoGacha.DropMap:
                MiaoGacha.DropMap[id] = []
            for _ in range(int(row[numCol])):
                MiaoGacha.DropMap[id].append(row[itemCol])
        del GachaDrop

    @classmethod
    def __InitGachaLibrary(cls):
        GachaLibrary = XlDeal("MiaoGacha.xlsx", "MiaoGachaPackLibrary")
        idCol = GachaLibrary.GetColIndex("ID", 2)
        listCol = GachaLibrary.GetColIndex("ModuleList", 2)
        rewardCol = GachaLibrary.GetColIndex("Reward", 2)
        for row in GachaLibrary.data[3:]:
            if row[idCol] is not None:
                id = int(row[idCol])
                MiaoGacha.LibraryMap[id] = {}
                MiaoGacha.LibraryMap[id]["ModuleList"] = row[listCol]
                MiaoGacha.LibraryMap[id]["Reward"] = row[rewardCol]
        GachaLibrary.CloseBook()
        del GachaLibrary
