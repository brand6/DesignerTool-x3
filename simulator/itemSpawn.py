import math

from xlDeal import XlDeal


class ItemSpawn:
    leftWeekDays = 0
    leftMonthDays = 0
    ItemMap: dict[int, int] = {}

    def __init__(self, _rewardStr: str, _system: str, _unLockDay: int, _period=None, _getTimes=None, _costStr=None) -> None:
        self.getTimes = _getTimes if _getTimes is not None else 1
        self.period = _period if _period is not None else 0
        self.unLockDay = _unLockDay
        self.system = _system
        self.rewardList = self.SplitItemStr(_rewardStr)
        if _costStr is not None:
            self.costList = self.SplitItemStr(_costStr)
        else:
            self.costList = []

    def GetItemsByDay(self, day, lastDay) -> tuple[list[list[int]], list[list[int]]] | tuple[None, None]:
        """获取奖励的有效次数

        Args:
            day (int): 当前开服天数
        Returns:
            list[list[int]], list[list[int]]: 获得的奖励，消耗的道具
        """
        getTimes = 0
        if day >= self.unLockDay:
            if self.period == 0:  # 单次奖励
                if lastDay < self.unLockDay:
                    getTimes = self.getTimes
            else:
                match self.period:
                    case 7:
                        leftDay = self.leftWeekDays
                    case 30:
                        leftDay = self.leftMonthDays
                    case _:
                        leftDay = 1
                dayTimes = self.__GetPeriodRewardTimes(day, leftDay)
                lastDayTimes = self.__GetPeriodRewardTimes(lastDay, leftDay)
                getTimes = (dayTimes - lastDayTimes) * self.getTimes
        if getTimes > 0:
            return (
                self.GetItemListTimes(self.rewardList, getTimes),
                self.GetItemListTimes(self.costList, getTimes),
            )
        else:
            return (None, None)

    def __GetPeriodRewardTimes(self, day: int, leftDay: int) -> int:
        curCount = math.ceil((day + self.period - leftDay) / self.period)
        unlockCount = math.ceil((self.unLockDay + self.period - leftDay) / self.period)
        return curCount - unlockCount + 1

    @classmethod
    def GetItemListTimes(cls, items: list | str, getTimes=1) -> list[list[int]]:
        """获取多次itemList的奖励

        Args:
            itemList (_type_): 物品列表
            getTimes (_type_): 次数

        Returns:
            list: _description_
        """
        returnList = []
        if isinstance(items, str):
            items = cls.SplitItemStr(items)
        for item in items:
            returnList.append([item[0], item[1], item[2] * getTimes])
        return returnList

    @classmethod
    def ReduceNeedItem(cls, needItemMap, resMap) -> None:
        """扣除需要消耗的物品数量

        Args:
            needItemMap (_type_): 需要消耗的物品数量
            resMap (_type_): 拥有的物品数量

        """
        for item in needItemMap:
            resMap[item] -= needItemMap[item]

    @classmethod
    def CheckNeedItem(cls, needItemMap, resMap) -> bool:
        """检测是否满足需要消耗的物品数量

        Args:
            needItemMap (_type_): 需要消耗的物品数量
            resMap (_type_): 拥有的物品数量

        Returns:
            bool: 是否满足
        """
        for item in needItemMap:
            if item not in resMap or needItemMap[item] > resMap[item]:
                return False
        else:
            return True

    @classmethod
    def GetItemKeyAndValue(cls, items: str) -> tuple[str, int] | None:
        """获取item的key和数量

        Args:
            items (str): 道具字符串

        Returns:
            tuple[str, int] | None: key，value
        """
        items = cls.SplitItemStr(items)
        for item in items:
            itemKey = str(item[0]) + "=" + str(item[1])
            itemValue = int(item[2])
            return itemKey, itemValue

    @classmethod
    def GetItemNumMap(cls, items: str | list) -> dict[str, int]:
        """将items转为map

        Args:
            items (str | list): 道具字符串或list

        Returns:
            Dict: itemNumMap
        """
        if isinstance(items, str):
            items = cls.SplitItemStr(items)
        returnMap = {}
        for item in items:
            itemKey = str(item[0]) + "=" + str(item[1])
            if itemKey not in returnMap:
                returnMap[itemKey] = item[2]
            else:
                returnMap[itemKey] += item[2]
        return returnMap

    @classmethod
    def SplitItemStr(cls, itemStr: str) -> list[list[int]]:
        returnList = []
        itemList: list[str] = itemStr.split("|")
        for item in itemList:
            itemType, itemId, itemValue = item.split("=")
            returnList.append([int(itemType), int(itemId), int(itemValue)])

        itemList.clear()
        return returnList

    @classmethod
    def GetWeek(cls, days):
        """获取天数对应的周数

        Args:
            days (_type_): _description_

        Returns:
            _type_: _description_
        """
        return math.ceil((days + 7 - cls.leftWeekDays) / 7)

    @classmethod
    def GetMonth(cls, days):
        """获取天数对应的月份

        Args:
            days (_type_): _description_

        Returns:
            _type_: _description_
        """
        return math.ceil((days + 30 - cls.leftMonthDays) / 30)

    @classmethod
    def UpdateLeftWeekDays(cls, days) -> None:
        cls.leftWeekDays = days

    @classmethod
    def UpdateLeftMonthDays(cls, days) -> None:
        cls.leftMonthDays = days

    @classmethod
    def ItemInit(cls):
        Item = XlDeal("Item.xlsx", "Item")
        idCol = Item.GetColIndex("ID", 2)
        skipCol = Item.GetColIndex("SkipExport", 2)
        typeCol = Item.GetColIndex("Type", 2)
        for row in Item.data[3:]:
            if row[skipCol] is None:
                cls.ItemMap[int(row[idCol])] = int(row[typeCol])
        Item.CloseBook()
        del Item
