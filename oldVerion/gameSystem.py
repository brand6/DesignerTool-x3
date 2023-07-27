from itemSpawn import ItemSpawn


class GameSystem:
    def __init__(self, _name) -> None:
        self.name: str = _name
        self.itemPawns: list[ItemSpawn] = []
        self.itemMap: dict[str, int] = {}

    def AddItemSpawn(self, itemSpawn: ItemSpawn) -> None:
        """给系统添加道具添加器

        Args:
            itemSpawn (ItemSpawn): 道具添加器
        """
        self.itemPawns.append(itemSpawn)

    def GetItemsByDay(self, day, lastDay) -> dict[str, int]:
        """根据天数从道具添加器获取奖励

        Args:
            day (int): 开服天数
        """
        self.itemMap.clear()
        for itemSpawn in self.itemPawns:
            rewardList, costList = itemSpawn.GetItemsByDay(day, lastDay)
            if rewardList is not None:
                self.AddItemListToItemMap(rewardList, self.itemMap)
            if costList is not None:
                self.AddItemListToItemMap(costList, self.itemMap, False)
        return self.itemMap

    @classmethod
    def AddItemListToItemMap(cls, itemList, itemMap, isAdd=True) -> None:
        """将道具list添加到道具map

        Args:
            itemList (_type_): 道具list
            itemMap (_type_): 道具map
            isAdd (bool, optional): 添加或减少. Defaults to True.
        """
        for reward in itemList:
            key = str(reward[0]) + "=" + str(reward[1])
            if isAdd is False:
                reward[2] = -reward[2]
            if key not in itemMap:
                itemMap[key] = reward[2]
            else:
                itemMap[key] += reward[2]
