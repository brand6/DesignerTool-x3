from itemSpawn import ItemSpawn
from lovePoint import E_LoveType
from xlDeal import XlDeal


class Card:
    # cardId:{'ManType','PosType','FormationTag','Quality','ExpMode','StarID',
    #         'AwakeID','PhaseMode','CardRewardGroup','SuitRewardGroup'}
    CardMap: dict[int, dict[str, any]] = {}
    # ExpMode:[exp1,...]
    LevelMap: dict[int, list[int]] = {}
    # StarId:[{'LevelLimit','PlayerLevel','ItemCost','GoldCost'},...]
    StarMap: dict[int, list[dict[str, any]]] = {}
    # awakeId:{'AwakeNeedLv','NeedItem','NeedGold'}
    AwakeMap: dict[int, dict[str, any]] = {}
    # phaseId:[cost1,...]
    PhaseMap: dict[int, list[int]] = {}
    # groupId:{'Type','Param','Reward'}
    RewardMap: dict[int, dict[str, any]] = {}
    # quality:{'BreakReward','FragmentNum','ComposeMoneyCost'}
    RareMap: dict[int, dict[str, any]] = {}

    def __init__(self, _cardId, _player) -> None:
        if len(Card.CardMap) == 0:
            Card.CardInit()
        self.cardId = _cardId
        self.player = _player
        self.cardLev = 1
        self.cardStar = 0
        self.cardAwake = 0
        self.cardPhase = 0
        self.man = Card.CardMap[_cardId]["ManType"]
        self.rare = Card.CardMap[_cardId]["Quality"]
        self.tag = Card.CardMap[_cardId]["FormationTag"]
        self.pos = Card.CardMap[_cardId]["PosType"]
        self.expMode = Card.CardMap[_cardId]["ExpMode"]
        self.starId = Card.CardMap[_cardId]["StarID"]
        self.awakeId = Card.CardMap[_cardId]["AwakeID"]
        self.phaseMode = Card.CardMap[_cardId]["PhaseMode"]
        self.cardReward = Card.CardMap[_cardId]["CardRewardGroup"]
        self.suitReward = Card.CardMap[_cardId]["SuitRewardGroup"]
        self._maxLev = Card.StarMap[self.starId][self.cardStar]["LevelLimit"]

    def LevelUp(self, resMap: dict) -> tuple[bool, str]:
        """卡牌升级

        Args:
            resMap (dict): 玩家拥有的资源Map
        Returns:
            bool: 升级是否成功，失败原因
        """
        if self.cardLev >= self._maxLev:
            return False, "reachMaxLev"
        else:
            needExp = Card.LevelMap[self.expMode][self.cardLev - 1]
            if "201=0" not in resMap or resMap["201=0"] < needExp:
                return False, "lackExp"
            else:
                resMap["201=0"] -= needExp
                self.cardLev += 1
                self.player.loveLevList[self.man].AddLoveExp(E_LoveType.CardLevelUp, self.rare, self.cardLev)
                return True, ""

    def StarUp(self, playerLev: int, resMap: dict) -> tuple[bool, str]:
        """卡牌突破

        Args:
            playerLev (int): 玩家等级
            resMap (dict): 玩家拥有的资源

        Returns:
            tuple[bool, str]:  突破是否成功，失败原因
        """
        starMap = Card.StarMap[self.starId][self.cardStar]
        if self.cardStar == len(Card.StarMap[self.starId]) - 1:
            return False, "reachMaxStar"
        elif self.cardLev < self._maxLev:
            return False, "unMatchMaxLev"
        elif playerLev < starMap["PlayerLevel"]:
            return False, "unMatchPlayerLev"
        elif "1=1" not in resMap or resMap["1=1"] < starMap["GoldCost"]:
            return False, "lackGold"
        else:
            needItems = ItemSpawn.GetItemNumMap(starMap["ItemCost"])
            if ItemSpawn.CheckNeedItem(needItems, resMap) is False:
                return False, "lackItem"
            else:
                resMap["1=1"] -= starMap["GoldCost"]
                ItemSpawn.ReduceNeedItem(needItems, resMap)
                self.cardStar += 1
                self.player.loveLevList[self.man].AddLoveExp(E_LoveType.CardStarUp, self.rare, self.cardStar)
                self._maxLev = Card.StarMap[self.starId][self.cardStar]["LevelLimit"]
                return True, ""

    def AwakeUp(self, resMap: dict) -> tuple[bool, str]:
        """卡牌觉醒

        Args:
            resMap (dict): 玩家拥有的资源

        Returns:
            tuple[bool, str]: 觉醒是否成功，失败原因
        """
        if self.cardAwake == 1:
            return False, "reachMaxAwake"
        elif self.cardLev < Card.AwakeMap[self.awakeId]["AwakeNeedLv"]:
            return False, "unMatchMaxLev"
        elif "1=1" not in resMap or resMap["1=1"] < Card.AwakeMap[self.awakeId]["NeedGold"]:
            return False, "lackGold"
        else:
            needItems = ItemSpawn.GetItemNumMap(Card.AwakeMap[self.awakeId]["NeedItem"])
            if ItemSpawn.CheckNeedItem(needItems, resMap) is False:
                return False, "lackItem"
            else:
                resMap["1=1"] -= Card.AwakeMap[self.awakeId]["NeedGold"]
                ItemSpawn.ReduceNeedItem(needItems, resMap)
                self.cardAwake += 1
                self.player.loveLevList[self.man].AddLoveExp(E_LoveType.CardAwakeUp, self.rare)
                return True, ""

    def PhaseUp(self, resMap: dict) -> tuple[bool, str]:
        if self.cardPhase == len(Card.PhaseMap[self.phaseMode]):
            return False, "reachMaxPhase"
        else:
            costItem = "203=" + str(self.cardId + 200000)
            if costItem not in resMap or resMap[costItem] < Card.PhaseMap[self.phaseMode][self.cardPhase]:
                return False, "lackItem"
            else:
                resMap[costItem] -= Card.PhaseMap[self.phaseMode][self.cardPhase]
                self.cardPhase += 1
                self.player.loveLevList[self.man].AddLoveExp(E_LoveType.CardPhaseUp, self.rare)
                return True, ""

    def SplitCard(self, resMap: dict) -> None:
        """没满阶时分解card成碎片，并尝试进阶，否则销毁成材料

        Args:
            resMap (dict): 玩家拥有的资源
        """

        # 判断是否已进阶满
        _, info = self.PhaseUp(resMap)
        if info == "reachMaxPhase":
            self.BreakCard(resMap)
        else:
            splitNum = Card.RareMap[self.rare]["FragmentNum"]
            itemKey = "203=" + str(self.cardId + 200000)
            if itemKey not in resMap:
                resMap[itemKey] = splitNum
            else:
                resMap[itemKey] += splitNum
            self.PhaseUp(resMap)

    def BreakCard(self, resMap) -> None:
        """销毁card成材料

        Args:
            resMap (dict): 玩家拥有的资源
        """
        breakReward = ItemSpawn.GetItemNumMap(Card.RareMap[self.rare]["BreakReward"])
        for item in breakReward:
            if item not in resMap:
                resMap[item] = breakReward[item]
            else:
                resMap[item] += breakReward[item]

    @classmethod
    def CardInit(cls) -> None:
        """从Card表读取数据"""
        cls.__InitCardMap()
        cls.__InitRewardMap()
        cls.__InitLevelMap()
        cls.__InitStarMap()
        cls.__InitAwakeMap()
        cls.__InitPhaseMap()
        cls.__InitRareMap()

    @classmethod
    def __InitCardMap(cls) -> None:
        CardBase = XlDeal("Card.xlsx", "CardBaseInfo")
        idCol = CardBase.GetColIndex("ID", 2)
        manCol = CardBase.GetColIndex("ManType", 2)
        posCol = CardBase.GetColIndex("PosType", 2)
        tagCol = CardBase.GetColIndex("FormationTag", 2)
        qualityCol = CardBase.GetColIndex("Quality", 2)
        expCol = CardBase.GetColIndex("ExpMode", 2)
        starCol = CardBase.GetColIndex("StarID", 2)
        awakeCol = CardBase.GetColIndex("AwakeID", 2)
        phaseCol = CardBase.GetColIndex("PhaseMode", 2)
        cardRewardCol = CardBase.GetColIndex("CardRewardGroup", 2)
        suitRewardCol = CardBase.GetColIndex("SuitRewardGroup", 2)
        for row in CardBase.data[3:]:
            cardId = int(row[idCol])
            Card.CardMap[cardId] = {}
            Card.CardMap[cardId]["ManType"] = int(row[manCol])
            Card.CardMap[cardId]["PosType"] = int(row[posCol])
            Card.CardMap[cardId]["FormationTag"] = int(row[tagCol])
            Card.CardMap[cardId]["Quality"] = int(row[qualityCol])
            Card.CardMap[cardId]["ExpMode"] = int(row[expCol])
            Card.CardMap[cardId]["StarID"] = int(row[starCol])
            Card.CardMap[cardId]["AwakeID"] = int(row[awakeCol])
            Card.CardMap[cardId]["PhaseMode"] = int(row[phaseCol])
            Card.CardMap[cardId]["CardRewardGroup"] = int(row[cardRewardCol]) if row[cardRewardCol] is not None else None
            Card.CardMap[cardId]["SuitRewardGroup"] = int(row[suitRewardCol]) if row[suitRewardCol] is not None else None
        del CardBase

    @classmethod
    def __InitRewardMap(cls) -> None:
        CardReward = XlDeal("Card.xlsx", "CardReward")
        groupCol = CardReward.GetColIndex("RewardGroup", 2)
        typeCol = CardReward.GetColIndex("Type", 2)
        paramCol = CardReward.GetColIndex("Param", 2)
        rewardCol = CardReward.GetColIndex("Reward", 2)
        for row in CardReward.data[3:]:
            groupId = int(row[groupCol])
            Card.RewardMap[groupId] = {}
            Card.RewardMap[groupId]["Type"] = row[typeCol]
            Card.RewardMap[groupId]["Param"] = row[paramCol]
            Card.RewardMap[groupId]["Reward"] = row[rewardCol]
        del CardReward

    @classmethod
    def __InitLevelMap(cls) -> None:
        CardLevel = XlDeal("Card.xlsx", "CardLevelExp")
        modeCol = CardLevel.GetColIndex("ExpMode", 2)
        expCol = CardLevel.GetColIndex("NextLvExp", 2)
        for row in CardLevel.data[3:]:
            modeId = int(row[modeCol])
            if modeId not in Card.LevelMap:
                Card.LevelMap[modeId] = []
            Card.LevelMap[modeId].append(row[expCol])
        del CardLevel

    @classmethod
    def __InitStarMap(cls) -> None:
        CardStar = XlDeal("Card.xlsx", "CardStar")
        starCol = CardStar.GetColIndex("StarID", 2)
        levLimitCol = CardStar.GetColIndex("LevelLimit", 2)
        playerLevCol = CardStar.GetColIndex("PlayerLevel", 2)
        itemCostCol = CardStar.GetColIndex("ItemCost", 2)
        goldCostCol = CardStar.GetColIndex("GoldCost", 2)
        for row in CardStar.data[3:]:
            starId = int(row[starCol])
            if starId not in Card.StarMap:
                Card.StarMap[starId] = []
            Card.StarMap[starId].append(
                {
                    "LevelLimit": row[levLimitCol],
                    "PlayerLevel": row[playerLevCol],
                    "ItemCost": row[itemCostCol],
                    "GoldCost": row[goldCostCol],
                }
            )
        del CardStar

    @classmethod
    def __InitAwakeMap(cls) -> None:
        CardAwake = XlDeal("Card.xlsx", "CardAwake")
        idCol = CardAwake.GetColIndex("AwakeID", 2)
        levCol = CardAwake.GetColIndex("AwakeNeedLv", 2)
        itemCol = CardAwake.GetColIndex("NeedItem", 2)
        goldCol = CardAwake.GetColIndex("NeedGold", 2)
        for row in CardAwake.data[3:]:
            if row[idCol] is not None:
                awakeId = int(row[idCol])
                Card.AwakeMap[awakeId] = {}
                Card.AwakeMap[awakeId]["AwakeNeedLv"] = row[levCol]
                Card.AwakeMap[awakeId]["NeedItem"] = row[itemCol]
                Card.AwakeMap[awakeId]["NeedGold"] = row[goldCol]
        del CardAwake

    @classmethod
    def __InitPhaseMap(cls) -> None:
        CardPhase = XlDeal("Card.xlsx", "CardPhase")
        modeCol = CardPhase.GetColIndex("Mode", 2)
        numCol = CardPhase.GetColIndex("CostSelfNum", 2)
        for row in CardPhase.data[3:]:
            if row[modeCol] is not None:
                modeId = int(row[modeCol])
                if modeId not in Card.PhaseMap:
                    Card.PhaseMap[modeId] = []
                Card.PhaseMap[modeId].append(row[numCol])
        del CardPhase

    @classmethod
    def __InitRareMap(cls) -> None:
        CardRare = XlDeal("Card.xlsx", "CardRare")
        idCol = CardRare.GetColIndex("Id", 2)
        breakCol = CardRare.GetColIndex("BreakReward", 2)
        numCol = CardRare.GetColIndex("FragmentNum", 2)
        goldCol = CardRare.GetColIndex("ComposeMoneyCost", 2)
        for row in CardRare.data[3:]:
            if row[idCol] is not None:
                rareId = int(row[idCol])
                Card.RareMap[rareId] = {}
                Card.RareMap[rareId]["BreakReward"] = row[breakCol]
                Card.RareMap[rareId]["FragmentNum"] = row[numCol]
                Card.RareMap[rareId]["ComposeMoneyCost"] = row[goldCol]
        CardRare.CloseBook()
        del CardRare
