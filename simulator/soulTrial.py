from itemSpawn import ItemSpawn
from lovePoint import E_LoveType
from xlDeal import XlDeal


class SoulTrial:
    # man*1000+floor:reward
    rewardMap: dict[int, str] = {}
    PowerMap = {}
    MaxFloor = 0

    def __init__(self, _man, _player) -> None:
        if len(SoulTrial.rewardMap) == 0:
            SoulTrial.SoulTrialInit()
        self.man = _man
        self.player = _player
        self.floor = 0

    def TryUp(self):
        base = self.player.developLevs[0]
        rareRate = 1
        phaseRate = 1
        for tag in range(1, 7):
            if len(self.player.dealTagCardList[tag]) > 0:
                card = self.player.dealTagCardList[tag][0]
                if self.player.cardMap[card].rare == 4:
                    rareRate += 0.06
                    if self.player.cardMap[card].cardPhase > 0:
                        phaseRate += 0.02 * self.player.cardMap[card].cardPhase
                else:
                    if self.player.cardMap[card].cardPhase > 0:
                        phaseRate += 0.01 * self.player.cardMap[card].cardPhase
        power = base * rareRate * phaseRate
        if self.floor < SoulTrial.MaxFloor and power >= SoulTrial.PowerMap[self.floor + 1]:
            self.floor += 1
            if self.man > 0:
                self.player.loveLevList[self.man].AddLoveExp(E_LoveType.SoulTrial, self.floor)
            self.player.GetNewItems(ItemSpawn.GetItemNumMap(SoulTrial.rewardMap[self.man * 1000 + self.floor]), "定向轨道")
            # print(f"man:{self.man},floor:{self.floor}")
        else:
            # print(f"{self.player.name}power{power},floor{self.floor+1},needPower{self.powerMap[self.floor + 1]}")
            pass

    @classmethod
    def SoulTrialInit(cls):
        SoulTrialReward = XlDeal("SoulTrial.xlsx", "SoulTrial")
        manCol = SoulTrialReward.GetColIndex("RoleID", 2)
        floorCol = SoulTrialReward.GetColIndex("Floor", 2)
        rewardCol = SoulTrialReward.GetColIndex("Reward", 2)
        for row in SoulTrialReward.data[3:]:
            id = int(row[manCol] * 1000 + row[floorCol])
            SoulTrial.rewardMap[id] = row[rewardCol]
        SoulTrialReward.CloseBook()
        del SoulTrialReward
