from cardGacha import Gacha
from dollCatcher import DollCatcher
from hangUp import HangUp
from itemSpawn import ItemSpawn
from lovePoint import E_LoveType
from miaoGacha import MiaoGacha
from player import Player


def simulate(player: Player, normalParaMap, day, playerOrder) -> None:
    # 根据日期进行各种计算
    lastDay = day - 1
    week = ItemSpawn.GetWeek(day)
    lastWeek = ItemSpawn.GetWeek(lastDay)
    player.UpdateItemsByDay(day, lastDay)
    # 尝试玩家升级
    player.levUp()
    # 更新日期相关的牵绊度
    for i in range(1, 6):
        if player.loveLevList[i] is not None:
            player.loveLevList[i].AddLoveExp(E_LoveType.DayChange, day)
    # 计算抽卡
    todayDrawTimes = 0
    # 活动池&新手池
    player.newGacha = int(normalParaMap["新手卡池"][playerOrder])
    if player.isNewFinish is False:
        gachaId = int(normalParaMap["新手卡池"][playerOrder])
    else:
        gachaId = int(normalParaMap["up卡池"][playerOrder])
    _ticketId = Gacha.AllMap[gachaId]["CostTicket"]
    _ticketType = ItemSpawn.ItemMap[_ticketId]
    drawTicket = str(_ticketType) + "=" + str(_ticketId)
    drawCostNum = Gacha.AllMap[gachaId]["Cost10"]
    while todayDrawTimes < normalParaMap["每日抽卡上限"][playerOrder]:
        if drawTicket in player.itemMap:
            lackNum = drawCostNum - player.itemMap[drawTicket]
        else:
            lackNum = drawCostNum
            player.itemMap[drawTicket] = 0
        if lackNum > 0:
            needRes = normalParaMap["抽卡券兑换消耗"][0] * lackNum
            if player.itemMap["2=2"] < needRes:
                break
            else:
                player.itemMap["2=2"] -= lackNum * normalParaMap["抽卡券兑换消耗"][0]
                player.itemMap[drawTicket] += lackNum
        player.DrawCard(gachaId, 1, is10=True)
        todayDrawTimes += 10
    # 常规池
    gachaId = int(normalParaMap["up卡池"][0])
    _ticketId = Gacha.AllMap[gachaId]["CostTicket"]
    _ticketType = ItemSpawn.ItemMap[_ticketId]
    drawTicket = str(_ticketType) + "=" + str(_ticketId)
    drawCostNum = Gacha.AllMap[gachaId]["Cost10"]
    while player.itemMap[drawTicket] >= drawCostNum:
        player.DrawCard(gachaId, 1, is10=True)
    # 计算挂机
    hangUpList = [2000, 1000]
    for exploreId in hangUpList:
        drawTicket, drawCostNum = ItemSpawn.GetItemKeyAndValue(HangUp.HangUpMap[exploreId]["Cost"])
        if drawTicket in player.itemMap:
            drawTimes = int(player.itemMap[drawTicket] / drawCostNum)
            player.HangUpExplore(exploreId, drawTimes)
            # print(f"{player.name:4} day{day:2} exploreId{exploreId} 探索次数={drawTimes:3}")
    # 计算养成
    player.UpdateDevelopCardList()
    player.DevelopCard()
    player.levUp()
    # 更新定向轨道
    for _ in range(int(normalParaMap["每日定向轨道次数"][0])):
        for trial in player.soulTrialList:
            if trial is not None:
                trial.TryUp()
    # 计算喵呜集卡
    if day >= normalParaMap["喵喵牌解锁天数"][0]:
        if len(player.miaoGachaMap) == 0 or week - lastWeek > 0:
            if len(player.miaoGachaMap) == 0:
                for pack in normalParaMap["常驻喵呜徽章"][0].split("|"):
                    packId = int(pack)
                    player.miaoGachaMap[packId] = MiaoGacha(packId, player)
            drawTimes = int(normalParaMap["每周喵喵牌次数"][playerOrder])
            for _ in range(drawTimes):
                minPack = 0
                minPackNum = 0
                for pack in player.miaoGachaMap:
                    packNum = len(player.miaoGachaMap[pack].getList)
                    if packNum != 9:
                        if minPack == 0 or packNum < minPackNum:
                            minPack = pack
                            minPackNum = packNum
                if minPack != 0:
                    player.miaoGachaMap[minPack].Draw()
            # 喵呜活动卡池 TODO

    # 计算娃娃机
    if day >= normalParaMap["娃娃机解锁天数"][0]:
        if len(player.dollMap) == 0 or week - lastWeek > 0:
            drawTimes = int(normalParaMap["每周娃娃机次数"][playerOrder])
            for _ in range(drawTimes):
                r = drawTimes % len(DollCatcher.DollCatcherMap[week])
                DollCatcher.DollCatcherMap[week][r].GetDoll(player)
