from outPutCardLev import OutPutCardLev
from player import Player


def simulateDevelop(player: Player, normalParaMap, day, playerOrder, rare, outPut: OutPutCardLev) -> None:
    # 根据日期进行各种计算
    lastDay = day - 1
    player.UpdateItemsByDay(day, lastDay)

    if day == 1:
        player.GetAllCard([1, 2, 5])

    # 计算养成
    resStr = ""
    resList = [
        "3=3",
        "201=0",
        "205=100101",
        "205=100201",
        "205=100301",
        "205=100102",
        "205=100202",
        "205=100302",
        "205=100103",
        "205=100203",
        "205=100303",
    ]
    for res in resList:
        if res in player.itemMap and player.itemMap[res] > 0:
            if resStr == "":
                resStr = res + "=" + str(int(player.itemMap[res]))
            else:
                resStr = resStr + "\n" + res + "=" + str(int(player.itemMap[res]))
    outPut.UpdateRes(resStr, day, playerOrder)
    developNum = outPut.GetDevelopNum(day)
    while developNum > len(player.targetCardList):
        player.AddTargetCard(rare)
    player.DevelopTargetCard()
    levMap = {}
    for card in player.targetCardList:
        if player.cardMap[card].cardLev not in levMap:
            levMap[player.cardMap[card].cardLev] = 1
        else:
            levMap[player.cardMap[card].cardLev] += 1
    outPut.UpdateData(levMap, day, playerOrder, rare)
    del levMap

    # 更新定向轨道
    for _ in range(int(normalParaMap["每日定向轨道次数"][0])):
        for trial in player.soulTrialList:
            if trial is not None:
                trial.TryUp()
