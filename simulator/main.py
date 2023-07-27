import sys
import traceback

import xlwings as xw
from card import Card
from cardGacha import Gacha
from dollCatcher import DollCatcher
from drop import Drop
from hangUp import HangUp
from initRes import InitRes
from itemSpawn import ItemSpawn
from lovePoint import LovePoint
from miaoGacha import MiaoGacha
from outPut import Output
from outPutCardLev import OutPutCardLev
from player import Player
from simulate import simulate
from simulateDevelop import simulateDevelop
from soulTrial import SoulTrial
from stage import Stage
from xlDeal import XlDeal


def main(app, tryTimes=1, runType=""):
    print("程序开始运行，运行结束后将自动关闭...")
    # 读取表格完成数据初始化
    XlDeal.app = app
    ItemSpawn.ItemInit()
    InitRes.InitResInit()
    Card.CardInit()
    Gacha.GachaInit()
    HangUp.HangUpInit()
    Drop.DropInit()
    Stage.StageInit()
    SoulTrial.SoulTrialInit()
    DollCatcher.DollCatcherInit()
    MiaoGacha.MiaoGachaInit()
    LovePoint.LovePointInit()
    Player.PlayerInit()
    # 获取参数表数据
    playerNum = 6
    paraXl = XlDeal("资源投放统计.xlsm", "参数设定")
    # 通用参数列
    normalParaTypeCol = paraXl.GetColIndex(["通用参数设定", "参数说明"])
    normalParaValueCol = paraXl.GetColIndex(["通用参数设定", "免费"])
    normalParaMap = {}
    # 转换参数列
    turnParaTypeBeforeCol = paraXl.GetColIndex(["材料转换表", "转换前材料"])
    turnParaTypeAfterCol = paraXl.GetColIndex(["材料转换表", "转换后材料"])
    turnParaRateCol = paraXl.GetColIndex(["材料转换表", "转换比例"])
    # 定向轨道参数列
    indexCol = paraXl.GetColIndex("编号")
    soulPowerCol = paraXl.GetColIndex("定向轨道难度")
    # 娃娃机和喵喵牌参数列
    weekCol = paraXl.GetColIndex("周")
    # miaoActivityCol = paraXl.GetColData(["喵呜徽章", "活动池"])
    # dollActivityCol = paraXl.GetColData(["娃娃机", "活动池"])
    dollNormalCol = paraXl.GetColIndex(["娃娃机", "常规池"])
    # 获取牵绊度参数列
    loveTypeCol = paraXl.GetColIndex(["牵绊度经验参数", "类型"])
    lovePara1Col = paraXl.GetColIndex(["牵绊度经验参数", "参数1"])
    lovePara2Col = paraXl.GetColIndex(["牵绊度经验参数", "参数2"])
    loveLimitCol = paraXl.GetColIndex(["牵绊度经验参数", "限制"])
    loveExpCol = paraXl.GetColIndex(["牵绊度经验参数", "经验"])
    for row in paraXl.data[2:]:
        # 通用参数
        if row[normalParaTypeCol] is not None:
            normalParaMap[row[normalParaTypeCol]] = []
            for i in range(playerNum):
                normalParaMap[row[normalParaTypeCol]].append(row[normalParaValueCol + i])
        # 资源转换参数
        if row[turnParaRateCol] is not None:
            Player.TurnParaMap[row[turnParaTypeBeforeCol]] = [row[turnParaTypeAfterCol], row[turnParaRateCol]]
        # 定向轨道参数
        if row[indexCol] is not None:
            SoulTrial.PowerMap[int(row[indexCol])] = row[soulPowerCol]
        # 喵喵牌和娃娃机活动池
        if row[weekCol] is not None:
            week = int(row[weekCol])
            DollCatcher.DollCatcherMap[week]: list[DollCatcher] = []
            if row[dollNormalCol] is not None:
                pools = row[dollNormalCol].split("|")
                for pool in pools:
                    catcherId, dropId = pool.split("=")
                    DollCatcher.DollCatcherMap[week].append(DollCatcher(int(catcherId), int(dropId)))
        # 牵绊度参数
        if row[loveTypeCol] is not None:
            key0 = str(int(row[loveTypeCol]))
            if row[lovePara1Col] is not None:
                key1 = []
                if isinstance(row[lovePara1Col], str) and "|" in row[lovePara1Col]:
                    for para in row[lovePara1Col].split("|"):
                        key1.append(para)
                else:
                    key1.append(str(int(row[lovePara1Col])))
                if row[lovePara2Col] is not None:
                    key2 = []
                    if isinstance(row[lovePara2Col], str) and "|" in row[lovePara2Col]:
                        for para in row[lovePara2Col].split("|"):
                            key2.append(para)
                    else:
                        key2.append(str(int(row[lovePara2Col])))
                    for k1 in key1:
                        for k2 in key2:
                            if key0 + "=" + k1 + "=" + k2 not in LovePoint.loveExpMap:
                                LovePoint.loveExpMap[key0 + "=" + k1 + "=" + k2] = row[loveExpCol]
                            else:
                                LovePoint.loveExpMap[key0 + "=" + k1 + "=" + k2] += row[loveExpCol]
                            if row[loveLimitCol] is not None:
                                LovePoint.loveLimitExpMap[key0 + "=" + k1 + "=" + k2] = row[loveLimitCol]
                else:
                    for k1 in key1:
                        if key0 + "=" + k1 not in LovePoint.loveExpMap:
                            LovePoint.loveExpMap[key0 + "=" + k1] = row[loveExpCol]
                        else:
                            LovePoint.loveExpMap[key0 + "=" + k1] += row[loveExpCol]
                        if row[loveLimitCol] is not None:
                            LovePoint.loveLimitExpMap[key0 + "=" + k1] = row[loveLimitCol]
            else:
                if key0 not in LovePoint.loveExpMap:
                    LovePoint.loveExpMap[key0] = row[loveExpCol]
                else:
                    LovePoint.loveExpMap[key0] += row[loveExpCol]
                if row[loveLimitCol] is not None:
                    LovePoint.loveLimitExpMap[key0] = row[loveLimitCol]
    SoulTrial.MaxFloor = normalParaMap["定向轨道层数"][0]
    # 保存娃娃机配置数据
    DollCatcher.CatcherAvgNum = normalParaMap["单局获取娃娃数"][0]
    DollCatcher.ChangeColorNum = normalParaMap["娃娃变色次数条件"][0]
    # 更新开服日期
    ItemSpawn.UpdateLeftWeekDays(normalParaMap["开服时本周剩余天数"][0])
    ItemSpawn.UpdateLeftMonthDays(normalParaMap["开服时本月剩余天数"][0])
    # 获取额外奖励参数
    rewardParaSysCol = paraXl.GetColIndex(["额外奖励配置", "奖励说明"])
    rewardParaUnlockCol = paraXl.GetColIndex(["额外奖励配置", "解锁日期"])
    rewardParaRepeatCol = paraXl.GetColIndex(["额外奖励配置", "循环周期"])
    rewardParaPlayerCol = paraXl.GetColIndex(["额外奖励配置", "免费"])
    # 获取奖励表数据
    rewardXl = XlDeal("资源投放统计.xlsm", "奖励投放")
    rewardData = rewardXl.data
    # 初始化输出
    match runType:
        case "LovePoint":
            outPut: Output = Output(tryTimes, playerNum, int(normalParaMap["模拟天数"][0]))
        case "CardLevel":
            outPut: OutPutCardLev = OutPutCardLev()
    # ------------------------------------------------------------------------------------------------------------------
    players: list[list[Player]] = [[] for _ in range(playerNum)]
    for i in range(tryTimes):
        print(f"开始模拟第{i+1}轮数据")
        for p in range(playerNum):
            playerName = paraXl.data[1][normalParaValueCol + p]
            player = Player(playerName, normalParaMap["培养男主"][p], normalParaMap["最低培养品质"][p])
            players[p].append(player)

            # 根据[参数设定]表创建奖励Spawn
            for row in paraXl.data[2:]:
                system = row[rewardParaSysCol]
                if system is not None and row[rewardParaUnlockCol] is not None:
                    player.AddItemSpawn(
                        ItemSpawn(row[rewardParaPlayerCol + p], system, row[rewardParaUnlockCol], row[rewardParaRepeatCol])
                    )
            # # 根据[奖励投放]表创建奖励Spawn
            for c in range(len(rewardData[0])):
                if rewardData[2][c] == "reward":
                    endCol = rewardXl.GetNotEmptyCol(c + 1, 1)
                    if endCol == len(rewardData[0]) - 1:
                        endCol += 1
                    systemCol = rewardXl.GetColIndex("系统", 2, c, endCol)
                    unlockDayCol = rewardXl.GetColIndex("解锁日期", 2, c, endCol)
                    repeatCol = rewardXl.GetColIndex("循环周期", 2, c, endCol)
                    getTimesCol = rewardXl.GetColIndex("获得次数", 2, c, endCol)
                    payCol = rewardXl.GetColIndex("付费", 2, c, endCol)
                    costCol = rewardXl.GetColIndex("cost", 2, c, endCol)

                    for r in range(4, len(rewardData)):
                        if (
                            rewardData[r][c] is not None
                            and rewardData[r][systemCol] is not None
                            and rewardData[r][unlockDayCol] is not None
                        ):
                            system = rewardData[r][systemCol]
                            payCheck = True
                            if payCol != -1 and rewardData[r][payCol] is not None:
                                if "+" in rewardData[r][payCol]:
                                    if p < int(rewardData[r][payCol][0]):
                                        payCheck = False
                                elif p != rewardData[r][payCol]:
                                    payCheck = False
                            if payCheck is True:
                                getTimes = 1 if getTimesCol == -1 else rewardData[r][getTimesCol]
                                period = 0 if repeatCol == -1 else rewardData[r][repeatCol]
                                cost = None if costCol == -1 else rewardData[r][costCol]
                                player.AddItemSpawn(
                                    ItemSpawn(rewardData[r][c], system, rewardData[r][unlockDayCol], period, getTimes, cost)
                                )
            player.GetNewItems(InitRes.GetInitRes(), "初始资源")
            player.GetNewItems(InitRes.GetShareRes(), "分享")

            for day in range(1, int(normalParaMap["模拟天数"][0] + 1)):
                match runType:
                    case "LovePoint":
                        simulate(player, normalParaMap, day, p)
                        if i == tryTimes - 1:
                            outPut.GetPlayerData(day, player, p)
                        else:
                            outPut.GetStatisticsData(day, player, p)
                    case "CardLevel":
                        simulateDevelop(player, normalParaMap, day, p, i + 2, outPut)

    outPut.UpdateTableData()


isDebug = True if sys.gettrace() else False
app = xw.App(visible=isDebug, add_book=False)
if not isDebug:
    match sys.argv[1]:
        case "LovePoint":
            tryTimes = 100
        case "CardLevel":
            tryTimes = 2

if isDebug:
    main(app, 2, "CardLevel")
else:
    try:
        main(app, tryTimes, sys.argv[1])
    except BaseException:
        traceback.print_exc()
        input("Error...按回车可安全退出！")
    finally:
        app.quit()
