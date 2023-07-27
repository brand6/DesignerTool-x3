import xlwings as xw
import math
from xlwings import Sheet
from common import common
import numpy as np

# 脚本说明：
# 用于解析奖励投放
#


def main():
    dataRng = xw.books.active.sheets['奖励投放'].used_range
    maxCol = dataRng.last_cell.column
    maxRow = dataRng.last_cell.row
    dataList = dataRng.value
    colList = dataList[2]

    paraSht: Sheet = xw.books.active.sheets['参数设定']
    paraDataList = np.array(paraSht.used_range.value)
    paraCols = common.getColBy2Para('通用参数设定', ['参数说明', '参数值'], paraDataList)
    dayCol = common.getDataOrder(paraDataList[0], '日期')
    dayList = paraDataList[2:, dayCol]
    paraMap = {}
    for row in paraDataList[2:]:
        paraMap[row[paraCols[0]]] = row[paraCols[1]]

    cardNumSht = xw.books.active.sheets['思念统计']
    cardNumData = cardNumSht.used_range.value
    cardNumStartCol = common.getColBy2Para('免费思念进阶', 'SSR', cardNumData)
    hungStartCol = common.getColBy2Para('普通挂机券数量', '免费', cardNumData)

    calcDataSht = xw.books.active.sheets['数据源']
    calcData = np.array(calcDataSht.used_range.value)
    cC_cardInfo_cols = common.getColBy3Para('Card.xlsx', 'CardBaseInfo', ['ID', 'Quality'], calcData)
    cardMap = {}  # 存储卡牌品质
    for row in calcData[3:]:
        cardMap[row[cC_cardInfo_cols[0]]] = row[cC_cardInfo_cols[1]]

    # 将奖励投放存入systemMap
    systemMap = {}  # {'每日任务'：[['player','解锁日期','循环周期','1=1',10],['player','解锁日期','循环周期','1=1',10]],'每周任务'...}
    for col in range(maxCol):
        if dataList[2][col] == 'reward':
            endCol = getEndCol(dataList, col)
            systemCol = getColOrder(colList, '系统', col, endCol)
            unlockDayCol = getColOrder(colList, '解锁日期', col, endCol)
            repeatCol = getColOrder(colList, '循环周期', col, endCol)
            getTimesCol = getColOrder(colList, '获得次数', col, endCol)
            payCol = getColOrder(colList, '付费', col, endCol)
            costCol = getColOrder(colList, 'cost', col, endCol)

            for row in range(4, maxRow):
                if dataList[row][col] is not None:
                    system = dataList[row][systemCol]
                    unlockDay = dataList[row][unlockDayCol]
                    if system is not None and unlockDay is not None:
                        if system not in systemMap:
                            systemMap[system] = []
                        if getTimesCol == -1 or dataList[row][getTimesCol] is None:
                            getTimes = 1
                        else:
                            getTimes = dataList[row][getTimesCol]
                        repeat = dataList[row][repeatCol] if repeatCol != -1 else None
                        payLevel = 0
                        payExtend = True
                        if payCol != -1 and dataList[row][payCol] is not None:
                            if '+' in dataList[row][payCol]:
                                payLevel = int(dataList[row][payCol][0])
                            else:
                                payLevel = int(dataList[row][payCol])
                                payExtend = False

                        for i in range(6):
                            if i == payLevel or (payExtend is True and i > payLevel):
                                rewardList = splitReward(dataList[row][col])
                                for reward in rewardList:
                                    for rList in systemMap[system]:
                                        if rList[0] == i and rList[1] == unlockDay and rList[2] == repeat and rList[
                                                3] == reward[0]:
                                            rList[4] += reward[1] * getTimes
                                            break
                                    else:
                                        systemMap[system].append([i, unlockDay, repeat, reward[0], reward[1]])
                                if costCol != -1 and dataList[row][costCol] is not None:
                                    costList = splitReward(dataList[row][costCol])
                                    for reward in costList:
                                        for rList in systemMap[system]:
                                            if rList[0] == i and rList[1] == unlockDay and rList[2] == repeat and rList[
                                                    3] == reward[0]:
                                                rList[4] -= reward[1] * getTimes
                                                break
                                        else:
                                            systemMap[system].append([i, unlockDay, repeat, reward[0], -reward[1]])

    # 将额外奖励存入systemMap
    unlockDayCol = common.getColBy2Para('额外奖励配置', '解锁日期', paraDataList)
    repeatCol = common.getColBy2Para('额外奖励配置', '循环周期', paraDataList)
    startCol = common.getColBy2Para('额外奖励配置', '免费', paraDataList)
    system = '额外奖励配置'
    systemMap[system] = []
    for row in paraDataList[2:]:
        if row[unlockDayCol] is not None:
            unlockDay = row[unlockDayCol]
            repeat = row[repeatCol]
            for i in range(6):
                rewardList = splitReward(row[startCol + i])
                for reward in rewardList:
                    for rList in systemMap[system]:
                        if rList[0] == i and rList[1] == unlockDay and rList[2] == repeat and rList[3] == reward[0]:
                            rList[4] += reward[1]
                            break
                    else:
                        systemMap[system].append([i, unlockDay, repeat, reward[0], reward[1]])

    # 根据系统统计奖励
    sysStatisticSht: Sheet = xw.books.active.sheets['系统奖励']
    sysStatisticData = sysStatisticSht.used_range.value
    itemTypeList = sysStatisticData[1]
    sMaxRow = len(sysStatisticData)
    sMaxCol = len(itemTypeList)
    if sMaxRow > 3:
        sysStatisticSht.range('4:' + str(sMaxRow)).clear()

    dayRewardList = []
    for day in dayList:
        if day is not None:
            for sys in systemMap:
                rewardList = [0] * sMaxCol
                rewardList[0] = day
                rewardList[1] = sys
                for item in systemMap[sys]:
                    unlockDay = item[1]
                    periodDay = item[2]
                    reward = item[3]
                    rewardNum = item[4]
                    for col in range(2, sMaxCol):
                        if itemTypeList[col] is not None:
                            itemType = itemTypeList[col]
                        player = (col - 2) % 6
                        if itemType == reward and player == item[0]:  # 奖励内容
                            rewardList[col] += getRewardCount(day, unlockDay, periodDay, rewardNum, paraMap)
                dayRewardList.append(rewardList)
    sysStatisticSht.cells(4, 1).value = dayRewardList

    # 转换道具类型
    exchangeCol = common.getColBy2Para('材料转换表', '转换前材料名称', paraDataList)
    exchangeList = paraSht.cells(2, exchangeCol + 1).expand('table').value
    for sys in systemMap:
        for item in systemMap[sys]:
            for e in exchangeList:
                if item[3] == e[1]:
                    item[3] = e[2]
                    item[4] *= e[3]

    # 根据奖励类型统计奖励
    totalStatisticSht = xw.books.active.sheets['奖励统计']
    totalStatisticData = totalStatisticSht.used_range.value
    itemTypeList = totalStatisticData[1]
    tMaxRow = len(totalStatisticData)
    tMaxCol = len(itemTypeList)
    if tMaxRow > 3:
        totalStatisticSht.range('4:' + str(tMaxRow)).clear()

    dayRewardList = []
    dayTicketList = []  # 抽卡券
    dayCardList = []  # 思念
    dayHungList = []  # 挂机券
    for day in dayList:
        if day is not None:
            rewardList = [0] * tMaxCol
            rewardList[0] = day
            cardList = [''] * 18
            hungList = [0] * 12
            for sys in systemMap:
                for item in systemMap[sys]:
                    unlockDay = item[1]
                    periodDay = item[2]
                    reward = item[3]
                    rewardNum = item[4]
                    for col in range(1, tMaxCol):
                        if itemTypeList[col] is not None:
                            itemType = itemTypeList[col]
                            if itemType == '2=2':
                                diamondCol = col
                            elif itemType == '400=400':
                                ticketCol = col
                        player = (col - 1) % 6
                        if player == item[0] and itemType == reward:  # 奖励内容
                            rewardList[col] += getRewardCount(day, unlockDay, periodDay, rewardNum, paraMap)

                    # 思念处理
                    if reward[:3] == '51=':
                        cardNum = int(getRewardCount(day, unlockDay, periodDay, rewardNum, paraMap))
                        if cardNum > 0:
                            cardId = int(reward[3:])
                            cardQuality = cardMap[cardId]
                            index = int(item[0] * 3 + 4 - cardQuality)
                            if cardList[index] == '':
                                cardList[index] = reward + '=' + str(cardNum)
                            else:
                                cardList[index] = cardList[0] + '|' + reward + '=' + str(cardNum)

                    # 挂机券
                    if reward == '75=600':
                        index = int(item[0])
                        hungList[index] += getRewardCount(day, unlockDay, periodDay, rewardNum, paraMap)
                    elif reward == '75=601':
                        index = int(6 + item[0])
                        hungList[index] += getRewardCount(day, unlockDay, periodDay, rewardNum, paraMap)
            dayCardList.append(cardList)
            dayRewardList.append(rewardList)
            dayHungList.append(hungList)
            # 抽卡券计算
            ticketList = [0] * 6
            maxNum = day * paraMap['每日抽卡上限']
            for i in range(6):
                ticketNum = rewardList[ticketCol + i]
                diamondNum = rewardList[diamondCol + i]
                ticketList[i] = getTicketNum(ticketNum, diamondNum, maxNum, paraMap)
            dayTicketList.append(ticketList)

    totalStatisticSht.cells(4, 1).value = dayRewardList
    drawSht = xw.books.active.sheets['抽卡统计']
    drawSht.cells(3, 3).value = dayTicketList
    cardNumSht.cells(3, cardNumStartCol + 1).value = dayCardList
    cardNumSht.cells(3, hungStartCol + 1).value = dayHungList
    input("数据处理完毕")


def getTicketNum(ticketNum, diamondNum, maxNum, paraMap):
    maxDiamondGet = diamondNum / paraMap['抽卡券兑换消耗']
    maxGet = int((ticketNum + maxDiamondGet) / 10) * 10
    if maxGet > maxNum:
        return maxNum
    else:
        return maxGet


def getRewardCount(day, unlockDay, periodDay, rewardNum, paraMap):
    """获取奖励的次数

        Args:
            day (_type_): 当前日期
            unlockDay (_type_): 奖励解锁日期
            periodDay (_type_): 奖励获取周期
            rewardNum (_type_): 奖励数量

        Returns:
            _type_: _description_
        """
    rNum = 0
    if unlockDay is not None and unlockDay <= day:  # 解锁日期
        if periodDay is None or periodDay == 0:  # 循环日期
            # 奖励1次
            rNum = rewardNum
        elif periodDay == 7:
            # 周奖励
            weekLeftDay = paraMap['开服时本周剩余天数']
            rNum = getRepeatCount(day, unlockDay, 7, weekLeftDay) * rewardNum
        elif periodDay == 30:
            # 月奖励
            monthLeftDay = paraMap['开服时本月剩余天数']
            rNum = getRepeatCount(day, unlockDay, 30, monthLeftDay) * rewardNum
        else:
            # 普通循环奖励
            rNum = math.ceil((day - unlockDay + 1) / periodDay) * rewardNum
    return rNum


def splitReward(rewardStr):
    """解析奖励的内容

    Args:
        rewardStr (_type_): [[1=1,10],[2=2,5]]
    """
    returnList = []
    rewardList = rewardStr.split('|')
    for reward in rewardList:
        loc = reward.rfind('=')
        returnList.append([reward[:loc], float(reward[loc + 1:])])
    return returnList


def getColOrder(colList, colStr, start=0, end=-1):
    """从开始位置向后查找符合的列位置

    Args:
        colList (_type_): 查找的列表
        colStr (_type_): 查找的列名
        start (int, optional): 开始位置. Defaults to 0.

    Returns:
        int: 列的位置
    """
    if end == -1:
        end = len(colList)
    for i in range(start, end):
        if colList[i] == colStr:
            return i
    else:
        return -1


def getRepeatCount(day, unLockDay, period, leftDay):
    """获取周/月奖励的有效次数

    Args:
        day (int): 当前开服天数
        unLockDay (int): 奖励解锁天数
        period (int): 奖励获得周期
        leftDay (int): 开服对应本周/月剩余天数
    """
    if day >= unLockDay:
        curCount = math.ceil((day + period - leftDay) / period)
        unlockCount = math.ceil((unLockDay + period - leftDay) / period)
        return curCount - unlockCount + 1
    else:
        return 0


def getEndCol(dataList, startCol):
    for col in range(startCol + 1, len(dataList[1])):
        if dataList[1][col] is not None:
            return col
    else:
        return col
