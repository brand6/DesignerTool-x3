import xlwings as xw
import math
from xlwings import Sheet
from common import common
import numpy as np

# 脚本说明：
# 用于解析奖励投放
#


def main(isPay=False):
    """
    Args:
        isPay (bool, optional): 是否付费. Defaults to False.
    """
    dataRng = xw.books.active.sheets['奖励数据表'].used_range
    dataList = dataRng.value
    colList = dataList[2]

    paraSht: Sheet = xw.books.active.sheets['参数']
    if isPay is True:
        paraCol = 2
        sysStatisticSht: Sheet = xw.books.active.sheets['付费系统奖励']
        totalStatisticSht = xw.books.active.sheets['付费奖励统计']
        developSht = xw.books.active.sheets['付费养成统计']
    else:
        paraCol = 1
        sysStatisticSht: Sheet = xw.books.active.sheets['系统奖励']
        totalStatisticSht = xw.books.active.sheets['奖励统计']
        developSht = xw.books.active.sheets['养成统计']

    paraDataList = paraSht.cells(1, 1).expand('table').value
    weekLeftDay = common.getRowData('开服本周剩余日期', 0, paraCol, paraDataList)
    monthLeftDay = common.getRowData('开服本月剩余日期', 0, paraCol, paraDataList)
    payLevel = common.getRowData('付费级别', 0, paraCol, paraDataList)

    extraRewardCol = common.getDataColOrder(paraSht.used_range.value, '每日额外养成资源')
    extraRewardList = paraSht.cells(1, extraRewardCol + 1).expand('table').value

    def getRewardCount(day, unlockDay, periodDay, rewardNum):
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
                rNum = getRepeatCount(day, unlockDay, 7, weekLeftDay) * rewardNum
            elif periodDay == 30:
                # 月奖励
                rNum = getRepeatCount(day, unlockDay, 30, monthLeftDay) * rewardNum
            else:
                # 普通循环奖励
                rNum = math.ceil((day - unlockDay + 1) / periodDay) * rewardNum
        return rNum

    maxCol = dataRng.last_cell.column
    maxRow = dataRng.last_cell.row

    # 将奖励数据存入systemMap
    systemMap = {}  # {'每日任务'：[['解锁日期','循环周期','1=1',10],['解锁日期','循环周期','1=1',10]],'每周任务'...}
    for col in range(maxCol):
        if dataList[2][col] == 'reward':
            endCol = getEndCol(dataList, col)
            payCol = getColOrder(colList, '付费', col, endCol)
            systemCol = getColOrder(colList, '系统', col, endCol)
            unlockDayCol = getColOrder(colList, '解锁日期', col, endCol)
            repeatCol = getColOrder(colList, '循环周期', col, endCol)
            getTimesCol = getColOrder(colList, '获得次数', col, endCol)

            for row in range(4, maxRow):
                if dataList[row][col] is not None:
                    payExtend = False
                    if '+' in str(dataList[row][payCol]):
                        pay = int(dataList[row][payCol][0])
                        payExtend = True
                    else:
                        pay = dataList[row][payCol]
                    system = dataList[row][systemCol]
                    unlockDay = dataList[row][unlockDayCol]
                    repeat = dataList[row][repeatCol]
                    if getTimesCol == -1 or dataList[row][getTimesCol] is None:
                        getTimes = 1
                    else:
                        getTimes = dataList[row][getTimesCol]
                    if system is not None and unlockDay is not None:
                        if pay is None or (isPay is False and pay == 0) \
                                or (isPay is True and pay == payLevel) or (isPay is True and payExtend is True and pay > payLevel):
                            if system not in systemMap:
                                systemMap[system] = []
                            rewardList = splitReward(dataList[row][col])
                            for reward in rewardList:
                                for rList in systemMap[system]:
                                    if rList[0] == unlockDay and rList[1] == repeat and rList[2] == reward[0]:
                                        rList[3] += reward[1] * getTimes
                                        break
                                else:
                                    systemMap[system].append([unlockDay, repeat, reward[0], reward[1]])

    paraColList = paraSht.used_range.value[0]
    dayCol = getColOrder(paraColList, '日期')
    dayList = paraSht.cells(2, dayCol + 1).expand('down').value
    zeroRoundList = paraSht.cells(2, dayCol + 2).expand('down').value
    gemRoundList = paraSht.cells(2, dayCol + 3).expand('down').value
    costPowerList = paraSht.cells(2, dayCol + 4).expand('down').value

    # 根据系统统计奖励
    sysStatisticData = sysStatisticSht.used_range.value
    itemTypeList = sysStatisticData[1]
    maxRow = len(sysStatisticData)
    maxCol = len(itemTypeList)
    if maxRow > 2:
        sysStatisticSht.range('3:' + str(maxRow)).clear()

    dayRewardList = []
    for day in dayList:
        for sys in systemMap:
            rewardList = [0] * maxCol
            rewardList[0] = day
            rewardList[1] = sys
            for item in systemMap[sys]:
                unlockDay = item[0]
                periodDay = item[1]
                reward = item[2]
                rewardNum = item[3]
                for col in range(maxCol):
                    if itemTypeList[col] is not None and itemTypeList[col] == reward:  # 奖励内容
                        rewardList[col] += getRewardCount(day, unlockDay, periodDay, rewardNum)
            dayRewardList.append(rewardList)
    sysStatisticSht.cells(3, 1).value = dayRewardList

    # 转换道具类型
    exchangeCol = getColOrder(paraColList, '转换前材料')
    exchangeList = paraSht.cells(2, exchangeCol + 1).expand('table').value
    for sys in systemMap:
        for item in systemMap[sys]:
            for e in exchangeList:
                if item[2] == e[0]:
                    item[2] = e[1]
                    item[3] *= e[2]

    # 根据奖励类型统计奖励
    totalStatisticData = totalStatisticSht.used_range.value
    itemTypeList = totalStatisticData[1]
    maxRow = len(totalStatisticData)
    maxCol = len(itemTypeList)
    if maxRow > 2:
        totalStatisticSht.range('3:' + str(maxRow)).clear()

    dayRewardList = []
    for day in dayList:
        rewardList = [0] * maxCol
        rewardList[0] = day
        for sys in systemMap:
            for item in systemMap[sys]:
                unlockDay = item[0]
                periodDay = item[1]
                reward = item[2]
                rewardNum = item[3]
                for col in range(maxCol):
                    if itemTypeList[col] is not None and itemTypeList[col] == reward:  # 奖励内容
                        rewardList[col] += getRewardCount(day, unlockDay, periodDay, rewardNum)

        for row in extraRewardList:
            for col in range(maxCol):
                if itemTypeList[col] is not None and itemTypeList[col] == row[3]:  # 奖励内容
                    rewardList[col] += row[paraCol] * day

        dayRewardList.append(rewardList)
    totalStatisticSht.cells(3, 1).value = dayRewardList

    # 计算养成进度
    developScoreNum = common.getRowData('养成搭档数', 0, paraCol, paraDataList)
    developQuality = common.getRowData('养成品质', 0, paraCol, paraDataList)
    buyExpTimes = common.getRowData('经验本购买次数', 0, paraCol, paraDataList)

    developData = np.array(xw.books.active.sheets['养成数据'].used_range.value)
    scoreLevResCol = common.getColBy2Para(developQuality + '搭档', '升级经验', developData)
    cardLevResCol = common.getColBy2Para(developQuality + '思念', '升级经验', developData)
    gemLevCol = common.getColBy2Para('星核', '升级经验', developData)
    scoreBreakGoldCol = common.getColBy2Para(developQuality + '搭档', '金币', developData)
    scoreBreakRes1Col = common.getColBy2Para(developQuality + '搭档', '材料1', developData)
    scoreBreakRes2Col = common.getColBy2Para(developQuality + '搭档', '材料2', developData)
    scoreBreakRes3Col = common.getColBy2Para(developQuality + '搭档', '材料3', developData)
    cardBreakGoldCol = common.getColBy2Para(developQuality + '思念', '金币', developData)
    cardBreakRes1Col = common.getColBy2Para(developQuality + '思念', '材料1', developData)
    cardBreakRes2Col = common.getColBy2Para(developQuality + '思念', '材料2', developData)
    cardBreakRes3Col = common.getColBy2Para(developQuality + '思念', '材料3', developData)
    roundPowerExp = common.getColBy2Para('资源本经验转换率', '体力转经验', developData)
    roundPowerGold = common.getColBy2Para('资源本经验转换率', '体力转金币', developData)
    roundPowerRes1 = common.getColBy2Para('资源本经验转换率', '体力转材料1', developData)
    roundPowerRes2 = common.getColBy2Para('资源本经验转换率', '体力转材料2', developData)
    roundPowerRes3 = common.getColBy2Para('资源本经验转换率', '体力转材料3', developData)
    roundPowerGemExp = common.getColBy2Para('资源本经验转换率', '体力转芯核经验', developData)

    developList = []  # 保存输出的数据
    leftResourceMap = {}  # 保存未用完的资源
    lastList = [0, 1, 0, 1, 0, 0]  # 保存上一条的数据,【0日期，1搭档等级，2搭档突破，3思念等级，4思念突破，5星核等级】

    for row in range(len(dayList)):
        if row == 0:
            days = 1
            for col in range(maxCol):
                if itemTypeList[col] is not None:
                    leftResourceMap[itemTypeList[col]] = dayRewardList[row][col]
        else:
            days = dayList[row] - dayList[row - 1]
            for col in range(maxCol):
                if itemTypeList[col] is not None:
                    leftResourceMap[itemTypeList[col]] += dayRewardList[row][col] - dayRewardList[row - 1][col]

        # 处理经验本
        expPara = common.toNum(developData[1 + int(zeroRoundList[row])][roundPowerExp])  # 体力转换系数
        goldPara = common.toNum(developData[1 + int(zeroRoundList[row])][roundPowerGold])
        res1Para = common.toNum(developData[1 + int(zeroRoundList[row])][roundPowerRes1])
        res2Para = common.toNum(developData[1 + int(zeroRoundList[row])][roundPowerRes2])
        res3Para = common.toNum(developData[1 + int(zeroRoundList[row])][roundPowerRes3])
        gemExpPara = common.toNum(developData[1 + int(common.toNum(gemRoundList[row]))][roundPowerGemExp])

        if lastList[1] < 80:
            leftResourceMap['200=0'] += expPara * (buyExpTimes + 1) * 10 * days
            leftResourceMap['3=3'] -= 8 * (buyExpTimes + 1) * 10 * days
        if lastList[3] < 80:
            leftResourceMap['201=0'] += expPara * (buyExpTimes + 1) * 10 * days
            leftResourceMap['3=3'] -= 8 * (buyExpTimes + 1) * 10 * days
        leftResourceMap['3=3'] -= int(common.toNum(costPowerList[row]))

        # 处理搭档
        breakTimes = lastList[2]
        for lev in range(lastList[1], 81):
            if lev == (breakTimes + 1) * 10 and breakTimes < 7:  # 突破
                needGold = common.toNum(developData[2 + breakTimes][scoreBreakGoldCol]) * developScoreNum
                needRes1 = common.toNum(developData[2 + breakTimes][scoreBreakRes1Col]) * developScoreNum
                needRes2 = common.toNum(developData[2 + breakTimes][scoreBreakRes2Col]) * developScoreNum
                needRes3 = common.toNum(developData[2 + breakTimes][scoreBreakRes3Col]) * developScoreNum
                if leftResourceMap['1=1'] < needGold:
                    needPower = math.ceil((needGold - leftResourceMap['1=1']) / goldPara) * 8
                    costPower = min(math.floor(leftResourceMap['3=3'] / 8) * 8, needPower)
                    leftResourceMap['1=1'] += costPower / 8 * goldPara
                    leftResourceMap['3=3'] -= costPower
                if leftResourceMap['202=10001'] < needRes1:
                    needPower = math.ceil((needRes1 - leftResourceMap['202=10001']) / res1Para) * 8
                    costPower = min(math.floor(leftResourceMap['3=3'] / 8) * 8, needPower)
                    leftResourceMap['202=10001'] += costPower / 8 * res1Para
                    leftResourceMap['202=10002'] += costPower / 8 * res2Para
                    leftResourceMap['202=10003'] += costPower / 8 * res3Para
                    leftResourceMap['3=3'] -= costPower
                if leftResourceMap['202=10002'] < needRes2:
                    needPower = math.ceil((needRes2 - leftResourceMap['202=10002']) / res2Para) * 8
                    costPower = min(math.floor(leftResourceMap['3=3'] / 8) * 8, needPower)
                    leftResourceMap['202=10001'] += costPower / 8 * res1Para
                    leftResourceMap['202=10002'] += costPower / 8 * res2Para
                    leftResourceMap['202=10003'] += costPower / 8 * res3Para
                    leftResourceMap['3=3'] -= costPower
                if leftResourceMap['202=10003'] < needRes3:
                    needPower = math.ceil((needRes3 - leftResourceMap['202=10003']) / res3Para) * 8
                    costPower = min(math.floor(leftResourceMap['3=3'] / 8) * 8, needPower)
                    leftResourceMap['202=10001'] += costPower / 8 * res1Para
                    leftResourceMap['202=10002'] += costPower / 8 * res2Para
                    leftResourceMap['202=10003'] += costPower / 8 * res3Para
                    leftResourceMap['3=3'] -= costPower

                if leftResourceMap['1=1'] >= needGold and leftResourceMap['202=10001'] >= needRes1 and leftResourceMap[
                        '202=10002'] >= needRes2 and leftResourceMap['202=10003'] >= needRes3:
                    leftResourceMap['1=1'] -= needGold
                    leftResourceMap['202=10001'] -= needRes1
                    leftResourceMap['202=10002'] -= needRes2
                    leftResourceMap['202=10003'] -= needRes3
                    breakTimes += 1
                else:
                    break

            if lev < (breakTimes + 1) * 10 and lev < 80:  # 升级
                scoreNeedRes = common.toNum(developData[2 + lev][scoreLevResCol]) * developScoreNum
                if leftResourceMap['200=0'] >= scoreNeedRes:
                    leftResourceMap['200=0'] -= scoreNeedRes
                else:
                    break
        lastList[1] = lev
        lastList[2] = breakTimes

        # 处理思念
        breakTimes = lastList[4]
        for lev in range(lastList[3], 81):
            if lev == (breakTimes + 1) * 10 and breakTimes < 7:  # 突破
                needGold = common.toNum(developData[2 + breakTimes][cardBreakGoldCol]) * developScoreNum * 6
                needRes1 = common.toNum(developData[2 + breakTimes][cardBreakRes1Col]) * developScoreNum * 6
                needRes2 = common.toNum(developData[2 + breakTimes][cardBreakRes2Col]) * developScoreNum * 6
                needRes3 = common.toNum(developData[2 + breakTimes][cardBreakRes3Col]) * developScoreNum * 6
                if leftResourceMap['1=1'] < needGold:
                    needPower = math.ceil((needGold - leftResourceMap['1=1']) / goldPara) * 8
                    costPower = min(math.floor(leftResourceMap['3=3'] / 8) * 8, needPower)
                    leftResourceMap['1=1'] += costPower / 8 * goldPara
                    leftResourceMap['3=3'] -= costPower
                if leftResourceMap['205=100031'] < needRes1:
                    needPower = math.ceil((needRes1 - leftResourceMap['205=100031']) / res1Para) * 8
                    costPower = min(math.floor(leftResourceMap['3=3'] / 8) * 8, needPower)
                    leftResourceMap['205=100031'] += costPower / 8 * res1Para
                    leftResourceMap['205=100032'] += costPower / 8 * res2Para
                    leftResourceMap['205=100033'] += costPower / 8 * res3Para
                    leftResourceMap['3=3'] -= costPower
                if leftResourceMap['205=100032'] < needRes2:
                    needPower = math.ceil((needRes2 - leftResourceMap['205=100032']) / res2Para) * 8
                    costPower = min(math.floor(leftResourceMap['3=3'] / 8) * 8, needPower)
                    leftResourceMap['205=100031'] += costPower / 8 * res1Para
                    leftResourceMap['205=100032'] += costPower / 8 * res2Para
                    leftResourceMap['205=100033'] += costPower / 8 * res3Para
                    leftResourceMap['3=3'] -= costPower
                if leftResourceMap['205=100033'] < needRes3:
                    needPower = math.ceil((needRes3 - leftResourceMap['205=100033']) / res3Para) * 8
                    costPower = min(math.floor(leftResourceMap['3=3'] / 8) * 8, needPower)
                    leftResourceMap['205=100031'] += costPower / 8 * res1Para
                    leftResourceMap['205=100032'] += costPower / 8 * res2Para
                    leftResourceMap['205=100033'] += costPower / 8 * res3Para
                    leftResourceMap['3=3'] -= costPower

                if leftResourceMap['1=1'] >= needGold and leftResourceMap['205=100031'] >= needRes1 and leftResourceMap[
                        '205=100032'] >= needRes2 and leftResourceMap['205=100033'] >= needRes3:
                    leftResourceMap['1=1'] -= needGold
                    leftResourceMap['205=100031'] -= needRes1
                    leftResourceMap['205=100032'] -= needRes2
                    leftResourceMap['205=100033'] -= needRes3
                    breakTimes += 1
                else:
                    break

            if lev < (breakTimes + 1) * 10 and lev < 80:  # 升级
                cardNeedRes = common.toNum(developData[2 + lev][cardLevResCol]) * developScoreNum * 6
                if leftResourceMap['201=0'] >= cardNeedRes:
                    leftResourceMap['201=0'] -= cardNeedRes
                else:
                    break
        lastList[3] = lev
        lastList[4] = breakTimes

        # 处理星核
        if dayList[row] >= 3:
            for lev in range(lastList[5], 26):
                gemNeedRes = common.toNum(developData[2 + lev][gemLevCol]) * developScoreNum * 4
                needPower = math.ceil((gemNeedRes - leftResourceMap['301=0']) / gemExpPara) * 10
                costPower = min(math.floor(leftResourceMap['3=3'] / 10) * 10, needPower)
                leftResourceMap['301=0'] += costPower / 10 * gemExpPara
                leftResourceMap['3=3'] -= costPower
                if leftResourceMap['301=0'] >= gemNeedRes:
                    leftResourceMap['301=0'] -= gemNeedRes
                else:
                    break
            lastList[5] = lev
        developList.append(
            [dayList[row], lastList[1], lastList[2], lastList[3], lastList[4], lastList[5], leftResourceMap['3=3']])
    developSht.cells(2, 1).value = developList


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
