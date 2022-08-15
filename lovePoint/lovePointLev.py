import xlwings as xw
from xlwings import Range
from xlwings import Sheet
from lovePoint import lovePointDraw

getDataList = lovePointDraw.getDataList
getParaDict = lovePointDraw.getParaDict
getItemsDict = lovePointDraw.getItemsDict
getDaysDict = lovePointDraw.getDaysDict
getlevsDict = lovePointDraw.getlevsDict
setRangeData = lovePointDraw.setRangeData
playExtendNum = lovePointDraw.playExtendNum  # 玩家种类扩展数量（玩家种类数-1）
paraBeginLine = lovePointDraw.paraBeginLine  # 数据开始行（excel行数-1）
levBeginLine = 7  # 数据开始行（excel行数-1）

# 脚本说明：
# 牵绊度等级计算
#


def main():
    levSht: Sheet = xw.books.active.sheets['亲密度验算']
    paraSht: Sheet = xw.books.active.sheets['验算参数']
    developSht: Sheet = xw.books.active.sheets['养成数据']
    levRange: Range = levSht.used_range
    paraRange: Range = paraSht.used_range
    developRange: Range = developSht.used_range

    itemList = getDataList(paraRange, "养成参数", -1)
    itemDict = getItemsDict(itemList)
    strList = ['爬塔', '养成上限', '养成资源价值']
    for i in range(1, len(itemList)):
        strList.append(itemList[i][0] + '数量')
    dayDict = getDaysDict(paraRange.value, strList)
    levDict = getlevsDict(developRange.value, ['升级经验价值', '突破价值', '等级战力', '突破战力'])
    l_towerList = getDataList(levRange, '爬塔层数', playExtendNum)
    l_dayList = getDataList(levRange, "天数", 0)
    limitDict = getLimitDict(levRange.value)

    dayCount = len(l_dayList)
    playerCount = playExtendNum + 1
    fightMapList = []  # 出战对象，优先养满
    collectMapList = []  # 收集对象
    fightScoreNumList = []
    fightCardNumList = []
    sumLevList = []
    sumBreakList = []
    powerList = []
    stageList = []

    for n in range(dayCount):
        fList = []
        cList = []
        fsList = []
        fcList = []
        slList = []
        sbList = []
        plist = []
        sList = []
        for m in range(playerCount):
            fList.append({})
            cList.append({})
            fsList.append(0)
            fcList.append(0)
            slList.append({})
            sbList.append({})
            plist.append(0)
            sList.append({})
        fightMapList.append(fList)
        collectMapList.append(cList)
        fightScoreNumList.append(fsList)
        fightCardNumList.append(fcList)
        sumLevList.append(slList)
        sumBreakList.append(sbList)
        powerList.append(plist)
        stageList.append(sList)

    # 计算数量
    print('开始处理数量')
    for i in range(1, len(itemList)):
        item = itemList[i][0]
        itemNumStr = item + "数量"
        itemStageStr = item + "深化"
        l_numList = getDataList(levRange, itemNumStr, playExtendNum)  # 牵绊度系统计算的数量
        l_stageList = getDataList(levRange, itemStageStr, playExtendNum)
        print(item + '处理中...')
        for n in range(levBeginLine, dayCount):
            day = l_dayList[n]
            limitNum = getLimitNum(limitDict, day, itemNumStr)
            limitStage = getLimitNum(limitDict, day, itemStageStr) * limitNum
            for m in range(playerCount):
                num = int(dayDict[day][itemNumStr][m])
                stageNum = 0
                if num > limitNum:
                    stageNum = num - limitNum
                    num = limitNum
                if stageNum > 0:
                    stageNum = limitStage if stageNum > limitStage else stageNum
                l_numList[n][m] = num  # 养成数量
                l_stageList[n][m] = stageNum  # 深化次数
                if num > 0:
                    stageList[n][m][item] = int(stageNum / num)
                else:
                    stageList[n][m][item] = 0

                fightNum = 0
                collectNum = 0
                if num > 0:  # 计算出战列表和收集列表
                    if i < 3:  # 搭档
                        if fightScoreNumList[n][m] < itemDict[item]['出战数量']:  # 出战列表有空位
                            fightMaxNum = itemDict[item]['出战数量'] - fightScoreNumList[n][m]
                            if num > fightMaxNum:  # 数量有溢出，可进收藏列表
                                fightNum = fightMaxNum
                                collectNum = num - fightMaxNum
                            else:
                                fightNum = num
                            fightScoreNumList[n][m] = fightScoreNumList[n][m] + fightNum
                        else:
                            collectNum = num
                    else:  # 羁绊
                        if fightCardNumList[n][m] < itemDict[item]['出战数量']:  # 出战列表有空位
                            fightMaxNum = itemDict[item]['出战数量'] - fightCardNumList[n][m]
                            if num > fightMaxNum:
                                fightNum = fightMaxNum
                                collectNum = num - fightMaxNum
                            else:
                                fightNum = num
                            fightCardNumList[n][m] = fightCardNumList[n][m] + fightNum
                        else:
                            collectNum = num
                fightMapList[n][m][item] = fightNum
                collectMapList[n][m][item] = collectNum
        setRangeData(l_numList, levSht, levBeginLine + 1, itemNumStr, levBeginLine)
        setRangeData(l_stageList, levSht, levBeginLine + 1, itemStageStr, levBeginLine)

    # 计算养成等级
    print('开始计算养成进度')
    for n in range(levBeginLine, dayCount):
        day = l_dayList[n]
        limitPower = getLimitNum(limitDict, day, '爬塔层数')
        for m in range(playerCount):
            limitLev = int(dayDict[day]['养成上限'][m])
            hasValue = dayDict[day]['养成资源价值'][m]
            collectValue = 0
            for lev in range(1, limitLev + 1):
                breakTimes = int((lev - 1) / 10)
                cosValue = 0
                if breakTimes > 0:
                    for item in fightMapList[n][m]:
                        cosValue += levDict[breakTimes]['突破价值'][item] * fightMapList[n][m][item]
                    if cosValue > hasValue:
                        breakTimes -= 1
                        break
                for item in fightMapList[n][m]:
                    cosValue += levDict[lev]['升级经验价值'][item] * fightMapList[n][m][item]
                if cosValue > hasValue:
                    lev -= 1
                    break
            else:
                collectValue = hasValue - cosValue

            for item in fightMapList[n][m]:
                sumLevList[n][m][item] = fightMapList[n][m][item] * lev
                sumBreakList[n][m][item] = fightMapList[n][m][item] * breakTimes
                power = fightMapList[n][m][item] * levDict[lev]['等级战力'][item]
                if breakTimes > 0:
                    power += fightMapList[n][m][item] * levDict[breakTimes]['突破战力'][item]
                power *= 1 + stageList[n][m][item] * itemDict[item]['深化强度']
                powerList[n][m] += power
            l_towerList[n][m] = dayDict[min(day * 10, int(powerList[n][m]) - 10, limitPower)]['爬塔'][0]

            if collectValue > 0:  # 养成材料有溢出
                for lev in range(1, limitLev + 1):
                    breakTimes = int(lev / 10)
                    breakTimes = 7 if breakTimes == 8 else breakTimes
                    cosValue = 0
                    if breakTimes > 0:
                        for item in collectMapList[n][m]:
                            cosValue += levDict[breakTimes]['突破价值'][item] * collectMapList[n][m][item]
                        if cosValue > collectValue:
                            breakTimes -= 1
                            break
                    for item in collectMapList[n][m]:
                        cosValue += levDict[lev]['升级经验价值'][item] * collectMapList[n][m][item]
                    if cosValue > collectValue:
                        lev -= 1
                        break
                for item in collectMapList[n][m]:
                    sumLevList[n][m][item] += collectMapList[n][m][item] * lev
                    sumBreakList[n][m][item] += collectMapList[n][m][item] * breakTimes
    setRangeData(l_towerList, levSht, levBeginLine + 1, '爬塔层数', levBeginLine)

    print('开始输出养成进度')
    for i in range(1, len(itemList)):
        item = itemList[i][0]
        itemLevStr = item + "等级"
        itemBreakStr = item + "突破"
        l_LevList = getDataList(levRange, itemLevStr, playExtendNum)
        l_BreakList = getDataList(levRange, itemBreakStr, playExtendNum)
        print(item + '处理中...')
        for n in range(levBeginLine, dayCount):
            for m in range(playerCount):
                l_LevList[n][m] = sumLevList[n][m][item]
                l_BreakList[n][m] = sumBreakList[n][m][item]

        setRangeData(l_LevList, levSht, levBeginLine + 1, itemLevStr, levBeginLine)
        setRangeData(l_BreakList, levSht, levBeginLine + 1, itemBreakStr, levBeginLine)


def getLimitNum(limitDict: map, day, colStr):
    isTest = True  # 是否取测试上限
    limitNum = 999
    if isTest:
        return limitDict['二测'][colStr]
    else:
        for key in limitDict:
            if not isinstance(key, str) and key <= day:
                if limitDict[key][colStr] < limitNum:
                    limitNum = limitDict[key][colStr]
        return limitNum


def getLimitDict(levList):
    """获取数量限制字典

    Args:
        levList (_type_): _description_

    Returns:
        _type_: _description_
    """
    limitDict = {}
    limitCount = 0
    for i in range(len(levList)):
        if levList[i][0] == '资源上限':
            limitCount += 1
        if limitCount > 1:
            break
        elif limitCount > 0:
            dayDict = {}
            for j in range(len(levList[0])):
                if levList[0][j] is not None and levList[0][j] != '':
                    dayDict[levList[0][j]] = levList[i][j]
            limitDict[levList[i][0]] = dayDict
    return limitDict
