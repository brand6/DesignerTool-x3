import random
from common.common import isNumber


# 脚本说明：
# 23-4-11搭档up活动池抽卡掉率
#
def main():
    drawTimes = 10000000
    resultMap = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 's': 0, 'c': 0}
    rateList = [120, 0, 120, 600, 9160]  # ssr-score,ssr-card,sr-score,sr-card,r-card
    upCount = 80
    upPara = 1000
    guaranteeMap = {2: 0, 10: 0, upCount: 0}  # 2抽SSR必出活动角色保底，10抽必出SR+保底

    for _ in range(drawTimes):
        curRateList = getSSRUpRateList(guaranteeMap[upCount], rateList, upPara, upCount)
        if guaranteeMap[10] >= 9:
            curRateList = getSRUpRateList(guaranteeMap[10], curRateList)
        Draw(curRateList, resultMap, guaranteeMap, upCount)

    totalNum = 0
    for k in resultMap:
        if isNumber(k):
            totalNum += resultMap[k]
        resultMap[k] = str(round(resultMap[k] / drawTimes * 100, 2)) + '%'

    print(str(totalNum / 10000) + ' 万抽结果')
    print(resultMap)


def Draw(rateList, resultMap, guaranteeMap, upCount):
    rndNum = random.randint(1, 10000)
    for j in range(len(rateList)):
        if rndNum <= rateList[j]:
            if j == 0:
                getSSR_S(resultMap, guaranteeMap, upCount)
            elif j == 1:
                getSSR_C(resultMap, guaranteeMap, upCount)
            elif j == 2:
                getSR_S(resultMap, guaranteeMap, upCount)
            elif j == 3:
                getSR_C(resultMap, guaranteeMap, upCount)
            else:
                getR_C(resultMap, guaranteeMap, upCount)
            break
        else:
            rndNum -= rateList[j]


def getSSR_S(resultMap, guaranteeMap, upCount):
    resultMap[0] += 1
    guaranteeMap[upCount] = 0
    guaranteeMap[10] = 0

    if guaranteeMap[2] == 1:
        resultMap['s'] += 1
        guaranteeMap[2] = 0
    else:
        rnd = random.randint(1, 10000)
        if rnd <= 5000:
            resultMap['s'] += 1
            guaranteeMap[2] = 0
        else:
            guaranteeMap[2] += 1


def getSSR_C(resultMap, guaranteeMap, upCount):
    resultMap[1] += 1
    guaranteeMap[upCount] = 0
    guaranteeMap[10] = 0
    if guaranteeMap[2] == 1:
        resultMap['c'] += 1
        guaranteeMap[2] = 0
    else:
        rnd = random.randint(0, 9999)
        if rnd < 5000:
            resultMap['c'] += 1
            guaranteeMap[2] = 0
        else:
            guaranteeMap[2] += 1


def getSR_S(resultMap, guaranteeMap, upCount):
    resultMap[2] += 1
    guaranteeMap[upCount] += 1
    guaranteeMap[10] = 0


def getSR_C(resultMap, guaranteeMap, upCount):
    resultMap[3] += 1
    guaranteeMap[upCount] += 1
    guaranteeMap[10] = 0


def getR_C(resultMap, guaranteeMap, upCount):
    resultMap[4] += 1
    guaranteeMap[upCount] += 1
    guaranteeMap[10] += 1


def getSRUpRateList(count, rateList):
    """连续未抽出sr+，sr概率提升

    Args:
        count (_type_): 连续未出SR+的次数
        rateList (_type_): 基础出货率

    Returns:
        _type_: sr概率提升后的出货率
    """
    curRateList = [0] * len(rateList)
    addSum = rateList[2] + rateList[3]
    if count >= 9 and addSum > 0:
        for i in range(len(rateList)):
            if i < 2:
                curRateList[i] = rateList[i]
            elif i < 4:
                curRateList[i] = rateList[i] + rateList[i] / addSum * rateList[4]
            else:
                curRateList[i] = 0
    else:
        for i in range(len(rateList)):
            curRateList[i] = rateList[i]
    return curRateList


def getSSRUpRateList(count, rateList, upPara, upCount):
    """连续未抽出ssr，ssr概率提升

    Args:
        count (_type_): 连续未出SSR的次数
        rateList (_type_): 基础出货率
        upPara (_type_): 提升权重
        upCount (_type_): 提升基准值

    Returns:
        _type_: ssr概率提升后的出货率
    """
    curRateList = [0] * len(rateList)
    if count >= upCount:
        addSum = rateList[0] + rateList[1]
        minusSum = rateList[2] + rateList[3] + rateList[4]
        changeWeight = (count - upCount + 1) * upPara
        if changeWeight > minusSum:
            changeWeight = minusSum
        for i in range(len(rateList)):
            if i < 2:
                curRateList[i] = rateList[i] * (1 + 1 / addSum * changeWeight)
            else:
                curRateList[i] = rateList[i] * (1 - 1 / minusSum * changeWeight)
    else:
        for i in range(len(rateList)):
            curRateList[i] = rateList[i]
    return curRateList
