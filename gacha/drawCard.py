from gacha import drawRole
from common.common import isNumber


# 脚本说明：
# 23-4-11思念up活动池抽卡掉率
#
def main():
    drawTimes = 10000000
    resultMap = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 's': 0, 'c': 0}
    rateList = [0, 120, 120, 600, 9160]  # ssr-score,ssr-card,sr-score,sr-card,r-card
    upCount = 40
    upPara = 1000
    guaranteeMap = {2: 0, 10: 0, upCount: 0}  # 2抽SSR必出活动角色保底，10抽必出SR+保底

    for _ in range(drawTimes):
        curRateList = drawRole.getSSRUpRateList(guaranteeMap[upCount], rateList, upPara, upCount)
        if guaranteeMap[10] >= 9:
            curRateList = drawRole.getSRUpRateList(guaranteeMap[10], curRateList)
        drawRole.Draw(curRateList, resultMap, guaranteeMap, upCount)

    totalNum = 0
    for k in resultMap:
        if isNumber(k):
            totalNum += resultMap[k]
        resultMap[k] = str(round(resultMap[k] / drawTimes * 100, 2)) + '%'

    print(str(totalNum / 10000) + ' 万抽结果')
    print(resultMap)
