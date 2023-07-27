import random


# 脚本说明：
# 旧版抽卡掉率
#
def main():
    drawTimes = 10000000
    rateList = [40, 80, 240, 480, 9160]  # ssr-score,ssr-card,sr-score,sr-card,r-card
    SSRRateList = [3333, 6667]  # ssr-score,ssr-card
    SRRateList = [40, 80, 3293, 6587]  # ssr-score,ssr-card,sr-score,sr-card
    upCount = 50
    upPara = [333, 667, -24, -49, -927]  # 50抽不出SSR，开始概率递增，每次递增10%
    SRUpPara = [333, 667, -333, -667]  # 保底时的概率递增
    guaranteeMap = {3: 0, 60: 0, 10: 0}  # 3抽SSR必出角色保底，60抽必出SSR保底，10抽必出SR+保底
    resultMap = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0}

    def getCurRateList(count, rateList, upPara):
        curRateList = [0] * len(rateList)
        if count > upCount:
            for i in range(len(rateList)):
                curRateList[i] = rateList[i] + upPara[i] * (count - upCount)
        else:
            for i in range(len(rateList)):
                curRateList[i] = rateList[i]
        return curRateList

    def getSSR_S():
        resultMap[0] += 1
        guaranteeMap[3] = 0
        guaranteeMap[60] = 0
        guaranteeMap[10] = 0

    def getSSR_C():
        if guaranteeMap[3] >= 2:
            getSSR_S()
        else:
            resultMap[1] += 1
            guaranteeMap[3] += 1
            guaranteeMap[60] = 0
            guaranteeMap[10] = 0

    def getSR_S():
        resultMap[2] += 1
        guaranteeMap[60] += 1
        guaranteeMap[10] = 0

    def getSR_C():
        resultMap[3] += 1
        guaranteeMap[60] += 1
        guaranteeMap[10] = 0

    def getR_C():
        resultMap[4] += 1
        guaranteeMap[60] += 1
        guaranteeMap[10] += 1

    def Draw(rateList):
        rndNum = random.randint(0, 10000)
        for j in range(len(rateList)):
            if rndNum <= rateList[j]:
                if j == 0:
                    getSSR_S()
                elif j == 1:
                    getSSR_C()
                elif j == 2:
                    getSR_S()
                elif j == 3:
                    getSR_C()
                else:
                    getR_C()
                break
            else:
                rndNum -= rateList[j]

    for i in range(drawTimes):
        if guaranteeMap[60] >= 59:
            Draw(SSRRateList)
        elif guaranteeMap[10] >= 9:
            curRateList = getCurRateList(guaranteeMap[60], SRRateList, SRUpPara)
            Draw(curRateList)
        else:
            curRateList = getCurRateList(guaranteeMap[60], rateList, upPara)
            Draw(curRateList)

    totalNum = 0
    for k in resultMap:
        totalNum += resultMap[k]
        resultMap[k] = str(round(resultMap[k] / drawTimes * 100, 2)) + '%'

    print(str(totalNum / 10000) + ' 万抽结果')
    print(resultMap)
