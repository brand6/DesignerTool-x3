import random


# 脚本说明：
# 用于模拟统计抽卡掉率
#
def main():
    drawTimes = 10000000
    rateList = [75, 150, 450, 900, 8425]
    guaranteeMap = {3: 0, 60: 0, 10: 0}
    resultMap = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0}

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

    def Draw(limit=5):
        maxNum = 0
        for r in rateList[:limit]:
            maxNum += r
        rndNum = random.randint(0, maxNum)
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
            Draw(2)
        elif guaranteeMap[10] >= 9:
            Draw(4)
        else:
            Draw(5)

    totalNum = 0
    for k in resultMap:
        totalNum += resultMap[k]
        resultMap[k] = str(round(resultMap[k] / drawTimes * 100, 2)) + '%'

    print(str(totalNum / 10000) + ' 万抽结果')
    print(resultMap)
