import random


def testWinRate(funType):
    testTimes = 100000
    cellList = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    winTimes = [0] * len(cellList)
    for i in range(len(cellList)):
        for _ in range(testTimes):
            cellCount = cellList[i]
            winFlag = True
            funFlag = True
            while True:
                num = random.randint(1, 6)  # 扔骰子
                if funType == -1 and winFlag is False and funFlag is True:  # 对方数字减半
                    funFlag = False
                    num = int(num / 2)
                    if num < 1:
                        num = 1
                elif funType > 0 and winFlag is True and funFlag is True:  # 我方加步数
                    funFlag = False
                    num = num + funType

                if num >= cellCount:
                    if winFlag is True:
                        winTimes[i] += 1
                    break
                else:
                    cellCount -= num
                    winFlag = not winFlag

    for i in range(len(winTimes)):
        print(i + 2, winTimes[i] / testTimes)


# -1:对方减半，其他数字：我方骰子增加的数字
testWinRate(-1)
