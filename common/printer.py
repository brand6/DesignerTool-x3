import colorama
import sys
from time import time


class Printer:
    colorDict = {'red': '31m', 'green': '32m', 'yellow': '33m', 'blue': '34m', 'white': '37m'}

    def __init__(self, compareTime=0, compareColor='yellow'):
        self.startTime = time()
        self.compareTime = compareTime
        self.compareColor = compareColor
        colorama.init(autoreset=True)

    # 附带颜色打印
    @classmethod
    def printColor(cls, content, colorStr='', skipLines=0):
        if colorStr in cls.colorDict:
            print("\033[1;" + cls.colorDict[colorStr] + content + "\033[0m ")
        else:
            print(content)

        while skipLines > 0:
            print("")
            skipLines = skipLines - 1

    # 打印时间
    @classmethod
    def printTime(cls, desc, time=time(), precision=2, colorStr='', skipLines=0):
        time = round(time, precision)
        cls.printColor(desc + str(time) + '秒', colorStr, skipLines)

    # 设置计时的开始时间
    def setStartTime(self, content='', colorStr='', skipLines=0):
        if content != '':
            self.printColor(content, colorStr, skipLines)
        self.startTime = time()

    # 设置计时的比较时间
    def setCompareTime(self, compareTime, compareColor='yellow'):
        self.compareTime = compareTime
        self.compareColor = compareColor

    # 打印间隔时间[compareTime：用于比较的时间，colorTime：比较值大于这个值的时高亮]
    def printGapTime(self, desc, precision=2, colorStr='', skipLines=0, colorTime=1, isCompare=False):
        gapTime = time() - self.startTime
        if isCompare and abs(gapTime - self.compareTime) > colorTime:
            self.printTime(desc, gapTime, precision, self.compareColor, skipLines)
        else:
            self.printTime(desc, gapTime, precision, colorStr, skipLines)
        return gapTime

    # 打印进度
    def printProgress(self, i, maxNum):
        sys.stdout.write('\r%d / %d' % (i, maxNum))
        sys.stdout.flush()
