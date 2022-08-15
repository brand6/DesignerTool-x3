import xlwings as xw
import random

# region 全局变量
gridType = {1: "红", 2: "黄", 3: "蓝", 4: "绿"}  # 格子颜色
maxCardNum = 10  # 最大 手牌数
averageScore = 4.1  # 摸牌平均得分（一次摸牌对应的分值）
exceptScore = 3.5  # 期望得分（常规情况下，得分大于该值时出牌）
funCardExceptPower = 6.5  # 功能牌强度期望


class strategy():
    抢翻倍格 = '抢翻倍格'
    最大得分 = '最大得分'


class role():
    ST = 'ST'
    YS = 'YS'
    RY = 'RY'


class AIStatus():
    普通 = '普通'
    放水 = '放水'
    屯牌 = '屯牌'
    随性 = '随性'


class funCard():
    冰冻 = '冰冻'
    跳过 = '跳过'
    毁天灭地 = '毁天灭地'
    发财 = '发财'
    复制 = '复制'  # 判断弃牌堆价值，决定出牌顺序
    贪心 = '贪心'
    兴奋剂 = '兴奋剂'
    窥视 = '窥视'
    赌狗喵 = '赌狗喵'
    暗言术 = '暗言术•滚'
    拆迁喵 = '拆迁喵'
    否决 = '否决'


# 功能牌出牌顺序
funCardOrderList = [
    funCard.跳过,
    funCard.贪心,
    funCard.冰冻,
    funCard.毁天灭地,
    funCard.窥视,
    funCard.发财,
    funCard.赌狗喵,
    funCard.暗言术,
    funCard.拆迁喵,
    funCard.复制,
    funCard.兴奋剂,
]

# endregion


class MiaoCard():
    # region 成员变量
    sht = None  # 数据表格对象
    winLog = ''  # 批量数据
    # 牌局设定
    isBasic = True  # 标记基础版/功能版
    strategy = ''  # AI出牌策略
    AIRole = role.ST  # AI对应男主
    AIStatus = ''  # AI状态
    firstMove = ''  # 先手
    # 状态相关数据
    PLScore = 0  # 玩家当前得分
    AIScore = 0  # AI当前得分
    roundTimes = 0  # 游戏内行动次数
    nextMove = ''  # 下回合行动对象
    isPL = False  # 当前行动对象是否玩家
    AIFrozen = False  # 冰冻：跳过功能牌出牌
    AISkip = False  # 跳过数字牌阶段
    AIHigh = False  # 出2张数字牌
    PLFrozen = False
    PLSkip = False
    PLHigh = False
    # 获胜次数相关
    AIWinTimes = 0
    PLWinTimes = 0
    peaceTimes = 0
    allRound = 0
    allAction = 0
    # 数字牌相关
    cardLib = []  # 基础牌库
    cardList = []  # AI手牌
    PLCardList = []  # 玩家手牌
    # 功能牌相关
    funCardLib = []  # 功能牌库
    usedFunCardLib = []  # 功能牌弃牌堆
    funCardPowerMap = {}  # 功能牌对应强度
    funCardList = []  # AI功能手牌
    PLFunCardList = []  # 玩家功能手牌
    # 日志相关
    log = []  # AI行动日志
    PLLog = []  # 玩家行动日志
    getScore = 0  # 本次行动得分，只用于打印Log
    # 最大得分牌相关
    maxScore = 0
    maxCard = ''
    maxPutIndex = -1
    # 最优翻倍牌相关
    doubleScore = 0
    doubleCard = ''
    doublePutIndex = -1
    # 场上牌数据相关
    gridTypes = []  # 存储牌局格子类型
    gridData = []  # 存储牌局出牌信息
    emptyNum = 0  # 空格数量
    doubleEmptyNum = 0  # 空翻倍格数量
    emptyDict = {}  # 各类空格数量
    exceptMaxScore = 0  # 最大可能得分
    enemyCardInGrid = []  # 卡牌数据为元组（得分，牌面数字，牌面颜色，位置）
    myCardInGrid = []

    # endregion

    # region excel相关
    def __init__(self, _sht=None):
        """初始化，从excel内读取数据

        Args:
            _sht (sheet, optional): 是否指定sht. Defaults to None.
        """
        self.sht = xw.books.active.sheets.active if _sht is None else _sht
        self.dataRng = self.sht.used_range.value
        # 牌局设定
        self.strategy = self.dataRng[0][12]
        self.AIRole = self.dataRng[1][12]
        self.isBasic = self.dataRng[2][12]
        self.firstMove = self.dataRng[0][14]
        # 游戏状态
        self.nextMove = self.dataRng[2][8]
        self.roundTimes = self.dataRng[2][10]
        self.PLScore = self.dataRng[3][8] if self.dataRng[3][8] is not None else 0
        self.AIScore = self.dataRng[4][8] if self.dataRng[4][8] is not None else 0
        self.isPL = False if self.nextMove == 'AI' else True
        # 格子数据
        self.gridTypes = []
        self.gridData = []
        for i in range(1, 9):
            self.gridTypes.append(self.dataRng[i][4])
            self.gridData.append(self.dataRng[i][5])
        # 牌库数据
        self.GetNotNoneRowList(self.cardLib, 1, 40, 1)
        # 数字手牌数据
        self.GetNotNoneColList(self.cardList, 5, 8, 17)
        self.GetNotNoneColList(self.PLCardList, 6, 8, 17)
        if not self.isBasic:
            self.AIStatus = self.dataRng[1][14]
            self.PLFrozen = self.dataRng[3][10]
            self.AIFrozen = self.dataRng[4][10]
            self.PLSkip = self.dataRng[3][12]
            self.AISkip = self.dataRng[4][12]
            self.PLHigh = self.dataRng[3][14]
            self.AIHigh = self.dataRng[4][14]
            self.GetNotNoneRowList(self.funCardLib, 1, 28, 2)  # 功能牌库
            self.GetNotNoneRowList(self.usedFunCardLib, 11, 38, 3)  # 功能弃牌库
            self.GetNotNoneColList(self.funCardList, 7, 8, 17)  # 功能手牌
            self.GetNotNoneColList(self.PLFunCardList, 8, 8, 17)
            # 功能牌强度
            for i in range(1, 13):
                self.funCardPowerMap[self.dataRng[i][19]] = self.dataRng[i][20]

    def ShowChange(self):
        """在excel内显示改变后的数据
        """
        # 状态数据
        self.dataRng[0][8] += self.PLWinTimes
        self.dataRng[1][8] += self.AIWinTimes
        self.dataRng[2][8] = self.nextMove
        self.dataRng[3][8] = self.PLScore
        self.dataRng[4][8] = self.AIScore
        self.dataRng[1][10] += self.peaceTimes
        self.dataRng[2][10] = self.roundTimes
        # 手牌数据
        self.SetColList(self.cardList, 5, 8, 17)
        self.SetColList(self.PLCardList, 6, 8, 17)
        # 牌堆数据
        self.SetRowList(self.cardLib, 1, 40, 1)
        if not self.isBasic:
            self.dataRng[3][10] = self.PLFrozen
            self.dataRng[4][10] = self.AIFrozen
            self.dataRng[3][12] = self.PLSkip
            self.dataRng[4][12] = self.AISkip
            self.dataRng[1][14] = self.AIStatus
            self.dataRng[3][14] = self.PLHigh
            self.dataRng[4][14] = self.AIHigh
            self.SetColList(self.funCardList, 7, 8, 17)
            self.SetColList(self.PLFunCardList, 8, 8, 17)
            self.SetRowList(self.funCardLib, 1, 28, 2)
            self.SetRowList(self.usedFunCardLib, 11, 38, 3)
        # 格子数据
        self.SetRowList(self.gridTypes, 1, 8, 4)
        self.SetRowList(self.gridData, 1, 8, 5)
        # 日志数据
        if len(self.log) > 0:
            self.setLogList(self.log, 7)
        if len(self.PLLog) > 0:
            self.setLogList(self.PLLog, 13)
        self.sht.range("A1").value = self.dataRng

    def PrintWinLog(self):
        """更新批量运行结果数据
        """
        runTimes = self.sht.range("K1").value
        winLog = []
        winLog.append(self.firstMove)
        winLog.append(self.AIRole)
        winLog.append(self.strategy)
        winLog.append(self.AIStatus)
        winLog.append(round(self.PLWinTimes / runTimes, 2))
        winLog.append(round(self.AIWinTimes / runTimes, 2))
        winLog.append(round(self.peaceTimes / runTimes, 2))
        winLog.append(round(self.allRound / runTimes, 1))
        winLog.append(round(self.allAction / runTimes, 1))
        for i in range(16, len(self.dataRng)):
            if self.dataRng[i][19] is None:
                break
        self.sht.range("T" + str(i + 1)).value = winLog

    def setLogList(self, logList: list, col: int):
        """更新log数据

        Args:
            logList (list): 显示的log列表
            col (int): log显示的excel的列
        """
        firstEmptyRow = 10
        for r in range(firstEmptyRow, len(self.dataRng)):
            if self.dataRng[r][col] is None:
                firstEmptyRow = r
                break
        else:
            firstEmptyRow = r + 1
        # 数据超出原来列表时增加新的行
        newLines = firstEmptyRow + len(logList) - len(self.dataRng)
        if newLines > 0:
            colNum = len(self.dataRng[0])
            for i in range(newLines):
                self.dataRng.append([None] * colNum)
        for r in range(len(logList)):
            self.dataRng[firstEmptyRow + r][col] = logList[r]

    def SetRowList(self, dataList: list, startRow: int, endRow: int, col: int):
        """按行更新数据

        Args:
            dataList (list): 需更新的数据列表
            startRow (int): 数据开始行
            endRow (int): 数据结束行
            col (int): 数据所在列
        """
        for i in range(startRow, endRow + 1):
            index = i - startRow
            if index < len(dataList):
                self.dataRng[i][col] = dataList[index]
            else:
                self.dataRng[i][col] = None

    def SetColList(self, dataList: list, row: int, startCol: int, endCol: int):
        """按列更新数据

        Args:
            dataList (list): 需更新的数据列表
            row (int): 数据所在行
            startCol (int): 数据开始列
            endCol (int): 数据结束列
        """
        for i in range(startCol, endCol + 1):
            index = i - startCol
            if index < len(dataList):
                self.dataRng[row][i] = dataList[index]
            else:
                self.dataRng[row][i] = None

    def GetNotNoneRowList(self, dataList: list, startRow: int, endRow: int, col: int):
        """按行获取非空数据

        Args:
            dataList (list): 获取数据的列表
            startRow (int): 数据开始行
            endRow (int): 数据结束行
            col (int): 数据所在列
        """
        dataList.clear()
        for i in range(startRow, endRow + 1):
            if self.dataRng[i][col] is not None:
                dataList.append(self.dataRng[i][col])
            else:
                break

    def GetNotNoneColList(self, dataList: list, row: int, startCol: int, endCol: int):
        """按列获取非空数据

        Args:
            dataList (list): 获取数据的列表
            row (int): 数据所在行
            startCol (int): 数据开始列
            endCol (int): 数据结束列
        """
        dataList.clear()
        for j in range(startCol, endCol + 1):
            if self.dataRng[row][j] is not None:
                dataList.append(self.dataRng[row][j])
            else:
                break

    def Log(self, _str: str, isPL: bool = None):
        """添加到行动日志列表

        Args:
            _str (str): 日志内容
            isPL (bool, optional): 是否指定对象. Defaults to None.
        """
        isPL = self.isPL if isPL is None else isPL
        logStr = str(int(self.roundTimes)) + '-' + _str
        self.log.append(logStr) if not isPL else self.PLLog.append(logStr)

    def ChangeWinData(self):
        """修改获胜数据
        """
        self.allRound += self.roundTimes
        if self.AIScore > self.PLScore:
            self.AIWinTimes += 1
            self.Log('AI 获胜，AI得分：' + str(self.AIScore) + '，玩家得分：' + str(self.PLScore))
        elif self.AIScore < self.PLScore:
            self.PLWinTimes += 1
            self.Log('玩家 获胜，AI得分：' + str(self.AIScore) + '，玩家得分：' + str(self.PLScore))
        else:
            self.peaceTimes += 1
            self.Log('平局，得分：' + str(self.AIScore))

# endregion

# region 初始化相关

    def GameBegin(self, clearLog=True):
        """开始一轮新的游戏
        """
        self.InitGrid()
        self.InitData(clearLog)
        self.GetCard()
        self.GetCard()
        self.GetCard(True) if self.firstMove == 'AI' else self.GetCard()
        self.GetCard(True)
        self.GetCard(True)
        if not self.isBasic:
            self.GetFunCard()
            self.GetFunCard(True) if self.firstMove == 'AI' else self.GetFunCard()
            self.GetFunCard(True)

    def InitGrid(self):
        """初始化格子类型和数据
        """
        colorList = [1, 1, 2, 2, 3, 3, 4, 4]
        self.gridTypes = []  # 存储格子类型
        self.gridData = []  # 存储格子数据
        colorCount = 0
        rnd = 0
        for i in range(8):
            self.gridData.append(None)
            if colorCount < 4:
                rnd = random.randint(0, len(colorList) - 1)
                self.gridTypes.append(gridType[colorList[rnd]])
                colorList.remove(colorList[rnd])
                colorCount += 1
            else:
                self.gridTypes.append(None)

    def InitData(self, clearLog=True):
        """初始化游戏数据
        """
        self.AIScore = 0
        self.PLScore = 0
        self.roundTimes = 0
        self.emptyNum = 8
        self.nextMove = 'AI' if self.firstMove == 'AI' else '玩家'
        self.isPL = False
        self.AIFrozen = False
        self.AISkip = False
        self.AIHigh = False
        self.PLFrozen = False
        self.PLSkip = False
        self.PLHigh = False
        # 初始化牌库
        self.cardLib = []
        for i in range(1, 41):
            self.cardLib.append(self.dataRng[i][0])
        if not self.isBasic:
            self.usedFunCardLib = []
            self.funCardLib = []
            for i in range(1, 29):
                self.funCardLib.append(self.dataRng[i + 41][0])
        # 清空手牌
        self.cardList = []
        self.PLCardList = []
        self.funCardList = []
        self.PLFunCardList = []
        # 清空日志
        if clearLog:
            for i in range(10, len(self.dataRng)):
                self.dataRng[i][7] = None
                self.dataRng[i][13] = None

# endregion

# region 数字牌决策相关

    def DoAction(self):
        """ 行动决策
        """
        self.allAction += 1
        if self.isPL:
            cardList = self.PLCardList
            enemyCardList = self.cardList
            funCardList = self.PLFunCardList
            score = self.PLScore
            enemyScore = self.AIScore
            self.nextMove = 'AI'
        else:
            cardList = self.cardList
            enemyCardList = self.PLCardList
            funCardList = self.funCardList
            score = self.AIScore
            enemyScore = self.PLScore
            self.nextMove = '玩家'

        self.GetEmptyGridNum()
        self.GetExceptScore()
        self.GetMaxScoreCard()
        if self.isPL and (self.PLSkip or self.PLHigh):
            if self.PLSkip:
                self.Log('跳过状态，跳过数字牌阶段')
            elif self.PLHigh:
                self.Log('兴奋状态，可出2张牌')
                self.PutCard()
                self.GetMaxScoreCard()
                self.PutCard()
            self.PLSkip = False
            self.PLHigh = False
            return
        elif not self.isPL and (self.AISkip or self.AIHigh):
            if self.AISkip:
                self.Log('跳过状态，跳过数字牌阶段')
            elif self.AIHigh:
                self.Log('兴奋状态，可出2张牌')
                self.PutCard()
                self.GetMaxScoreCard()
                self.PutCard()
            self.AISkip = False
            self.AIHigh = False
            return

        # 无空格时直接返回
        if self.emptyNum == 0:
            return
        # 无手牌时摸牌
        if len(cardList) == 0:
            self.Log("无手牌时：摸牌")
            return self.GetCard()
        # 手牌将满时出牌
        if len(cardList) + len(funCardList) >= maxCardNum - 1:
            self.Log("手牌将满时：出牌")
            return self.PutCard()
        # 1空时出牌能赢则出牌，否则摸牌
        if self.emptyNum == 1:
            if score + self.maxScore >= enemyScore:
                self.Log("1空时出牌能赢或平：出牌")
                return self.PutCard()
            elif self.maxScore >= self.exceptMaxScore:
                self.Log("基础版1空时拥有最大分牌：出牌")
                return self.PutCard()
            else:
                self.Log("1空时出牌会输：摸牌")
                return self.GetCard()
        # 数字牌数>(空格数+1)/2， 对方数字牌数不足，有机会就抢占更多的格子
        if len(cardList) > (self.emptyNum + 1) / 2 and \
           (self.emptyNum % 2 == 0 and len(enemyCardList) < (self.emptyNum - 1) / 2 or len(enemyCardList) < (self.emptyNum - 4) / 2):
            if (self.maxScore == 0 and score < enemyScore):
                self.Log("抢格子不可得分且积分落后：摸牌")
                return self.GetCard()
            else:
                # 可得分抢格子
                self.Log("抢格子可得分：出牌", )
                return self.PutCard()
        # 手牌数足够时，不出牌会被对方多占一格：出牌
        if self.emptyNum % 2 == 1 and len(enemyCardList) > (self.emptyNum - 1) / 2 and len(cardList) > (self.emptyNum - 1) / 2:
            if (self.maxScore == 0 and score < enemyScore):
                self.Log("即将被抢格子，但不可得分且积分落后：摸牌")
                return self.GetCard()
            else:
                # 可得分防止被抢格子
                self.Log("即将被抢格子，可得分：出牌")
                return self.PutCard()
        # 2空时，如果出牌后积分不占优势则摸牌
        elif self.emptyNum == 2:
            # 出牌后积分>对方+得分期望
            if self.maxScore + score > enemyScore + exceptScore:
                self.Log("2空时，出牌后积分>对方期望积分：出牌")
                return self.PutCard()
            else:
                self.Log("2空时，出牌后积分<=对方期望积分:摸牌")
                return self.GetCard()
        # 抢翻倍格
        elif self.doubleScore > 0 and self.strategy == strategy.抢翻倍格:
            self.Log("抢翻倍格：出牌")
            return self.PutCard(isDouble=1)
        # 可得分大于期望值时出牌
        elif self.maxScore > exceptScore:
            self.Log("可得分大于期望值：出牌")
            return self.PutCard()
        # 可得分小于期望值时摸牌
        else:
            self.Log("可得分小于期望值：摸牌")
            return self.GetCard()

    def GetMaxScoreCard(self, gridData=None, cardList=None):
        """获取最优得分相关数据（不可得分时出最小牌）

        Args:
            gridData (list, optional): 是否指定牌局数据. Defaults to None.\n
            cardList (list,optional)：是否指定卡牌list. Defaults to None.
        Returns:
            score(int): 最大得分
        Updates:
            self.maxPutIndex self.doublePutIndex \n
            self.MaxCard self.doubleCard \n
            self.maxScore self.doubleScore \n
        """
        # 最大得分牌
        self.maxScore = 0
        self.maxCard = ''
        self.maxPutIndex = -1
        # 最优翻倍牌
        self.doubleScore = 0
        self.doubleCard = ''
        self.doublePutIndex = -1

        if gridData is None:
            gridData = self.gridData
        if cardList is None:
            cardList = self.cardList if not self.isPL else self.PLCardList
        if len(cardList) == 0:
            return 0
        checkList = []
        doubleColorNum = 2
        for i in range(len(gridData)):
            if gridData[i] is None:
                if self.gridTypes[i] is None and 'normal' not in checkList:
                    checkList.append('normal')
                    # 普通格最大得分 = 最大牌
                    num = int(cardList[0][1])
                    # 空格>3时：YS普通格最大得分 = 第二大牌
                    # if not self.isPL and self.emptyNum >3 and self.AIRole == role.YS and len(cardList)>1:
                    #    num = int(cardList[1][1])
                    if num > self.maxScore:
                        self.maxPutIndex = i
                        self.maxCard = cardList[0]
                        self.maxScore = num
                elif self.gridTypes[i] is not None and self.gridTypes[i] not in checkList:
                    checkList.append(self.gridTypes[i])
                    for c in cardList:
                        num = int(c[1]) * 2
                        if c[0] == self.gridTypes[i]:
                            # 满足条件更新最优翻倍牌
                            if num > 0 and self.emptyDict[c[0]] < doubleColorNum or \
                               num > self.doubleScore and self.emptyDict[c[0]] == doubleColorNum:
                                self.doubleCard = c
                                self.doubleScore = num
                                self.doublePutIndex = i
                                doubleColorNum = self.emptyDict[c[0]]
                            # 翻倍格得分=普通格时，优先出翻倍格
                            if num >= self.maxScore:
                                self.maxCard = c
                                self.maxScore = num
                                self.maxPutIndex = i
                            break
                    # 翻倍格不可得分,设置为最小的牌
                    else:
                        if self.maxPutIndex == -1:
                            self.maxCard = c
                            self.maxPutIndex = i
                        if self.doublePutIndex == -1:
                            self.doubleCard = c
                            self.doublePutIndex = i
        return self.maxScore

    def GetExceptScore(self):
        """获取最大可能得分（判断牌堆情况）
        """
        self.exceptMaxScore = 0
        cardList = self.cardLib
        checkList = []
        for i in range(len(self.gridData)):
            if self.gridData[i] is None:
                if self.gridTypes[i] is None and 'normal' not in checkList:
                    checkList.append('normal')
                    for c in cardList:
                        num = int(c[1])
                        if num == 6:
                            self.exceptMaxScore = num
                            break
                        elif num > self.exceptMaxScore:
                            self.exceptMaxScore = num
                elif self.gridTypes[i] is not None and self.gridTypes[i] not in checkList:
                    checkList.append(self.gridTypes[i])
                    for c in cardList:
                        num = int(c[1]) * 2
                        if c[0] == self.gridTypes[i]:
                            if num == 12:
                                self.exceptMaxScore = num
                                break
                            elif num > self.exceptMaxScore:
                                self.exceptMaxScore = num
                    if self.exceptMaxScore == 12:
                        break

    def GetEmptyGridNum(self):
        """获取空格数，空翻倍格数
        """
        self.emptyNum = 0
        self.doubleEmptyNum = 0
        self.emptyDict = {'normal': 0, "红": 0, "黄": 0, "蓝": 0, "绿": 0}
        for i in range(len(self.gridData)):
            if self.gridData[i] is None:
                self.emptyNum += 1
                if self.gridTypes[i] is not None:
                    self.doubleEmptyNum += 1
                    self.emptyDict[self.gridTypes[i]] += 1
                else:
                    self.emptyDict['normal'] += 1

    def PutCard(self, index=-1, card='', isDouble=-1, color=''):
        """
        出牌逻辑（可得分时最大得分出牌，不可得分时出最小牌）\n
        指定位置出指定牌 传入index,card \n
        普通格出牌 isDouble = 0 \n
        翻倍格出牌 isDouble = 1 \n
        指定翻倍格颜色出牌 isDouble = 1,color = 颜色 \n
        任意格出牌 无参 \n
        Args:
            index (int, optional): 指定出牌位置. Defaults to -1.
            card (str, optional): 指定出的牌. Defaults to ''.
            isDouble (int, optional):0：出牌普通格 1：出牌翻倍格. Defaults to -1.
            color (str, optional): 指定翻倍格颜色出牌. Defaults to ''.
        """
        if self.strategy == strategy.抢翻倍格 and self.doubleScore > 0 and self.emptyDict['normal'] != 1:
            isDouble = 1
        cardList = self.cardList if not self.isPL else self.PLCardList

        # 指定位置出指定牌
        if index != -1 and card != '':
            self.AddScore(card, index)
        # 普通格出牌
        elif isDouble == 0:
            if self.emptyNum - self.doubleEmptyNum > 0:
                card = cardList[0]
                for i in range(len(self.gridData)):
                    if self.gridData[i] is None and self.gridTypes[i] is None:
                        self.AddScore(card, i)
                        break
            else:
                self.Log('不存在空的普通格')
        # 翻倍格出牌
        elif isDouble == 1:
            if self.doubleEmptyNum > 0:
                if color == '':
                    self.AddScore(self.doubleCard, self.doublePutIndex)
                else:
                    # 指定翻倍格颜色出牌
                    for i in range(len(self.gridData)):
                        if self.gridTypes[i] == color and self.gridData[i] is not None:
                            for card in cardList:
                                if card[0] == color:
                                    self.AddScore(card, i)
                                    break
                                else:
                                    self.AddScore(card, i)
                                break
                    else:
                        self.Log('不存在空的' + color + '色翻倍格')
            else:
                self.Log('不存在空的翻倍格')
        # 任意格出牌
        else:
            self.AddScore(self.maxCard, self.maxPutIndex)
        # 更新空格数量
        self.emptyNum -= 1

    def AddScore(self, card: str, index: int):
        """得分，并将手牌移到牌局

        Args:
            card (str): 出数字牌
            index (int): 出牌位置
        """
        if card == '':
            return self.Log('异常！无牌可出')
        if self.gridTypes[index] is None:
            self.getScore = int(card[1])
        elif self.gridTypes[index] == card[0]:
            self.getScore = 2 * int(card[1])
        if not self.isPL:
            self.cardList.remove(card)
            self.AIScore += self.getScore
            self.gridData[index] = 'AI-' + card + '-' + str(self.getScore) + '分'
        else:
            self.PLCardList.remove(card)
            self.PLScore += self.getScore
            self.gridData[index] = '玩家-' + card + '-' + str(self.getScore) + '分'
        self.Log("出牌位置" + str(index) + "，得分" + str(self.getScore))

# endregion

# region 抽数字牌相关函数

    def GetCard(self, isPL: bool = None):
        """抽基础牌

        Args:
            isPL (bool, optional): 是否指定对象. Defaults to None.
        """
        isPL = self.isPL if isPL is None else isPL
        if not isPL:
            if self.AIRole == role.ST:
                self.STGetCard()
            elif self.AIRole == role.YS:
                self.YSGetCard()
            elif self.AIRole == role.RY:
                self.RYGetCard()
        else:
            self.PLGetCard()

    def STGetCard(self):
        """ST摸基础牌
        """
        if self.isBasic:
            # 首张牌加强
            if self.roundTimes == 0 and len(self.cardList) == 0:
                card = self.GetCardBetween(5, 6)
            # 特殊抽牌概率
            elif random.random() < 0.5:
                card = self.GetCardBetween(5, 6) if random.random() < 0.6 else self.GetCardBetween(1, 2)
            else:
                card = self.GetCardBetween()
        else:
            card = self.GetCardBetween()
        if len(self.cardList) + len(self.funCardList) < maxCardNum:
            self.InsertCard(card, False)

    def YSGetCard(self):
        """YS摸基础牌
        """
        if self.isBasic:
            # 首张牌加强
            if self.roundTimes == 0 and len(self.cardList) == 0:
                card = self.GetCardBetween(4, 6)
            else:
                card = self.GetCardBetween()
        else:
            card = self.GetCardBetween()
        if len(self.cardList) + len(self.funCardList) < maxCardNum:
            self.InsertCard(card, False)

    def RYGetCard(self):
        """RY摸基础牌
        """
        if self.isBasic:
            # 初始摸牌削弱
            if self.roundTimes == 0:
                if len(self.cardList) == 0:
                    card = self.GetCardBetween(1, 3)
                elif len(self.cardList) == 1:
                    card = self.GetCardBetween(1, 4)
                else:
                    card = self.GetCardBetween(1, 5)
            elif self.roundTimes > 3:
                card = self.getCardFrom(3, 3)
            else:
                card = self.GetCardBetween()
        else:
            card = self.GetCardBetween()
        if len(self.cardList) + len(self.funCardList) < maxCardNum:
            self.InsertCard(card, False)

    def PLGetCard(self):
        """玩家摸基础牌
        """
        if self.isBasic:
            if len(self.PLCardList) >= 2:
                card = self.GetCardBetween(4, 6) if random.random() < 0.7 else self.GetCardBetween()
            else:
                card = self.GetCardBetween()
        else:
            card = self.GetCardBetween()
        if len(self.PLCardList) + len(self.PLFunCardList) < maxCardNum:
            self.InsertCard(card, True)

    def GetCardBetween(self, min=1, max=6):
        """从范围内抽牌，最多尝试10次，10次都失败则返回最后一次抽到的牌 \n
        Args:
            min (int, optional): 最小数字. Defaults to 1.
            max (int, optional): 最大数字. Defaults to 6.
        Returns:
            卡牌
        """
        tryTimes = 10
        while (tryTimes > 0):
            rnd = random.randint(0, len(self.cardLib) - 1)
            card = self.cardLib[rnd]
            tryTimes -= 1
            if int(card[1]) >= min and int(card[1]) <= max:
                break
        return card

    def getCardFrom(self, minTimes: int, maxTimes: int):
        """抽N张牌返回最大基础牌

        Args:
            minTimes (int): 最小抽牌数
            maxTimes (int): 最大抽牌数

        Returns:
            str: 卡牌
        """
        times = random.randint(minTimes, maxTimes)
        cardNum = 0
        while (times > 0):
            times -= 1
            rnd = random.randint(0, len(self.cardLib) - 1)
            if int(self.cardLib[rnd][1]) > cardNum:
                card = self.cardLib[rnd]
                cardNum = int(card[1])
        return card

    def InsertCard(self, card, isPL=None, isFromLib=True):
        """插入抽到的卡到手牌

        Args:
            card (str): 卡牌
            isPL (bool, optional): 是否指定对象. Defaults to None.
            isFromLib (bool, optional): 是否从牌堆抽牌. Defaults to True.
        """
        isPL = self.isPL if isPL is None else isPL
        cardList = self.cardList if not isPL else self.PLCardList
        for i in range(len(cardList)):
            if int(card[1]) > int(cardList[i][1]):
                cardList.insert(i, card)
                break
        else:
            cardList.append(card)
        if isFromLib:
            self.cardLib.remove(card)
        self.Log("获得数字牌:" + card, isPL)
# endregion

# region 功能牌决策相关

    def DoFun(self):
        """开始功能牌决策逻辑
        """

        if self.isBasic:
            return
        if self.isPL and self.PLFrozen:
            self.Log('冰冻状态，跳过功能牌出牌')
            self.PLFrozen = False
        elif not self.isPL and self.AIFrozen:
            self.Log('冰冻状态，跳过功能牌出牌')
            self.AIFrozen = False
        else:
            self.GetCardsInGrid()
            self.GetEmptyGridNum()
            funCardList = self.PLFunCardList if self.isPL else self.funCardList
            while True:
                card, value, index = self.GetNextFunCard(funCardList)
                if card is None:
                    break
                else:
                    self.allAction += 1
                    if self.UseFunCard(card, value):
                        self.EnforceFunCard(card, index)

    def GetNextFunCard(self, funCardList):
        """获得下一张执行的功能牌，返回None时无可用的功能牌

        Args:
            funCardList (_type_): 功能牌列表

        Returns:
            card: 功能牌
            value: 功能牌价值
            index: 作用位置
        """
        if len(funCardList) == 0:
            return None, None, None
        else:
            # 默认顺序
            for card in funCardOrderList:
                if card in funCardList:
                    isUse, value, index = self.CheckFunCard(card)
                    if isUse:
                        return card, value, index
            else:
                return None, None, None

    def CheckFunCard(self, card: funCard):
        """功能牌是否出牌判定

        Args:
            card (funCard): 功能牌
        returns:
            isUse,value,index: 是否出牌，出牌价值，出牌位置
        """
        if card == funCard.跳过:
            return self.CheckFunCardSkip()
        if card == funCard.贪心:
            return self.CheckFunCardGreed()
        if card == funCard.冰冻:
            return self.CheckFunCardFreeze()
        if card == funCard.毁天灭地:
            return self.CheckFunCardDropNumCards()
        if card == funCard.窥视:
            return self.CheckFunCardPeep()
        if card == funCard.发财:
            return self.CheckFunCardRich()
        if card == funCard.赌狗喵:
            return self.CheckFunCardGamble()
        if card == funCard.暗言术:
            return self.CheckFunCardShadow()
        if card == funCard.拆迁喵:
            return self.CheckFunCardDemolition()
        if card == funCard.复制:
            return self.CheckFunCardCopy()
        if card == funCard.兴奋剂:
            return self.CheckFunCardHigh()

    def UseFunCard(self, card, value=None, isPL=None, extraStr=''):
        """使用指定功能牌

        Args:
            card (funCard): 功能牌
            cardPower (float, optional): 功能牌强度. Defaults to None.
            isPL (bool, optional): 是否指定对象. Defaults to None.
            extraStr (str, optional): 附带Log输出. Defaults to ''.

        Returns:
            bool: 是否使用成功（经过否决判定）
        """
        isPL = self.isPL if isPL is None else isPL
        self.Log('使用功能牌：' + card + extraStr, isPL)
        self.usedFunCardLib.append(card)
        if not isPL:
            self.funCardList.remove(card)
        else:
            self.PLFunCardList.remove(card)
        if card != funCard.否决:
            return self.Veto(card, value)

    def EnforceFunCard(self, card: funCard, index: int):
        """功能牌效果生效

        Args:
            card (funCard): 功能牌
        """
        if card == funCard.跳过:
            self.EnforceFunCardSkip()
        if card == funCard.贪心:
            self.EnforceFunCardGreed()
        if card == funCard.冰冻:
            self.EnforceFunCardFreeze()
        if card == funCard.毁天灭地:
            self.EnforceFunCardDropNumCards()
        if card == funCard.窥视:
            self.EnforceFunCardPeep()
        if card == funCard.发财:
            self.EnforceFunCardRich()
        if card == funCard.赌狗喵:
            self.EnforceFunCardGamble(index)
        if card == funCard.暗言术:
            self.EnforceFunCardShadow(index)
        if card == funCard.拆迁喵:
            self.EnforceFunCardDemolition(index)
        if card == funCard.复制:
            self.EnforceFunCardCopy()
        if card == funCard.兴奋剂:
            self.EnforceFunCardHigh()

    def GetCardsInGrid(self):
        """获得场上的牌状况，数据保存至

        self.myCardInGrid

        self.enemyCardInGrid
        """
        self.myCardInGrid = []
        self.enemyCardInGrid = []
        for i in range(len(self.gridData)):
            if self.gridData[i] is not None:
                cardInfo = self.GetCardInfo(i)
                if 'AI' in self.gridData[i]:
                    if not self.isPL:
                        self.myCardInGrid.append(cardInfo)
                    else:
                        self.enemyCardInGrid.append(cardInfo)
                else:
                    if self.isPL:
                        self.myCardInGrid.append(cardInfo)
                    else:
                        self.enemyCardInGrid.append(cardInfo)

    def GetCardInfo(self, i):
        """获得指定格子位置数字牌数据 (score,cardNum,cardColor,i)

        Args:
            i (int): 格子的序号
        Returns:
            tuple: (卡牌得分, 牌面数字, 牌面颜色, 位置)
        """
        cardIndex = self.gridData[i].find('-')
        scoreIndex = self.gridData[i].rfind('-')
        scoreEndIndex = self.gridData[i].rfind('分')
        card = self.gridData[i][cardIndex + 1:cardIndex + 3]
        score = int(self.gridData[i][scoreIndex + 1:scoreEndIndex])
        return (score, int(card[1]), card[0], i)

    def GetCardMaxScore(self, card: str):
        """获取特定卡最大可得分

        Args:
            card (str): 指定数字牌

        Returns:
            int: 最大可得分
        """
        maxScore = 0
        for i in range(len(self.gridData)):
            if self.gridData[i] is None:
                score = 2 * int(card[1]) if self.gridTypes == card[0] else int(card[1])
                if score > maxScore:
                    maxScore = score
        return maxScore

    def GetUsedFunCardValue(self):
        """获取弃牌堆复制价值

        Returns:
            float: 弃牌堆复制价值
        """
        usedCardNum = len(self.usedFunCardLib)
        startNum = 0 if usedCardNum > 5 else usedCardNum - 5
        cardValue = 0
        for i in range(startNum, usedCardNum):
            cardValue += self.funCardPowerMap[self.usedFunCardLib[i]]
        return cardValue / (usedCardNum - startNum)

    def IsStock(self):
        """是否应该继续屯牌

        Returns:
            bool: 是否应该继续屯牌
        """
        if self.emptyNum < 4:
            return False
        elif len(self.cardList) + len(self.funCardList) > maxCardNum - 3:
            return False
        else:
            return True

# region 功能牌效果

    def Veto(self, card, value=None):
        """功能牌：否决 判定
        """
        # 否决，返回false代表被否决
        value = self.funCardPowerMap[card] if value is None else value
        vetoTag = True
        if not self.isPL:
            myFunList = self.funCardList
            enemyFunList = self.PLFunCardList
            enemyCardList = self.PLCardList
        else:
            myFunList = self.PLFunCardList
            enemyFunList = self.funCardList
            enemyCardList = self.cardList
        nextRound = 1
        # 玩家出功能牌：【AI屯牌策略：价值高时出否决】【AI随性策略：价值不为零出否决】【AI普通策略：价值稍高出否决】
        # Ai出功能牌：价值稍高时玩家出否决
        if self.isPL and (self.AIStatus == AIStatus.屯牌 and value > funCardExceptPower or self.AIStatus == AIStatus.随性 and value > 0 or self.AIStatus == AIStatus.普通 and value > exceptScore) \
           or (not self.isPL and value > exceptScore):
            while nextRound > 0:
                nextRound = 0
                if funCard.否决 in enemyFunList and vetoTag:
                    nextRound = 1
                    vetoTag = False
                    self.UseFunCard(funCard.否决, None, not self.isPL, extraStr='，作用目标：' + card)
                if funCard.否决 in myFunList and not vetoTag:
                    if card != funCard.窥视 or card == funCard.窥视 and len(enemyFunList) + len(enemyCardList) > 0:
                        nextRound = 1
                        vetoTag = True
                        self.UseFunCard(funCard.否决, extraStr='，作用目标：否决-' + card)
        return vetoTag

    def CheckFunCardFreeze(self):
        """检测是否出功能牌：冰冻

        Returns:
            isUse,value,index: 是否出牌,出牌价值,出牌位置
        """
        value = self.funCardPowerMap[funCard.冰冻]
        enemyFrozen = self.AIFrozen if self.isPL else self.PLFrozen
        enemyFunCardList = self.funCardList if self.isPL else self.PLFunCardList
        # 对方已被冰冻，或对方无功能牌可出时无效
        if enemyFrozen or len(enemyFunCardList) + len(self.funCardLib) == 0:
            return False, 0, None
        # AI放水时不出
        elif not self.isPL and self.AIStatus == AIStatus.放水:
            return False, value, None
        else:
            return True, value, None

    def EnforceFunCardFreeze(self):
        """功能牌：冰冻 生效

        冰冻：跳过功能牌出牌阶段，不影响获得牌
        """
        if self.isPL:
            self.AIFrozen = True
        else:
            self.PLFrozen = True
        self.Log(funCard.冰冻 + '生效，对方跳过下次功能牌出牌')

    def CheckFunCardSkip(self):
        """检测是否出功能牌：跳过

        Returns:
            isUse,value,index: 是否出牌,出牌价值,出牌位置
        """
        value = self.funCardPowerMap[funCard.跳过]
        enemySkip = self.AISkip if self.isPL else self.PLSkip
        # 对方已被跳过时不出
        if enemySkip:
            return False, 0, None
        # AI放水时不出
        elif not self.isPL and self.AIStatus == AIStatus.放水:
            return False, value, None
        else:
            return True, value, None

    def EnforceFunCardSkip(self):
        """功能牌：跳过 生效

        跳过：跳过数字牌阶段
        """
        self.Log(funCard.跳过 + '生效，对方跳过下次数字牌阶段')
        if self.isPL:
            self.AISkip = True
        else:
            self.PLSkip = True

    def CheckFunCardDropNumCards(self):
        """检测是否出功能牌：毁天灭地

        Returns:
            isUse,value,index: 是否出牌,出牌价值,出牌位置
        """
        if self.isPL:
            cardList = self.PLCardList
            enemyCardList = self.cardList
            score = self.PLScore
            enemyScore = self.AIScore
            isHigh = self.PLHigh
            funCardList = self.PLFunCardList
        else:
            cardList = self.cardList
            enemyCardList = self.PLCardList
            score = self.AIScore
            enemyScore = self.PLScore
            isHigh = self.AIHigh
            funCardList = self.funCardList

        self.GetMaxScoreCard()
        value = (averageScore - self.maxScore) + (len(enemyCardList) - len(cardList)) * averageScore
        # 对方无数字牌时不出
        if len(enemyCardList) == 0:
            return False, 0, None
        # 出数字牌即可获胜时不出
        elif len(cardList) > 0 and self.emptyNum == 1 and score + self.maxScore >= enemyScore:
            return False, value, None
        # AI放水时不出
        elif not self.isPL and self.AIStatus == AIStatus.放水:
            return False, value, None
        # 可得分较高，或有兴奋状态时不出
        elif self.maxScore > exceptScore or isHigh:
            return False, value, None
        # 手牌比对方多(都是小牌)，且没有发财时不出
        elif len(cardList) > len(enemyCardList) and funCard.发财 not in funCardList:
            return False, value, None
        else:
            return True, value, None

    def EnforceFunCardDropNumCards(self):
        """功能牌：毁天灭地 生效

        毁天灭地：双方丢弃所有数字手牌
        """
        self.Log(funCard.毁天灭地 + '生效，双方丢弃所有数字牌')
        AIDrop = ''
        PLDrop = ''
        for c in self.cardList:
            AIDrop = AIDrop + '-' + c
        for c in self.PLCardList:
            PLDrop = PLDrop + '-' + c
        if AIDrop == '':
            self.Log('AI无数字手牌')
        else:
            self.Log('AI丢弃牌' + AIDrop)
        if PLDrop == '':
            self.Log('玩家无数字手牌')
        else:
            self.Log('玩家丢弃牌' + PLDrop)
        self.cardList = []
        self.PLCardList = []

    def CheckFunCardGreed(self):
        """检测是否要出功能牌：贪心

        Returns:
            isUse,value,index: 是否出牌,出牌价值,出牌位置
        """
        cardList = self.PLCardList if self.isPL else self.cardList
        funCardList = self.PLFunCardList if self.isPL else self.funCardList
        value = funCardExceptPower * 2 - exceptScore
        # 牌堆无牌不使用
        if len(self.funCardLib) == 0:
            return False, 0, None
        # 游戏即将结束时使用
        elif self.emptyNum < 3:
            return True, value, None
        # 手牌满了，不使用
        elif len(cardList) + len(funCardList) == maxCardNum:
            return False, 0, None
        # 屯牌：没有可丢弃的小牌不使用
        elif not self.isPL and self.AIStatus == AIStatus.屯牌:
            hasMinCard = False
            for card in cardList:
                if self.GetCardMaxScore(card) < exceptScore:
                    hasMinCard = True
                    break
            for card in funCardList:
                if card == funCard.冰冻:
                    hasMinCard = True
                    break
            if hasMinCard or not self.IsStock():
                return True, value, None
            else:
                return False, 0, None
        else:
            return True, value, None

    def EnforceFunCardGreed(self):
        """功能牌：贪心 生效

        贪心：获得两张功能牌，之后从手牌中选择一张丢弃
        """
        if len(self.funCardLib) == 1:
            self.GetFunCard()
        else:
            self.GetFunCard()
            self.GetFunCard()
        baseList = self.cardList if not self.isPL else self.PLCardList
        funList = self.funCardList if not self.isPL else self.PLFunCardList
        cardList = []
        minValue = 100
        isHigh = self.PLHigh if self.isPL else self.AIHigh
        if not isHigh or isHigh and (len(baseList) > 2 or funCard.发财 in funList):
            for card in baseList:
                cardList.append(card)
        for card in funList:
            cardList.append(card)
        for card in cardList:
            if card in self.funCardPowerMap:
                cardValue = self.funCardPowerMap[card]
            else:
                cardValue = self.GetCardMaxScore(card)
            if cardValue < minValue:
                minValue = cardValue
                minCard = card
        if minCard in self.funCardPowerMap:
            funList.remove(minCard)
            self.usedFunCardLib.append(minCard)
        else:
            baseList.remove(minCard)
        self.Log('弃掉手牌：' + minCard)

    def CheckFunCardPeep(self):
        """检测是否出功能牌：窥视

        Returns:
            isUse,value,index: 是否出牌,出牌价值,出牌位置
        """
        enemyCardList = self.PLCardList if not self.isPL else self.cardList
        enemyFunCardList = self.PLFunCardList if not self.isPL else self.funCardList
        value = funCardExceptPower
        # 对方没手牌时不使用
        if len(enemyCardList) + len(enemyFunCardList) == 0:
            return False, 0, None
        else:
            return True, value, None

    def EnforceFunCardPeep(self):
        """功能牌：窥视 生效

        窥视：查看对方随机3张手牌，并丢弃其中一张
        """
        baseList = self.cardList if self.isPL else self.PLCardList
        funList = self.funCardList if self.isPL else self.PLFunCardList
        cardList = []
        for card in baseList:
            cardList.append(card)
        for card in funList:
            cardList.append(card)
        maxCardNum = 3 if len(cardList) > 3 else len(cardList)
        seeList = []
        while len(seeList) < maxCardNum:
            rnd = random.randint(0, len(cardList) - 1)
            if rnd not in seeList:
                seeList.append(rnd)
        maxValue = 0
        maxCard = ''
        cardStr = ''
        for i in seeList:
            cardStr = cardStr + '-' + cardList[i]
            if cardList[i] in self.funCardPowerMap:
                cardValue = self.funCardPowerMap[cardList[i]]
            else:
                cardValue = self.GetCardMaxScore(cardList[i])
            if cardValue > maxValue:
                maxValue = cardValue
                maxCard = cardList[i]
        if maxCard in self.funCardPowerMap:
            funList.remove(maxCard)
            self.usedFunCardLib.append(maxCard)
        else:
            baseList.remove(maxCard)
        self.Log('查看对方手牌' + cardStr)
        self.Log('弃掉对方手牌：' + maxCard)

    def CheckFunCardCopy(self):
        """检测是否出功能牌：复制

        Returns:
            isUse,value,index: 是否出牌,出牌价值,出牌位置
        """
        usedCardNum = len(self.usedFunCardLib)
        startNum = 0 if usedCardNum < 6 else usedCardNum - 5
        if usedCardNum == 0:
            return False, 0, None
        else:
            value = 0
            copyTag = True
            demolitionTag = False
            for i in range(startNum, len(self.usedFunCardLib)):
                card = self.usedFunCardLib[i]
                value += self.funCardPowerMap[card]
                if card != funCard.复制:
                    copyTag = False
                if card == funCard.拆迁喵:
                    demolitionTag = True
            value = value / (usedCardNum - startNum)
            if copyTag:  # 牌堆里全是复制，不出牌
                return False, 0, None
            elif not self.isPL and self.AIStatus == AIStatus.屯牌:
                if value > funCardExceptPower or demolitionTag or not self.IsStock():
                    return True, value, None
                else:
                    return False, 0, None
            else:
                if value > exceptScore:
                    return True, value, None
                else:
                    return False, 0, None

    def EnforceFunCardCopy(self):
        """功能牌：复制 生效

        复制：从双方打出的最近5张功能牌中随机抽走一张加入己方手牌，弃牌堆内至少2张牌才能使用
        """
        usedCardNum = len(self.usedFunCardLib)
        startNum = 0 if usedCardNum < 6 else usedCardNum - 5
        while (True):
            rnd = random.randint(startNum, usedCardNum - 1)
            card = self.usedFunCardLib[rnd]
            if card != funCard.复制:
                break

        self.InsertFunCard(card, isFromLib=False)
        for i in range(usedCardNum - 1, -1, -1):
            if self.usedFunCardLib[i] == card:
                del (self.usedFunCardLib[i])
                break

    def CheckFunCardHigh(self):
        """检测是否出功能牌：兴奋剂

        Returns:
            isUse,value,index: 是否出牌,出牌价值,出牌位置
        """
        if self.isPL:
            cardList = self.PLCardList
            enemyCardList = self.cardList
            funCardList = self.PLFunCardList
            isHigh = self.PLHigh
            isSkip = self.PLSkip
            isEnemySkip = self.AISkip
            score = self.PLScore
            enemyScore = self.AIScore
        else:
            cardList = self.cardList
            enemyCardList = self.PLCardList
            funCardList = self.funCardList
            isHigh = self.AIHigh
            isSkip = self.AISkip
            isEnemySkip = self.PLSkip
            score = self.AIScore
            enemyScore = self.PLScore

        # 被跳过，或已有兴奋，或数字牌过少时，或空格过少时不使用
        if isSkip or isHigh or \
           len(cardList) < 2 and funCard.发财 not in funCardList or \
           self.emptyNum < 2 and funCard.拆迁喵 not in funCardList:
            return False, 0, None
        else:
            self.GetMaxScoreCard()
            tempGrid = []
            tempCardList = []
            for i in range(len(self.gridData)):
                if i != self.maxPutIndex:
                    tempGrid.append(self.gridData[i])
                else:
                    tempGrid.append('First Card')
            for card in cardList:
                tempCardList.append(card)
            tempCardList.remove(self.maxCard)
            value = self.GetMaxScoreCard(tempGrid, tempCardList)
            tempGrid.clear()
            tempCardList.clear()

            self.GetMaxScoreCard()
            # 出牌会输时不出
            if self.emptyNum == 2 and value + self.maxScore + score < enemyScore:
                return False, value, None
            # 出牌后，轮到对方出牌会输时不出
            elif self.emptyNum == 3 and value + self.maxScore + score <= enemyScore and len(
                    enemyCardList) > 0 and not isEnemySkip:
                return False, value, None
            # 屯牌策略：价值较高，或手牌过多，或空格过少时出牌
            elif not self.isPL and self.AIStatus == AIStatus.屯牌:
                if value > exceptScore or (value > 0 and not self.IsStock()):
                    return True, value, None
                else:
                    return False, value, None
            else:
                if value > 0:
                    return True, value, None
                else:
                    return False, value, None

    def EnforceFunCardHigh(self):
        """功能牌：兴奋剂 生效

        兴奋剂：本回合可打出两张数字牌
        """
        self.Log(funCard.兴奋剂 + '生效，可出2张数字牌')
        if self.isPL:
            self.PLHigh = True
        else:
            self.AIHigh = True

    def CheckFunCardRich(self):
        """检测是否出功能牌：发财

        Returns:
            isUse,value,index: 是否出牌,出牌价值,出牌位置
        """
        cardList = self.PLCardList if self.isPL else self.cardList
        funCardList = self.PLFunCardList if self.isPL else self.funCardList
        isSkip = self.PLSkip if self.isPL else self.AISkip
        value = averageScore * 2
        # 被跳过，或手牌已满时不使用
        if isSkip or len(cardList) + len(funCardList) == maxCardNum:
            return False, 0, None
        else:
            return True, value, None

    def EnforceFunCardRich(self):
        """功能牌：发财 生效

        发财：获得2张数字牌
        """
        self.Log(funCard.发财 + '生效，获得2张数字牌')
        self.GetCard()
        self.GetCard()

    def CheckFunCardGamble(self):
        """检测是否出功能牌：赌狗喵

        Returns:
            isUse,value,index: 是否出牌,出牌价值,出牌位置
        """
        value = 0
        index = -1
        # 放水时变化非0的最小牌
        if not self.isPL and self.AIStatus == AIStatus.放水:
            for cardInfo in self.enemyCardInGrid:  # cardInfo: (score,cardNum,cardColor,i)
                cardValue = cardInfo[0] * 0.75
                if value == 0 or cardValue > 0 and cardValue < value:
                    value = cardValue
                    index = cardInfo[3]
            if value > 0:
                return True, value, index
            else:
                return False, value, None
        else:
            for cardInfo in self.enemyCardInGrid:  # cardInfo: (score,cardNum,cardColor,i)
                cardValue = cardInfo[0] * 0.75
                if cardValue > value:
                    value = cardValue
                    index = cardInfo[3]
            if not self.isPL and self.AIStatus == AIStatus.屯牌:
                # 屯牌策略：价值较高，手牌过多，空格过少时出牌
                if value > funCardExceptPower or (value > 0 and not self.IsStock()):
                    return True, value, index
                else:
                    return False, value, None
            elif value > exceptScore:
                return True, value, index
            else:
                return False, value, None

    def EnforceFunCardGamble(self, index):
        """功能牌：赌狗喵 生效

        赌狗喵:指定一个格子，随机改变颜色（可与当前颜色相同）
        """
        cardInfo = self.GetCardInfo(index)  # (score,cardNum,cardColor,i)
        rnd = random.randint(1, 4)
        scoreChange = 0
        # 原来是普通格
        if cardInfo[0] == cardInfo[1]:
            if gridType[rnd] == cardInfo[2]:
                cardScore = cardInfo[1] * 2
                scoreChange = cardInfo[1]
            else:
                cardScore = 0
                scoreChange = -cardInfo[1]
        # 原来是翻倍格
        else:
            if gridType[rnd] == cardInfo[2]:
                cardScore = cardInfo[1] * 2
                if cardInfo[0] > 0:
                    scoreChange = 0
                else:
                    scoreChange = cardScore
            else:
                cardScore = 0
                if cardInfo[0] > 0:
                    scoreChange = -cardInfo[0]
                else:
                    scoreChange = 0

        cardStr = 'AI-' if self.isPL else '玩家-'
        cardStr = cardStr + cardInfo[2] + str(cardInfo[1]) + '-' + str(cardScore) + '分'
        self.gridTypes[index] = gridType[rnd]
        self.gridData[index] = cardStr
        if self.isPL:
            self.AIScore += scoreChange
        else:
            self.PLScore += scoreChange
        self.GetCardsInGrid()
        if scoreChange > 0:
            self.Log('位置' + str(index) + '赌狗喵生效，颜色变为' + gridType[rnd] + '色，对方积分增加' + str(scoreChange))
        elif scoreChange < 0:
            self.Log('位置' + str(index) + '赌狗喵生效，颜色变为' + gridType[rnd] + '色，对方积分减少' + str(-scoreChange))
        else:
            self.Log('位置' + str(index) + '赌狗喵生效，颜色变为' + gridType[rnd] + '色，对方积分未发生改变')

    def CheckFunCardShadow(self):
        """检测是否出功能牌：暗言术

        Returns:
            isUse,value,index: 是否出牌,出牌价值,出牌位置
        """
        value = 0
        index = -1
        # 放水时变化非1的最小牌
        if not self.isPL and self.AIStatus == AIStatus.放水:
            for cardInfo in self.enemyCardInGrid:  # cardInfo: (score,cardNum,cardColor,i)
                cardValue = cardInfo[0] - 1 if cardInfo[0] == cardInfo[1] else cardInfo[0] - 2
                cardValue = 0 if cardValue < 0 else cardValue
                if value == 0 or cardValue > 0 and cardValue < value:
                    value = cardValue
                    index = cardInfo[3]
            if value > 0:
                return True, value, index
            else:
                return False, value, None
        else:
            for cardInfo in self.enemyCardInGrid:  # cardInfo: (score,cardNum,cardColor,i)
                cardValue = cardInfo[0] - 1 if cardInfo[0] == cardInfo[1] else cardInfo[0] - 2
                cardValue = 0 if cardValue < 0 else cardValue
                if cardValue > value:
                    value = cardValue
                    index = cardInfo[3]
            if not self.isPL and self.AIStatus == AIStatus.屯牌:
                # 屯牌策略：价值较高，或手牌过多，或空格过少时出牌
                if value > funCardExceptPower or (not self.IsStock() and value > 0):
                    return True, value, index
                else:
                    return False, value, None
            elif value > exceptScore:
                return True, value, index
            else:
                return False, value, None

    def EnforceFunCardShadow(self, index):
        """功能牌：暗言术 生效

        暗言术•滚：使对方的一个已放置喵变为1分（当前格子效果维持不变）
        """
        cardInfo = self.GetCardInfo(index)
        cardStr = 'AI-' if self.isPL else '玩家-'
        if self.gridTypes[index] is None:
            scoreChange = cardInfo[0] - 1
            cardStr = cardStr + cardInfo[2] + '1-1分'
        elif self.gridTypes[index] == cardInfo[2]:
            scoreChange = cardInfo[0] - 2
            cardStr = cardStr + cardInfo[2] + '1-2分'
        else:
            scoreChange = 0
            cardStr = cardStr + cardInfo[2] + '1-0分'
        self.gridData[index] = cardStr
        if self.isPL:
            self.AIScore -= scoreChange
        else:
            self.PLScore -= scoreChange
        self.GetCardsInGrid()
        self.Log(funCard.暗言术 + "作用位置" + str(index) + "，对方得分减少" + str(scoreChange))

    def CheckFunCardDemolition(self):
        """检测是否出功能牌：拆迁

        Returns:
            isUse,value,index: 是否出牌，出牌价值，出牌位置
        """
        value = 0
        index = -1
        # 放水时拆小牌
        if not self.isPL and self.AIStatus == AIStatus.放水:
            for cardInfo in self.enemyCardInGrid:  # cardInfo: (score,cardNum,cardColor,i)
                if value == 0 or cardInfo[0] > 0 and cardInfo[0] < value:
                    value = cardInfo[0]
                    index = cardInfo[3]
            if value > 0:
                return True, value, index
            else:
                return False, value, None
        else:
            # 拆格子效果 = 目标牌得分 + 拆之后可得分 - 拆之前可得分 + 摸牌均分
            beforeScore = self.GetMaxScoreCard()  # 拆迁前最大可得分
            for cardInfo in self.enemyCardInGrid:
                cardStr = self.gridData[cardInfo[3]]  # 暂存拆迁牌数据
                self.gridData[cardInfo[3]] = None
                afterScore = self.GetMaxScoreCard()
                self.gridData[cardInfo[3]] = cardStr
                removeValue = cardInfo[0] + afterScore - beforeScore + averageScore
                if removeValue > value:
                    value = removeValue
                    index = cardInfo[3]
            if not self.isPL and self.AIStatus == AIStatus.屯牌:
                # 屯牌策略：价值较高，或手牌过多，或空格过少时出牌
                if value > funCardExceptPower + averageScore or (value > 0 and not self.IsStock()):
                    return True, value, index
                else:
                    return False, value, None
            elif value > exceptScore + averageScore:
                return True, value, index
            else:
                return False, value, None

    def EnforceFunCardDemolition(self, index):
        """功能牌：拆迁 生效

        拆迁：拆掉并丢弃一个已放置的喵
        """
        cardInfo = self.GetCardInfo(index)
        if self.isPL:
            self.AIScore -= cardInfo[0]
        else:
            self.PLScore -= cardInfo[0]

        self.gridData[index] = None
        self.GetCardsInGrid()
        self.GetEmptyGridNum()
        self.Log(funCard.拆迁喵 + "作用位置" + str(index) + "，对方得分减少" + str(cardInfo[0]))

# endregion

# region 抽功能牌

    def GetFunCard(self, isPL=None):
        """抽功能牌

        Args:
            isPL (bool, optional): 是否指定对象. Defaults to None.
        """
        isPL = self.isPL if isPL is None else isPL
        if not isPL:
            if self.AIRole == role.ST:
                self.STGetFunCard()
            elif self.AIRole == role.YS:
                self.YSGetFunCard()
            elif self.AIRole == role.RY:
                self.RYGetFunCard()
        else:
            self.PLGetFunCard()

    def STGetFunCard(self):
        """ST模功能牌
        """
        card = self.GetFunCardBetween()
        if len(self.cardList) + len(self.funCardList) == maxCardNum:
            self.usedFunCardLib.append(card)
        else:
            self.InsertFunCard(card, False)

    def YSGetFunCard(self):
        """YS模功能牌
        """
        card = self.GetFunCardBetween()
        if len(self.cardList) + len(self.funCardList) == maxCardNum:
            self.usedFunCardLib.append(card)
        else:
            self.InsertFunCard(card, False)

    def RYGetFunCard(self):
        """RY模功能牌
        """
        card = self.GetFunCardBetween()
        if len(self.cardList) + len(self.funCardList) == maxCardNum:
            self.usedFunCardLib.append(card)
        else:
            self.InsertFunCard(card, False)

    def PLGetFunCard(self):
        """玩家摸功能牌
        """
        card = self.GetFunCardBetween()
        if len(self.PLCardList) + len(self.PLFunCardList) == maxCardNum:
            self.usedFunCardLib.append(card)
        else:
            self.InsertFunCard(card, True)

    def GetFunCardBetween(self, min: int = 0, max: int = 99):
        """从强度范围内抽功能牌，最多尝试10次，10次都失败则返回最后一次抽到的牌 \n
        Args:
            min (int, optional): 最小强度. Defaults to 1.
            max (int, optional): 最大强度. Defaults to 99.
        Returns:
            功能牌(str)
        """
        tryTimes = 10
        while (tryTimes > 0):
            rnd = random.randint(0, len(self.funCardLib) - 1)
            card = self.funCardLib[rnd]
            tryTimes -= 1
            if self.funCardPowerMap[card] >= min and self.funCardPowerMap[card] <= max:
                break
        return card

    def GetFunCardFrom(self, minTimes: int, maxTimes: int):
        """抽N张牌返回最大强度功能牌

        Args:
            minTimes (int): 最小抽牌数
            maxTimes (int): 最大抽牌数

        Returns:
            功能牌(str)
        """
        times = random.randint(minTimes, maxTimes)
        cardNum = 0
        while (times > 0):
            times -= 1
            rnd = random.randint(0, len(self.funCardLib) - 1)
            if self.funCardPowerMap[self.funCardLib[rnd]] > cardNum:
                card = self.funCardLib[rnd]
                cardNum = self.funCardPowerMap[self.funCardLib[rnd]]
        return card

    def InsertFunCard(self, card: str, isPL: bool = None, isFromLib: bool = True):
        """将功能牌放入手牌末尾

        Args:
            card (str): 功能牌
            isPL (bool, optional): 是否指定对象. Defaults to None.
            isFromLib (bool, optional): 是否从牌库获得手牌. Defaults to True.
        """
        isPL = self.isPL if isPL is None else isPL
        cardList = self.funCardList if not isPL else self.PLFunCardList
        if len(cardList) < maxCardNum:
            cardList.append(card)
        if isFromLib:
            self.funCardLib.remove(card)
        self.Log("获得功能牌：" + card, isPL)


# endregion
