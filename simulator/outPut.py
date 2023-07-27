from lovePoint import E_LoveType
from player import Player
from xlDeal import XlDeal


class Output:
    def __init__(self, _tryTimes, _playerNum=6, _maxDay=7):
        self.tryTimes = _tryTimes
        self.playNum = _playerNum
        self.maxDay = _maxDay
        self.manColNum = 0
        # 定义输出数据
        self.detailXl = XlDeal("资源投放统计.xlsm", "数据总览")
        self.StatisticsXl = XlDeal("资源投放统计.xlsm", "总览统计")
        self.manDetailXlMap: dict[int, XlDeal] = {}
        self.manDetailXlMap[1] = XlDeal("资源投放统计.xlsm", "ST牵绊度")
        self.manDetailXlMap[2] = XlDeal("资源投放统计.xlsm", "ys牵绊度")
        self.manDetailXlMap[5] = XlDeal("资源投放统计.xlsm", "ry牵绊度")
        self.manStatisticsXlMap: dict[int, XlDeal] = {}
        self.manStatisticsXlMap[1] = XlDeal("资源投放统计.xlsm", "ST统计")
        self.manStatisticsXlMap[2] = XlDeal("资源投放统计.xlsm", "YS统计")
        self.manStatisticsXlMap[5] = XlDeal("资源投放统计.xlsm", "RY统计")

        self.detailList = []  # 存放【数据总览】表数据
        self.detailKeys = []  # 存放【数据总览】表列名
        self.manDetailMap: dict[int, any] = {}  # 存放【man牵绊度】表数据
        self.manDetailKeys = []  # 存放【man牵绊度】表列名
        self.statisticsList = []  # 存放【总览统计】表数据
        self.manStatisticsMap: dict[int, any] = {}  # 存放【man统计】表数据
        self.InitData()

    def InitData(self):
        self.enumlist = []
        for i in E_LoveType:
            if i.label not in self.enumlist:
                self.enumlist.append(i.label)
        self.manColNum = len(self.enumlist) + 2
        self.detailList.append(self.detailKeys)
        self.statisticsList = [[0] * (self.playNum * 3 + 1) for _ in range(self.maxDay + 2)]  # 存放【总览统计】表数据
        for m in range(1, 6):
            if m != 3 and m != 4:
                self.manDetailMap[m] = []
                self.manDetailMap[m].append(self.manDetailKeys)
                self.manStatisticsMap[m] = [[0] * (self.manColNum * 2) for _ in range(self.playNum * self.maxDay + 1)]

        self.statisticsList[0][1] = "ST牵绊度"
        self.statisticsList[0][7] = "YS牵绊度"
        self.statisticsList[0][13] = "RY牵绊度"
        self.statisticsList[1][0] = "日期"
        for m in self.manStatisticsMap:
            for i in range(len(self.enumlist) + 2):
                if i == 0:
                    self.manStatisticsMap[m][0][i] = "日期"
                    self.manStatisticsMap[m][0][i + self.manColNum] = "日期"
                elif i == 1:
                    self.manStatisticsMap[m][0][i] = "玩家"
                    self.manStatisticsMap[m][0][i + self.manColNum] = "玩家"
                else:
                    self.manStatisticsMap[m][0][i] = self.enumlist[i - 2]
                    self.manStatisticsMap[m][0][i + self.manColNum] = self.enumlist[i - 2]

    def UpdateTableData(self):
        self.GetAverageStatisticsData()
        self.detailXl.UpdateTableData(self.detailList, isClear=True)
        self.StatisticsXl.UpdateTableData(self.statisticsList, isClear=True)
        for m in self.manDetailXlMap:
            for row in self.manDetailMap[m]:
                while len(row) < len(self.manDetailKeys):
                    row.append(None)
            self.manDetailXlMap[m].UpdateTableData(self.manDetailMap[m], isClear=True)
        for m in self.manStatisticsXlMap:
            self.manStatisticsXlMap[m].UpdateTableData(self.manStatisticsMap[m], isClear=True)

    def GetPlayerData(self, day, player: Player, playerOrder):
        self.GetDetailPlayerStatus(day, player)
        self.GetDetailManLoveExp(day, player)
        self.GetStatisticsData(day, player, playerOrder)

    def GetDetailPlayerStatus(self, day, player: Player):
        # 数据总览数据
        detailLineMap = {}
        detailLineMap["日期"] = day
        detailLineMap["玩家"] = player.name
        detailLineMap["玩家等级"] = player.playerLev
        detailLineMap["ST牵绊度"] = player.loveLevList[1].loveLev
        detailLineMap["YS牵绊度"] = player.loveLevList[2].loveLev
        detailLineMap["RY牵绊度"] = player.loveLevList[5].loveLev
        detailLineMap["抽卡次数"] = player.totalCardGachaTimes
        detailLineMap["剩余钻石"] = player.itemMap["2=2"]
        detailLineMap["SSR思念数量"] = len(player.rareCardList[4])
        detailLineMap["SR思念数量"] = len(player.rareCardList[3])
        detailLineMap["R思念数量"] = len(player.rareCardList[2])
        detailLineMap["SSR思念进阶次数"] = player.GetTotalCardPhase(4)
        detailLineMap["SR思念进阶次数"] = player.GetTotalCardPhase(3) - player.GetTotalCardPhase(4)
        detailLineMap["R思念进阶次数"] = player.GetTotalCardPhase(2) - player.GetTotalCardPhase(3)
        rare = ""
        level = ""
        phase = ""
        man = ""
        for i in range(player.developNumPerTag):
            for tag in range(1, 7):
                if len(player.dealTagCardList[tag]) > i:
                    card = player.dealTagCardList[tag][i]
                    if rare == "":
                        rare = str(player.cardMap[card].rare)
                        level = str(player.cardMap[card].cardLev)
                        phase = str(player.cardMap[card].cardPhase)
                        man = str(player.cardMap[card].man)
                    else:
                        rare = rare + "|" + str(player.cardMap[card].rare)
                        level = level + "|" + str(player.cardMap[card].cardLev)
                        phase = phase + "|" + str(player.cardMap[card].cardPhase)
                        man = man + "|" + str(player.cardMap[card].man)
        detailLineMap["培养男主"] = player.developMan
        detailLineMap["最低培养品质"] = player.developRare
        detailLineMap["最高思念等级"] = player.developLevs[0]
        detailLineMap["最强思念卡男主"] = man
        detailLineMap["最强思念卡品质"] = rare
        detailLineMap["最强思念卡等级"] = level
        detailLineMap["最强思念卡进阶"] = phase
        if len(self.detailKeys) == 0:
            for key in detailLineMap:
                self.detailKeys.append(key)
        rowList = []
        for key in self.detailKeys:
            rowList.append(detailLineMap[key])
        self.detailList.append(rowList)
        del detailLineMap

    def GetDetailManLoveExp(self, day, player: Player):
        # man牵绊度数据
        for m in self.manDetailMap:
            manMap = {}
            manMap["日期"] = day
            manMap["玩家"] = player.name
            manMap["玩家等级"] = player.playerLev
            manMap["牵绊度等级"] = player.loveLevList[m].loveLev
            manMap["牵绊度经验"] = player.loveLevList[m].exp
            for key in self.enumlist:
                if key in player.loveLevList[m].loveTypeExpMap:
                    manMap[key] = player.loveLevList[m].loveTypeExpMap[key]
                else:
                    manMap[key] = 0
            cardList = [[] for _ in range(5)]
            for card in player.manCardList[m]:
                cardList[player.cardMap[card].rare].append(card)
            manMap["SSR思念"] = self.GetListStr(cardList[4])
            manMap["SR思念"] = self.GetListStr(cardList[3])
            del cardList
            for key in player.loveLevList[m].loveDetailTimesMap:
                manMap[key] = player.loveLevList[m].loveDetailTimesMap[key]

            for key in manMap:
                if key not in self.manDetailKeys:
                    self.manDetailKeys.append(key)
            rowList = []
            for key in self.manDetailKeys:
                if key in manMap:
                    rowList.append(manMap[key])
                else:
                    rowList.append(None)
            self.manDetailMap[m].append(rowList)
            del manMap

    def GetStatisticsData(self, day, player: Player, playerOrder):
        # 处理【总览统计】表
        row = day + 1
        stCol = playerOrder + 1
        ysCol = stCol + self.playNum
        ryCol = ysCol + self.playNum
        self.statisticsList[row][0] = day
        self.statisticsList[row][stCol] += player.loveLevList[1].loveLev
        self.statisticsList[row][ysCol] += player.loveLevList[2].loveLev
        self.statisticsList[row][ryCol] += player.loveLevList[5].loveLev
        if day == 1:
            self.statisticsList[1][stCol] = player.name
            self.statisticsList[1][ysCol] = player.name
            self.statisticsList[1][ryCol] = player.name

        # 处理【man统计】表
        row1 = day + self.maxDay * playerOrder
        row2 = self.playNum * (day - 1) + playerOrder + 1
        for m in self.manStatisticsMap:
            self.manStatisticsMap[m][row1][0] = day
            self.manStatisticsMap[m][row1][1] = player.name
            self.manStatisticsMap[m][row2][self.manColNum] = day
            self.manStatisticsMap[m][row2][1 + self.manColNum] = player.name
            for i in range(len(self.enumlist)):
                key = self.enumlist[i]
                exp = player.loveLevList[m].loveTypeExpMap[key] if key in player.loveLevList[m].loveTypeExpMap else 0
                self.manStatisticsMap[m][row1][2 + i] += exp
                self.manStatisticsMap[m][row2][2 + self.manColNum + i] += exp

    def GetAverageStatisticsData(self):
        for r in range(2, len(self.statisticsList)):
            for c in range(1, len(self.statisticsList[0])):
                self.statisticsList[r][c] /= self.tryTimes

        for m in self.manStatisticsMap:
            for r in range(1, len(self.manStatisticsMap[m])):
                for c in range(len(self.enumlist)):
                    self.manStatisticsMap[m][r][2 + c] /= self.tryTimes
                    self.manStatisticsMap[m][r][2 + self.manColNum + c] /= self.tryTimes

    @classmethod
    def GetListStr(cls, _list: list):
        content = ""
        for v in _list:
            if content == "":
                content = str(v)
            else:
                content = content + "|" + str(v)
        return content
