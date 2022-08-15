import sys
import xlwings as xw
from miaoCard.miaoCard import MiaoCard


def main():
    app = xw.apps.active
    try:
        app.screen_updating = False
        sht = app.books.active.sheets.active
        runTimes = int(sht.range("K1").value)
        self = MiaoCard(sht)
        for i in range(10, len(self.dataRng)):  # 清除旧日志
            self.dataRng[i][7] = None
            self.dataRng[i][13] = None
        print("********  先手方：" + self.firstMove + "  ********")
        print("********  AI男主：" + self.AIRole + "  ********")
        print("********  AI基础牌策略：" + self.strategy + "  ********")
        if not self.isBasic:
            print("********  AI功能牌策略：" + self.AIStatus + "  ********")
        print("********  模拟打牌次数：" + str(runTimes) + "  ********")

        for i in range(runTimes):
            self.GameBegin(False)
            while (self.emptyNum > 0):
                self.roundTimes += 1
                self.isPL = False if self.nextMove == 'AI' else True
                if not self.isBasic:
                    if len(self.funCardLib) > 0:
                        self.GetFunCard()
                    self.DoFun()
                self.DoAction()
            self.ChangeWinData()
            sys.stdout.write('\rAI获胜次数：%d      玩家获胜次数：%d      平局次数：%d' % (self.AIWinTimes, self.PLWinTimes, self.peaceTimes))
        if runTimes < 10:
            self.ShowChange()
        else:
            self.PrintWinLog()
            sht.range("I1").value += self.PLWinTimes
            sht.range("I2").value += self.AIWinTimes
            sht.range("K2").value += self.peaceTimes
    finally:
        app.screen_updating = True
