import xlwings as xw
from miaoCard.miaoCard import MiaoCard


def main():
    app = xw.apps.active
    try:
        app.screen_updating = False
        self = MiaoCard()
        self.isPL = False if self.nextMove == 'AI' else True
        self.GetEmptyGridNum()
        if self.emptyNum > 0:
            self.log = []
            self.PLLog = []
            self.roundTimes += 1
            if not self.isBasic:
                if len(self.funCardLib) > 0:
                    self.GetFunCard()
                self.DoFun()
            self.DoAction()
            if self.emptyNum == 0:
                self.ChangeWinData()
            self.ShowChange()
    finally:
        app.screen_updating = True
