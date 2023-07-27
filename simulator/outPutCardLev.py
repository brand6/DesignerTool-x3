from xlDeal import XlDeal


class OutPutCardLev:
    def __init__(self) -> None:
        self.xl = XlDeal("资源投放统计.xlsm", "养成模拟")
        self.dataList = self.xl.data

    def UpdateTableData(self):
        self.xl.UpdateTableData(self.dataList)

    def GetDevelopNum(self, day):
        return self.dataList[day + 1][1]

    def UpdateRes(self, resStr, day, playerOrder):
        self.dataList[day + 1][2 + playerOrder] = resStr

    def UpdateData(self, levMap, day, playerOrder, rare):
        levStr = ""
        for lev in levMap:
            if levStr == "":
                levStr = str(lev) + "=" + str(levMap[lev])
            else:
                levStr = levStr + "\n" + str(lev) + "=" + str(levMap[lev])
        self.dataList[day + 1][(rare - 1) * 6 + playerOrder + 2] = levStr
