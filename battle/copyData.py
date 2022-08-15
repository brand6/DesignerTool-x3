import xlwings as xw
from common.workApp import WorkApp
from common.printer import Printer
from common import common


# 脚本说明：
# 用于战斗数值模板表同步表格数据
#
def main():
    printer = Printer()
    tablePath = common.getTablePath()

    dataSht = xw.sheets.active
    lastCell = dataSht.used_range.last_cell
    dataSht.range(dataSht.cells(4, 1), lastCell).clear_contents()
    wb = None
    sht = None
    lastC = None

    with WorkApp() as app:
        for i in range(1, lastCell.column + 1):
            if dataSht.cells(1, i).value is not None:
                if wb is not None:
                    wb.close()
                printer.setStartTime("开始打开来源表格:" + dataSht.cells(1, i).value, 'green')
                wb = app.books.open(tablePath + "\\" + dataSht.cells(1, i).value, 0)
                printer.setCompareTime(printer.printGapTime("来源表格打开完毕，耗时:"))

                printer.printColor("数据处理中...")
            if dataSht.cells(2, i).value is not None:
                sht = wb.sheets[dataSht.cells(2, i).value]
                lastC = sht.used_range.last_cell
            for j in range(1, lastC.column + 1):
                if sht.cells(3, j).raw_value == dataSht.cells(3, i).raw_value:
                    copyValue = sht.range(sht.cells(4, j), sht.cells(lastC.row, j)).value
                    dataSht.cells(4, i).options(transpose=True).value = copyValue
        wb.close()
