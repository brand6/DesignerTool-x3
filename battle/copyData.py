import xlwings as xw
from common.printer import Printer
from common import common


# 脚本说明：
# 用于同步配置表的数据
#
def main():
    printer = Printer()
    tablePath = common.getTablePath()

    dataSht = xw.sheets.active
    lastCell = dataSht.used_range.last_cell
    if lastCell.row > 3:
        dataSht.range(dataSht.cells(4, 1), lastCell).clear_contents()
    wb = None
    sht = None
    lastC = None

    for i in range(1, lastCell.column + 1):
        # 更新表名
        if dataSht.cells(1, i).value is not None:
            if wb is not None:  # 关闭上一个处理的表格
                wb.close()
            printer.setStartTime("开始打开来源表格:" + dataSht.cells(1, i).value, 'green')
            wb = xw.books.open(tablePath + "\\" + dataSht.cells(1, i).value, 0)
            printer.setCompareTime(printer.printGapTime("来源表格打开完毕，耗时:"))

        # 更新sht名
        if dataSht.cells(2, i).value is not None:
            sht = wb.sheets[dataSht.cells(2, i).value]
            lastC = sht.used_range.last_cell

        # 复制数据
        for j in range(1, lastC.column + 1):
            if sht.cells(3, j).raw_value == dataSht.cells(3, i).raw_value:
                copyValue = sht.range(sht.cells(4, j), sht.cells(lastC.row, j)).value
                dataSht.cells(4, i).options(transpose=True).value = copyValue
                break

    wb.close()
    input("数据更新完毕...")
