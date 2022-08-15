import xlwings as xw
from common.printer import Printer
from common import common
from common.workBook import WorkBook

getDataOrder = common.getDataOrder
toStr = common.toStr
toInt = common.toInt


# 脚本说明：
# 用于牵绊度随机奖励配置
#
def main():
    itemMap = {
        '短信': '20',
        '语音电话': '21',
        '朋友圈': '22',
    }

    printer = Printer()
    tablePath = common.getTablePath()

    app = xw.App(visible=False, add_book=False)
    try:
        printer.setStartTime("开始处理数据...", 'green')
        app.screen_updating = False
        app.display_alerts = False
        dataSht = xw.sheets.active
        dataRng = dataSht.used_range.value

        dropWb = WorkBook(tablePath, "Drop.xlsx", app, 'Drop')
        dropCols = dropWb.sheet.used_range.raw_value[2]
        idCol = common.getDataOrder(dropCols, 'ID') + 1
        groupCol = common.getDataOrder(dropCols, 'GroupID') + 1
        levCol = common.getDataOrder(dropCols, 'LevelID') + 1
        descCol = common.getDataOrder(dropCols, 'Description') + 1
        itemCol = common.getDataOrder(dropCols, 'Item') + 1
        weightCol = common.getDataOrder(dropCols, 'Weight') + 1
        timesCol = common.getDataOrder(dropCols, 'MaxTime') + 1

        for j in range(len(dataRng[0])):
            count = [0, 0, 0, 0]
            if dataRng[0][j] is not None:
                item = dataRng[0][j]
                itemType = itemMap[item]
            if j % 3 == 0:
                role = dataRng[1][j]
                groupId = dataRng[1][j + 1]
                for row in range(3, len(dataRng)):
                    itemID = dataRng[row][j]
                    if itemID is not None:
                        level = int(dataRng[row][j + 1])
                        count[level] += 1
                        id = groupId * 10000 + level * 1000 + count[level]
                        dataRow = dropWb.findRowById(id, idCol)
                        if dataRow == -1:
                            dataRow = dropWb.insertRowById(id, idCol)
                        dropWb.cells(dataRow, groupCol).raw_value = groupId
                        dropWb.cells(dataRow, levCol).raw_value = level
                        dropWb.cells(dataRow, descCol).raw_value = '牵绊度-' + role + item
                        dropWb.cells(dataRow, itemCol).raw_value = itemType + '=' + toStr(itemID) + '=1'
                        dropWb.cells(dataRow, weightCol).raw_value = 500
                        dropWb.cells(dataRow, timesCol).raw_value = 1

    finally:
        app.screen_updating = True
        app.display_alerts = True
        dropWb.close(True)
        app.quit()
        printer.setCompareTime(printer.printGapTime("数据处理完毕，耗时:"))