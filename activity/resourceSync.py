import xlwings as xw
from common.workApp import WorkApp
from common.printer import Printer
from common import common

# 脚本说明：
# 用于读取奖励投放相关的配置表
#


def main():
    printer = Printer()
    tablePath = common.getTablePath()

    dataSht = xw.sheets.active
    lastCell = dataSht.used_range.last_cell
    dataList = dataSht.used_range.value
    wbMap = {}  # [wb1Name:wb1,wb2Name:wb2]
    shtMap = {}  # {wb1:[sht1,sht2],wb2:[sht3,sht4]}
    checkColMap = {}  # {wb1&sht1:[col1,col2],wb2&sht2:[col3,col4]}
    findColMap = {}  # {wb1&sht1:[col1,col2],wb2&sht2:[col3,col4]}
    keyColMap = {}
    resultColMap = {}  # {wb1&sht1:[col1,col2],wb2&sht2:[col3,col4]}
    resultMap = {}  # {col1:[id1,id2],col2:[id3,id4]}
    wb = None

    with WorkApp() as app:
        printer.setStartTime("开始打开数据表格:", 'green')
        for i in range(lastCell.column):
            if dataList[0][i] is not None:
                wbName = dataList[0][i]
                wb = app.books.open(tablePath + "\\" + wbName, 0)
                wbMap[wbName] = wb
                shtMap[wbName] = []

            if dataList[1][i] is not None:
                shtName = dataList[1][i]
                shtMap[wbName].append(shtName)
                shtData = wbMap[wbName].sheets[shtName].used_range.value
                colKey = wbName + shtName
                checkColMap[colKey] = []  # 目标表格：key字段的位置
                keyColMap[colKey] = []  # 写入表格：key字段的位置
                findColMap[colKey] = []  # 目标表格：查询的字段的位置
                resultColMap[colKey] = []  # 写入表格：查询的字段的位置

            keyStr = dataList[2][i]
            colStr = dataList[3][i]
            if keyStr == 'key':
                colOrder = common.getDataColOrder(shtData, colStr, 2)
                if colOrder != -1:
                    checkColMap[colKey].append(colOrder)
                    keyColMap[colKey].append(i)
            elif colStr is not None:
                colOrder = common.getDataColOrder(shtData, colStr, 2)
                if colOrder != -1:
                    findColMap[colKey].append(colOrder)
                    resultColMap[colKey].append(i)
                    resultMap[i] = []
                else:
                    print(wbName + "找不到字段：" + colStr)
        printer.setCompareTime(printer.printGapTime("表格打开完毕，耗时:"))
        printer.setStartTime("开始同步表格数据:", 'green')
        for wbName in wbMap:
            for shtName in shtMap[wbName]:
                findData = wbMap[wbName].sheets[shtName].used_range.value
                colKey = wbName + shtName
                for row in range(4, lastCell.row):
                    if dataList[row][keyColMap[colKey][0]] is not None and dataList[row][keyColMap[colKey][0]] != 0:
                        checkIds = []
                        for i in keyColMap[colKey]:
                            checkIds.append(dataList[row][i])
                        rowData = common.getRowData(checkIds, checkColMap[colKey], findColMap[colKey], findData)
                        if rowData == [None] * len(findColMap[colKey]):
                            print(colKey, '缺少key-id：', checkIds, '或奖励内容为空')
                        for i in range(len(findColMap[colKey])):
                            resultMap[resultColMap[colKey][i]].append(rowData[i])

                    else:
                        for i in range(len(findColMap[colKey])):
                            resultMap[resultColMap[colKey][i]].append(None)
        printer.setCompareTime(printer.printGapTime("数据同步完毕，耗时:"))
    printer.setStartTime("开始写入数据:", 'green')
    for key in resultMap:
        dataSht.cells(5, key + 1).options(transpose=True).value = resultMap[key]
    printer.setCompareTime(printer.printGapTime("数据写入完毕，耗时:"))
    input("脚本执行完毕")