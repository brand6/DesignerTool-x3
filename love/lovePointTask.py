import xlwings as xw
from common.printer import Printer
from common import common

getDataOrder = common.getDataOrder
toStr = common.toStr
toInt = common.toInt


# 脚本说明：
# 用于牵绊度任务配置
#
def main():
    # 男主开关 1导出 0不导出
    roleMap = [None, 1, 1, 0, 0, 1]

    # 不同男主替换
    replaceMap = {
        '{0}': (None, '0', '1', '2', '3', '4'),
        '{1}': (None, '1', '2', '3', '4', '5'),
        '{2}': (None, '2', '3', '4', '5', '6'),
        '{3}': (None, '3', '4', '5', '6', '7'),
        '{5}': (None, '5', '6', '7', '8', '9'),
        '{09}': (None, '09', '10', '11', '12', '13'),
        '{光}': (None, '光', '冰', '3', '4', '火'),
        '{1004}': (None, '1004', '1005', '1007', '1008', '1006'),
    }

    try:
        printer = Printer()

        printer.setStartTime("开始处理数据...")
        dataSht = xw.sheets.active
        dataRng = dataSht.used_range
        dataValues = dataRng.value
        dv_columnData = dataRng.rows[0].value
        cv_columnData = xw.sheets['任务配置'].used_range.rows[0].value
        skipCol = getDataOrder(cv_columnData, 'SkipExport')

        # 根据不同的男主替换文本
        def replaceStr(content, roleId):
            for k in replaceMap:
                if k in content:
                    return content.replace(k, replaceMap[k][roleId])
            else:
                return content

        # 初始化taskValues
        taskValues = [[None for j in range(len(cv_columnData))] for i in range(5 * len(dataValues))]
        for j in range(len(cv_columnData)):
            taskValues[0][j] = cv_columnData[j]

        # taskValues赋值
        for r in range(1, len(dataValues)):
            # 首个单元格不为空时，处理本行数据
            if dataValues[r][0] is not None:
                for i in range(1, 6):
                    # 处理不导出字段
                    for d in range(len(dv_columnData)):
                        if dv_columnData[d] is not None:
                            c = getDataOrder(cv_columnData, dv_columnData[d])
                            if not isinstance(dataValues[r][d], str):
                                taskValues[(r - 1) * 5 + i][c] = dataValues[r][d]
                            else:
                                taskValues[(r - 1) * 5 + i][c] = replaceStr(dataValues[r][d], i)
                    if roleMap[i] == 0:
                        taskValues[(r - 1) * 5 + i][skipCol] = 1

        xw.sheets['任务配置'].cells(1, 1).value = taskValues
        xw.sheets['任务配置'].activate()

    finally:
        printer.setCompareTime(printer.printGapTime("数据处理完毕，耗时:"))
