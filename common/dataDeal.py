import copy
import re
import numpy as np
from datetime import datetime


class DataDeal():
    numChar = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
    kinChar = ['', '十', '百', '千', '万']

    def __init__(self, rngData=None, colRow=0, mainCol=None):
        self.rngData = rngData
        self.colRow = colRow  # 列名所在行
        self.mainCol = mainCol  # 主键/主键列表
        self.rngCheckValue = ''  # 上一个处理的数据
        self.oriCheckValue = ''  # 上一个处理的源数据
        self.rngCheckRow = 0  # 上一个处理的行
        self.oriCheckRow = 0  # # 上一个处理的源行
        self.lastInsertId = 0
        self.transId = 0
        self.insertMap = {}
        self.repeatLog = 0

    def setData(self, rngData):
        self.rngData = rngData

    def getData(self):
        return self.rngData

    # 打印重复信息
    def setRepeatLog(self, tag=1):
        self.repeatLog = tag
        self.repeatMap = {}

    # 设置数据源
    def setOriData(self, oriData, oriRow=0):
        self.oriData = oriData
        self.oriRow = oriRow

    # 设置数据的值
    def setRowData(self, checkId, setMap={}, checkCol=0, dataType=0, extraId='', extraCol=1, defRow=-1, insertType=0):
        checkData = self.rngData
        if dataType != 0:
            checkData = self.oriData

        if extraId == '':
            dataRow = self.getRow(checkId, checkCol, dataType)
        else:
            dataRow = self.getRow2(checkId, checkCol, extraId, extraCol, dataType)
        if dataRow == -1:
            if insertType == 0:
                dataRow = self.getInsertRow(checkId, checkCol)
            else:
                dataRow = self.getLastAddRow()
            checkData.insert(dataRow, copy.deepcopy(checkData[dataRow - 1]))

        for col in setMap:
            colOrder = self.getColOrder(col, dataType, defRow)
            checkData[dataRow][colOrder] = setMap[col]

    # 根据列名获得列的位置
    def getColOrder(self, checkCol, dataType=0, defRow=-1):
        if self.isNumber(checkCol):
            return int(checkCol)
        else:
            checkData = self.rngData
            checkRow = self.colRow
            if dataType != 0:
                checkData = self.oriData
                checkRow = self.oriRow

            if defRow != -1:
                checkRow = defRow

            for order in range(len(checkData[checkRow])):
                if checkCol == checkData[checkRow][order]:
                    return order
            else:
                return -1

    # 获得插入到数据最后的位置行
    def getLastAddRow(self):
        for row in range(len(self.rngData) - 1, 0, -1):
            if self.isEmpty(self.rngData[row][0]) is False:
                return row + 1

    # 获得插入的位置行
    def getInsertRow(self, checkValue, checkCol=0):
        if checkValue == self.lastInsertId:
            self.lastInsertRow += 1
            return self.lastInsertRow

        r = len(self.rngData)
        if self.isNumber(checkValue):
            checkCol = self.getColOrder(checkCol)
            checkValue = float(checkValue)
            if checkValue == self.lastInsertId + 1:
                self.lastInsertRow += 1
                return self.lastInsertRow
            for row in range(len(self.rngData) - 1, 0, -1):
                compareValue = self.rngData[row][checkCol]
                if compareValue is not None and (self.isNumber(compareValue) is False or float(compareValue) <= checkValue):
                    r = row + 1
                    break
        else:
            for row in range(len(self.rngData) - 1, 0, -1):
                if self.rngData[row][checkCol] is not None:
                    r = row + 1
                    break

        self.lastInsertRow = r
        self.lastInsertId = checkValue
        return self.lastInsertRow

    # 获得数据对应行
    def getRow(self, checkValue, checkCol=0, dataType=0):
        checkCol = self.getColOrder(checkCol, dataType)
        checkValue = self.toStr(checkValue)
        checkData = self.rngData
        if dataType != 0:
            checkData = self.oriData

        for row in range(len(checkData)):
            compareValue = checkData[row][checkCol]
            if self.toStr(compareValue) == checkValue:
                return row
        else:
            return -1

    # 获得数据对应的所有行
    def getRows(self, checkValue, checkCol=0, dataType=0):
        checkCol = self.getColOrder(checkCol, dataType)
        checkValue = self.toStr(checkValue)
        returnList = []
        checkData = self.rngData
        if dataType != 0:
            checkData = self.oriData

        for row in range(len(checkData)):
            compareValue = checkData[row][checkCol]
            if self.toStr(compareValue) == checkValue:
                returnList.append(row)

        return returnList

    # 根据两个参数获得行
    def getRow2(self, checkValue, checkCol, checkValue2, checkCol2, dataType=0):
        checkCol = self.getColOrder(checkCol, dataType)
        checkCol2 = self.getColOrder(checkCol2, dataType)
        checkValue = self.toStr(checkValue)
        checkValue2 = self.toStr(checkValue2)

        startRow = 0
        checkData = self.rngData
        if dataType == 0:
            if self.rngCheckValue == checkValue and len(checkData) > 100000:
                if self.rngCheckRow == -1:
                    return -1
                else:
                    startRow = self.rngCheckRow
            else:
                self.rngCheckValue = checkValue
        else:
            checkData = self.oriData
            if self.oriCheckValue == checkValue and len(checkData) > 100000:
                if self.oriCheckRow == -1:
                    return -1
                else:
                    startRow = self.oriCheckRow
            else:
                self.oriCheckValue = checkValue

        for row in range(startRow, len(checkData)):
            compareValue = checkData[row][checkCol]
            compareValue2 = checkData[row][checkCol2]
            if self.toStr(compareValue) == checkValue and self.toStr(compareValue2) == checkValue2:
                if dataType == 0:
                    self.rngCheckRow = row
                else:
                    self.oriCheckRow = row
                return row
        else:
            if dataType == 0:
                self.rngCheckRow = -1
            else:
                self.oriCheckRow = -1
            return -1

    # 根据三个参数获得行
    def getRow3(self, checkValue, checkCol, checkValue2, checkCol2, checkValue3, checkCol3, dataType=0,):  # yapf:disable
        checkCol = self.getColOrder(checkCol, dataType)
        checkCol2 = self.getColOrder(checkCol2, dataType)
        checkCol3 = self.getColOrder(checkCol3, dataType)
        checkValue = self.toStr(checkValue)
        checkValue2 = self.toStr(checkValue2)
        checkValue3 = self.toStr(checkValue3)

        checkData = self.rngData
        startRow = 0
        if dataType == 0:
            if self.rngCheckValue == checkValue and len(checkData) > 100000:
                if self.rngCheckRow == -1:
                    return -1
                else:
                    startRow = self.rngCheckRow
            else:
                self.rngCheckValue = checkValue
        else:
            checkData = self.oriData
            if self.oriCheckValue == checkValue and len(checkData) > 100000:
                if self.oriCheckRow == -1:
                    return -1
                else:
                    startRow = self.oriCheckRow
            else:
                self.oriCheckValue = checkValue

        for row in range(startRow, len(checkData)):
            compareValue = checkData[row][checkCol]
            compareValue2 = checkData[row][checkCol2]
            compareValue3 = checkData[row][checkCol3]
            if self.toStr(compareValue) == checkValue and self.toStr(compareValue2) == checkValue2 and self.toStr(compareValue3) == checkValue3:  # yapf:disable
                if dataType == 0:
                    self.rngCheckRow = row
                else:
                    self.oriCheckRow = row
                return row
        else:
            if dataType == 0:
                self.rngCheckRow = -1
            else:
                self.oriCheckRow = -1
            return -1

    # 查找数据
    def getRowData(self, checkValue, returnCol, checkCol=0, dataType=0):
        checkCol = self.getColOrder(checkCol, dataType)
        returnCol = self.getColOrder(returnCol, dataType)
        checkValue = self.toStr(checkValue)

        checkData = self.rngData
        if dataType != 0:
            checkData = self.oriData

        for row in range(len(checkData)):
            compareValue = checkData[row][checkCol]
            if self.toStr(compareValue) == checkValue:
                return checkData[row][returnCol]
        else:
            return -1

    # 获得复制对应的列
    def __getCopyColumn(self):
        copyColList = []
        for i in range(len(self.rngData[0])):
            if self.rngData[0][i] is not None:
                for j in range(len(self.oriData[0])):
                    if self.rngData[0][i] == self.oriData[0][j]:
                        copyColList.append(j)
                        break
                else:
                    copyColList.append(-1)
            elif self.rngData[1][i] is not None:
                for j in range(len(self.oriData[1])):
                    if self.rngData[1][i] == self.oriData[1][j]:
                        copyColList.append(j)
                        break
                else:
                    copyColList.append(-1)
            else:
                copyColList.append(-1)
        self.copyColList = copyColList

    # 复制单行数据
    def __copyRowData(self, copyId, oriRow, tarRow, checkCol, insertType=0):
        # 如果数据不存在则新增行
        if tarRow == -1:
            self.insertMap[copyId] = 1
            tarRow = self.getInsertRow(copyId, checkCol)
            if insertType != 0:
                self.rngData.insert(tarRow, copy.deepcopy(self.rngData[tarRow - 1]))
            else:
                insertList = []
                for i in range(len(self.rngData[tarRow - 1])):
                    insertList.append(None)
                self.rngData.insert(tarRow, insertList)
        elif self.repeatLog == 1:  # 目标表格已有该id时打印
            if copyId not in self.insertMap and copyId not in self.repeatMap:
                self.repeatMap[copyId] = 1
                print('目标表格存在该id:' + self.toStr(copyId))

        for i in range(len(self.copyColList)):
            if self.copyColList[i] != -1:
                self.rngData[tarRow][i] = self.oriData[oriRow][self.copyColList[i]]

    # 复制数据
    def copyData(self, copyId, checkCol):
        if not hasattr(self, 'copyColList'):
            self.__getCopyColumn()
        if type(checkCol).__name__ == 'list':
            oriCol = []
            for col in checkCol:
                oriCol.append(self.getColOrder(col, 1))
            oriRows = self.getRows(copyId, oriCol[0], 1)

            if len(checkCol) == 2:
                for r in oriRows:
                    tarRow = self.getRow2(copyId, checkCol[0], self.oriData[r][oriCol[1]], checkCol[1])
                    self.__copyRowData(copyId, r, tarRow, checkCol[0])
            elif len(checkCol) == 3:
                for r in oriRows:
                    tarRow = self.getRow3(copyId, checkCol[0], self.oriData[r][oriCol[1]], checkCol[1],
                                          self.oriData[r][oriCol[2]], checkCol[2])
                    self.__copyRowData(copyId, r, tarRow, checkCol[0])
        else:
            oriRow = self.getRow(copyId, checkCol, 1)
            tarRow = self.getRow(copyId, checkCol)
            self.__copyRowData(copyId, oriRow, tarRow, checkCol)

    # 获得主键列
    def getMainCol(self, dataType=0):
        if self.mainCol is None:
            checkCol = None
            checkData = self.rngData
            if dataType != 0:
                checkData = self.oriData

            if checkData[1][0] == 'map:int':
                checkCol = checkData[0][0]
            elif '#' not in checkData[2][0]:
                for col in checkData[2]:
                    if 'key' in col:
                        if checkCol is None:
                            checkCol = [col]
                        else:
                            checkCol.append(col)
            self.mainCol = checkCol
        return self.mainCol

    # 修改指定列数据
    def updateData(self, checkId, checkCol, updateCol, updateStr=''):
        rows = []
        if type(checkCol).__name__ == 'list':
            rows = self.getRows(checkId, checkCol[0])
        else:
            rows = self.getRows(checkId, checkCol)

        if rows != []:
            for r in rows:
                self.rngData[r][updateCol] = updateStr
            return len(rows)
        else:
            return 0

    # 复制所有数据
    def copyAllData(self, checkCol, startRow=0):
        for row in self.oriData[startRow:]:
            self.copyData(row[checkCol], checkCol)

    def setTransColumn(self, rowList=[]):
        transColList = []
        for i in range(len(self.rngData[0])):
            if self.rngData[0][i] is not None:
                for j in range(len(self.oriData[0])):
                    rngCol = self.rngData[0][i]
                    oriCol = self.oriData[0][j]
                    if rngCol + "#E" == oriCol:
                        transColList.append(j)
                        break
                    elif rngCol == oriCol:
                        if '#' in oriCol or oriCol in rowList:
                            transColList.append(j)
                            break
                else:
                    transColList.append(-1)
            else:
                transColList.append(-1)
        self.transColList = transColList

    # 获得翻译的列
    def __getTransColumn(self):
        transColList = []
        for i in range(len(self.rngData[0])):
            if self.rngData[0][i] is not None:
                for j in range(len(self.oriData[0])):
                    if self.rngData[0][i] + "#E" == self.oriData[0][j]:
                        transColList.append(j)
                        break
                    elif self.rngData[0][i] == self.oriData[0][j]:
                        if '#' in self.rngData[0][i]:
                            transColList.append(j)
                            break
                else:
                    transColList.append(-1)
            else:
                transColList.append(-1)
        self.transColList = transColList

    # 翻译单行数据
    def __transRowData(self, transId, oriRow, tarRow, checkCol):
        checkCol = self.getColOrder(checkCol)
        # 如果数据不存在则插入一行
        if tarRow == -1:
            tarRow = self.getInsertRow(transId, checkCol)
            noneList = copy.deepcopy(self.rngData[tarRow - 1])
            for i in range(len(noneList)):
                noneList[i] = None
            self.rngData.insert(tarRow, noneList)
            self.rngData[tarRow][checkCol] = transId

        for i in range(len(self.transColList)):
            if self.transColList[i] != -1:
                self.rngData[tarRow][i] = self.oriData[oriRow][self.transColList[i]]

    # 翻译数据
    def transData(self, transId, checkCol):
        self.transId = transId
        if not hasattr(self, 'transColList'):
            self.__getTransColumn()
        if type(checkCol).__name__ == 'list':
            oriCol = []
            for col in checkCol:
                oriCol.append(self.getColOrder(col, 1))
            oriRows = self.getRows(transId, oriCol[0], 1)

            if len(checkCol) == 2:
                for r in oriRows:
                    tarRow = self.getRow2(transId, checkCol[0], self.oriData[r][oriCol[1]], checkCol[1])
                    self.__transRowData(transId, r, tarRow, checkCol[0])
            elif len(checkCol) == 3:
                for r in oriRows:
                    tarRow = self.getRow3(transId, checkCol[0], self.oriData[r][oriCol[1]], checkCol[1],
                                          self.oriData[r][oriCol[2]], checkCol[2])
                    self.__transRowData(transId, r, tarRow, checkCol[0])
        else:
            oriRow = self.getRow(transId, checkCol, 1)
            tarRow = self.getRow(transId, checkCol)
            self.__transRowData(transId, oriRow, tarRow, checkCol)

    # 获取空数据区域
    def getBlankRange(self):
        rowNum = 0
        tag = False
        for row in self.rngData:
            for cell in row:
                if cell is not None:
                    break
            else:
                np.delete(self.rngData, rowNum, axis=0)
                tag = True
                continue
            rowNum += 1
        return tag

    # 处理单元格格式
    def strCheck(self):
        for col in range(len(self.rngData[0])):
            if self.rngData[0][col] is not None:
                for row in range(len(self.rngData)):
                    if self.rngData[row][col] is not None:
                        if self.isNumber(self.rngData[row][col]) is True:
                            pass
                        # 处理时间格式
                        elif type(self.rngData[row][col]).__name__ == 'datetime':
                            self.rngData[row][col] = "'" + datetime.strftime(self.rngData[row][col], r'%Y/%m/%d %H:%M:%S')
                        # 处理带逗号和冒号的数字，文本格式的时间
                        elif re.match(r"[^']\d*[:|,|/]", self.rngData[row][col]) is not None:
                            self.rngData[row][col] = "'" + self.rngData[row][col][0:]
                        # 处理'开头的文本
                        elif re.match(r"'[^']", self.rngData[row][col]) is not None:
                            self.rngData[row][col] = "'" + self.rngData[row][col][0:]

    # 重新生成连续的uid
    def uidRebuild(self, idCol=0):
        orderNum = 1
        for row in range(len(self.rngData)):
            if self.isNumber(self.rngData[row][idCol]):
                self.rngData[row][idCol] = orderNum
                orderNum += 1

    # 获得最大id
    def getMaxId(self, checkCol, dataType=0, condition=None, conditionCol=None):
        checkCol = self.getColOrder(checkCol, dataType)
        if condition is not None:
            condition = self.toStr(condition)
        if conditionCol is not None:
            conditionCol = self.getColOrder(conditionCol, dataType)

        maxId = 0
        checkData = self.rngData
        if dataType != 0:
            checkData = self.oriData

        for row in range(len(checkData)):
            compareValue = self.toInt(checkData[row][checkCol])

            if condition is None:
                if compareValue > maxId:
                    maxId = compareValue
            else:
                conditionValue = self.toStr(checkData[row][conditionCol])
                if compareValue > maxId and condition == conditionValue:
                    maxId = compareValue

        return maxId

    # 增加列
    def addCol(self, colName):
        self.rngData[0].append(colName)
        for row in self.rngData[1:]:
            row.append(None)
        return self.getColOrder(colName)
