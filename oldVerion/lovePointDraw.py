import xlwings as xw
from xlwings import Range
from xlwings import Sheet

# 脚本说明：
# 根据抽卡数计算获得搭档&羁绊数量
#

playExtendNum = 3  # 玩家种类扩展数量（玩家分类数-1）
paraBeginLine = 3  # 数据开始行（excel行数-1）


def main():

    paraSht = xw.books.active.sheets['验算参数']
    paraRange = paraSht.used_range
    paraList = getDataList(paraRange, "基础参数", -1)
    paraDict = getParaDict(paraList)
    roleNum = paraDict["男主数量"]
    hangNumList = [-1, 0, 0, paraDict["挂机SSR羁绊"], paraDict["挂机SR羁绊"], paraDict["挂机R羁绊"]]
    ItemList = getDataList(paraRange, "养成参数", -1)
    itemsDict = getItemsDict(ItemList)
    timesList = getDataList(paraRange, "抽卡次数", playExtendNum)
    dayList = getDataList(paraRange, "日期", 0)

    for i in range(1, len(ItemList)):
        numList = []
        item = ItemList[i][0]
        drawRate = itemsDict[item]['抽卡概率']
        drawExcept = itemsDict[item]['抽卡期望']
        initNum = itemsDict[item]['初始数量']
        for m in range(paraBeginLine, len(timesList)):
            rowList = []
            for n in range(len(timesList[m])):
                drawTimes = timesList[m][n]
                firstCardExcept = drawExcept * 2  # 第二次出货为想要男主的卡
                if drawTimes < firstCardExcept:
                    num = drawTimes / firstCardExcept
                else:
                    num = 1 + (drawTimes - firstCardExcept) * drawRate / roleNum
                num = num + initNum + hangNumList[i] * dayList[m] / roleNum
                if num > 100:
                    num = int(num)
                elif num > 10:
                    num = round(num, 1)
                else:
                    num = round(num, 2)
                rowList.append(num)
            numList.append(rowList)
        setRangeData(numList, paraSht, paraBeginLine + 1, item + '数量')


def getlevsDict(developList: list, colList: list):
    """将等级参数转为字典

    Args:
        developList (list): 养成数据表的数据

    Returns:
        dict: 参数字典
    """

    colNum = len(developList[0])
    for i in range(len(developList[0])):
        if developList[0][i] == '等级':
            levCol = i

    levsDict = {}
    for i in range(3, len(developList)):
        levDict = {}
        for j in range(1, colNum):
            dataStr = developList[0][j]
            if dataStr in colList:
                dataMap = {}
                for m in range(5):
                    dataMap[developList[1][j + m]] = developList[i][j + m]
                levDict[dataStr] = dataMap
        levsDict[developList[i][levCol]] = levDict
    return levsDict


def getDaysDict(paraList: list, colList: list):
    """将日期参数转为字典

    Args:
        paraList (list): 验算参数表的数据
        colList (list): 需转为字典的列名

    Returns:
        dict: 参数字典
    """

    colNum = len(paraList[0])
    for i in range(len(paraList[0])):
        if paraList[0][i] == '日期':
            dayCol = i

    daysDict = {}
    for i in range(paraBeginLine, len(paraList)):
        dayDict = {}
        for j in range(dayCol + 1, colNum):
            dataStr = paraList[0][j]
            if dataStr in colList:
                dataList = []
                for m in range(playExtendNum + 1):
                    dataList.append(paraList[i][j + m])
                dayDict[dataStr] = dataList
        daysDict[paraList[i][dayCol]] = dayDict
    return daysDict


def getItemsDict(itemList: list):
    """将养成参数转为字典

    Args:
        itemList (list): _description_

    Returns:
        dict: 参数字典
    """
    itemsDict = {}
    for j in range(1, len(itemList[0])):
        itemDict = {}
        for i in range(1, len(itemList)):
            itemDict[itemList[i][0]] = itemList[i][j]
        itemsDict[itemList[0][j]] = itemDict
    return itemsDict


def getParaDict(paraList: list):
    """将参数转为字典

    Args:
        paraList (list): 参数列表

    Returns:
        dict: 参数字典
    """
    paraDict = {}
    for i in range(1, len(paraList)):
        paraDict[paraList[i][0]] = paraList[i][1]
    return paraDict


def getDataList(dataRange: Range, dataCol: str, expandNum: int = 0):
    """根据标题获取参数块

    Args:
        dataRange (Range): 数据range
        dataCol (str): _description_
        expandType (int): 扩展列数，-1=table
    Returns:
        _type_: _description_
    """

    for cell in dataRange.rows[0]:
        if cell.value == dataCol:
            if expandNum == -1:
                return cell.expand('table').value
            else:
                expandRow = dataRange.last_cell.row - 1
                return Range(cell, cell.offset(expandRow, expandNum)).value


def setRangeData(dataList: list, sht: Sheet, beginRow: int, colStr: str, beginLine=0):
    """将数组写入表格

    Args:
        dataList (_type_): 待写入的数组
        sht (Sheet): 表格
        beginRow (int): 表格写入开始行
        colStr (str): 列名
        beginLine(int): dataList 数据开始行
    """
    shtDataList = sht.used_range.value
    for i in range(len(shtDataList[0])):
        if shtDataList[0][i] == colStr:
            dataCol = i + 1
    sht.cells(beginRow, dataCol).value = dataList[beginLine:]
