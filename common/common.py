import os
from tkinter import filedialog
import xlwings as xw


# 退出隐藏的app
def quitHideApp():
    for app in xw.apps:
        if app.visible is False:
            print('quit unVisible app')
            app.quit()


# 获取操作的配置表目录
def getTablePath():
    configPath = os.path.dirname(__file__) + r'\config.txt'
    if os.path.exists(configPath):
        with open(configPath, 'r', encoding='GBK') as file:
            tablePath = file.read()
            if os.path.exists(tablePath):
                return tablePath

    with open(configPath, 'w', encoding='GBK') as file:
        tablePath = selectDirectory("选择项目Program文件夹")
        loc = tablePath.index('Program')

        tablePath = tablePath[:loc] + r'Program\Binaries\Tables\OriginTable'
        file.write(tablePath)
        return tablePath


def selectDirectory(_title="选择文件夹"):
    selectPath = filedialog.askdirectory(title=_title).replace('/', '\\')
    return selectPath


def getDataOrder(columnData, columnName):
    """根据列名查找在列表或元组中的顺序

    Args:
        columnData (_type_): 数组或元组
        columnName (_type_): 列名

    Returns:
        _type_: _description_
    """
    for i in range(len(columnData)):
        if columnData[i] == columnName:
            return i
    else:
        return -1


# 根据列名获得列的位置
def getDataColOrder(dataList, colStr, checkRow=0):
    for order in range(len(dataList[checkRow])):
        if colStr == dataList[checkRow][order]:
            return order
    else:
        return -1


# 查找所在行，未查到时返回插入位置
def getRangeRow(findRange, findValue):
    findData = findRange.value
    insertRow = 1
    for i in range(len(findData)):
        if findData[i] == findValue:
            return i, insertRow
        elif isNumber(findData[i]) and findData[i] < findValue:
            insertRow = i + 2
    else:
        return -1, insertRow


def getListRow(findList, findValue):
    for i in range(len(findList)):
        if findList[i] == findValue:
            return i
    else:
        return -1


# 转为字符串
def toStr(content):
    if content is None:
        return ''
    elif isNumber(content):
        intC = toInt(content)
        if intC == round(float(content), 10):
            return str(intC)
        else:
            return str(content)
    else:
        return str(content)


# 转为数值用于计算，非数值返回0
def toNum(content):
    if isNumber(content):
        return float(content)
    else:
        return 0


# 转为整数，四舍五入，非数值返回-1
def toInt(content):
    if isNumber(content):
        return int(round(float(content), 0))
    else:
        return -1


# 判断是否有效
def isNumberValid(content, checkNum=0):
    if not isNumber(content):
        return False
    elif float(content) > checkNum:
        return True
    else:
        return False


# 判断是否数字
def isNumber(content):
    try:
        int(content)
        return True
    except BaseException:
        return False


# 判断是否为空
def isEmpty(content):
    if content == '' or content is None:
        return True
    else:
        return False


# 重载分隔操作
def split(content, sep):
    if content is None:
        return []
    else:
        return str.split(toStr(content), sep)


# 调用bat文件
def callBat(filePath, fileName):
    os.chdir(filePath)
    if fileName[-3:] != "bat":
        fileName = fileName + ".bat"

    os.system(r'start cmd.exe /k' + filePath + "\\" + fileName)
