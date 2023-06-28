import os
from tkinter import filedialog
import xlwings as xw


def getDataOrders(findList: list, findValue):
    """在一维list或元组内查找数据所在的所有行，未找到时返回空列表

    Args:
        findList (list): 数组或元组
        findValue (_type_): 查找的数据

    Returns:
        _type_: _description_
    """
    returnList = []
    for i in range(len(findList)):
        if findList[i] == findValue:
            returnList.append(i)
        elif toStr(findList[i]) == toStr(findValue):
            returnList.append(i)

    return returnList


def getDataOrder(findList: list, findValue):
    """在一维list或元组内查找数据所在的index，未找到时返回-1

    Args:
        findList (list): 数组或元组
        findValue (_type_): 查找的数据

    Returns:
        _type_: _description_
    """
    for i in range(len(findList)):
        if findList[i] == findValue:
            return i
        elif toStr(findList[i]) == toStr(findValue):
            return i
    else:
        return -1


def getColBy3Para(wbName: str, shtName: str, colName: str | list, findData: list):
    """传入3行参数，获取列名所在的位置，可一次查找shtName下的多个colName，未找到的列返回-1

    Args:
        wbName (str): 表名
        shtName (str): sht名
        colName (str | list): 列名，可传入str或list
        propertyData (list): 属性数据
    """

    _wbName = None
    _shtName = None
    if isinstance(colName, list):
        colList = [-1] * len(colName)
        for col in range(len(findData[0])):
            if findData[0][col] is not None:
                _wbName = findData[0][col]
            if findData[1][col] is not None:
                _shtName = findData[1][col]
            for i in range(len(colName)):
                if wbName == _wbName and shtName == _shtName and colName[i] == findData[2][col]:
                    colList[i] = col
        return colList
    else:
        for col in range(len(findData[0])):
            if findData[0][col] is not None:
                _wbName = findData[0][col]
            if findData[1][col] is not None:
                _shtName = findData[1][col]
            if wbName == _wbName and shtName == _shtName and colName == findData[2][col]:
                return col
        else:
            return -1


def getColBy2Para(classify: str, colName: str | list, findData: list):
    """传入2行参数，获取列名所在的位置，可一次查找classify下的多个colName，未找到的列返回-1

    Args:
        classify (str): 分类名
        colName (str| list): 列名，可传入str或list
        roleData (list): 属性数据
    """
    _classify = None
    if isinstance(colName, list):
        colList = [-1] * len(colName)
        for col in range(len(findData[0])):
            if findData[0][col] is not None:
                _classify = findData[0][col]
            for i in range(len(colName)):
                if classify == _classify and colName[i] == findData[1][col]:
                    colList[i] = col
        return colList
    else:
        for col in range(len(findData[0])):
            if findData[0][col] is not None:
                _classify = findData[0][col]
            if classify == _classify and colName == findData[1][col]:
                return col
        else:
            return -1


def getDataColOrder(findData: list, colStr: str | list, checkRow=0):
    """在二维数组中根据列名获得列的位置，未找到时返回-1

    Args:
        findData (list): 二维数组
        colStr (str): 列名
        checkRow (int, optional): 列所在行. Defaults to 0.

    Returns:
        _type_: _description_
    """
    if isinstance(colStr, list):
        colList = [-1] * len(colStr)
        for i in range(len(colStr)):
            for order in range(len(findData[checkRow])):
                if colStr[i] == findData[checkRow][order]:
                    colList[i] = order
                    break
        return colList
    else:
        for order in range(len(findData[checkRow])):
            if colStr == findData[checkRow][order]:
                return order
        else:
            return -1


def getListData(ids: list, idCol: int, returnCol: list, findData: list, isNum=True):
    """根据id列表，获取表数据

    Args:
        ids (list): 玩家等级
        idCol (int): 属性表等级所在列
        returnCol (list): 返回的数据列名
        findData (list): 属性表数据
        isNum (bool): 返回的数据是否为数字

    Returns:
        list: 属性列表
    """
    returnNum = 1
    if isinstance(returnCol, list):
        returnNum = len(returnCol)
    returnList = []
    for i in range(len(ids)):
        if isNum:
            rList = [0] * returnNum
        else:
            rList = [None] * returnNum
        rList = getRowData(ids[i], idCol, returnCol, findData)
        returnList.append(rList)
    return returnList


def getRowData(id, idCol, returnCol, findData: list):
    """根据id查找数据

    Args:
        id (): 用于查询的id，可传入str/list
        idCol (): id所在列，可传入str/list
        returnCol (): 返回数据所在列，可传入str/list
        findData (list): 属性数据

    Returns:
        property: 属性值
    """
    returnNum = 1
    if isinstance(id, list):
        for row in findData:
            matchTag = True
            for i in range(len(id)):
                if id[i] != row[idCol[i]] and toStr(id[i]) != toStr(row[idCol[i]]):
                    matchTag = False
                    break
            if matchTag:
                if isinstance(returnCol, list):
                    returnNum = len(returnCol)
                    returnList = []
                    for col in returnCol:
                        returnList.append(row[col])
                    return returnList
                else:
                    return row[returnCol]
    else:
        for row in findData:
            if row[idCol] == id:
                if isinstance(returnCol, list):
                    returnNum = len(returnCol)
                    returnList = []
                    for col in returnCol:
                        returnList.append(row[col])
                    return returnList
                else:
                    return row[returnCol]
    if isinstance(returnCol, list):
        return [None] * returnNum
    else:
        return None


def getRangeRow(findRange, findValue):
    """在Range内查找目标所在行，未查到时返回插入位置

    Args:
        findRange (_type_): 查找的Range
        findValue (_type_): 查找的值

    Returns:
        (数据对应行位置，插入位置): _description_
    """
    findData = findRange.value
    insertRow = 1
    for i in range(len(findData)):
        if toStr(findData[i]) == toStr(findValue):
            return i, insertRow
        elif isNumber(findData[i]) and isNumber(findValue) and toNum(findData[i]) < toNum(findValue):
            insertRow = i + 2
    else:
        return int(-1), insertRow


def getListRow(findData, findValue, findCol=0):
    """在二维列表中查找数据所在行和插入位置

    Args:
        findList (_type_): 查找的二维数组
        findValue (_type_): 查找的id
        findCol (int, optional): id所在列. Defaults to 0.
    
    Returns:
        _type_: 数据所在行，插入行
    """
    insertRow = 1
    for i in range(len(findData)):
        checkValue = findData[i][findCol]
        if toStr(checkValue) == toStr(findValue):
            return i, insertRow
        elif isNumber(checkValue) and isNumber(findValue) and toNum(checkValue) < toNum(findValue):
            insertRow = i + 2
    else:
        return int(-1), insertRow


def toStr(content):
    """转为字符串，None会转为''

    Args:
        content (_type_): _description_

    Returns:
        _type_: _description_
    """
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


def toNum(content):
    """转为数值，非数值返回0

    Args:
        content (_type_): _description_

    Returns:
        _type_: _description_
    """
    if isNumber(content):
        return float(content)
    else:
        return 0


def toInt(content):
    """转为整数，四舍五入(处理excel数据莫名其妙变成很长小数的问题)，非数值返回-1

    Args:
        content (_type_): _description_

    Returns:
        _type_: _description_
    """
    if isNumber(content):
        return round(float(content))
    else:
        return -1


def isNumberValid(content, checkNum=0):
    """判断数字是否有效，>checkNum为有效

    Args:
        content (_type_): 支持字符串格式的数字
        checkNum (int, optional): 有效的条件. Defaults to 0.

    Returns:
        bool: _description_
    """
    if not isNumber(content):
        return False
    elif float(content) > checkNum:
        return True
    else:
        return False


def isNumber(content):
    """判断是否数字，None不是数字

    Args:
        content (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        float(content)
        return True
    except BaseException:
        return False


def isEmpty(content):
    """判断是否为空对象或空字符串

    Args:
        content (_type_): _description_

    Returns:
        _type_: _description_
    """
    if content == '' or content is None:
        return True
    else:
        return False


def split(content, sep):
    """重载分隔操作，空对象会转为空列表

    Args:
        content (_type_): 处理对象
        sep (_type_): 分隔符

    Returns:
        list[str]: _description_
    """
    if content is None:
        return []
    else:
        return str.split(toStr(content), sep)


def quitHideApp():
    """退出隐藏的app
    """
    for app in xw.apps:
        if app.visible is False:
            print('quit unVisible app')
            app.quit()


def getTablePath(newSelect=False):
    """获取操作的配置表目录

    Returns:
        str: 配置表目录路径
    """
    configPath = os.path.dirname(__file__) + r'\config.txt'
    if os.path.exists(configPath) and newSelect is False:
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
    """弹窗提示选择文件夹

    Args:
        _title (str, optional): 提示文字. Defaults to "选择文件夹".

    Returns:
        str: 文件夹路径
    """
    return filedialog.askdirectory(title=_title).replace('/', '\\')


def callBat(filePath, fileName):
    """调用bat文件

    Args:
        filePath (_type_): bat文件所在文件夹
        fileName (_type_): bat文件名
    """
    os.chdir(filePath)
    if fileName[-3:] != "bat":
        fileName = fileName + ".bat"

    os.system(r'start cmd.exe /k' + filePath + "\\" + fileName)
