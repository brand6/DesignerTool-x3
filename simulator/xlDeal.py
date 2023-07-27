import os
import sys
from tkinter import filedialog

import xlwings as xw
from commFun import Common
from pandas import DataFrame
from xlwings import Book, Sheet


class XlDeal:
    app: xw.App = None

    def __init__(self, _wbName: str, _shtName: str, path="", isReadOnly=True) -> None:
        self.wbName = _wbName
        self.shtName = _shtName
        self.isReadOnly = isReadOnly
        self.wb: Book = self.OpenBook(_wbName, path)
        self.sht: Sheet = self.wb.sheets[_shtName]
        self.data: list = self.sht.used_range.value
        self.pdData: DataFrame = None

    def GetColIndex(self, colName: str | list, startCheckRow: int = 0, startCol=0, endCol=-1) -> int:
        """根据列名查找列的索引，未找到时返回-1

        Args:
            colName (str | list): 传入list时，根据多个条件查找.\n
            startCheckRow (int, optional): 第一个列名所在的行. Defaults to 0.\n
            startCol:开始的列（包含本列）\n
            EndCol:结束的列（不包含本列）\n
        Returns:
            int: 列在数据中的索引
        """
        if endCol == -1:
            endCol = len(self.data[0])

        if isinstance(colName, list):
            checkRowNum = len(colName)
        else:
            colName = [colName]
            checkRowNum = 1

        compareList = [None] * checkRowNum
        for c in range(startCol, endCol):
            for r in range(startCheckRow, startCheckRow + checkRowNum):
                k = r - startCheckRow
                if self.data[r][c] is not None and self.data[r][c] != compareList[k]:
                    compareList[k] = self.data[r][c]  # 不为空时更新compareList数据
                if compareList[k] != colName[k]:
                    break  # 比较是否满足条件，不满足则比对下一列
            else:
                return c
        else:
            return -1

    def GetColData(self, col: str | list, startCheckRow: int = 0, startCol=0, endCol=-1):
        """根据条件查找列的数据

        Args:
            col (str | list): 传入list时，根据多个条件查找.\n
            startCheckRow (int, optional): 第一个列名所在的行. Defaults to 0.
        Returns:
            _type_: _description_
        """
        if isinstance(col, list) or isinstance(col, str):
            col = self.GetColIndex(col, startCheckRow, startCol, endCol)
        if col != -1:
            if self.pdData is None:
                self.pdData = DataFrame(self.data)
            return self.pdData.iloc[:, col]
        else:
            return -1

    def GetRowIndex(self, checkValue, checkCol=0, hasStrNum=False, startRow=0, endRow=-1) -> int:
        """根据条件查找行的索引，未找到时返回-1

        Args:
            checkValue (any|list): 查找的条件，可以是list\n
            checkCol (any, optional):查找的列，[列名|列索引]. Defaults to 0.\n
            hasStrNum (bool, optional): 是否可能有字符串类型的数字. Defaults to False.

        Returns:
            int: 在数据中的索引行
        """
        if endRow == -1:
            endRow = len(self.data)

        if isinstance(checkValue, list):
            checkColNum = len(checkValue)
        else:
            checkValue = [checkValue]
            checkCol = [checkCol]
            checkColNum = 1

        for c in range(len(checkCol)):
            col = Common.toInt(checkCol[c])
            if col == -1:
                checkCol[c] = self.GetColIndex(checkCol[c])
            else:
                checkCol[c] = col

        compareList = [None] * checkColNum
        for r in range(startRow, endRow):
            for i in range(len(checkCol)):
                c = checkCol[i]
                if self.data[r][c] is not None and self.data[r][c] != compareList[i]:
                    compareList[i] = self.data[r][c]  # 不为空时更新compareList数据
                if hasStrNum is True:
                    if Common.toStr(compareList[i]) != Common.toStr(checkValue[i]):
                        break  # 比较是否满足条件，不满足则比对下一行
                else:
                    if compareList[i] != checkValue[i]:
                        break  # 比较是否满足条件，不满足则比对下一行
            else:
                return r
        else:
            return -1

    def GetRowData(self, checkValue, returnCol, checkCol=0, hasStrNum=False, startRow=0, endRow=-1):
        """根据条件查找行的数据

        Args:
            checkValue (any|list): 查找的条件，可以是list\n
            returnCol (any|list): 返回的列，可以是list[列名|列索引]\n
            checkCol (any, optional):查找的列，[列名|列索引]. Defaults to 0.\n
            hasStrNum (bool, optional): 是否可能有字符串类型的数字. Defaults to False.\n

        Returns:
            _type_: _description_
        """
        row = self.GetRowIndex(checkValue, checkCol, hasStrNum, startRow, endRow)
        if isinstance(returnCol, list):
            returnList = []
            if row != -1:
                for c in range(len(returnCol)):
                    col = Common.toInt(returnCol[c])
                    if col == -1:
                        col = self.GetColIndex(returnCol[c])
                    returnList.append(self.data[row][col])
                return returnList
            else:
                return [-1] * len(returnCol)
        else:
            if row != -1:
                col = Common.toInt(returnCol)
                if col == -1:
                    col = self.GetColIndex(returnCol)
                return self.data[row][col]
            else:
                return -1

    def UpdateTableData(self, tableData: list, startRow=0, startCol=0, isClear=False) -> None:
        """更新表格数据

        Args:
            startRow (int): 更新起始的行index
            startCol (_type_): 更新起始的列index
            tableData (list): 更新的数据
        """
        if isClear is True:
            self.sht.clear_contents()
        startCol = Common.toInt(startCol)
        if startCol == -1:
            startCol = self.GetColIndex(startCol)
        self.sht.cells(startRow + 1, startCol + 1).value = tableData
        if self.data is not None:
            self.data.clear()
        self.data = self.sht.used_range.value
        if self.pdData is not None:
            del self.pdData
            self.pdData = DataFrame(self.data)

    def UpdateColData(self, colData: list, col) -> None:
        """更新列数据

        Args:
            col (_type_): 列名|列index
            colData (list): 列数据
        """
        col = Common.toInt(col)
        if col == -1:
            col = self.GetColIndex(col)
        self.sht.cells(1, col + 1).options(transpose=True).value = colData
        self.data.clear()
        self.data = self.sht.used_range.value
        if self.pdData is not None:
            del self.pdData
            self.pdData = DataFrame(self.data)

    def GetNotEmptyCol(self, startCol: int, checkRow: int) -> int:
        """获取从当前列开始往后数，第一个非空的列index，找不到时返回最后一列

        Args:
            startCol (int): 开始检测的列\n
            checkRow (int): 检测是否为空的行
        Returns:
            int: 第一个非空列的index
        """
        for col in range(startCol, len(self.data[checkRow])):
            if self.data[checkRow][col] is not None:
                return col
        else:
            return col

    def CloseBook(self, isSave=False) -> None:
        """关闭Excel

        Args:
            isSave (bool): 是否需要保存. Defaults to False.
        """
        if isSave is True:
            self.wb.save()
        self.wb.close()

    def OpenBook(self, _wbName: str, path="") -> Book:
        """打开Excel

        Args:
            _wbName (str): Excel的名字
            path(str):文件路径

        Returns:
            Book: _description_
        """
        if path == "":
            path = self.GetTablePath()
        if len(xw.apps) > 0:
            for wb in xw.books:
                if wb.name == _wbName:
                    return xw.books[_wbName]
            else:
                if XlDeal.app is not None:
                    for wb in XlDeal.app.books:
                        if wb.name == _wbName:
                            return XlDeal.app.books[_wbName]
                    else:
                        print("打开表格", _wbName)
                        return XlDeal.app.books.open(r"" + path + _wbName, update_links=False, read_only=self.isReadOnly)
                else:
                    print("打开表格", _wbName)
                    return xw.books.open(r"" + path + _wbName, update_links=False, read_only=self.isReadOnly)
        else:
            isDebug = True if sys.gettrace() else False
            app = xw.App(visible=isDebug, add_book=False)
            return app.books.open(r"" + path + _wbName, update_links=False, read_only=self.isReadOnly)

    @classmethod
    def GetTablePath(cls, newSelect=False) -> str:
        """获取操作的配置表目录

        Args:
            newSelect (bool, optional): 是否强制重选路径. Defaults to False.
        Returns:
            str: 配置表目录路径
        """

        configPath = os.path.dirname(__file__) + r"\config.txt"
        if os.path.exists(configPath) and newSelect is False:
            with open(configPath, "r", encoding="GBK") as file:
                tablePath = file.read()
                if os.path.exists(tablePath):
                    return tablePath

        with open(configPath, "w", encoding="GBK") as file:
            tablePath = cls.SelectDirectory("选择项目Program文件夹")
            loc = tablePath.index("Program")

            tablePath = tablePath[:loc] + "Program\\Binaries\\Tables\\OriginTable\\"
            file.write(tablePath)
            return tablePath

    @classmethod
    def SelectDirectory(cls, _title="选择文件夹") -> str:
        """弹窗提示选择文件夹

        Args:
            _title (str, optional): 提示文字. Defaults to "选择文件夹".
        Returns:
            str: 文件夹路径
        """
        return filedialog.askdirectory(title=_title).replace("/", "\\")
        return filedialog.askdirectory(title=_title).replace("/", "\\")
