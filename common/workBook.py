import xlwings as xl


class WorkBook(xl.Book):

    def __init__(self, path, name, app, defSht=0):
        if name != "":
            if name[-5:] != ".xlsx":
                name = name + ".xlsx"
            else:
                name = name

            WorkBook.closeByName(name)
            impl = app.books.open(path + "\\" + name, 0)
        else:
            impl = app.books.open(path, 0)
        self.impl = impl

        if defSht == -1:
            self.sheet = xl.sheets.active
        else:
            self.sheet = self.sheets[defSht]

        self.range = self.sheet.range
        self.cells = self.sheet.cells

    # 指定操作的sheet
    def setSht(self, defSht):
        self.sheet = self.sheets[defSht]

    # 根据id插入行
    def insertRowById(self, id, idCol):
        idCol = idCol - 1
        dataValue = self.sheet.used_range.raw_value
        for row in range(len(dataValue) - 1, -1, -1):
            if dataValue[row][idCol] < id:
                self.insertRow(row + 2)
                self.cells(row + 2, idCol + 1).raw_value = id
                return row + 2

    # 根据id查找行
    def findRowById(self, id, idCol):
        idCol = idCol - 1
        dataValue = self.sheet.used_range.raw_value
        for row in range(len(dataValue) - 1, -1, -1):
            if dataValue[row][idCol] == id:
                return row + 1
        else:
            return -1

    # 删除行
    def deleteRow(self, firstRow, lastRow=None):
        if lastRow is None:
            lastRow = firstRow
        delRows = str(int(firstRow)) + ':' + str(int(lastRow))
        self.sheet[delRows].delete()

    # 删除列
    def deleteCol(self, firstCol, lastCol=None):
        if lastCol is None:
            lastCol = firstCol

        delCols = self.getColumnAddress(firstCol) + ':' + self.getColumnAddress(lastCol)
        self.sheet[delCols].delete()

    # 插入行
    def insertRow(self, firstRow, lastRow=None):
        """_summary_

        Args:
            firstRow (_type_): xl中的行号，插入后该行原来的数据后移
            lastRow (_type_, optional): _description_. Defaults to None.
        """
        if lastRow is None:
            lastRow = firstRow
        insertStr = str(int(firstRow)) + ':' + str(int(lastRow))
        self.sheet[insertStr].insert()

    # 插入列
    def insertCol(self, firstCol, lastCol=None):
        if lastCol is None:
            lastCol = firstCol
        insertStr = self.getColumnAddress(firstCol) + ':' + self.getColumnAddress(lastCol)
        self.sheet[insertStr].insert()

    # 获得列名
    def getColumnAddress(self, colNum):
        colAddress = self.cells(1, colNum).get_address(True, False)
        colName = colAddress.split('$')[0]
        return colName

    # 获取表格指定数据
    def getRangeData(self, rangeStr, expendStr=''):
        sht = self.sheet
        if expendStr == '':
            return sht.range(rangeStr)
        else:
            return sht.range(rangeStr).expand(expendStr)

    # 获取表格全部数据
    def getData(self):
        return self.sheet.used_range.value

    # 写入表格数据
    def setData(self, dataValue):
        if self.sheet.api.FilterMode is True:
            self.sheet.api.ShowAllData()
        self.sheet.range('A1').value = dataValue

    # 关闭workbook
    def close(self, isSave=False):
        if isSave is True:
            self.save()
        if self.impl is not None:
            self.impl.close()

    # 获取表格指定数据
    @classmethod
    def getRangeByShtName(cls, shtName, rangeStr='', expendStr='', wb=''):  # yapf:disable
        if wb == '':
            wb = cls.getActiveWb()
        sht = wb.sheets(shtName)
        if expendStr == '':
            if rangeStr != '':
                return sht.range(rangeStr)
            else:
                return sht.used_range
        else:
            return sht.range(rangeStr).expand(expendStr)

    # 获取当前激活的表格
    @classmethod
    def getActiveSht(cls, wb=None):
        if wb is None:
            return xl.sheets.active
        else:
            return wb.sheets.active

    # 获取当前激活的工作簿
    @classmethod
    def getActiveWb(cls, app=None):
        if app is None:
            return xl.books.active
        else:
            return app.books.active

    # 通过名字关闭workbook
    @classmethod
    def closeByName(cls, wbName):
        if wbName[-5:] != ".xlsx":
            wbName = wbName + ".xlsx"

        for app in xl.apps:
            for wb in app.books:
                if wb.name == wbName:
                    wb.close()
