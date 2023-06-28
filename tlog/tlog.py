import xlwings as xw
import csv

titleList = [
    'Field',
    'Type',
    'Null',
    'Key',
    'Default',
    'Comment',
]

keyList = ['DtEventTime', 'Vopenid', 'VRoleID', 'OpenID']

# 类型对应Collation
typeCollationMap = {
    'int': 'NULL',
    'bigint': 'NULL',
    'varchar': 'utf8mb4_general_ci',
    'datetime': 'NULL',
    'float': 'NULL',
    'array': 'NULL',
}


def addRowData(dataList, field, type, comment):
    rowMap = {
        'Field': field,
        'Type': type,
        'Null': 'NO',
        'Key': 'MUL' if field in keyList else '',
        'Default': 'NULL',
        'Comment': comment,
    }
    dataList.append(rowMap)


def main():
    wb = xw.books.active
    folderPath = wb.fullname[:str.find(wb.fullname, wb.name)] + "Tlog数据表结构\\"

    activeSht = wb.sheets.active
    activeFirstRange = activeSht.range('B2').end('down')
    activeList = activeFirstRange.expand('table').value
    dataList = []

    if activeSht.range('D1').value != '基础':
        publicSht = wb.sheets['0.公共字段']
        publicFirstRange = publicSht.range('B2').end('down')
        publicList = publicFirstRange.expand('table').value
        for row in publicList[1:]:
            if row[0] is not None and row[0] != '':
                addRowData(dataList, row[1], row[2], row[3])
    for row in activeList[1:]:
        if row[0] is not None and row[0] != '':
            addRowData(dataList, row[1], row[2], row[3])

    csvName = activeSht.range('B1').value
    csvDesc = activeSht.range('A1').value
    if csvName is not None and csvName != '':
        with open(folderPath + csvName + r'@@' + csvDesc + '.csv', mode='w', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, titleList)
            writer.writeheader()
            writer.writerows(dataList)
