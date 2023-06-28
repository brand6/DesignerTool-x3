import xlwings as xw
import numpy as np
import os
from common import common
import csv


def main():
    configPath = common.getTablePath(True)
    curPath = os.path.abspath(__file__)
    loc = curPath.index('PythonScript')
    csvPath = curPath[:loc] + 'Tlog\\Tlog维度表\\'

    for ap in xw.apps:
        if ap.visible is False:
            ap.quit()
    app = xw.App(visible=False, add_book=False)

    # 处理Reason
    reasonWb = app.books.open(configPath + '\\Reason.xlsx', read_only=True)
    reasonSht = reasonWb.sheets['Reason']
    reasonData = np.array(reasonSht.used_range.value)
    reasonIdCol = common.getDataColOrder(reasonData, 'Id', 2)
    reasonDescCol = common.getDataColOrder(reasonData, 'FormatDesc', 2)
    reasonIds = reasonData[3:, reasonIdCol]
    reasonDescs = reasonData[3:, reasonDescCol]

    titleList = ['Reason', 'content']
    dataList = []
    for i in range(len(reasonIds)):
        rowMap = {
            'Reason': int(reasonIds[i]),
            'content': reasonDescs[i],
        }
        dataList.append(rowMap)
    writeToCsv(csvPath, 'MoneyFlow-Reason', titleList, dataList)
    dataList.clear()

    # 处理item
    itemWb = app.books.open(configPath + '\\Item.xlsx', read_only=True)
    itemSht = itemWb.sheets['Item']
    itemData = np.array(itemSht.used_range.value)
    itemIdCol = common.getDataColOrder(itemData, 'ID', 2)
    itemNameCol = common.getDataColOrder(itemData, 'Name', 2)
    itemIds = itemData[3:, itemIdCol]
    itemNames = itemData[3:, itemNameCol]

    titleList = ['TargetParam', 'content']
    dataList2 = []
    for i in range(len(itemIds)):
        rowMap = {
            'TargetParam': int(itemIds[i]),
            'content': nameReplace(itemNames[i]),
        }
        dataList2.append(rowMap)
    writeToCsv(csvPath, 'MoneyFlow-TargetParam', titleList, dataList2)
    dataList2.clear()

    app.quit()
    input("脚本执行完毕")


def nameReplace(itemName):
    replaceMap = {
        '{Role1}': '沈星回',
        '{Role2}': '黎深',
        '{Role3}': '',
        '{Role4}': '',
        '{Role5}': '祁煜',
    }

    for kStr in replaceMap.keys():
        if kStr in itemName:
            itemName = itemName.replace(kStr, replaceMap[kStr])
            break
    return itemName


def writeToCsv(folderPath, csvName, titleList, dataList):
    with open(folderPath + csvName + '.csv', mode='w', encoding='utf-8', newline='') as f:
        writer = csv.DictWriter(f, titleList)
        writer.writeheader()
        writer.writerows(dataList)
