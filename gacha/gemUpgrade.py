import random

import xlwings as xw

from common import common


def main():
    sht = xw.sheets.active
    dataList = sht.used_range.value
    maxRow = sht.used_range.last_cell.row

    subPropNameCol = common.getColBy2Para('副属性', '属性类型', dataList)
    subPropWeightCol = common.getColBy2Para('副属性', '权重', dataList)
    subPropsExpectRateCol = common.getColBy2Para('副属性升级', '预期次数', dataList)
    subPropsNumCol = common.getColBy2Para('副属性升级', '属性数量', dataList)
    subPropsCol = common.getColBy2Para('副属性升级', ['副属性1', '副属性2', '副属性3', '副属性4'], dataList)

    subPropNameList = []
    subPropWeightList = []
    subPropsList = []
    subPropsExpectRateList = []

    for i in range(2, maxRow):
        if dataList[i][subPropNameCol] is not None:
            subPropNameList.append(dataList[i][subPropNameCol])
            subPropWeightList.append(dataList[i][subPropWeightCol])

        if dataList[i][subPropsCol[0]] is not None:
            subPropsExpectRateList.append(0)
            tempList = []
            for j in range(4):
                if dataList[i][subPropsCol[j]] is not None:
                    tempList.append(dataList[i][subPropsCol[j]])
            subPropsList.append(tempList)

    tryTimes = 100000
    for _ in range(tryTimes):
        hasList = []
        times = random.randint(3, 4)  # 芯核初始副属性数量
        for __ in range(times):
            totalWeight = 0
            for m in range(len(subPropNameList)):  # 计算不放回抽取的总权重
                if subPropNameList[m] not in hasList:
                    totalWeight += subPropWeightList[m]

            rndNum = random.randrange(0, totalWeight)
            for m in range(len(subPropNameList)):
                if subPropNameList[m] not in hasList:
                    if rndNum < subPropWeightList[m]:
                        hasList.append(subPropNameList[m])
                        break
                    else:
                        rndNum -= subPropWeightList[m]

        for j in range(len(subPropsExpectRateList)):  # 是否满足属性组合
            for m in subPropsList[j]:
                if m not in hasList:
                    break
            else:
                subPropsExpectRateList[j] += 1

    # 输出结果
    for j in range(len(subPropsExpectRateList)):
        if subPropsExpectRateList[j] > 0:
            subPropsExpectRateList[j] = round(tryTimes / subPropsExpectRateList[j], 1)
    sht.cells(3, subPropsExpectRateCol + 1).options(transpose=True).value = subPropsExpectRateList
