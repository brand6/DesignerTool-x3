import xlwings as xw
from xlwings import Sheet
from common import common
import numpy as np


def main():
    dataList = xw.sheets.active.used_range.value

    programPath = common.getTablePath()
    hitPath = programPath + r'\Battle\BattleHitParam.xlsx'
    buffPath = programPath + r'\Battle\BattleBuff.xlsx'
    summonPath = programPath + r'\Battle\BattleSummon.xlsx'

    hitWb = xw.books.open(hitPath, update_links=False)
    buffWb = xw.books.open(buffPath, update_links=False)
    summonWb = xw.books.open(summonPath, update_links=False)

    hitSht: Sheet = hitWb.sheets['&HitParamConfig']
    buffSht: Sheet = buffWb.sheets['&BuffLevelConfig^']
    summonSht: Sheet = summonWb.sheets['&BattleSummon']

    hitList = np.array(hitSht.used_range.value)
    buffList = np.array(buffSht.used_range.value)
    summonList = np.array(summonSht.used_range.value)

    hitKeyCol = [common.getDataColOrder(dataList, 'H#HitParamID', 0), common.getDataColOrder(hitList, 'HitParamID', 2)]
    buffKeyCol = [common.getDataColOrder(dataList, 'B#ID', 0), common.getDataColOrder(buffList, 'ID', 2)]
    summonKeyCol = [common.getDataColOrder(dataList, 'S#Template', 0), common.getDataColOrder(summonList, 'Template', 2)]

    hitCols = []
    buffCols = []
    summonCols = []

    for i in range(len(dataList[0])):
        if dataList[0][i] is not None and '#' in dataList[0][i]:
            col = common.split(dataList[0][i], '#')[1]
            if 'H#' in dataList[0][i]:
                hitCols.append([common.getDataColOrder(dataList, dataList[0][i], 0), common.getDataColOrder(hitList, col, 2)])
            elif 'B#' in dataList[0][i]:
                buffCols.append([common.getDataColOrder(dataList, dataList[0][i], 0), common.getDataColOrder(buffList, col, 2)])
            elif 'S#' in dataList[0][i]:
                summonCols.append(
                    [common.getDataColOrder(dataList, dataList[0][i], 0),
                     common.getDataColOrder(summonList, col, 2)])

    for i in range(1, len(dataList)):
        if dataList[i][hitKeyCol[0]] is not None:
            copyData(dataList, hitList, i, hitKeyCol, hitCols)
        if dataList[i][buffKeyCol[0]] is not None:
            copyData(dataList, buffList, i, buffKeyCol, buffCols)
        if dataList[i][summonKeyCol[0]] is not None:
            copyData(dataList, summonList, i, summonKeyCol, summonCols)

    for col in hitCols:
        hitSht.cells(1, col[1] + 1).options(transpose=True).value = hitList[:, col[1]]
    for col in buffCols:
        buffSht.cells(1, col[1] + 1).options(transpose=True).value = buffList[:, col[1]]
    for col in summonCols:
        summonSht.cells(1, col[1] + 1).options(transpose=True).value = summonList[:, col[1]]

    hitWb.save()
    hitWb.close()
    buffWb.save()
    buffWb.close()
    summonWb.save()
    summonWb.close()


def copyData(dataList, targetList, row, keyCol, cols):
    findRow = common.getListRow(targetList, dataList[row][keyCol[0]], keyCol[1])[0]
    if findRow != -1:
        for col in cols:
            targetList[findRow][col[1]] = dataList[row][col[0]]
