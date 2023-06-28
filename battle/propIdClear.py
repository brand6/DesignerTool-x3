import xlwings as xw
from xlwings import Sheet
from common import common
import numpy as np


def main():
    wbPath = xw.books.active.fullname
    programPath = wbPath[:wbPath.rfind(r'\SpawnCsv')]
    levPath = programPath + r'\Binaries\Tables\OriginTable\Battle\BattleLevel.xlsx'
    propPath = programPath + r'\Binaries\Tables\OriginTable\Battle\BattleMonsterProperty.xlsx'

    levWb = xw.books.open(levPath, update_links=False)
    levSht: Sheet = levWb.sheets['&BattleLevelConfig^']
    levData = np.array(levSht.used_range.value)
    levIdCol = common.getDataColOrder(levData, 'ID', 2)

    propWb = xw.books.open(propPath, update_links=False)
    propSht: Sheet = propWb.sheets['&MonsterProperty^']
    propData = np.array(propSht.used_range.value)
    propIdCol = common.getDataColOrder(propData, 'ID', 2)

    dealLevelId = 0
    startRow = 0
    endRow = 0
    for r in range(len(propData) - 1, 2, -1):
        if propData[r][propIdCol] is not None:
            levelId = int(propData[r][propIdCol] / 100)
            if dealLevelId == 0:  # 第一行数据
                dealLevelId = levelId
                startRow = r + 1
                endRow = r + 1
            elif dealLevelId == levelId:  # 相同的关卡
                endRow = r + 1
            else:  # 不同的关卡
                # 结算上一关的数据
                checkLevelValid(levData, levIdCol, dealLevelId, propSht, startRow, endRow)

                # 新关卡的数据
                dealLevelId = levelId
                startRow = r + 1
                endRow = r + 1

    levWb.close()
    propWb.save()
    propWb.close()


def checkLevelValid(levData, levIdCol, levelId, propSht, startRow, endRow):
    """检测关卡是否有效，无效则删除对应数值id所在的行

    Args:
        levData (_type_): _description_
        levIdCol (_type_): _description_
        levelId (_type_): _description_
        propSht (_type_): _description_
        startRow (_type_): _description_
        endRow (_type_): _description_
    """
    if levelId > 0:
        levelRow, _ = common.getListRow(levData, levelId, levIdCol)
        if levelRow == -1:
            deleteRows = str(endRow) + ':' + str(startRow)
            propSht[deleteRows].delete('up')
            print('关卡', levelId, '不存在，删除对应数值id')
