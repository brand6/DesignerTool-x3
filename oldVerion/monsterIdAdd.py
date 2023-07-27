import xlwings as xw
from xlwings import Sheet
from common import common
import numpy as np
import os


def main():
    wbPath = xw.books.active.fullname
    programPath = wbPath[:wbPath.rfind(r'\SpawnCsv')]
    spawnPath = programPath + r'\SpawnCsv\stageSpawnPoint.csv'
    levPath = programPath + r'\Binaries\Tables\OriginTable\Battle\BattleLevel.xlsx'
    comPath = programPath + r'\Binaries\Tables\OriginTable\CommonStageEntry.xlsx'
    monsterPath = programPath + r'\Binaries\Tables\OriginTable\Battle\BattleMonster.xlsx'
    propPath = programPath + r'\Binaries\Tables\OriginTable\Battle\BattleMonsterProperty.xlsx'
    scriptPath = os.path.dirname(__file__)
    loc = scriptPath.find(r'\数值工具')
    calcPath = scriptPath[:loc] + r'\战斗数值\战斗数值模板-b.xlsm'

    spawnData = np.loadtxt(spawnPath, delimiter=',', dtype=str, encoding='utf-8-sig')
    spawnIdCol = 0
    spawnMonsterCol = 1
    spawnTemplateCol = 2
    spawnWaveCol = 3

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    levWb = app.books.open(levPath, update_links=False)
    levSht: Sheet = levWb.sheets['&BattleLevelConfig^']
    levData = np.array(levSht.used_range.value)
    levIdCol = common.getDataColOrder(levData, 'ID', 2)
    levMonsterCol = common.getDataColOrder(levData, 'MonsterUIDs', 2)
    levTemplateCol = common.getDataColOrder(levData, 'MonsterTemplateIDs', 2)
    levPropertyCol = common.getDataColOrder(levData, 'MonsterPropertyIDs', 2)
    levIds = levData[:, levIdCol]
    levMonsters = levData[:, levMonsterCol]
    levTemplates = levData[:, levTemplateCol]
    levPropertys = levData[:, levPropertyCol]

    comWb = app.books.open(comPath, update_links=False)
    comSht: Sheet = comWb.sheets['CommonStageEntry']
    comData = np.array(comSht.used_range.value)
    comIdCol = common.getDataColOrder(comData, 'ID', 2)
    comNameCol = common.getDataColOrder(comData, 'Name', 2)
    comShowCol = common.getDataColOrder(comData, 'MonsterForShow', 2)
    comIds = comData[:, comIdCol]
    comNames = comData[:, comNameCol]
    comShows = comData[:, comShowCol]

    monsterWb = app.books.open(monsterPath, update_links=False)
    monsterSht: Sheet = monsterWb.sheets['&MonsterTemplate^']
    monsterData = np.array(monsterSht.used_range.value)
    monsterIdCol = common.getDataColOrder(monsterData, 'ID', 2)
    monsterNameCol = common.getDataColOrder(monsterData, 'Name', 2)
    monsterTypeCol = common.getDataColOrder(monsterData, 'Type', 2)
    monsterShowCol = common.getDataColOrder(monsterData, 'ShowIndexNote', 2)
    monsterShieldCol = common.getDataColOrder(monsterData, 'ShieldMax', 2)
    monsterIds = monsterData[:, monsterIdCol]
    monsterShows = monsterData[:, monsterShowCol]
    monsterNames = monsterData[:, monsterNameCol]
    monsterTypes = monsterData[:, monsterTypeCol]
    monsterShields = monsterData[:, monsterShieldCol]

    propWb = app.books.open(propPath, update_links=False)
    propSht: Sheet = propWb.sheets['&MonsterProperty^']
    propData = np.array(propSht.used_range.value)
    propIdCol = common.getDataColOrder(propData, 'ID', 2)
    propNameCol = common.getDataColOrder(propData, 'Name', 2)
    propNoteCol = common.getDataColOrder(propData, 'Note', 2)
    propTypeCol = common.getDataColOrder(propData, 'NumType', 2)
    propTemplateCol = common.getDataColOrder(propData, 'TemplateID', 2)
    propShieldCol = common.getDataColOrder(propData, 'ShieldHpPara', 2)
    propUpdateColList = [propNameCol, propNoteCol, propTypeCol, propTemplateCol, propShieldCol]

    calcWb = app.books.open(calcPath, update_links=False)
    calcSht: Sheet = calcWb.sheets['关卡']
    calcData = np.array(calcSht.used_range.value)
    calcIdCol = common.getDataColOrder(calcData, '关卡id', 0)
    calcSkipCol = common.getDataColOrder(calcData, '跳过开关', 0)
    calcNameCol = common.getDataColOrder(calcData, '关卡名', 0)
    calcTemplateCol = common.getDataColOrder(calcData, '怪物模板ID', 0)
    calcWaveCol = common.getDataColOrder(calcData, '怪物组ID', 0)
    calcNumCol = common.getDataColOrder(calcData, '计算系数', 0)
    calcUpdateColList = [calcSkipCol, calcNameCol, calcTemplateCol, calcWaveCol, calcNumCol]

    for r in range(1, len(spawnData)):
        row = spawnData[r]
        stageId = int(row[spawnIdCol])
        monsterId = row[spawnMonsterCol]
        TemplateId = row[spawnTemplateCol]
        waveId = row[spawnWaveCol]
        templateMap = {}
        templateShowMap = {}

        comRow = common.getDataOrder(comIds, stageId)
        if comRow != -1:
            stageName = comNames[comRow]
        else:
            stageName = ''
            # print("commonStage表stageId 找不到：", stageId)
        # battleLevel表相关处理
        levRow = common.getDataOrder(levIds, stageId)
        if levRow != -1:
            levMonsters[levRow] = monsterId
            levTemplates[levRow] = TemplateId
            if TemplateId is not None and TemplateId != '':
                propertStr = ''
                templateList = TemplateId.split('|')
                for template in templateList:
                    if template in templateMap:
                        propertyId = templateMap[template]
                    else:
                        propertyId = stageId * 100 + len(templateMap) + 1
                        templateMap[template] = propertyId
                        # 更新属性表数据
                        monsterRow = common.getDataOrder(monsterIds, template)
                        if monsterRow != -1:
                            monsterName = monsterNames[monsterRow]
                            monsterType = monsterTypes[monsterRow]
                            monsterShield = monsterShields[monsterRow]
                            monsterShowTemplate = monsterShows[monsterRow] if int(monsterShows[monsterRow]) is not None else template  # yapf:disable
                            monsterShow = str(monsterShowTemplate) + '=' + str(propertyId)
                            if monsterType in templateShowMap:
                                templateShowMap[monsterType].append(monsterShow)
                            else:
                                templateShowMap[monsterType] = [monsterShow]
                        else:
                            print("monster表Template找不到：", template)
                            monsterName = ''
                            monsterType = 1
                            monsterShield = 0
                        propRow, insertRow = common.getListRow(propData, propertyId, propIdCol)
                        if propRow != -1:
                            propData[propRow][propNameCol] = monsterName
                            propData[propRow][propNoteCol] = stageName
                            propData[propRow][propTemplateCol] = template
                            propData[propRow][propTypeCol] = monsterType
                            propData[propRow][propShieldCol] = getShieldHpPara(monsterShield)
                        else:
                            # 保存数组的修改到表内，后面数据即将被重新赋值
                            updateShtData(propSht, propData, propUpdateColList)
                            if insertRow < 5:
                                insertRow = 4
                                propSht.range('4:4').insert('down')
                                propSht.range('5:5').copy(propSht.range('4:4'))
                                propSht.cells(4, propIdCol + 1).value = propertyId
                            else:
                                propSht.range(str(insertRow) + ':' + str(insertRow)).insert('down')
                                propSht.range(str(insertRow - 1) + ':' + str(insertRow - 1)).copy(propSht.range(str(insertRow) + ':' + str(insertRow)))  # yapf:disable
                                propSht.cells(insertRow, propIdCol + 1).value = propertyId

                            propData = np.array(propSht.used_range.value)
                            propData[insertRow - 1][propNameCol] = monsterName
                            propData[insertRow - 1][propNoteCol] = stageName
                            propData[insertRow - 1][propTemplateCol] = template
                            propData[insertRow - 1][propTypeCol] = monsterType
                            propData[insertRow - 1][propShieldCol] = getShieldHpPara(monsterShield)

                    propertStr = combineIds(propertStr, propertyId)
                levPropertys[levRow] = propertStr
            else:
                print("stage包含怪物列表为空：", stageId)
        else:
            print("battlelevel表stageId 找不到：", stageId)

        # commonStage表相关处理
        """显示优先级boss》精英》小怪，最多显示3个怪，超出数量时去掉优先级低的，顺序随意
        """
        if len(templateShowMap) > 0 and comRow != -1:
            showStr = ''
            showNum = 0
            for i in range(3, 0, -1):
                if i in templateShowMap and len(templateShowMap[i]) + showNum < 4:
                    showNum += len(templateShowMap[i])
                    for show in templateShowMap[i]:
                        showStr = combineIds(showStr, show)

            comShows[comRow] = showStr

        # 数据写入战斗数值模板
        calcRow, insertRow = common.getListRow(calcData, stageId, calcIdCol)
        if calcRow != -1:
            calcData[calcRow][calcNameCol] = stageName
            calcData[calcRow][calcTemplateCol] = TemplateId
            calcData[calcRow][calcWaveCol] = waveId
        else:
            # 保存数组的修改到表内，后面数据即将被重新赋值
            updateShtData(calcSht, calcData, calcUpdateColList)
            if insertRow < 3:
                insertRow = 2
                calcSht.range('2:2').insert('down')
                calcSht.range('3:3').copy(calcSht.range('2:2'))
                calcSht.cells(2, calcIdCol + 1).value = stageId
            else:
                calcSht.range(str(insertRow) + ':' + str(insertRow)).insert('down')
                calcSht.range(str(insertRow - 1) + ':' + str(insertRow - 1)).copy(calcSht.range(str(insertRow) + ':' + str(insertRow)))  # yapf:disable
                calcSht.cells(insertRow, calcIdCol + 1).value = stageId

            calcData = np.array(calcSht.used_range.value)
            calcData[insertRow - 1][calcSkipCol] = 1
            calcData[insertRow - 1][calcNameCol] = stageName
            calcData[insertRow - 1][calcTemplateCol] = TemplateId
            calcData[insertRow - 1][calcWaveCol] = waveId
            calcData[insertRow - 1][calcNumCol] = 1

    print("数据处理完毕，正在保存数据...")
    levSht.cells(1, levMonsterCol + 1).options(transpose=True).value = levMonsters
    levSht.cells(1, levTemplateCol + 1).options(transpose=True).value = levTemplates
    levSht.cells(1, levPropertyCol + 1).options(transpose=True).value = levPropertys
    levWb.save()
    levWb.close()

    comSht.cells(1, comShowCol + 1).options(transpose=True).value = comShows
    comWb.save()
    comWb.close()

    updateShtData(propSht, propData, propUpdateColList)
    propWb.save()
    propWb.close()

    updateShtData(calcSht, calcData, calcUpdateColList)
    calcWb.save()
    calcWb.close()

    monsterWb.close()
    app.quit()


def updateShtData(dataSht: Sheet, dataList, colList: list):
    """更新property表的数据列

    Args:
        propSht (Sheet): sheet
        propData (np.array): data
        colList (list): 更新的列，传入list
    """
    for col in colList:
        dataSht.cells(1, col + 1).options(transpose=True).value = dataList[:, col]


def combineIds(str1, id):
    if str1 == '':
        return common.toStr(id)
    else:
        return common.toStr(id) + '|' + str1


def getShieldHpPara(shieldNum):
    """根据芯核数计算血量系数

    Args:
        shieldNum (_type_): 芯核数

    Returns:
        _type_: 血量系数
    """
    if shieldNum == 0:
        return 1
    elif shieldNum == 1:
        return 1.2
    elif shieldNum == 2:
        return 1.2
    elif shieldNum == 4:
        return 1.2
    elif shieldNum == 6:
        return 1.2
    else:
        return 1
