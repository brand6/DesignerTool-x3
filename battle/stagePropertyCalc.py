import xlwings as xw
from xlwings import Sheet
from common import common
from common import printer
from battle import monsterIdSync
import numpy as np


def main():
    myPrinter = printer.Printer()
    myPrinter.printColor('~~~开始处理数据~~~', 'green')

    wb = xw.books.active
    roundSht: Sheet = wb.sheets['关卡']
    roundData = np.array(roundSht.used_range.value)
    roundIdCol = common.getDataColOrder(roundData, '关卡id', 0)
    roundSkipCol = common.getDataColOrder(roundData, '跳过开关', 0)
    roundMonsterCol = common.getDataColOrder(roundData, '怪物模板ID', 0)
    roundPropertyCol = common.getDataColOrder(roundData, '怪物数值ID', 0)
    roundWaveCol = common.getDataColOrder(roundData, '怪物组ID', 0)
    roundLevelCol = common.getDataColOrder(roundData, '关卡等级', 0)
    roundHpCol = common.getDataColOrder(roundData, '血量系数', 0)
    roundAtkCol = common.getDataColOrder(roundData, '攻击系数', 0)
    roundWaveHpCol = common.getDataColOrder(roundData, '每组总血量', 0)
    roundMonsterHpCol = common.getDataColOrder(roundData, '怪物血量详情', 0)
    roundWaveHpList = roundData[:, roundWaveHpCol]
    roundMonsterHpList = roundData[:, roundMonsterHpCol]

    tablePath = common.getTablePath()
    monsterWb = xw.books.open(tablePath + r'\Battle\BattleMonster.xlsx', 0)
    monsterSht = monsterWb.sheets['&MonsterTemplate^']
    monsterData = monsterSht.used_range.value
    monsterIDCol = common.getDataColOrder(monsterData, 'ID', 2)
    monsterTypeCol = common.getDataColOrder(monsterData, 'Type', 2)
    monsterCalcTypeCol = common.getDataColOrder(monsterData, 'BattleParadigm', 2)
    monsterTimeCol = common.getDataColOrder(monsterData, 'BattleTimes', 2)
    monsterEquipShieldCol = common.getDataColOrder(monsterData, 'EquipShield', 2)
    monsterShieldCol = common.getDataColOrder(monsterData, 'ShieldMax', 2)

    propertyWb = xw.books.open(tablePath + r'\Battle\BattleMonsterProperty.xlsx', 0)
    propertySht = propertyWb.sheets['&MonsterProperty^']
    propertyData = np.array(propertySht.used_range.value)
    propertyIDCol = common.getDataColOrder(propertyData, 'ID', 2)
    propertyTypeCol = common.getDataColOrder(propertyData, 'NumType', 2)
    propertyShieldCol = common.getDataColOrder(propertyData, 'ShieldHpPara', 2)
    propertyLevCol = common.getDataColOrder(propertyData, 'Level', 2)
    propertyHpCol = common.getDataColOrder(propertyData, 'MaxHPRate', 2)
    propertyAtkCol = common.getDataColOrder(propertyData, 'PhyAttackRate', 2)
    propertyTypeList = propertyData[:, propertyTypeCol]
    propertyShieldList = propertyData[:, propertyShieldCol]
    propertyLevList = propertyData[:, propertyLevCol]
    propertyHpList = propertyData[:, propertyHpCol]
    propertyAtkList = propertyData[:, propertyAtkCol]

    levelWb = xw.books.open(tablePath + r'\Battle\BattleLevel.xlsx', 0)
    levelSht = levelWb.sheets['&BattleLevelConfig^']
    levelData = np.array(levelSht.used_range.value)
    levelIDCol = common.getDataColOrder(levelData, 'ID', 2)
    levelPropertyCol = common.getDataColOrder(levelData, 'OfflineHeroPropertyIDs', 2)
    levelPropertyList = levelData[:, levelPropertyCol]

    monsterMap = {}
    calcTypeMap = {'女主普攻': 1, '双人全技能': 2.88}
    for row in range(1, len(roundData)):
        if roundData[row][roundSkipCol] is None:
            lRow = common.getListRow(levelData, roundData[row][roundIdCol], levelIDCol)[0]
            roundLev = roundData[row][roundLevelCol]
            levelPropertyList[lRow] = common.toStr(2000 + roundLev) + '|' + common.toStr(1000 + roundLev)

            monsterList = common.split(roundData[row][roundMonsterCol], '|')
            propertyList = common.split(roundData[row][roundPropertyCol], '|')
            waveList = common.split(roundData[row][roundWaveCol], '|')
            waveTimeMap = {}
            propertyMap = {}
            monsterTimeStr = ''
            waveTimeStr = ''
            for i in range(len(monsterList)):
                if monsterList[i] in monsterMap:
                    mRow = monsterMap[monsterList[i]]
                else:
                    mRow = common.getListRow(monsterData, monsterList[i], monsterIDCol)[0]
                    monsterMap[monsterList[i]] = mRow
                if monsterData[mRow][monsterTimeCol] is not None:
                    calcType = monsterData[mRow][monsterCalcTypeCol]
                    timePara = calcTypeMap['双人全技能']
                    if calcType in calcTypeMap:
                        timePara = calcTypeMap[calcType]
                    monsterTime = round(monsterData[mRow][monsterTimeCol] * timePara / calcTypeMap['双人全技能'], 1)
                    if monsterTimeStr == '':
                        monsterTimeStr = common.toStr(monsterTime)
                    else:
                        monsterTimeStr = monsterTimeStr + '|' + common.toStr(monsterTime)
                    if waveList[i] in waveTimeMap:
                        waveTimeMap[waveList[i]] = round(monsterTime + waveTimeMap[waveList[i]], 1)
                    else:
                        waveTimeMap[waveList[i]] = monsterTime
                    if propertyList[i] not in propertyMap:
                        monsterType = monsterData[mRow][monsterTypeCol]
                        monsterShield = monsterData[mRow][monsterShieldCol]
                        monsterEquipShield = monsterData[mRow][monsterEquipShieldCol]

                        pRow = common.getListRow(propertyData, propertyList[i], propertyIDCol)[0]
                        if pRow == -1:
                            print(propertyList[i], 'BattleMonsterProperty表中不存在该id')
                        else:
                            propertyTypeList[pRow] = monsterIdSync.getMonsterType(monsterType)
                            propertyShieldList[pRow] = monsterIdSync.getShieldHpPara(monsterEquipShield, monsterShield)
                            propertyLevList[pRow] = roundLev
                            if roundData[row][roundHpCol] is not None:
                                propertyHpList[pRow] = int(10 * roundData[row][roundHpCol] * monsterTime *
                                                           propertyShieldList[pRow])
                            if roundData[row][roundAtkCol] is not None:
                                propertyAtkList[pRow] = 1000 * roundData[row][roundAtkCol]
                else:
                    print(monsterList[i], '未填写战斗时长')
            for wave in waveTimeMap:
                if waveTimeStr == '':
                    waveTimeStr = common.toStr(wave) + '=' + common.toStr(waveTimeMap[wave])
                else:
                    waveTimeStr = waveTimeStr + '|' + common.toStr(wave) + '=' + common.toStr(waveTimeMap[wave])
            roundWaveHpList[row] = waveTimeStr
            roundMonsterHpList[row] = monsterTimeStr

    roundSht.cells(1, roundWaveHpCol + 1).options(transpose=True).value = roundWaveHpList
    roundSht.cells(1, roundMonsterHpCol + 1).options(transpose=True).value = roundMonsterHpList

    propertySht.cells(1, propertyTypeCol + 1).options(transpose=True).value = propertyTypeList
    propertySht.cells(1, propertyShieldCol + 1).options(transpose=True).value = propertyShieldList
    propertySht.cells(1, propertyLevCol + 1).options(transpose=True).value = propertyLevList
    propertySht.cells(1, propertyHpCol + 1).options(transpose=True).value = propertyHpList
    propertySht.cells(1, propertyAtkCol + 1).options(transpose=True).value = propertyAtkList
    propertyWb.save()
    propertyWb.close()

    levelSht.cells(1, levelPropertyCol + 1).options(transpose=True).value = levelPropertyList
    levelWb.save()
    levelWb.close()

    monsterWb.close()
    input("程序运行完毕，按回车键退出...")
