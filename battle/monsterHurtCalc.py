import xlwings as xw
import numpy as np
from xlwings import Sheet
from xlwings import Range
from common import common

getDataOrder = common.getDataOrder
getRowData = common.getRowData
toNum = common.toNum

# 脚本说明：
# 用于计算怪的伤害
#

maxStageCount = 10
maxSkillCount = 50


def main():
    activeSht: Sheet = xw.sheets.active
    usedRange: Range = activeSht.used_range
    dataList: list = usedRange.value
    colList: list = dataList[0]
    stage_idCol = getDataOrder(colList, 's怪物id')
    stage_stageCol = getDataOrder(colList, 's阶段')
    stage_timeCol = getDataOrder(colList, '阶段时长')
    stage_randUnitTimeCol = getDataOrder(colList, '随机单位时长')
    stage_prioritySkillTimeCol = getDataOrder(colList, '优先技能时长')
    stage_specialTimeCol = getDataOrder(colList, '特殊时长')
    stage_randTimeCol = getDataOrder(colList, '随机时长')
    stage_noCdrandTimeCol = getDataOrder(colList, '无CD随机时长')

    skill_idCol = getDataOrder(colList, '怪物id')
    skill_stageCol = getDataOrder(colList, '阶段')
    skill_weightCol = getDataOrder(colList, '随机权重')
    skill_timeCol = getDataOrder(colList, '技能时长')
    skill_CdCol = getDataOrder(colList, '配置CD')
    skill_realCdCol = getDataOrder(colList, '实际CD')
    skill_inheritCdCol = getDataOrder(colList, '继承CD')
    skill_prioritySkillCol = getDataOrder(colList, '优先技能')
    skill_randCdSkillCol = getDataOrder(colList, '随机技能')
    skill_realWeightCol = getDataOrder(colList, '普攻实际权重')

    for i in range(1, len(dataList)):
        for j in range(len(dataList[0])):
            if dataList[i][j] is None:
                dataList[i][j] = 0

    getRandOnceTime(dataList, stage_idCol, stage_stageCol, stage_randUnitTimeCol, skill_idCol, skill_stageCol, skill_weightCol,
                    skill_timeCol)
    getSkillCdTime(dataList, stage_idCol, stage_stageCol, stage_randUnitTimeCol, skill_idCol, skill_stageCol, skill_weightCol,
                   skill_timeCol, skill_CdCol, skill_realCdCol)
    getPrioritySkillCount(dataList, stage_idCol, stage_stageCol, stage_timeCol, skill_idCol, skill_stageCol, skill_weightCol,
                          skill_realCdCol, skill_inheritCdCol, skill_prioritySkillCol)
    getPrioritySkillTime(dataList, stage_idCol, stage_stageCol, stage_prioritySkillTimeCol, skill_idCol, skill_stageCol,
                         skill_timeCol, skill_prioritySkillCol)
    getRandTime(dataList, stage_timeCol, stage_specialTimeCol, stage_prioritySkillTimeCol, stage_randTimeCol)
    getRandCdSkillCount(dataList, stage_idCol, stage_stageCol, stage_randTimeCol, stage_randUnitTimeCol, skill_idCol,
                        skill_stageCol, skill_weightCol, skill_timeCol, skill_CdCol, skill_randCdSkillCol)
    getNoCdRandTime(dataList, stage_idCol, stage_stageCol, stage_randTimeCol, stage_noCdrandTimeCol, skill_idCol,
                    skill_stageCol, skill_timeCol, skill_randCdSkillCol)
    getSkillRealWeight(dataList, skill_idCol, skill_stageCol, skill_weightCol, skill_CdCol, skill_timeCol, skill_realWeightCol)

    nData = np.array(dataList)
    activeSht.cells(2, stage_randUnitTimeCol + 1).options(transpose=True).value = nData[1:, stage_randUnitTimeCol]
    activeSht.cells(2, stage_prioritySkillTimeCol + 1).options(transpose=True).value = nData[1:, stage_prioritySkillTimeCol]
    activeSht.cells(2, stage_randTimeCol + 1).options(transpose=True).value = nData[1:, stage_randTimeCol]
    activeSht.cells(2, stage_noCdrandTimeCol + 1).options(transpose=True).value = nData[1:, stage_noCdrandTimeCol]
    activeSht.cells(2, skill_realCdCol + 1).options(transpose=True).value = nData[1:, skill_realCdCol]
    activeSht.cells(2, skill_prioritySkillCol + 1).options(transpose=True).value = nData[1:, skill_prioritySkillCol]
    activeSht.cells(2, skill_randCdSkillCol + 1).options(transpose=True).value = nData[1:, skill_randCdSkillCol]
    activeSht.cells(2, skill_realWeightCol + 1).options(transpose=True).value = nData[1:, skill_realWeightCol]


def getSkillRealWeight(dataList, skill_idCol, skill_stageCol, skill_weightCol, skill_CdCol, skill_timeCol, skill_realWeightCol):
    """获取无CD技能的实际权重
    """
    for j in range(1, len(dataList)):
        monsterId = dataList[j][skill_idCol]
        stageId = dataList[j][skill_stageCol]
        if monsterId is not None and stageId > 0:
            if dataList[j][skill_weightCol] < 1 and dataList[j][skill_CdCol] < dataList[j][skill_timeCol]:
                startRow = 1 if j < maxSkillCount else j - maxSkillCount
                endRow = len(dataList) if len(dataList) < j + maxSkillCount else j + maxSkillCount
                totalWeight = 0
                for i in range(startRow, endRow):
                    if monsterId == dataList[i][skill_idCol] and stageId == dataList[i][skill_stageCol] and dataList[i][
                            skill_weightCol] < 1 and dataList[i][skill_CdCol] < dataList[i][skill_timeCol]:
                        totalWeight += dataList[i][skill_weightCol]
                dataList[j][skill_realWeightCol] = round(dataList[j][skill_weightCol] / totalWeight, 2)


def getNoCdRandTime(dataList, stage_idCol, stage_stageCol, stage_randTimeCol, stage_noCdrandTimeCol, skill_idCol,
                    skill_stageCol, skill_timeCol, skill_randCdSkillCol):
    """获取无CD技能随机时长
    """
    for i in range(1, len(dataList)):
        monsterId = dataList[i][stage_idCol]
        if monsterId is not None:
            getFlag = False
            unitTime = 0
            stageId = dataList[i][stage_stageCol]
            startRow = 1 if i < maxStageCount else i - maxStageCount
            endRow = len(dataList) if len(dataList) < startRow + maxSkillCount else startRow + maxSkillCount
            for j in range(startRow, endRow):
                if monsterId == dataList[j][skill_idCol] and stageId == dataList[j][skill_stageCol]:
                    getFlag = True
                    if dataList[j][skill_randCdSkillCol] is not None:
                        unitTime = unitTime + dataList[j][skill_randCdSkillCol] * dataList[j][skill_timeCol]
                elif getFlag:
                    break
            dataList[i][stage_noCdrandTimeCol] = dataList[i][stage_randTimeCol] - unitTime


def getRandCdSkillCount(dataList, stage_idCol, stage_stageCol, stage_randTimeCol, stage_randUnitTimeCol, skill_idCol,
                        skill_stageCol, skill_weightCol, skill_timeCol, skill_CdCol, skill_randCdSkillCol):
    """获取带CD的随机技能的释放次数

    带CD技能真实权重=1/(1/配置权重+cd/随机单位时间-1)
    """
    for j in range(1, len(dataList)):
        monsterId = dataList[j][skill_idCol]
        stageId = dataList[j][skill_stageCol]
        if monsterId is not None and stageId > 0:
            weightRate = dataList[j][skill_weightCol] if dataList[j][skill_weightCol] is not None else 0
            cd = dataList[j][skill_CdCol]
            if 0 < weightRate and weightRate < 1 and cd > dataList[j][skill_timeCol]:
                startRow = 1 if j < maxSkillCount else j - maxSkillCount
                endRow = len(dataList) if len(dataList) < j + maxSkillCount else j + maxSkillCount
                for i in range(startRow, endRow):
                    if monsterId == dataList[i][stage_idCol] and stageId == dataList[i][stage_stageCol]:
                        randTime = dataList[i][stage_randTimeCol]
                        onceTime = dataList[i][stage_randUnitTimeCol]
                        break
                realWeight = 1 / (1 / weightRate + cd / onceTime - 1)
                dataList[j][skill_randCdSkillCol] = round(randTime * realWeight / onceTime, 2)
            else:
                dataList[j][skill_randCdSkillCol] = 0
        else:
            dataList[j][skill_randCdSkillCol] = 0


def getRandTime(dataList, stage_timeCol, stage_specialTimeCol, stage_prioritySkillTimeCol, stage_randTimeCol):
    """获取随机时长
    """
    for i in range(1, len(dataList)):
        if dataList[i][stage_timeCol] is not None:
            dataList[i][stage_randTimeCol] = dataList[i][stage_timeCol] + toNum(dataList[i][stage_specialTimeCol]) - dataList[i][stage_prioritySkillTimeCol]  # yapf:disable


def getPrioritySkillTime(dataList, stage_idCol, stage_stageCol, stage_prioritySkillTimeCol, skill_idCol, skill_stageCol,
                         skill_timeCol, skill_prioritySkillCol):
    """获取优先释放技能的时长
    """
    for i in range(1, len(dataList)):
        monsterId = dataList[i][stage_idCol]
        if monsterId is not None:
            getFlag = False
            unitTime = 0
            stageId = dataList[i][stage_stageCol]
            startRow = 1 if i < maxStageCount else i - maxStageCount
            endRow = len(dataList) if len(dataList) < startRow + maxSkillCount else startRow + maxSkillCount
            for j in range(startRow, endRow):
                if monsterId == dataList[j][skill_idCol] and stageId == dataList[j][skill_stageCol]:
                    getFlag = True
                    if dataList[j][skill_prioritySkillCol] is not None:
                        unitTime = unitTime + dataList[j][skill_prioritySkillCol] * dataList[j][skill_timeCol]
                elif getFlag:
                    break
            dataList[i][stage_prioritySkillTimeCol] = unitTime


def getPrioritySkillCount(dataList, stage_idCol, stage_stageCol, stage_timeCol, skill_idCol, skill_stageCol, skill_weightCol,
                          skill_realCdCol, skill_inheritCdCol, skill_prioritySkillCol):
    """获取优先释放技能的次数
    """
    for j in range(1, len(dataList)):
        monsterId = dataList[j][skill_idCol]
        stageId = dataList[j][skill_stageCol]
        if monsterId is not None and stageId > 0:
            if dataList[j][skill_weightCol] == 1:
                startRow = 1 if j < maxSkillCount else j - maxSkillCount
                endRow = len(dataList) if len(dataList) < j + maxSkillCount else j + maxSkillCount
                for i in range(startRow, endRow):
                    if monsterId == dataList[i][stage_idCol] and stageId == dataList[i][stage_stageCol]:
                        skillCount = (dataList[i][stage_timeCol] - toNum(dataList[j][skill_inheritCdCol])) / dataList[j][skill_realCdCol]  # yapf:disable
                        dataList[j][skill_prioritySkillCol] = 1 + int(skillCount)
                        break
            else:
                dataList[j][skill_prioritySkillCol] = 0
        else:
            dataList[j][skill_prioritySkillCol] = 0


def getSkillCdTime(dataList, stage_idCol, stage_stageCol, stage_randUnitTimeCol, skill_idCol, skill_stageCol, skill_weightCol,
                   skill_timeCol, skill_CdCol, skill_realCdCol):
    """获取技能的实际CD
    """
    for j in range(1, len(dataList)):
        monsterId = dataList[j][skill_idCol]
        stageId = dataList[j][skill_stageCol]
        if monsterId is not None and stageId is not None and stageId > 0:
            if dataList[j][skill_weightCol] == 1:
                startRow = 1 if j < maxSkillCount else j - maxSkillCount
                endRow = len(dataList) if len(dataList) < j + maxSkillCount else j + maxSkillCount
                for i in range(startRow, endRow):
                    if monsterId == dataList[i][stage_idCol] and stageId == dataList[i][stage_stageCol]:
                        dataList[j][skill_realCdCol] = round(dataList[j][skill_CdCol] + dataList[i][stage_randUnitTimeCol], 2)
                        break
            elif dataList[j][skill_weightCol] is not None and dataList[j][skill_weightCol] > 0:
                dataList[j][skill_realCdCol] = round(max(dataList[j][skill_CdCol], dataList[j][skill_timeCol]), 2)


def getRandOnceTime(dataList, stage_idCol, stage_stageCol, stage_randUnitTimeCol, skill_idCol, skill_stageCol, skill_weightCol,
                    skill_timeCol):
    """获取随机的单位时长
    """

    for i in range(1, len(dataList)):
        monsterId = dataList[i][stage_idCol]
        if monsterId is not None:
            getFlag = False
            unitTime = 0
            stageId = dataList[i][stage_stageCol]
            startRow = 1 if i < maxStageCount else i - maxStageCount
            endRow = len(dataList) if len(dataList) < startRow + maxSkillCount else startRow + maxSkillCount
            for j in range(startRow, endRow):
                if monsterId == dataList[j][skill_idCol] and stageId == dataList[j][skill_stageCol]:
                    getFlag = True
                    if dataList[j][skill_weightCol] is not None and dataList[j][skill_weightCol] < 1:
                        unitTime = unitTime + dataList[j][skill_weightCol] * dataList[j][skill_timeCol]
                elif getFlag:
                    break
            dataList[i][stage_randUnitTimeCol] = round(unitTime, 2)
