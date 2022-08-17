import xlwings as xw
from xlwings import Sheet
from xlwings import Range
from common.printer import Printer
from common.common import getColBy2Para
from common.common import getDataColOrder
from common.common import getListData
from common.common import getRowData
from common.common import isNumberValid
from common.common import toNum

# 脚本说明：
# 用于战斗数值模板表计算关卡强度
#


def main():
    printer = Printer()
    printer.printColor('~~~开始处理数据~~~', 'green')

    wb = xw.books.active
    stageSht: Sheet = wb.sheets['关卡']
    attrSht: Sheet = wb.sheets['战斗属性']
    pSkillSht: Sheet = wb.sheets['角色技能']
    mSkillSht: Sheet = wb.sheets['怪物技能']
    deepSht: Sheet = wb.sheets['深化&套装']

    stageRange: Range = stageSht.used_range
    stageShtData = stageRange.value
    stageDataCols = getColBy2Para('关卡数据', ['关卡类型', '总输出', '总时长'], stageShtData)

    attrShtData = attrSht.used_range.value
    attrLevCol = getDataColOrder(attrShtData, '等级', 1)
    attrColsMap = {}
    attrColsMap['主线'] = getColBy2Para('主线', ['生命比', '攻击比'], attrShtData)
    attrColsMap['零点追踪'] = getColBy2Para('零点追踪', ['生命比', '攻击比'], attrShtData)
    attrColsMap[11] = getColBy2Para(11, ['生命', '攻击', '防御', '暴击', '暴伤'], attrShtData)
    attrColsMap[12] = getColBy2Para(12, ['生命', '攻击', '防御', '暴击', '暴伤'], attrShtData)
    attrColsMap[13] = getColBy2Para(13, ['生命', '攻击', '防御', '暴击', '暴伤'], attrShtData)
    attrColsMap['男主'] = getColBy2Para('男主', ['生命', '攻击', '防御', '暴击', '暴伤'], attrShtData)

    # 获得玩家属性数据
    playerLevCol = getColBy2Para('玩家', '等级', stageShtData)
    playerLevIds = stageRange.columns[playerLevCol].value
    pAtrrList = getListData(playerLevIds, attrLevCol, attrColsMap['男主'], attrShtData)

    # 获得男女主技能数据
    pSkillShtData = pSkillSht.used_range.value
    pSkillIdCol = getDataColOrder(pSkillShtData, '角色id', 0)
    pSkillDataCol = getDataColOrder(pSkillShtData, ['生命加成', '预期DPS'], 0)
    playerWeaponCol = getColBy2Para('玩家', '武器id', stageShtData)
    playerWeaponIds = stageRange.columns[playerWeaponCol].value
    weaponSkillList = getListData(playerWeaponIds, pSkillIdCol, pSkillDataCol, pSkillShtData)
    playerScoreCol = getColBy2Para('玩家', '搭档id', stageShtData)
    playerScoreIds = stageRange.columns[playerScoreCol].value
    scoreSkillList = getListData(playerScoreIds, pSkillIdCol, pSkillDataCol, pSkillShtData)

    # 获得男主深化数据
    deepShtData = deepSht.used_range.value
    deepCheckCol = getDataColOrder(deepShtData, ['score', 'lev'], 0)
    deepDataCol = getDataColOrder(deepShtData, ['累计伤害加成', '累计承受伤害'], 0)
    playerDeepCol = getColBy2Para('玩家', ['搭档id', '深化'], stageShtData)
    playerDeepLevs = stageRange.columns[playerDeepCol[1]].value
    playerScoreDeeps = []
    for i in range(len(playerScoreIds)):
        playerScoreDeeps.append([playerScoreIds[i], playerDeepLevs[i]])
    deepDataList = getListData(playerScoreDeeps, deepCheckCol, deepDataCol, deepShtData)
    # 计算玩家属性[dps输出，hp，防御，承受伤害]
    playerAtrrList = getPlayerFightAttr(pAtrrList, weaponSkillList, scoreSkillList, deepDataList)

    # 获取怪物数据
    mSkillShtData = mSkillSht.used_range.value
    mSkillIdCol = getDataColOrder(mSkillShtData, '怪物id', 0)
    mSkillDataCol = getDataColOrder(mSkillShtData, ['生命加成', '预期DPS'], 0)
    monstersDataCols = []
    monstersDataCols.append(getColBy2Para('怪1', ['波次', 'ID', '数量', '属性', '等级', '输出', '生存'], stageShtData))
    monstersDataCols.append(getColBy2Para('怪2', ['波次', 'ID', '数量', '属性', '等级', '输出', '生存'], stageShtData))
    monstersDataCols.append(getColBy2Para('怪3', ['波次', 'ID', '数量', '属性', '等级', '输出', '生存'], stageShtData))
    monstersAtrrList = []  # 怪物属性[波次，数量，有效生命，dps输出，lev]

    # 获取怪物属性数据
    for r in range(2, len(stageShtData)):
        monsterAtrrList = []
        for i in range(len(monstersDataCols)):
            rList = []
            pList = [0] * 5  # [波次，数量，有效生命，dps输出，lev]
            attrType = stageShtData[r][monstersDataCols[i][3]]
            monsterLev = stageShtData[r][monstersDataCols[i][4]]
            if isNumberValid(monsterLev):
                # 获取怪物等级对应的属性
                rList = getRowData(monsterLev, attrLevCol, attrColsMap[attrType], attrShtData)
                # 获取关卡对应的属性比例
                stageType = stageShtData[r][stageDataCols[0]]
                changeList = getRowData(monsterLev, attrLevCol, attrColsMap[stageType], attrShtData)
                for j in range(len(changeList)):
                    rList[j] *= changeList[j]
                # 获取怪物技能
                monsterId = stageShtData[r][monstersDataCols[i][1]]
                monsterSkillList = getRowData(monsterId, mSkillIdCol, mSkillDataCol, mSkillShtData)
                # 计算怪物最终属性
                wave = stageShtData[r][monstersDataCols[i][0]]
                num = stageShtData[r][monstersDataCols[i][2]]
                playerLev = stageShtData[r][playerLevCol]
                pList = getMonsterFightAttr(rList, monsterSkillList, playerLev, wave, num, monsterLev)
            monsterAtrrList.append(pList)
        monstersAtrrList.append(monsterAtrrList)

    # 计算数据
    totalList, resultList = getFightResult(playerAtrrList, monstersAtrrList)
    stageSht.cells(3, stageDataCols[1] + 1).value = totalList
    for i in range(len(monstersDataCols)):
        stageSht.cells(3, monstersDataCols[i][5] + 1).value = resultList[i]
    printer.printGapTime("关卡怪物时长处理完毕，耗时:")

    printer.printColor('~~~所有数据处理完毕~~~', 'green')


def getMonsterFightAttr(attrList, skillList, playerLev, wave, num, lev):
    """根据属性计算获得[波次，数量，有效生命，dps输出，等级]

    Args:
        attrList (_type_): 属性['生命', '攻击', '防御', '暴击', '暴伤']
        skillList (_type_): 技能效果['生命加成', '预期DPS']
        playerLev (_type_): 对方等级
        wave (_type_): 怪物波次
        num (_type_): 怪物数量
        lev (_type_): 怪物等级
    """
    reduceHurt = toNum(attrList[2]) / (toNum(attrList[2]) + 1000 + toNum(playerLev) * 12)
    effectHp = toNum(attrList[0]) * (1 + toNum(skillList[0])) / (1 - reduceHurt)
    dps = toNum(skillList[1]) * toNum(attrList[1]) * (toNum(attrList[3]) * toNum(attrList[4]) + 1 - toNum(attrList[3]))
    return [wave, num, int(effectHp), round(dps, 2), lev]


def getPlayerFightAttr(attrList, WeaponSkillList, scoreSkillList, deepList):
    """根据属性计算获得[dps输出，hp，防御，承受伤害]

    Args:
        attrList (_type_):  属性['生命', '攻击', '防御', '暴击', '暴伤']
        WeaponSkillList (_type_): 技能效果['生命加成', '预期DPS']
        scoreSkillList (_type_): 技能效果['生命加成', '预期DPS']
        deepList (_type_): 深化效果['伤害加成','承受伤害']
    """
    returnList = []
    for i in range(2, len(attrList)):
        rList = [0] * 4
        rList[0] = toNum(attrList[i][1]) * (toNum(attrList[i][3]) * toNum(attrList[i][4]) + 1 - toNum(attrList[i][3]))
        rList[0] *= (toNum(WeaponSkillList[i][1]) + toNum(scoreSkillList[i][1])) * (1 + toNum(deepList[i][0]))
        rList[0] = round(rList[0], 2)
        rList[1] = int(toNum(attrList[i][0]) * (2 + toNum(scoreSkillList[i][0]) + toNum(WeaponSkillList[i][0])))
        rList[2] = toNum(attrList[i][2])
        rList[3] = toNum(deepList[i][1])
        returnList.append(rList)
    return returnList


def getFightResult(playerAtrrList, monstersAtrrList):
    """计算战斗结果

    Args:
        playerAtrrList (_type_): 玩家属性[dps输出，hp，防御，承受伤害]
        monstersAtrrList (_type_): 怪物属性[[波次，数量，有效生命，dps输出，lev][..][..]]
    """
    # 同一波怪的数量对应生命系数
    numHpMap = {1: 1, 2: 0.8, 3: 0.7, 4: 0.625, 5: 0.56, 6: 0.5, 7: 0.45, 8: 0.4}

    totalList = []  # [总输出，总生存]
    resultList = [[], [], []]  # [[输出,生存][..][..]]
    for i in range(len(playerAtrrList)):
        tList = [0, 0]
        waveNumMap = getWaveNumMap(monstersAtrrList[i])
        for j in range(len(monstersAtrrList[i])):
            mAttr = monstersAtrrList[i][j]
            if mAttr[1] > 0:
                num = waveNumMap[mAttr[0]]
                mHp = mAttr[1] * mAttr[2] * numHpMap[num]
                mDps = mAttr[1] * mAttr[3]
                mTime = round(mHp / playerAtrrList[i][0], 1)  # 怪物存活时间
                pSufferHurt = 1 - playerAtrrList[i][2] / (playerAtrrList[i][2] + 1000 + 12 * mAttr[4])
                pSufferHurt *= playerAtrrList[i][3]
                mHurt = round(mDps * mTime * pSufferHurt / playerAtrrList[i][1], 2)  # 怪物造成伤害占玩家总生命（score+pl）的比例
                tList[0] += mHurt
                tList[1] += mTime
                resultList[j].append([mHurt, mTime])
            else:
                resultList[j].append([0, 0])
        totalList.append(tList)

    return totalList, resultList


def getWaveNumMap(monstersAtrrList):
    """获取波次对应怪数量

    Args:
        monstersAtrrList (_type_): 怪物数据

    Returns:
        _type_: 波次对应怪数量字典
    """
    waveNumMap = {}
    for i in range(1, 4):
        for attr in monstersAtrrList:
            if attr[0] == i:
                if i in waveNumMap:
                    waveNumMap[i] += attr[1]
                else:
                    waveNumMap[i] = attr[1]
    return waveNumMap
