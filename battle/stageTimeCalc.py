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
from common.common import toStr

# 脚本说明：
# 用于战斗数值模板表计算关卡强度
#


def main():
    printer = Printer()
    printer.printColor('~~~开始处理数据~~~', 'green')

    wb = xw.books.active
    stageSht: Sheet = wb.sheets['关卡']
    attrSht: Sheet = wb.sheets['战斗属性']
    pSkillSht: Sheet = wb.sheets['深化&套装']
    mSkillSht: Sheet = wb.sheets['怪物强度']
    # deepSht: Sheet = wb.sheets['深化&套装']

    stageRange: Range = stageSht.used_range
    stageShtData = stageRange.value
    stageDataCols = getColBy2Para('关卡数据', ['总输出', '总时长'], stageShtData)

    attrShtData = attrSht.used_range.value
    attrLevCol = getDataColOrder(attrShtData, '等级', 1)
    attrColsMap = {}
    attrColsMap[1] = getColBy2Para(1, ['生命', '攻击', '防御', '暴击', '暴伤'], attrShtData)
    attrColsMap[2] = getColBy2Para(2, ['生命', '攻击', '防御', '暴击', '暴伤'], attrShtData)
    attrColsMap[3] = getColBy2Para(3, ['生命', '攻击', '防御', '暴击', '暴伤'], attrShtData)
    attrColsMap[11] = getColBy2Para(11, ['生命', '攻击', '防御', '暴击', '暴伤'], attrShtData)
    attrColsMap['男主'] = getColBy2Para('男主', ['生命', '攻击', '防御', '暴击', '暴伤'], attrShtData)
    attrColsMap['玩家'] = getColBy2Para('玩家', ['生命', '攻击', '防御', '暴击', '暴伤'], attrShtData)

    # 获得玩家属性数据
    playerLevCol = getColBy2Para('玩家', '等级', stageShtData)
    playerLevIds = stageRange.columns[playerLevCol].value
    pAtrrList = getListData(playerLevIds, attrLevCol, attrColsMap['男主'], attrShtData)
    p2AtrrList = getListData(playerLevIds, attrLevCol, attrColsMap['玩家'], attrShtData)

    # 获得玩家技能数据
    pSkillShtData = pSkillSht.used_range.value
    pSkillIndexCol = getDataColOrder(pSkillShtData, 'index', 0)
    pSkillDataCol = getDataColOrder(pSkillShtData, ['累计承受伤害', '累计伤害加成'], 0)
    playerScoreCol = getColBy2Para('玩家', '搭档id', stageShtData)
    playerDeepCol = getColBy2Para('玩家', '深化', stageShtData)
    playerScoreList = stageRange.columns[playerScoreCol].value
    playerDeepList = stageRange.columns[playerDeepCol].value

    # 计算玩家属性[dps输出，hp，防御]
    playerAtrrList = getPlayerFightAttr(pAtrrList, p2AtrrList, playerScoreList, playerDeepList, pSkillShtData, pSkillIndexCol,
                                        pSkillDataCol)

    # 获取怪物数据
    mSkillShtData = mSkillSht.used_range.value
    mSkillIdCol = getDataColOrder(mSkillShtData, '怪物id', 0)
    mSkillDataCol = getDataColOrder(mSkillShtData, ['生命比', '攻击比', 'Dps', 'type'], 0)
    monstersDataCols = []
    monstersDataCols.append(getColBy2Para('怪1', ['波次', 'ID', '数量', '属性', '攻击比', '生命比', '等级', '输出', '生存'], stageShtData))
    monstersDataCols.append(getColBy2Para('怪2', ['波次', 'ID', '数量', '属性', '攻击比', '生命比', '等级', '输出', '生存'], stageShtData))
    monstersDataCols.append(getColBy2Para('怪3', ['波次', 'ID', '数量', '属性', '攻击比', '生命比', '等级', '输出', '生存'], stageShtData))
    monstersDataCols.append(getColBy2Para('怪4', ['波次', 'ID', '数量', '属性', '攻击比', '生命比', '等级', '输出', '生存'], stageShtData))
    monstersDataCols.append(getColBy2Para('怪5', ['波次', 'ID', '数量', '属性', '攻击比', '生命比', '等级', '输出', '生存'], stageShtData))
    monstersAtrrList = []  # 怪物属性[波次，数量，有效生命，dps输出，lev]

    # 获取怪物属性数据
    for r in range(2, len(stageShtData)):
        monsterAtrrList = []
        for i in range(len(monstersDataCols)):
            rList = []
            pList = [0] * 5  # [波次，数量，有效生命，dps输出，lev]
            attrType = stageShtData[r][monstersDataCols[i][3]]
            atkPercent = stageShtData[r][monstersDataCols[i][4]]
            hpPercent = stageShtData[r][monstersDataCols[i][5]]
            monsterLev = stageShtData[r][monstersDataCols[i][6]]
            if isNumberValid(monsterLev):
                # 获取怪物等级对应的属性['生命', '攻击', '防御', '暴击', '暴伤']
                rList = getRowData(monsterLev, attrLevCol, attrColsMap[attrType], attrShtData)
                rList[0] *= hpPercent / 1000
                rList[1] *= atkPercent / 1000
                rList[3] /= 1000
                rList[4] /= 1000
                # 获取怪物技能
                monsterId = stageShtData[r][monstersDataCols[i][1]]
                monsterSkillList = getRowData(monsterId, mSkillIdCol, mSkillDataCol, mSkillShtData)
                # 计算怪物最终属性
                wave = stageShtData[r][monstersDataCols[i][0]]
                num = stageShtData[r][monstersDataCols[i][2]]
                playerLev = stageShtData[r][playerLevCol]
                if monsterSkillList == [None]:
                    print('[怪物强度]表内缺少怪物id', monsterId)
                    monsterSkillList = [1, 1, 1, 1]
                pList = getMonsterFightAttr(rList, monsterSkillList, playerLev, wave, num, monsterLev)
            monsterAtrrList.append(pList)
        monstersAtrrList.append(monsterAtrrList)

    # 计算数据[dps输出，hp，防御],[波次，数量，有效生命，dps输出，lev]
    totalList, resultList = getFightResult(playerAtrrList, monstersAtrrList)
    stageSht.cells(3, stageDataCols[0] + 1).value = totalList
    for i in range(len(monstersDataCols)):
        stageSht.cells(3, monstersDataCols[i][7] + 1).value = resultList[i]
    printer.printGapTime("关卡怪物时长处理完毕，耗时:")

    printer.printColor('~~~所有数据处理完毕~~~', 'green')


def getPlayerFightAttr(attrList, p2AtrrList, playerScoreList, playerDeepList, pSkillShtData, pSkillIndexCol, pSkillDataCol):
    """根据属性计算获得[dps输出，hp，防御]

    Args:
        attrList (_type_):  属性['生命', '攻击', '防御', '暴击', '暴伤']
        playerScoreList (_type_): scoreid
        playerDeepList (_type_): score深化等级
        pSkillShtData (_type_): 技能表数据
        pSkillIndexCol (_type_): 技能索引index
        pSkillDataCol (_type): 技能数据['累计承受伤害','累计伤害加成']
    """
    returnList = []
    for i in range(2, len(attrList)):
        skillData = getRowData(toStr(playerScoreList[i]) + toStr(playerDeepList[i]), pSkillIndexCol, pSkillDataCol, pSkillShtData)  # yapf:disable
        if skillData == [None]:
            print('[深化&套装]表内缺少角色索引', toStr(playerScoreList[i]) + toStr(playerDeepList[i]))
            skillData = [1, 4]
        rList = [0] * 3
        hp = (toNum(attrList[i][0]) + toNum(p2AtrrList[i][0])) / 2
        atk = (toNum(attrList[i][1]) + toNum(p2AtrrList[i][1])) / 2
        defence = (toNum(attrList[i][2]) + toNum(p2AtrrList[i][2])) / 2
        critRate = (toNum(attrList[i][3]) + toNum(p2AtrrList[i][3])) / 2 / 1000
        critValue = (toNum(attrList[i][4]) + toNum(p2AtrrList[i][4])) / 2 / 1000 + 1.5

        rList[0] = round(atk * (critRate * critValue + 1 - critRate) * toNum(skillData[1]), 2)
        rList[1] = int(hp * toNum(skillData[0]))
        rList[2] = toNum(defence)
        returnList.append(rList)
    return returnList


def getMonsterFightAttr(attrList, skillList, playerLev, wave, num, lev):
    """根据属性计算获得[波次，数量，有效生命，dps输出，等级]

    Args:
        attrList (_type_): 属性['生命', '攻击', '防御', '暴击', '暴伤']
        skillList (_type_): 技能效果['生命比','攻击比','统计预期Dps','type']
        playerLev (_type_): 对方等级
        wave (_type_): 怪物波次
        num (_type_): 怪物数量
        lev (_type_): 怪物等级
    """
    reduceHurt = toNum(attrList[2]) / (toNum(attrList[2]) + 1600 + toNum(playerLev) * 20)
    effectHp = toNum(attrList[0]) * toNum(skillList[0]) / (1 - reduceHurt)
    if skillList[3] > 1:  # 精英和boss带核（不匹配tag）损失20%生命
        effectHp /= 1.2
    dps = toNum(skillList[1]) * toNum(skillList[2]) * toNum(attrList[1]) * (toNum(attrList[3]) * toNum(attrList[4]) + 1 - toNum(attrList[3]))  # yapf:disable
    return [wave, num, int(effectHp), round(dps, 2), lev]


def getFightResult(playerAtrrList, monstersAtrrList):
    """计算战斗结果

    Args:
        playerAtrrList (_type_): 玩家属性[dps输出，hp，防御]
        monstersAtrrList (_type_): 怪物属性[[波次，数量，有效生命，dps输出，lev][..][..]]
    """
    totalList = []  # [总输出，总生存]
    resultList = [[], [], [], [], []]  # [[输出,生存][..][..]]
    for i in range(len(playerAtrrList)):
        tList = [0, 2]  # 2秒结算时间,每波附加3秒移动时间
        effectiveHpList, totalAtkTimeHp, maxWave = getWaveEffectiveHp(monstersAtrrList[i])
        tList[1] += maxWave * 3
        for j in range(len(monstersAtrrList[i])):
            mAttr = monstersAtrrList[i][j]
            if mAttr[1] > 0:
                mTime = round(effectiveHpList[j] / playerAtrrList[i][0], 1)  # 怪物存活时间
                fTime = round(totalAtkTimeHp[j] / playerAtrrList[i][0], 1)  # 怪物输出时间
                pSufferHurt = 1 - playerAtrrList[i][2] / (playerAtrrList[i][2] + 1600 + 20 * mAttr[4])
                mHurt = round(mAttr[3] * fTime * pSufferHurt / playerAtrrList[i][1], 2)  # 怪物造成伤害占玩家总生命（score+pl）的比例
                tList[0] += mHurt
                tList[1] += mTime
                resultList[j].append([mHurt, mTime])
            else:
                resultList[j].append([0, 0])
        totalList.append(tList)

    return totalList, resultList


def getWaveEffectiveHp(monstersAtrrList):
    """获取波次中怪物的有效血量和有效输出血量
    一波怪的总生命计算，按生命倒叙排序后
    hp1+(hp2^2*0.8+hp3^2*0.6+hp4^2*0.4+hp5^2*0.2)/hp1

    Args:
        monstersAtrrList (_type_): 怪物属性[[波次，数量，有效生命，dps输出，lev][..][..]]
    
    Returns:
        _type_: 有效生存血量，有效输出血量（同类多个怪输出和)
    """

    monsterNum = len(monstersAtrrList)  # 非实际数量,同波同种怪物算作1
    effectiveHpList = [0] * monsterNum
    totalAtkTimeHpList = [0] * monsterNum  # 计算如果有多个相同怪时，总的输出时间
    timeHpList = [0] * monsterNum
    orderList = [1] * monsterNum  # 怪物血量第几高
    maxWave = 0
    for i in range(monsterNum):
        checkMonster = monstersAtrrList[i]
        if checkMonster[0] > 0 and checkMonster[1] > 0:
            wave = checkMonster[0]
            if wave > maxWave:
                maxWave = wave

            maxHp = checkMonster[2]
            for j in range(monsterNum):
                if j != i:
                    attr = monstersAtrrList[j]
                    if attr[0] == wave:
                        if attr[2] > maxHp:
                            maxHp = attr[2]
                        if attr[2] > checkMonster[2]:
                            orderList[i] += attr[1]
                        elif attr[2] == checkMonster[2] and j < i:
                            orderList[i] += attr[1]
            tTime = 0
            for m in range(int(checkMonster[1]) - 1, -1, -1):  # 处理同一波内相同的怪
                fTime = max(0, (1.2 - 0.2 * (orderList[i] + m)))  # 单个怪存活时间
                tTime += fTime  # 单个怪输出时间：自身存活时间+比自己先死的怪的时间
                effectiveHpList[i] += checkMonster[2] / maxHp * checkMonster[2] * fTime  # 所有怪存活时间
                totalAtkTimeHpList[i] += checkMonster[2] / maxHp * checkMonster[2] * tTime  # 所有怪输出时间

    # 计算有效输出时间，假设生命少的怪先死。输出时间=自身存活时间+比自己先死的怪的存活时间
    for i in range(monsterNum):
        timeHpList[i] = totalAtkTimeHpList[i]
        for j in range(monsterNum):
            if j != i and monstersAtrrList[i][0] == monstersAtrrList[j][0]:
                if orderList[j] > orderList[i]:
                    timeHpList[i] += effectiveHpList[i]
    return effectiveHpList, timeHpList, maxWave


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
