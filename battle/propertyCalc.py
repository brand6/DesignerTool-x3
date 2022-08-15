import xlwings as xw
from xlwings import Sheet
from xlwings import Range
from common import common


def main():
    activeWb = xw.books.active
    roleSht: Sheet = activeWb.sheets['男主数值']
    propertySht: Sheet = activeWb.sheets['【附】表格数据']
    propertyData = propertySht.used_range.value
    roleRng: Range = roleSht.used_range
    roleData = roleRng.value

    # 搭档属性
    scoreLevCol = getRoleCol('搭档', '等级', roleData)
    scoreIdCol = getRoleCol('搭档', 'id', roleData)
    scoreLevs = roleRng.columns[scoreLevCol].value
    scoreIds = roleRng.columns[scoreIdCol].value
    scorePropertys = getScoreProperty(scoreIds, scoreLevs, propertyData)

    # 思念属性
    cardLevCol = getRoleCol('思念', '等级', roleData)
    cardPhaseCol = getRoleCol('思念', '进阶', roleData)
    cardIdCol = getRoleCol('思念', ['id1', 'id2', 'id3', 'id4', 'id5', 'id6'], roleData)
    cardLevs = roleRng.columns[cardLevCol].value
    cardPhases = roleRng.columns[cardPhaseCol].value
    cardIds = []
    for i in range(len(cardLevs)):
        idList = []
        for j in range(len(cardIdCol)):
            idList.append(roleData[i][cardIdCol[j]])
        cardIds.append(idList)
    cardPropertys = getCardProperty(cardIds, cardLevs, cardPhases, propertyData)
    # 思念：计算属性加成
    for i in range(len(scorePropertys)):
        for j in range(3):
            cardPropertys[i][j] += scorePropertys[i][j] * cardPropertys[i][j + 5] / 1000

    # 芯核属性
    gemLevCol = getRoleCol('芯核', '等级', roleData)
    gemIdCol = getRoleCol('芯核', ['id1', 'id2', 'id3', 'id4'], roleData)
    gemMainAttrCol = getRoleCol('芯核', ['主属性1', '主属性2', '主属性3', '主属性4'], roleData)
    gemSubAttrIdCol = getRoleCol('芯核', ['副属性1', '副属性2', '副属性3', '副属性4'], roleData)
    gemLevs = roleRng.columns[gemLevCol].value
    gemIds = []
    gemMainAttrs = []
    gemSubAttrs = []
    for i in range(len(gemLevs)):
        idList = []
        mianList = []
        subList = []
        for j in range(len(gemIdCol)):
            idList.append(roleData[i][gemIdCol[j]])
            mianList.append(roleData[i][gemMainAttrCol[j]])
            subList.append(roleData[i][gemSubAttrIdCol[j]])
        gemIds.append(idList)
        gemMainAttrs.append(mianList)
        gemSubAttrs.append(subList)
    gemPropertys = getGemProperty(gemIds, gemLevs, gemMainAttrs, gemSubAttrs, propertyData)
    # 芯核：计算属性加成
    for i in range(len(scorePropertys)):
        for j in range(3):
            gemPropertys[i][j] += scorePropertys[i][j] * gemPropertys[i][j + 5] / 1000

    # 牵绊度属性
    loveLevCol = getRoleCol('牵绊度', '等级', roleData)
    loveLevs = roleRng.columns[loveLevCol].value
    lovePropertys = getLoveProperty(loveLevs, propertyData)

    # 计算属性和
    for i in range(len(scorePropertys)):
        for j in range(len(scorePropertys[0])):
            if common.isNumber(scorePropertys[i][j]):
                scorePropertys[i][j] += int(cardPropertys[i][j]) + int(gemPropertys[i][j]) + lovePropertys[i][j]
    resultCol = getRoleCol('面板属性', '生命', roleData)
    roleSht.cells(3, resultCol + 1).value = scorePropertys


def getScoreProperty(ids: list, levs: list, propertyData: list):
    """获取搭档的属性

    Args:
        ids (list): score的id
        levs (list): score的等级
        propertyData (list): 属性表数据

    return:
        propertys: 属性列表[[生命，攻击，防御，暴击，暴伤]]
    """

    propertys = []
    levIdCol = getPropertyCol('SCore.xlsx', 'SCoreLevel', 'ID', propertyData)
    levColList = getPropertyCol('SCore.xlsx', 'SCoreLevel', ['PropMaxHP', 'PropPhyAtk', 'PropPhyDef', 'PropCritVal'],
                                propertyData)
    starIdCol = getPropertyCol('SCore.xlsx', 'SCoreStar', 'ID', propertyData)
    starColList = getPropertyCol('SCore.xlsx', 'SCoreStar', ['AddMaxHP', 'AddPhyAtk', 'AddPhyDef'], propertyData)
    awakeIdCol = getPropertyCol('SCore.xlsx', 'SCoreAwake', 'ID', propertyData)
    awakeColList = getPropertyCol('SCore.xlsx', 'SCoreAwake', ['AddMaxHP', 'AddPhyAtk', 'AddPhyDef'], propertyData)

    for i in range(2, len(ids)):
        if ids[i] is not None and common.isNumber(ids[i]):
            pList = [0] * 5  # 本数组会输出，放在判断内，排除异常数据
            # 升级属性
            levId = ids[i] * 1000 + levs[i]
            tList = getPropertyData(levId, levIdCol, levColList, propertyData)
            for j in range(len(tList)):
                pList[j] += tList[j]
            # 突破属性
            starId = ids[i] * 1000 + int((levs[i] - 1) / 10) + 1
            tList = getPropertyData(starId, starIdCol, starColList, propertyData)
            for j in range(len(tList)):
                pList[j] += common.toNum(tList[j])
            # 觉醒属性
            if levs[i] == 80:
                awakeId = ids[i]
                tList = getPropertyData(awakeId, awakeIdCol, awakeColList, propertyData)
                for j in range(len(tList)):
                    pList[j] += tList[j]
            # 添加属性值到列表
            propertys.append(pList)
    return propertys


def getCardProperty(ids: list, levs: list, phases: list, propertyData: list):
    """获取思念的属性

    Args:
        ids (list): card的id
        levs (list): card的等级
        propertyData (list): 属性表数据

    return:
        propertys: 属性列表[[生命，攻击，防御，暴击，暴伤，生命加成，攻击加成，防御加成]]
    """
    pMap = {1: 0, 2: 1, 3: 2, 4: 3, 6: 4, 101: 5, 102: 6, 103: 7}

    propertys = []
    baseIdCol = getPropertyCol('Card.xlsx', 'CardBaseInfo', 'ID', propertyData)
    baseTemplateCol = getPropertyCol('Card.xlsx', 'CardBaseInfo', 'Template', propertyData)
    baseStarCol = getPropertyCol('Card.xlsx', 'CardBaseInfo', 'StarID', propertyData)
    basePhaseCol = getPropertyCol('Card.xlsx', 'CardBaseInfo', 'PhaseMode', propertyData)
    baseSpAttrCol = getPropertyCol('Card.xlsx', 'CardBaseInfo', ['SpecialAttrType', 'SpecialAttrValue'], propertyData)
    baseColList = getPropertyCol('Card.xlsx', 'CardBaseInfo', ['MaxHPRate', 'PhyAttackRate', 'PhyDefenceRate'], propertyData)
    tempIdColList = getPropertyCol('Card.xlsx', 'CardTemplate', ['Template', 'Level'], propertyData)
    tempColList = getPropertyCol('Card.xlsx', 'CardTemplate', ['PropMaxHP', 'PropPhyAtk', 'PropPhyDef'], propertyData)
    starIdColList = getPropertyCol('Card.xlsx', 'CardStar', ['StarID', 'StarLevel'], propertyData)
    starColList = getPropertyCol('Card.xlsx', 'CardStar', ['PropMaxHP', 'PropPhyAtk', 'PropPhyDef'], propertyData)
    phaseIdColList = getPropertyCol('Card.xlsx', 'CardPhase', ['Mode', 'PhaseID'], propertyData)
    phaseColList = getPropertyCol('Card.xlsx', 'CardPhase', ['MaxHPUP', 'PhyAtkUP', 'PhyDefUP'], propertyData)

    for i in range(2, len(ids)):
        pList = [0] * 8  # 放在判断外，保证数据行不错位
        if levs[i] is not None and levs[i] > 0:
            for id in ids[i]:
                if id is not None and common.isNumber(id):
                    rList = [0] * 8
                    # 升级属性
                    templateId = getPropertyData(id, baseIdCol, baseTemplateCol, propertyData)
                    levId = levs[i]
                    tList = getPropertyData([templateId, levId], tempIdColList, tempColList, propertyData)
                    for j in range(len(tList)):
                        rList[j] += common.toInt(tList[j])
                    # 突破属性
                    starId = getPropertyData(id, baseIdCol, baseStarCol, propertyData)
                    starLev = int((levId - 1) / 10) + 1
                    tList = getPropertyData([starId, starLev], starIdColList, starColList, propertyData)
                    for j in range(len(tList)):
                        rList[j] += common.toInt(tList[j])
                    # 属性比例
                    rateList = getPropertyData(id, baseIdCol, baseColList, propertyData)
                    for j in range(len(tList)):
                        rList[j] *= common.toInt(rateList[j]) / 1000
                    # 进阶属性
                    if phases[i] is not None and phases[i] > 0:
                        phaseId = getPropertyData(id, baseIdCol, basePhaseCol, propertyData)
                        rateList = getPropertyData([phaseId, phases[i]], phaseIdColList, phaseColList, propertyData)
                        for j in range(len(tList)):
                            rList[j] *= (1 + common.toInt(rateList[j]) / 1000)
                    # 附加属性
                    spAttrList = getPropertyData(id, baseIdCol, baseSpAttrCol, propertyData)
                    if spAttrList[0] in pMap:
                        attrValueList = spAttrList[1].split('|')
                        rList[pMap[spAttrList[0]]] += int(attrValueList[starLev - 1])
                    # 本行属性更新
                    for j in range(len(pList)):
                        pList[j] += rList[j]
        # 添加属性值到列表
        propertys.append(pList)
    return propertys


def getGemProperty(ids: list, levs: list, mainAttrs: list, subAttrs: list, propertyData: list):
    """获取芯核的属性

    Args:
        ids (list): gem的id
        levs (list): gem的等级
        mainAttrs(list): gem的主属性
        subAttrs(list): gem的副属性
        propertyData (list): 属性表数据

    return:
        propertys: 属性列表[[生命，攻击，防御，暴击，暴伤，生命加成，攻击加成，防御加成]]
    """
    pNamemap = {'生命值': 1, '攻击值': 2, '防御值': 3, '暴击': 4, '暴伤': 6, '生命加成': 101, '攻击加成': 102, '防御加成': 103}
    pMap = {1: 0, 2: 1, 3: 2, 4: 3, 6: 4, 101: 5, 102: 6, 103: 7}

    propertys = []
    baseIdCol = getPropertyCol('GemCore.xlsx', 'GemCoreBaseInfo', 'ItemID', propertyData)
    baseMainDropCol = getPropertyCol('GemCore.xlsx', 'GemCoreBaseInfo', 'MainAttrGroup', propertyData)
    baseSubDropCol = getPropertyCol('GemCore.xlsx', 'GemCoreBaseInfo', 'SubAttrGroup', propertyData)
    baseSuitCol = getPropertyCol('GemCore.xlsx', 'GemCoreBaseInfo', 'SuitID', propertyData)
    suitIdCol = getPropertyCol('GemCore.xlsx', 'GemCoreSuit', 'SuitID', propertyData)
    suitAttrCol = getPropertyCol('GemCore.xlsx', 'GemCoreSuit', 'AttrNum', propertyData)
    dropIdCol = getPropertyCol('GemCore.xlsx', 'GemCoreAttrDrop', ['AttrGroupID', 'Attr'], propertyData)
    dropAttrIdCol = getPropertyCol('GemCore.xlsx', 'GemCoreAttrDrop', 'AttrID', propertyData)
    attrIdColList = getPropertyCol('GemCore.xlsx', 'GemCoreAttr', ['AttrID', 'CountMin'], propertyData)
    attrValueCol = getPropertyCol('GemCore.xlsx', 'GemCoreAttr', 'AttrMax', propertyData)

    for i in range(2, len(ids)):
        pList = [0] * 8  # 放在判断外，保证数据行不错位
        if levs[i] is not None and levs[i] > 0:
            for j in range(len(ids[0])):
                if ids[i][j] is not None and common.isNumber(ids[i][j]):
                    # 主属性
                    mainDropId = getPropertyData(ids[i][j], baseIdCol, baseMainDropCol, propertyData)
                    mainAttrType = pNamemap[mainAttrs[i][j]]
                    mainAttrId = getPropertyData([mainDropId, mainAttrType], dropIdCol, dropAttrIdCol, propertyData)
                    lev1Value = getPropertyData([mainAttrId, 1], attrIdColList, attrValueCol, propertyData)
                    lev2Value = getPropertyData([mainAttrId, 2], attrIdColList, attrValueCol, propertyData)
                    if lev2Value is None:
                        pList[pMap[mainAttrType]] += lev1Value * levs[i]
                    else:
                        pList[pMap[mainAttrType]] += lev1Value + lev2Value * (levs[i] - 1)
                    # 副属性
                    for m in range(len(subAttrs[0])):
                        subDropId = getPropertyData(ids[i][j], baseIdCol, baseSubDropCol, propertyData)
                        subAttrType = pNamemap[subAttrs[i][m]]
                        subAttrId = getPropertyData([subDropId, subAttrType], dropIdCol, dropAttrIdCol, propertyData)
                        lev1Value = getPropertyData([subAttrId, 1], attrIdColList, attrValueCol, propertyData)
                        lev2Value = getPropertyData([subAttrId, 2], attrIdColList, attrValueCol, propertyData)
                        if m == 0:  # 芯核升级时提示副属性1
                            lev = 1 + int(levs[i] / 5)
                        else:
                            lev = 1
                        if lev2Value is None:
                            pList[pMap[subAttrType]] += lev1Value * lev  # 4个芯核的副属性相同
                        else:
                            pList[pMap[subAttrType]] += lev1Value + lev2Value * (lev - 1)
                    # 套装
                    if j == 0:
                        suitId = getPropertyData(ids[i][j], baseIdCol, baseSuitCol, propertyData)
                        suitAttr = getPropertyData(suitId, suitIdCol, suitAttrCol, propertyData).split('=')
                        suitAttrType = int(suitAttr[0])
                        suitAttrValue = int(suitAttr[1])
                        pList[pMap[suitAttrType]] += suitAttrValue

        # 添加属性值到列表
        propertys.append(pList)
    return propertys


def getLoveProperty(levs: list, propertyData: list):
    """获取牵绊度的属性

    Args:
        levs (list): 牵绊度的等级
        propertyData (list): 属性表数据

    return:
        propertys: 属性列表[[生命，攻击，防御，暴击，暴伤]]
    """

    propertys = []
    levIdCol = getPropertyCol('LovePointLevel.xlsx', 'LovePointLevel', 'ID', propertyData)
    levColList = getPropertyCol('LovePointLevel.xlsx', 'LovePointLevel', ['PropMaxHP', 'PropPhyAtk', 'PropPhyDef'],
                                propertyData)

    for i in range(2, len(levs)):
        pList = [0] * 5  # 放在判断外，保证数据行不错位
        if levs[i] is not None and common.isNumberValid(levs[i]):
            # 升级属性
            tList = getPropertyData(levs[i], levIdCol, levColList, propertyData)
            for j in range(len(tList)):
                pList[j] += tList[j]
            # 添加属性值到列表
            propertys.append(pList)
    return propertys


def getPropertyData(id, idCol, propertyCol, propertyData: list):
    """根据id获取属性数据

    Args:
        id (): 用于查询的id，可传入str/list
        idCol (): id所在列，可传入str/list
        propertyCol (): 属性所在列，可传入str/list
        propertyData (list): 属性数据

    Returns:
        property: 属性值
    """

    if isinstance(id, list):
        for row in propertyData:
            matchTag = True
            for i in range(len(id)):
                if id[i] != row[idCol[i]]:
                    matchTag = False
                    break
            if matchTag:
                if isinstance(propertyCol, list):
                    returnList = []
                    for col in propertyCol:
                        returnList.append(row[col])
                    return returnList
                else:
                    return row[propertyCol]
    else:
        for row in propertyData:
            if row[idCol] == id:
                if isinstance(propertyCol, list):
                    returnList = []
                    for col in propertyCol:
                        returnList.append(row[col])
                    return returnList
                else:
                    return row[propertyCol]
    return None


def getPropertyCol(wbName: str, shtName: str, colName, propertyData: list):
    """获取列名所在的位置

    Args:
        wbName (str): 表名
        shtName (str): sht名
        colName (): 列名，可传入str或list
        propertyData (list): 属性数据
    """

    _wbName = None
    _shtName = None
    if isinstance(colName, list):
        colList = [None] * len(colName)
        for col in range(len(propertyData[0])):
            if propertyData[0][col] is not None:
                _wbName = propertyData[0][col]
            if propertyData[1][col] is not None:
                _shtName = propertyData[1][col]
            for i in range(len(colName)):
                if wbName == _wbName and shtName == _shtName and colName[i] == propertyData[2][col]:
                    colList[i] = col
        return colList
    else:
        for col in range(len(propertyData[0])):
            if propertyData[0][col] is not None:
                _wbName = propertyData[0][col]
            if propertyData[1][col] is not None:
                _shtName = propertyData[1][col]
            if wbName == _wbName and shtName == _shtName and colName == propertyData[2][col]:
                return col
        else:
            return None


def getRoleCol(classify: str, colName: str, roleData: list):
    """获取列名所在的位置

    Args:
        classify (str): 分类名
        colName (str): 列名，可传入str或list
        roleData (list): 属性数据
    """
    _classify = None
    if isinstance(colName, list):
        colList = [None] * len(colName)
        for col in range(len(roleData[0])):
            if roleData[0][col] is not None:
                _classify = roleData[0][col]
            for i in range(len(colName)):
                if classify == _classify and colName[i] == roleData[1][col]:
                    colList[i] = col
        return colList
    else:
        for col in range(len(roleData[0])):
            if roleData[0][col] is not None:
                _classify = roleData[0][col]
            if classify == _classify and colName == roleData[1][col]:
                return col
        else:
            return None
