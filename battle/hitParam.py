import xlwings as xw
import numpy as np
from common import common
"""
技能伤害来源：伤害盒+子弹+法术场
技能伤害盒：
    技能关联多个伤害盒使用相同数值id：内部分配权重（此类伤害盒限定为Once类型）
    技能关联多个伤害盒使用不同数值id：分开计算伤害

子弹伤害盒：导出数据格式【命中伤害;爆炸伤害】
    技能关联多个子弹使用相同子弹id：只能造成单次伤害
    技能关联多个子弹使用不同子弹id：可造成多次伤害

法术场：
    可造成多次伤害

"""
"""导出需求
技能导出增加：伤害盒数值id
伤害盒导出：增加技能id
"""

###
# 用于检查技能的配置是否有错误
###


def main():
    dataWb = xw.books.active
    programPath = dataWb.sheets['路径'].range('D2').value

    hitPath = r"\Binaries\Tables\OriginTable\Battle\BattleHitParam.xlsx"
    boxPath = r"\TimelineCsv\战斗数值伤害盒导出.csv"
    skillPath = r"\TimelineCsv\战斗数值技能导出工具.csv"
    hitWb = xw.books.open(programPath + hitPath, update_links=False)
    boxWb = xw.books.open(programPath + boxPath)
    skillWb = xw.books.open(programPath + skillPath)

    # 技能强度表
    dataList = dataWb.sheets['规范化数据输入&技能强度计算'].used_range.value
    dataSkillCol = common.getDataColOrder(dataList, '技能id')
    dataHurtCol = common.getDataColOrder(dataList, '最终skill分配值')

    # 技能导出表
    skillList = skillWb.sheets['战斗数值技能导出工具'].used_range.value
    skillIdCol = common.getDataColOrder(skillList, '技能ID')
    skillBoxCol = common.getDataColOrder(skillList, '技能关联伤害盒')
    skillBulletCol = common.getDataColOrder(skillList, '子弹关联伤害盒')
    skillFieldCol = common.getDataColOrder(skillList, '法术场伤害盒ID')

    # 伤害盒导出表
    boxList = boxWb.sheets['战斗数值伤害盒导出'].used_range.value
    boxIdCol = common.getDataColOrder(boxList, '伤害盒ID')
    boxHitIdCol = common.getDataColOrder(boxList, '伤害盒关联伤害ID')
    boxHitTypeCol = common.getDataColOrder(boxList, '伤害盒判定方式')
    boxHitParaCol = common.getDataColOrder(boxList, '周期伤害参数')

    # 伤害数值填写表
    hitSht = hitWb.sheets['&HitParamConfig']
    hitList = np.array(hitSht.used_range.value)
    hitIdCol = common.getDataColOrder(hitList, 'HitParamID', 2)
    hitDamageCol = common.getDataColOrder(hitList, 'TargetDamageAtkRatio', 2)
    hitIds = hitList[:, hitIdCol]
    hitDamages = hitList[:, hitDamageCol]

    skillHurtMap = {}
    hurtIdMap = {}
    for i in range(1, len(dataList)):
        if dataList[i][dataSkillCol] is not None:
            skillId = dataList[i][dataSkillCol]
            if skillId in skillHurtMap and skillHurtMap[skillId] != dataList[i][dataHurtCol]:
                print(skillId, '[skillId]存在重复技能id，但数值强度不同')
            else:
                skillHurtMap[skillId] = dataList[i][dataHurtCol]
                hitBoxList, bulletBoxList, fieldBoxList = getSkillHitId(skillList, skillId, skillIdCol, skillBoxCol,
                                                                        skillBulletCol, skillFieldCol)
                hitIdMap = {}
                skillTimes = 0
                for box in hitBoxList:
                    hitId, _skillTimes = getSkillHitTimes(boxList, box, boxIdCol, boxHitIdCol, boxHitTypeCol, boxHitParaCol,
                                                          hitIdMap)
                    skillTimes = _skillTimes
                for box in bulletBoxList:
                    hitId, _skillTimes = getSkillHitTimes(boxList, box, boxIdCol, boxHitIdCol, boxHitTypeCol, boxHitParaCol,
                                                          hitIdMap)
                    skillTimes = _skillTimes
                for box in fieldBoxList:
                    hitId, _skillTimes = getSkillHitTimes(boxList, box, boxIdCol, boxHitIdCol, boxHitTypeCol, boxHitParaCol,
                                                          hitIdMap)
                    skillTimes = _skillTimes

                if skillTimes > 0:
                    hitDamage = round(dataList[i][dataHurtCol] / skillTimes, 0)
                    if hitId not in hurtIdMap:
                        hurtIdMap[hitId] = [skillId]
                        updateSkillDamage(hitIds, hitDamages, hitId, hitDamage)
                    elif skillId not in hurtIdMap[hitId]:
                        hurtIdMap[hitId].append(skillId)
                        print(hitId, '[伤害数值id]被多个技能复用', hurtIdMap[hitId])

    hitSht.cells(1, hitDamageCol + 1).options(transpose=True).value = hitDamages
    hitWb.save()
    hitWb.close()
    boxWb.close()
    skillWb.close()
    input('程序运行结束')


def updateSkillDamage(hitIds, hitDamages, hitId, hitDamage):
    for i in range(len(hitIds)):
        if hitIds[i] == hitId:
            hitDamages[i] = hitDamage
            return
    else:
        print(hitId, '：[HitParamID]在&HitParamConfig中无该id')


# 计算技能次数
def getSkillHitTimes(boxList, boxId, boxIdCol, hitIdCol, hitTypeCol, hitParaCol, hitIdMap):
    for i in range(len(boxList)):
        if boxList[i][boxIdCol] == boxId:
            if boxList[i][hitTypeCol] == 'ActorCDCount' or boxList[i][hitTypeCol] == 'PeriodCount':
                paraList = common.split(boxList[i][hitParaCol], '=')
                for j in range(len(paraList)):
                    paraList[j] = common.toNum(paraList[j])
                if paraList[0] > 0:
                    times = int(paraList[2] / paraList[0])
                    if times > paraList[1]:
                        times = int(paraList[1])
                    if times < 1:
                        times = 1
                else:
                    times = 1
            else:
                times = 1
            # 检查是否有数值id相同，次数不同的情况
            hitId = boxList[i][hitIdCol]
            if hitId in hitIdMap:
                if hitIdMap[hitId] != times:
                    print(boxId, '[伤害盒]所属技能复用数值id，但伤害次数不同')
                return hitId, 0
            elif len(hitIdMap) > 0 and hitId > 0:
                print(boxId, '[伤害盒]所属技能存在多个伤害数值id')
                return hitId, 0
            elif hitId > 0:
                hitIdMap[hitId] = times
                return hitId, times
            else:
                return hitId, 0
    else:
        print(boxId, '[伤害盒]导出csv中不存在该id')
        return -1, -1


# 获取技能伤害盒
def getSkillHitId(skillList, skillId, skillIdCol, skillBoxCol, skillBulletCol, skillFieldCol):
    hitBoxList = []  # 权重伤害盒
    bulletBoxList = []  # 子弹伤害盒
    fieldBoxList = []  # 法术场伤害盒
    for i in range(len(skillList)):
        if skillList[i][skillIdCol] == skillId:
            if skillList[i][skillBoxCol] is not None:
                if common.isNumber(skillList[i][skillBoxCol]):
                    if skillList[i][skillBoxCol] != 0:
                        hitBoxList.append(skillList[i][skillBoxCol])
                else:
                    tList = common.split(skillList[i][skillBoxCol], '|')
                    for box in tList:
                        if box != '0':
                            hitBoxList.append(common.toInt(box))

            if skillList[i][skillBulletCol] is not None:
                bulletList = common.split(skillList[i][skillBulletCol], '|')
                for bullet in bulletList:
                    hurtBoxList = common.split(bullet, ';')
                    for box in hurtBoxList:
                        if box != '0':
                            bulletBoxList.append(common.toInt(box))

            if skillList[i][skillFieldCol] is not None:
                fieldList = common.split(skillList[i][skillBulletCol], '|')
                for field in fieldList:
                    hurtBoxList = common.split(field, ';')
                    for box in hurtBoxList:
                        if box != '0':
                            fieldBoxList.append(common.toInt(box))
            if len(hitBoxList) > 0 and len(bulletBoxList) > 0:
                print(skillId, '[技能id]同时存在打击盒伤害和子弹伤害')
            if len(bulletBoxList) > 0 and len(fieldBoxList) > 0:
                print(skillId, '[技能id]同时存在子弹伤害和法术场伤害')
            if len(hitBoxList) > 0 and len(fieldBoxList) > 0:
                print(skillId, '[技能id]同时存在打击盒伤害和法术场伤害')
            return hitBoxList, bulletBoxList, fieldBoxList
    else:
        return hitBoxList, bulletBoxList, fieldBoxList
