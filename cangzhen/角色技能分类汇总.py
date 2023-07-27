# 用于【战斗导出整合】表格，对monster，male，weapon的 角色id - slottype - [skillid]的表格解析
# 输出到程序导出表 格式转为 角色身份 - 角色id - slottype - skillid方便进一步处理

import xlwings as xw
import os

filePath = r'F:\P4_x3_project_batlemain\Program\Binaries\Tables\OriginTable\Battle'

BattleWeaponFileName = 'BattleWeapon.xlsx'
BattleActorFileName = 'BattleActor.xlsx'
BattleMonsterFileName = 'BattleMonster.xlsx'

战斗导出总表 = xw.Book(r'F:\P4_x3_doc\x3\策划文档\数值\战斗数值\工具箱_苍真.xlsb')
女主配置表 = xw.Book(os.path.join(filePath, BattleWeaponFileName))
男主配置表 = xw.Book(os.path.join(filePath, BattleActorFileName))
怪物配置表 = xw.Book(os.path.join(filePath, BattleMonsterFileName))

女主技能表 = 女主配置表.sheets('&WeaponLogicConfig^').used_range.value  # BattleWeapon.xlsx / &WeaponLogicConfig^
男主技能表 = 男主配置表.sheets('&MaleActorConfig^').used_range.value  # BattleActor.xlsx / &MaleActorConfig^
怪物技能表 = 怪物配置表.sheets('&MonsterTemplate^').used_range.value  # BattleMonster.xlsx / &MonsterTemplate^

女主配置表.close()
男主配置表.close()
怪物配置表.close()

keylist = [
    'AttackIDs', 'FillSkillIDs', 'ActiveSkillIDs', 'DodgeSkillIDs',
    'PassiveSkillIDs', 'CoopSkillIDs', 'FemaleCoopSkillIDs', 'UltraSkillIDs'
]

# 普攻,填充技能,主动技能,闪避技能,
# 被动技能,共鸣技能,女主自身共鸣技能,爆发技能,
# 闪避 / 防御技能

sheetlist = [女主技能表, 男主技能表, 怪物技能表]
namelist = ["女主", "搭档", "怪物"]


def GetDataRows(sheet):  # 获得表格有效行数 (表格不脏才能这么干)
    目标行数 = len(sheet) - 3  # 去掉3行表头
    return (目标行数)


def GetDataColumns(sheet):  # 获得表格有效列数
    目标列数 = len(sheet[0])
    return (目标列数)


def GetTargetCol(sheet, rowNum, key):  # 获得指定的key的列号，如果找不到就返回0
    for colCount in range(0, GetDataColumns(sheet)):
        if (key == sheet[rowNum - 1][colCount]):
            break
    if (colCount >= GetDataColumns(sheet) - 1):
        colCount = 0
    return (colCount)


def 拼出一行来(技能表, 表名, 行号, keylist, key号):
    matrix = []
    a = 技能表[行号 + 2][0]  # 角色id列
    b = keylist[key号]  # 通过key号获得key名
    c = GetTargetCol(技能表, 3, b)  # 通过key名获得列号

    matrix.append(表名)  # 表名 占位
    matrix.append(a)  # 角色id
    matrix.append(b)  # 技能类型
    if (c == 0):
        matrix.append("")  # 技能id组 占位
    else:
        d = 技能表[行号 + 2][c]  # 技能id组
        matrix.append(d)  # 技能id组 占位
    return matrix

def 拼出一块来(技能表, 表名, keylist):
    matrix = []
    a = GetDataRows(技能表)
    b = len(keylist)+1
    for c in range(0, a):
        for d in range(0, b - 1):
            e = 拼出一行来(技能表, 表名, c + 1, keylist, d)
            matrix.append(e)
    return matrix

def 多块拼起来(sheetlist, namelist, keylist):
    matrix = []
    a = len(sheetlist)
    for b in range(0, a):
        表名 = namelist[b]
        c = 拼出一块来(sheetlist[b], 表名, keylist)
        matrix.extend(c)
    return matrix


pass
f = 多块拼起来(sheetlist, namelist, keylist)

pass

# ————————————————整表完成，后面开始拓表————————————————

总行数 = len(f)
总列数 = len(f[0])
新行数 = 0
汇总表 = []

for a in range(0, 总行数 - 1):
    单行组 = f[a]
    项 = 单行组[总列数-1] #最后一列是项
    if (isinstance(项, str)):
        项组 = 项.split("|")
    else:
        if (项 == ""):
            项组 = [0]
        else:
            项组 = [项]

    if (isinstance(项组, list)):
        项数 = len(项组)
    else:
        项数 = 1

    pass

    拼接表 = []
    for b in range(0, 项数):
        拼接表1 = []
        拼接表1.append(单行组[0])
        拼接表1.append(单行组[1])
        拼接表1.append(单行组[2])
        拼接表1.append(项组[b])

        拼接表.append(拼接表1)
        pass
    汇总表.extend(拼接表)
    pass

pass
xw.sheets("角色技能分类汇总_程序导出").range('A1').value = 汇总表
print(len(汇总表))
print('hello world')
