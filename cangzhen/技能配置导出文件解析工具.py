
import csv
import xlwings as xw
filePath = 'F:\\P4_x3_project_batlemain\\Program\\TimelineCsv\\技能配置导出.csv'
outPath = 'F:\\P4_x3_project_batlemain\\Program\\TimelineCsv\\skillData.xlsx'
# 读取文件，以二维列表形式存入data
data = []
with open(filePath, 'r', encoding='utf-8') as f:
    reader = csv.reader(f)
    data = []
    for row in reader:
        data.append(row)
pass

column = len(data[0])
名称列 = data[0]

列_技能时长 = 名称列.index('技能时长')
列_连招区间 = 名称列.index('连招区间')
列_对应连招开始时间 = 名称列.index('对应连招开始时间')
列_打断区间 = 名称列.index('打断区间')
列_打断区间开始时间 = 名称列.index('打断区间开始时间')

pass

data数据区 = data[1:]

连招名称列表 = []
打断名称列表 = []

def 拆掉list每一项的前n个子分隔符(list, 子分隔符, 拆几个): #a=b=c|a=b=c
    拼接用list = []
    result = ''
    for a in list: # a=b=c,a=b=c
        list_1 = a.split(子分隔符) # a,b,c
        list_1.pop(拆几个-1)  # b,c
        result = '='.join(list_1) #b=c
        拼接用list.append(result) #b=c|b=c
    return(拼接用list)

def 对list查询字符并用字符分装(list,表头字符,子分隔符):
    输出list=[]
    for a in list:
        if 表头字符 in a:  # Skill=Attack;Fill;Active;Passive;Coop;Support;Dodge;Ultra;Gemcore;ScorePhase;Card
            # 有等号，要加前缀
            # 拆出等号前后
            b = a.split(表头字符)
            prefix = b[0]  # skill
            suffix = b[1]  #attack;dash;power
            suffixList = suffix.split(子分隔符) # (attack,dash,power)
            拼接用list = []
            for c in suffixList:
               拼接用list.append(f'{prefix}={c}') #(skill=attack,skill=dash,skill=power
        else:    # Attack;Fill
            # 无等号，不用加前缀
            拼接用list = a.split(子分隔符)

        输出list.append(拼接用list)
    return(输出list)


    pass

def 一个二级list一个一级list对齐拓为一级list(list1,list2,list1_1,list2_1):
    pass
    if list1 == [['']]:
        pass
    else:
        for a,b in zip(list1,list2):
            for c in a:
                list1_1.append(c)
                list2_1.append(b)

def 两list对齐后查询插入新表list(list1,list2,targetList1,targetList2):
    if list1 == ['']:
        pass
    else:
        for i, j in zip(list1, list2):
            if i in targetList1: # 已有表头
                k = list1.index(i) # 取位置
                if j < targetList2[k]: # 和已有数据比大小
                    targetList2[k] = j # 小就替换
                else:
                    pass
                    # TODO:
            else:
                targetList1.append(i) # 无已有表头，直接屁股上加表头
                targetList2.append(j) # 无已有表头，直接屁股上加数据
        pass

def 去重去大拼接函数(表头清单, 表头数据组, 数据数据组):
    二维盒子 = []
    for i, j in zip(表头数据组, 数据数据组): # 行
        for m, n in zip(i, j): # 项
            一维盒子 = []
            if m in 表头清单:
                pass
            else:
                pass

def getTitleCol(titleList, name):
    for i in range(len(titleList)):
        if titleList[i] == name:
            return i

# ——————————————————————————————————————————————————————————————————

新增连招表头行 = []
新增打断表头行 = []

连招表头数据组 = []
连招数据数据组 = []
打断表头数据组 = []
打断数据数据组 = []
pass

colList = ['id']
writeList = []
for row in data数据区:
    rowList = [0]*len(colList)
    rowList[0] = row[0]
    # 第一组拆解
    连招区间list = row[列_连招区间].split('|')
    连招区间list_1 = 拆掉list每一项的前n个子分隔符(连招区间list, '=', 1)
    对应连招开始时间list = row[列_对应连招开始时间].split('|')

    for l in range(len(连招区间list_1)):
        if 连招区间list_1[l] not in colList:
            colList.append(连招区间list_1[l])
            time = float(对应连招开始时间list[l]) if 对应连招开始时间list[l] != '' else 0
            rowList.append(time)
        else:
            index = getTitleCol(colList,连招区间list_1[l])
            time = float(对应连招开始时间list[l]) if 对应连招开始时间list[l] != '' else 0
            if rowList[index] > time and 0 < time or rowList[index] == 0:
                rowList[index] = time


    连招表头数据组.append(连招区间list_1)
    连招数据数据组.append(对应连招开始时间list)

    pass

    # 第二组拆解
    打断区间list = row[列_打断区间].split('|')
    打断区间list_1 = 对list查询字符并用字符分装(打断区间list,'=',';')
    打断区间开始时间list = row[列_打断区间开始时间].split('|')
    打断区间list_2 = []
    打断区间开始时间list_1 = []
    一个二级list一个一级list对齐拓为一级list(打断区间list_1,打断区间开始时间list,打断区间list_2,打断区间开始时间list_1)

    for l in range(len(打断区间list_2)):
        if 打断区间list_2[l] not in colList:
            colList.append(打断区间list_2[l])
            time = float(打断区间开始时间list_1[l]) if 打断区间开始时间list_1[l] != '' else 0
            rowList.append(time)
        else:
            index = getTitleCol(colList,打断区间list_2[l])
            time = float(打断区间开始时间list_1[l]) if 打断区间开始时间list_1[l] != '' else 0
            if rowList[index] > time and 0 < time or rowList[index] == 0:
                rowList[index] = time
    writeList.append(rowList)
    打断表头数据组.append(打断区间list_2)
    打断数据数据组.append(打断区间开始时间list_1)
    pass

pass
outList = []
outList.append(colList)
for l in writeList:
    oList = [0]*len(colList)
    for j in range(len(l)):
        oList[j] = l[j]
    outList.append(oList)

xl = xw.books.open(outPath)
sht = xl.sheets[0]
sht.range('A1').value = outList
xl.save()

pass

# 散行数据准备好，开始拼接
