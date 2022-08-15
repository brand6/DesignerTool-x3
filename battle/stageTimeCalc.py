import xlwings as xw
from common import common
from common.printer import Printer

# 脚本说明：
# 用于战斗数值模板表计算关卡强度
#


def main():
    printer = Printer()
    toStr = common.toStr
    isNum = common.isNumber
    getDataOrder = common.getDataOrder

    printer.printColor('~~~开始处理数据~~~', 'green')
    app = xw.apps.active
    try:
        app.screen_updating = False
        app.display_alerts = False
        wb = xw.books.active
        stageSht = wb.sheets['关卡']
        calcSht = wb.sheets['验算']

        stageRange = stageSht.used_range
        sv_columnData = stageRange.rows[1].value
        sv_stageIds = stageRange.columns[getDataOrder(sv_columnData, '关卡id')].value
        sv_stageTypes = stageRange.columns[getDataOrder(sv_columnData, '关卡类型')].value
        sv_pLevs = stageRange.columns[getDataOrder(sv_columnData, '等级')].value
        sv_m1IDs = stageRange.columns[getDataOrder(sv_columnData, 'ID1')].value
        sv_m1Porps = stageRange.columns[getDataOrder(sv_columnData, '属性1')].value
        sv_m1Levs = stageRange.columns[getDataOrder(sv_columnData, '等级1')].value
        sv_m1Nums = stageRange.columns[getDataOrder(sv_columnData, '数量1')].value
        sv_m2IDs = stageRange.columns[getDataOrder(sv_columnData, 'ID2')].value
        sv_m2Porps = stageRange.columns[getDataOrder(sv_columnData, '属性2')].value
        sv_m2Levs = stageRange.columns[getDataOrder(sv_columnData, '等级2')].value
        sv_m2Nums = stageRange.columns[getDataOrder(sv_columnData, '数量2')].value
        sv_m3IDs = stageRange.columns[getDataOrder(sv_columnData, 'ID3')].value
        sv_m3Porps = stageRange.columns[getDataOrder(sv_columnData, '属性3')].value
        sv_m3Levs = stageRange.columns[getDataOrder(sv_columnData, '等级3')].value
        sv_m3Nums = stageRange.columns[getDataOrder(sv_columnData, '数量3')].value

        calcRange = calcSht.used_range
        cv_columnData = calcRange.rows[1].value
        c_index_col = getDataOrder(cv_columnData, '索引')
        c_stageId_col = getDataOrder(cv_columnData, '关卡')
        c_stageType_col = getDataOrder(cv_columnData, '关卡类型')
        c_mID_col = getDataOrder(cv_columnData, 'ID2')
        c_indexs = calcRange.columns[c_index_col]

        # 查找所在行，未查到时返回插入位置
        def getDataRow(findRange, findValue):
            findData = findRange.value
            insertRow = 1
            for i in range(len(findData)):
                if findData[i] == findValue:
                    return i, insertRow
                elif isNum(findData[i]) and findData[i] < findValue:
                    insertRow = i + 2
            else:
                return -1, insertRow

        def updateCalcRange():
            global calcRange
            global c_indexs
            calcRange = calcSht.used_range
            c_indexs = calcRange.columns[c_index_col]

        # 验算表插入新的数据
        def addCalcData(stageId, monsterID, stageType):
            c_row, insertRow = getDataRow(c_indexs, stageId * 100000 + monsterID)
            if c_row == -1:
                print('Insert Line' + toStr(insertRow))
                calcSht.api.Rows(insertRow).Insert()
                calcSht.api.Rows(insertRow - 1).Copy(calcSht.api.Rows(insertRow))
                c_row = insertRow - 1
                updateCalcRange()
                calcRange.columns[c_stageId_col][c_row].value = stageId
                calcRange.columns[c_mID_col][c_row].value = monsterID
                calcRange.columns[c_stageType_col][c_row].value = stageType

        printer.setStartTime("正在检查验算表数据行是否完整...")
        # 检测验算表数据行是否完整
        for s_row in range(2, len(stageRange.rows)):
            if sv_m1IDs[s_row] is not None:
                addCalcData(sv_stageIds[s_row], sv_m1IDs[s_row], sv_stageTypes[s_row])
            if sv_m2IDs[s_row] is not None:
                addCalcData(sv_stageIds[s_row], sv_m2IDs[s_row], sv_stageTypes[s_row])
            if sv_m3IDs[s_row] is not None:
                addCalcData(sv_stageIds[s_row], sv_m3IDs[s_row], sv_stageTypes[s_row])
        printer.setCompareTime(printer.printGapTime("验算表数据行处理完毕，耗时:"))

        c_pLev_col = getDataOrder(cv_columnData, '等级1')
        c_mProp_col = getDataOrder(cv_columnData, '属性2')
        c_mLev_col = getDataOrder(cv_columnData, '等级2')
        cv_pLevs = calcRange.columns[c_pLev_col].value
        cv_mProps = calcRange.columns[c_mProp_col].value
        cv_mLevs = calcRange.columns[c_mLev_col].value

        # 验算表数据输入
        def updateCalcData(stageId, monsterID, pLev, mProp, mLev):  # yapf:disable
            c_row, _ = getDataRow(c_indexs, stageId * 100000 + monsterID)
            cv_pLevs[c_row] = pLev
            cv_mProps[c_row] = mProp
            cv_mLevs[c_row] = mLev

        printer.setStartTime("正在更新验算表数据...")
        # 更新验算表数据
        for s_row in range(2, len(stageRange.rows)):
            if sv_m1IDs[s_row] is not None:
                updateCalcData(sv_stageIds[s_row], sv_m1IDs[s_row], sv_pLevs[s_row], sv_m1Porps[s_row], sv_m1Levs[s_row])
            if sv_m2IDs[s_row] is not None:
                updateCalcData(sv_stageIds[s_row], sv_m2IDs[s_row], sv_pLevs[s_row], sv_m2Porps[s_row], sv_m2Levs[s_row])
            if sv_m3IDs[s_row] is not None:
                updateCalcData(sv_stageIds[s_row], sv_m3IDs[s_row], sv_pLevs[s_row], sv_m3Porps[s_row], sv_m3Levs[s_row])

        calcSht.cells(1, c_pLev_col + 1).options(transpose=True).value = cv_pLevs
        calcSht.cells(1, c_mProp_col + 1).options(transpose=True).value = cv_mProps
        calcSht.cells(1, c_mLev_col + 1).options(transpose=True).value = cv_mLevs

        printer.printGapTime("验算表数据处理完毕，耗时:")

        printer.setStartTime("正在更新关卡怪物时长...")
        sv_m1Waves = stageRange.columns[getDataOrder(sv_columnData, '波次1')].value
        sv_m2Waves = stageRange.columns[getDataOrder(sv_columnData, '波次2')].value
        sv_m3Waves = stageRange.columns[getDataOrder(sv_columnData, '波次3')].value

        s_m1Hurt_col = getDataOrder(sv_columnData, '输出1')
        s_m1Time_col = getDataOrder(sv_columnData, '生存1')
        s_m2Hurt_col = getDataOrder(sv_columnData, '输出2')
        s_m2Time_col = getDataOrder(sv_columnData, '生存2')
        s_m3Hurt_col = getDataOrder(sv_columnData, '输出3')
        s_m3Time_col = getDataOrder(sv_columnData, '生存3')
        s_allHurt_col = getDataOrder(sv_columnData, '总输出')
        s_allTime_col = getDataOrder(sv_columnData, '总时长')
        sv_m1Hurts = stageRange.columns[s_m1Hurt_col].value
        sv_m1Times = stageRange.columns[s_m1Time_col].value
        sv_m2Hurts = stageRange.columns[s_m2Hurt_col].value
        sv_m2Times = stageRange.columns[s_m2Time_col].value
        sv_m3Hurts = stageRange.columns[s_m3Hurt_col].value
        sv_m3Times = stageRange.columns[s_m3Time_col].value
        sv_allHurts = stageRange.columns[s_allHurt_col].value
        sv_allTimes = stageRange.columns[s_allTime_col].value

        c_pTimeLows = calcRange.columns[getDataOrder(cv_columnData, '低系数时长1')].value
        cv_mTimeLows = calcRange.columns[getDataOrder(cv_columnData, '低系数时长2')].value

        # 返回怪物生存时间，怪物dps
        def getMonsterTime(stageId, monsterID):
            c_row, _ = getDataRow(c_indexs, stageId * 100000 + monsterID)
            return c_pTimeLows[c_row], round(1.0 / cv_mTimeLows[c_row], 4)

        # 更新关卡怪物时长
        for s_row in range(2, len(stageRange.rows)):
            if sv_m1IDs[s_row] is not None:
                sv_m1Times[s_row], sv_m1Hurts[s_row] = getMonsterTime(sv_stageIds[s_row], sv_m1IDs[s_row])
            if sv_m2IDs[s_row] is not None:
                sv_m2Times[s_row], sv_m2Hurts[s_row] = getMonsterTime(sv_stageIds[s_row], sv_m2IDs[s_row])
            if sv_m3IDs[s_row] is not None:
                sv_m3Times[s_row], sv_m3Hurts[s_row] = getMonsterTime(sv_stageIds[s_row], sv_m3IDs[s_row])

            stageTime = 0
            stageHurt = 0
            for i in range(1, 4):
                waveTime = 0
                waveNum = 0
                if sv_m1Waves[s_row] == i:
                    for _ in range(int(sv_m1Nums[s_row])):
                        waveTime += sv_m1Times[s_row] / 2 * 0.7**waveNum
                        waveNum += 1
                if sv_m2Waves[s_row] == i:
                    for _ in range(int(sv_m2Nums[s_row])):
                        waveTime += sv_m2Times[s_row] / 2 * 0.7**waveNum
                        waveNum += 1
                if sv_m3Waves[s_row] == i:
                    for _ in range(int(sv_m3Nums[s_row])):
                        waveTime += sv_m3Times[s_row] / 2 * 0.7**waveNum
                        waveNum += 1
                waveHurt = 0
                if sv_m1Waves[s_row] == i:
                    waveHurt += sv_m1Hurts[s_row] * sv_m1Nums[s_row] * waveTime
                if sv_m2Waves[s_row] == i:
                    waveHurt += sv_m2Hurts[s_row] * sv_m2Nums[s_row] * waveTime
                if sv_m3Waves[s_row] == i:
                    waveHurt += sv_m3Hurts[s_row] * sv_m3Nums[s_row] * waveTime
                stageTime += waveTime
                stageHurt += waveHurt
            sv_allHurts[s_row] = round(stageHurt, 2)
            sv_allTimes[s_row] = stageTime

        stageSht.cells(1, s_m1Hurt_col + 1).options(transpose=True).value = sv_m1Hurts
        stageSht.cells(1, s_m1Time_col + 1).options(transpose=True).value = sv_m1Times
        stageSht.cells(1, s_m2Hurt_col + 1).options(transpose=True).value = sv_m2Hurts
        stageSht.cells(1, s_m2Time_col + 1).options(transpose=True).value = sv_m2Times
        stageSht.cells(1, s_m3Hurt_col + 1).options(transpose=True).value = sv_m3Hurts
        stageSht.cells(1, s_m3Time_col + 1).options(transpose=True).value = sv_m3Times
        stageSht.cells(1, s_allHurt_col + 1).options(transpose=True).value = sv_allHurts
        stageSht.cells(1, s_allTime_col + 1).options(transpose=True).value = sv_allTimes
        printer.printGapTime("关卡怪物时长处理完毕，耗时:")

        printer.printColor('~~~所有数据处理完毕~~~', 'green')
    finally:
        app.screen_updating = True
        app.display_alerts = True
