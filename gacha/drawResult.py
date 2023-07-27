import random

import numpy as np
import xlwings as xw

from common import common


def main():
    programPath = common.getTablePath()
    calcWb = xw.books["资源投放统计.xlsm"]
    gachaWb = xw.books.open(programPath + r"\Gacha.xlsx", update_links=False)

    calcDrawSht = calcWb.sheets["抽卡统计"]
    calcDrawData = np.array(calcDrawSht.used_range.value)
    cD_times_cols = common.getColBy2Para(
        "抽取次数", ["免费", "微R", "小R", "中R", "大R", "超R"], calcDrawData
    )
    cD_startWrite_col = common.getColBy2Para("免费思念进阶", "SSR", calcDrawData)
    cD_startCount_col = common.getColBy2Para("免费思念数量", "SSR", calcDrawData)
    cD_startRes_col = common.getColBy2Para("抽卡附属资源", "免费", calcDrawData)
    cD_gacha_Col = common.getDataColOrder(calcDrawData, "卡池ID", 1)
    writeData = []  # 输出的card数组
    countData = []  # 输出的计数数组
    resData = []  # 输出的资源数组
    for i in range(len(calcDrawData) - 2):
        writeData.append([None] * (len(cD_times_cols) * 3))
        countData.append([None] * (len(cD_times_cols) * 3))
        resData.append([None] * len(cD_times_cols))

    calcDataSht = calcWb.sheets["数据源"]
    calcData = np.array(calcDataSht.used_range.value)
    cC_cardInfo_cols = common.getColBy3Para(
        "Card.xlsx", "CardBaseInfo", ["ID", "Quality"], calcData
    )
    cardMap = {}  # 存储卡牌品质
    for row in calcData[3:]:
        cardMap[row[cC_cardInfo_cols[0]]] = row[cC_cardInfo_cols[1]]

    cC_cardRare_col = common.getColBy3Para(
        "Card.xlsx", "CardRare", ["Id", "BreakReward"], calcData
    )
    cardRareMap = {}
    for row in calcData[3:]:
        cardRareMap[row[cC_cardRare_col[0]]] = row[cC_cardRare_col[1]]

    gachaAllData = gachaWb.sheets["GachaAll"].used_range.value
    gA_id_col = common.getDataOrder(gachaAllData[2], "ID")
    gA_rule_col = common.getDataOrder(gachaAllData[2], "Rule")
    gA_reward_col = common.getDataOrder(gachaAllData[2], "Reward")

    gachaRuleData = gachaWb.sheets["GachaRule"].used_range.value
    gR_id_col = common.getDataOrder(gachaRuleData[2], "ID")
    gR_count_col = common.getDataOrder(gachaRuleData[2], "CountID")
    gR_type_col = common.getDataOrder(gachaRuleData[2], "RuleType")
    gR_para1_col = common.getDataOrder(gachaRuleData[2], "Param1")
    gR_para2_col = common.getDataOrder(gachaRuleData[2], "Param2")
    gR_para3_col = common.getDataOrder(gachaRuleData[2], "Param3")
    gR_group_col = common.getDataOrder(gachaRuleData[2], "RuleTypeGroup")
    gR_priority_col = common.getDataOrder(gachaRuleData[2], "Priority")
    gR_drop_col = common.getDataOrder(gachaRuleData[2], "Drop")

    gachaCountData = gachaWb.sheets["GuaranteeCount"].used_range.value
    gC_id_col = common.getDataOrder(gachaCountData[2], "ID")
    gC_type_col = common.getDataOrder(gachaCountData[2], "CountType")
    gC_para1_col = common.getDataOrder(gachaCountData[2], "Param1")
    countMap = {}  # 存储计数类型
    for row in gachaCountData[3:]:
        countMap[row[gC_id_col]] = {}
        countMap[row[gC_id_col]]["type"] = row[gC_type_col]
        countMap[row[gC_id_col]]["para1"] = str(row[gC_para1_col])

    gachaDropData = gachaWb.sheets["GachaDrop"].used_range.value
    gD_skip_col = common.getDataOrder(gachaDropData[2], "SkipExport")
    gD_group_col = common.getDataOrder(gachaDropData[2], "GroupID")
    gD_weight_col = common.getDataOrder(gachaDropData[2], "Weight")
    gD_item_col = common.getDataOrder(gachaDropData[2], "ItemID")

    dropMap = {}  # 存储掉落组数据
    for row in gachaDropData[3:]:
        if row[gD_skip_col] is None:
            dropId = row[gD_group_col]
            weight = row[gD_weight_col]
            cardId = int(str.split(row[gD_item_col], "=")[1])
            if dropId not in dropMap:
                dropMap[dropId] = {}
                dropMap[dropId]["TotalWeight"] = weight
                dropMap[dropId]["weightList"] = [weight]
                dropMap[dropId]["cardIdList"] = [cardId]

            else:
                dropMap[dropId]["TotalWeight"] += weight
                dropMap[dropId]["weightList"].append(weight)
                dropMap[dropId]["cardIdList"].append(cardId)

    granteeMap = {}  # 存储保底计数
    ruleMap = {}  # 存储规则数据
    orderMap = {}  # 存储规则的优先级
    resultMap = {}  # 存储抽到的卡
    cardCountMap = {}  # 存储抽到的卡数量
    resMap = {}  # 存储抽卡附属资源
    for playerCol in cD_times_cols:
        granteeMap[playerCol] = {}
        resultMap[playerCol] = {}
        cardCountMap[playerCol] = {}
        resMap[playerCol] = {}
        for i in range(2, 5):
            resultMap[playerCol][i] = {}
            cardCountMap[playerCol][i] = 0

        for row in range(2, len(calcDrawData)):
            gachaId = calcDrawData[row][cD_gacha_Col]
            if gachaId is not None:
                gachaReward = common.getRowData(
                    gachaId, gA_id_col, gA_reward_col, gachaAllData
                )
                # 初始化卡池数据
                if gachaId not in orderMap:
                    orderMap[gachaId] = []
                    orderList = []
                    gachaRules = common.getRowData(
                        gachaId, gA_id_col, gA_rule_col, gachaAllData
                    )
                    if gachaRules is not None:
                        rules = str.split(gachaRules, "|")
                        for rule in rules:
                            rule = int(rule)
                            ruleMap[rule] = {}
                            # 解析卡池规则的参数
                            ruleParas = common.getRowData(
                                rule,
                                gR_id_col,
                                [
                                    gR_count_col,
                                    gR_type_col,
                                    gR_para1_col,
                                    gR_para2_col,
                                    gR_para3_col,
                                ],
                                gachaRuleData,
                            )
                            ruleMap[rule]["countId"] = ruleParas[0]
                            ruleMap[rule]["type"] = ruleParas[1]
                            ruleMap[rule]["para1"] = ruleParas[2]
                            ruleMap[rule]["para2"] = ruleParas[3]
                            ruleMap[rule]["para3"] = ruleParas[4]

                            # 解析卡池规则的掉落组
                            ruleDrops = common.getRowData(
                                rule, gR_id_col, gR_drop_col, gachaRuleData
                            )
                            if ruleDrops is not None:
                                drops = str.split(ruleDrops, "|")
                                totalWeight = 0
                                weightList = []
                                dropIdList = []
                                for drop in drops:
                                    dropId, dropWeight = str.split(drop, "=")
                                    dropId = int(dropId)
                                    dropWeight = int(dropWeight)
                                    totalWeight += dropWeight
                                    weightList.append(dropWeight)
                                    dropIdList.append(dropId)

                                ruleMap[rule]["TotalWeight"] = totalWeight
                                ruleMap[rule]["weightList"] = weightList
                                ruleMap[rule]["dropIdList"] = dropIdList

                            else:
                                print("规则id", rule, "的掉落配置无效")

                            # 解析规则的优先级
                            ruleGroup, rulePriority = common.getRowData(
                                rule,
                                gR_id_col,
                                [gR_group_col, gR_priority_col],
                                gachaRuleData,
                            )
                            ruleOrder = ruleGroup * 100 - rulePriority  # 数字小的优先
                            if len(orderList) == 0:
                                orderMap[gachaId].append(rule)
                                orderList.append(ruleOrder)
                            else:
                                for i in range(len(orderList)):
                                    if ruleOrder < orderList[i]:
                                        orderMap[gachaId].insert(i, rule)
                                        orderList.insert(i, ruleOrder)

                    else:
                        print("卡池id", gachaId, "的规则配置无效")

                # 初始化保底计数
                if gachaId not in granteeMap[playerCol]:
                    granteeMap[playerCol][gachaId] = {}
                    granteeMap[playerCol][gachaId]["drawTimes"] = 0  # 该卡池的当前抽取次数
                    for rule in ruleMap.keys():
                        countId = ruleMap[rule]["countId"]
                        if (
                            countId is not None
                            and countId not in granteeMap[playerCol][gachaId]
                        ):
                            granteeMap[playerCol][gachaId][countId] = 0

                # 开始抽取card
                if row == 2:
                    drawTimes = int(calcDrawData[row][playerCol])
                else:
                    drawTimes = int(calcDrawData[row][playerCol]) - int(
                        calcDrawData[row - 1][playerCol]
                    )
                resType, resId, resValue = str.split(gachaReward, "=")
                resKey = resType + "=" + resId
                if resKey not in resMap[playerCol]:
                    resMap[playerCol][resKey] = int(resValue) * drawTimes
                else:
                    resMap[playerCol][resKey] += int(resValue) * drawTimes

                for i in range(drawTimes):
                    granteeMap[playerCol][gachaId]["drawTimes"] += 1
                    dropId = 0
                    for rule in orderMap[gachaId]:
                        countId = ruleMap[rule]["countId"]
                        ruleType = ruleMap[rule]["type"]
                        para1 = ruleMap[rule]["para1"]
                        para2 = ruleMap[rule]["para2"]
                        para3 = ruleMap[rule]["para3"]
                        totalWeight = ruleMap[rule]["TotalWeight"]
                        weightList = ruleMap[rule]["weightList"]
                        dropIdList = ruleMap[rule]["dropIdList"]
                        if countId is not None:
                            granteeCount = granteeMap[playerCol][gachaId][countId]

                        if ruleType == 1:
                            if (
                                granteeMap[playerCol][gachaId]["drawTimes"]
                                == ruleMap[rule]["para1"]
                                and granteeCount < para2
                            ):
                                dropId = getRandomResult(
                                    totalWeight, weightList, dropIdList
                                )
                                break
                        elif ruleType == 2:
                            if granteeCount >= para1:
                                dropId = getRandomResult(
                                    totalWeight, weightList, dropIdList
                                )
                                break
                            pass
                        elif ruleType == 4:
                            if granteeCount >= para1:
                                changeWeight = (granteeCount + 1 - para1) * para2
                                adjustWeight(
                                    totalWeight,
                                    weightList,
                                    dropIdList,
                                    changeWeight,
                                    para3,
                                )
                                dropId = getRandomResult(
                                    totalWeight, weightList, dropIdList
                                )
                                break
                        elif ruleType == 0:
                            dropId = getRandomResult(
                                totalWeight, weightList, dropIdList
                            )
                            break
                        else:
                            print(rule, "规则类型无效，因为类型不是1/2/4！")

                    # 从drop抽取card
                    cardId = getRandomResult(
                        dropMap[dropId]["TotalWeight"],
                        dropMap[dropId]["weightList"],
                        dropMap[dropId]["cardIdList"],
                    )
                    cardQuality = int(cardMap[cardId])
                    cardCountMap[playerCol][cardQuality] += 1
                    if cardId in resultMap[playerCol][cardQuality]:
                        resultMap[playerCol][cardQuality][cardId] += 1
                    else:
                        resultMap[playerCol][cardQuality][cardId] = 1

                    # 更新保底计数
                    cardQuality = str(cardQuality)
                    for countId in granteeMap[playerCol][gachaId].keys():
                        if countId != "drawTimes" and countId != "":
                            if countMap[countId]["type"] == 1:
                                if cardQuality in countMap[countId]["para1"]:
                                    granteeMap[playerCol][gachaId][countId] += 1
                            elif countMap[countId]["type"] == 2:
                                if cardQuality not in countMap[countId]["para1"]:
                                    granteeMap[playerCol][gachaId][countId] += 1
                                else:
                                    granteeMap[playerCol][gachaId][countId] = 0
                            else:
                                print(countId, "保底计数无效，因为类型不是1或2！")

                # 统计
                for i in range(2, 5):
                    writeStr = ""
                    for card in resultMap[playerCol][i].keys():
                        cardNum = resultMap[playerCol][i][card]
                        if cardNum > 4:
                            resultMap[playerCol][i][card] = 4
                            num = 4
                            turnNum = cardNum - 4
                        else:
                            num = cardNum
                            turnNum = 0

                        if writeStr == "":
                            writeStr = str(card) + "=" + str(num)
                        else:
                            writeStr = writeStr + "|" + str(card) + "=" + str(num)

                        if turnNum > 0:
                            turnReward = cardRareMap[i]
                            resType, resId, resValue = str.split(turnReward, "=")
                            resKey = resType + "=" + resId
                            if resKey not in resMap[playerCol]:
                                resMap[playerCol][resKey] = int(resValue) * turnNum
                            else:
                                resMap[playerCol][resKey] += int(resValue) * turnNum
                    offset = (playerCol - cD_times_cols[0]) * 3 + 4 - i
                    writeData[row - 2][offset] = writeStr
                    countData[row - 2][offset] = cardCountMap[playerCol][i]
                resStr = ""
                for key in resMap[playerCol]:
                    if resStr == "":
                        resStr = key + "=" + str(resMap[playerCol][key])
                    else:
                        resStr = resStr + "|" + key + "=" + str(resMap[playerCol][key])
                resData[row - 2][playerCol - cD_times_cols[0]] = resStr
    calcDrawSht.cells(3, cD_startWrite_col + 1).value = writeData
    calcDrawSht.cells(3, cD_startCount_col + 1).value = countData
    calcDrawSht.cells(3, cD_startRes_col + 1).value = resData
    gachaWb.close()


def getRandomResult(totalWeight, weightList, dropIdList):
    rndValue = random.randint(1, totalWeight)
    for i in range(len(weightList)):
        if rndValue <= weightList[i]:
            return dropIdList[i]
        else:
            rndValue -= weightList[i]


def adjustWeight(totalWeight, weightList, dropIdList, changeWeight, upDrop):
    index = -1
    extraWeight = 0
    for i in range(len(dropIdList)):
        if dropIdList[i] == upDrop:
            index = i
            extraWeight = weightList[i]
            weightList[i] += changeWeight
            break
    for i in range(len(weightList)):
        if i != index:
            weightList[i] = (
                weightList[i]
                / (totalWeight - extraWeight)
                * (totalWeight - weightList[index])
            )
