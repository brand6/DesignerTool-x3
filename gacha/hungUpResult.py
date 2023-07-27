import random

import xlwings as xw

from common import common


def main():
    programPath = common.getTablePath()

    hungWb = xw.books.open(programPath + r"\HangUpReward.xlsx", update_links=False)
    hungSht = hungWb.sheets["HangUpReward"]
    hungData = hungSht.used_range.value
    h_id_col = common.getDataColOrder(hungData, "ExploreID", 2)
    h_reward_col = common.getDataColOrder(hungData, "Reward", 2)
    h_exRewardProb_col = common.getDataColOrder(hungData, "ExRewardProb", 2)
    h_exReward_col = common.getDataColOrder(hungData, "ExReward", 2)
    h_exGuarantee_col = common.getDataColOrder(hungData, "GuaranteeTimes", 2)

    normalDrop = common.getRowData(1000, h_id_col, h_reward_col, hungData)
    nTotalWeight, nDropList, nWeightList = getDropList(normalDrop)
    normalERate = common.getRowData(1000, h_id_col, h_exRewardProb_col, hungData)
    if normalERate is not None:
        normalEDrop = common.getRowData(1000, h_id_col, h_exReward_col, hungData)
        normalEGuarantee = common.getRowData(
            1000, h_id_col, h_exGuarantee_col, hungData
        )
        nGuaranteeList = [0] * 6

    advanceDrop = common.getRowData(2000, h_id_col, h_reward_col, hungData)
    aTotalWeight, aDropList, aWeightList = getDropList(advanceDrop)
    advanceERate = common.getRowData(2000, h_id_col, h_exRewardProb_col, hungData)
    if advanceERate is not None:
        advanceEDrop = common.getRowData(2000, h_id_col, h_exReward_col, hungData)
        advanceEGuarantee = common.getRowData(
            2000, h_id_col, h_exGuarantee_col, hungData
        )
        aGuaranteeList = [0] * 6

    dropWb = xw.books.open(programPath + r"\Drop.xlsx", update_links=False)
    dropSht = dropWb.sheets["Drop"]
    dropData = dropSht.used_range.value
    d_id_col = common.getDataColOrder(dropData, "GroupID", 2)
    d_skip_col = common.getDataColOrder(dropData, "SkipExport", 2)
    d_item_col = common.getDataColOrder(dropData, "Item", 2)
    d_weight_col = common.getDataColOrder(dropData, "Weight", 2)

    dropMap = {}
    for row in dropData[3:]:
        if row[d_skip_col] is None:
            dropId = int(row[d_id_col])
            if dropId not in dropMap:
                dropMap[dropId] = {}
                dropMap[dropId]["TotalWeight"] = row[d_weight_col]
                dropMap[dropId]["weightList"] = [row[d_weight_col]]
                dropMap[dropId]["itemList"] = [row[d_item_col]]

            else:
                dropMap[dropId]["TotalWeight"] += row[d_weight_col]
                dropMap[dropId]["weightList"].append(row[d_weight_col])
                dropMap[dropId]["itemList"].append(row[d_item_col])

    calcWb = xw.books["资源投放统计.xlsm"]

    calcDataSht = calcWb.sheets["数据源"]
    calcData = calcDataSht.used_range.value
    cC_cardInfo_cols = common.getColBy3Para(
        "Card.xlsx", "CardBaseInfo", ["ID", "Quality"], calcData
    )
    cardMap = {}  # 存储卡牌品质
    for row in calcData[3:]:
        cardMap[row[cC_cardInfo_cols[0]]] = row[cC_cardInfo_cols[1]]

    cC_cardRare_col = common.getColBy3Para(
        "Card.xlsx",
        "CardRare",
        ["Id", "BreakReward", "FragmentNum", "ComposeMoneyCost"],
        calcData,
    )
    cardRareMap = {}
    for row in calcData[3:]:
        cardRareMap[row[cC_cardRare_col[0]]] = [
            row[cC_cardRare_col[1]],
            row[cC_cardRare_col[2]],
            row[cC_cardRare_col[3]],
        ]

    calcDrawSht = calcWb.sheets["思念统计"]
    calcDrawData = calcDrawSht.used_range.value
    cD_maxRow = len(calcDrawData)
    cD_maxCol = len(calcDrawData[0])
    cD_cardNum_col = common.getColBy2Para("免费思念进阶", "SSR", calcDrawData)
    cD_hangNum_col = common.getColBy2Para("普通挂机券数量", "免费", calcDrawData)
    cD_res_col = common.getColBy2Para("溢出资源", "免费", calcDrawData)

    cells = calcDrawSht.cells
    dayList = calcDrawSht.range(cells(3, 1), cells(cD_maxRow, 1)).value
    cardNumList = calcDrawSht.range(
        cells(3, cD_cardNum_col + 1), cells(cD_maxRow, cD_maxCol)
    ).value
    hangNumList = calcDrawSht.range(
        cells(3, cD_hangNum_col + 1), cells(cD_maxRow, cD_hangNum_col + 12)
    ).value
    resList = calcDrawSht.range(
        cells(3, cD_res_col + 1), cells(cD_maxRow, cD_res_col + 6)
    ).value

    resultMap = {}
    resMap = {}
    for player in range(6):
        resultMap[player] = {}
        resMap[player] = {}
        for i in range(2, 5):
            resultMap[player][i] = {}
        for r in range(len(dayList)):
            if dayList[r] is not None:
                nTimes = int(hangNumList[r][player])
                for i in range(nTimes):
                    if normalERate is not None:  #
                        nGuaranteeList[player] += 1
                        getE = False
                        if nGuaranteeList[player] >= normalEGuarantee:
                            nGuaranteeList[player] = 0
                            getE = True
                        else:
                            rndValue = random.randint(1, 1000)
                            if rndValue < normalERate:
                                nGuaranteeList[player] = 0
                                getE = True

                        if getE is True:
                            item = getRandomResult(
                                dropMap[normalEDrop]["TotalWeight"],
                                dropMap[normalEDrop]["weightList"],
                                dropMap[normalEDrop]["itemList"],
                            )
                            if item == "75=601=1":
                                hangNumList[r][player + 6] += 1
                            else:
                                addCardToMap(resultMap, item, cardMap, player)
                    nDropId = getRandomResult(nTotalWeight, nWeightList, nDropList)
                    item = getRandomResult(
                        dropMap[nDropId]["TotalWeight"],
                        dropMap[nDropId]["weightList"],
                        dropMap[nDropId]["itemList"],
                    )
                    addCardToMap(resultMap, item, cardMap, player)

                aTimes = int(hangNumList[r][player + 6])
                for i in range(aTimes):
                    if advanceERate is not None:  #
                        aGuaranteeList[player] += 1
                        getE = False
                        if aGuaranteeList[player] >= advanceEGuarantee:
                            aGuaranteeList[player] = 0
                            getE = True
                        else:
                            rndValue = random.randint(1, 1000)
                            if rndValue < advanceERate:
                                aGuaranteeList[player] = 0
                                getE = True

                        if getE is True:
                            item = getRandomResult(
                                dropMap[advanceEDrop]["TotalWeight"],
                                dropMap[advanceEDrop]["weightList"],
                                dropMap[advanceEDrop]["itemList"],
                            )
                            addCardToMap(resultMap, item, cardMap, player)
                    aDropId = getRandomResult(aTotalWeight, aWeightList, aDropList)
                    item = getRandomResult(
                        dropMap[aDropId]["TotalWeight"],
                        dropMap[aDropId]["weightList"],
                        dropMap[aDropId]["itemList"],
                    )
                    addCardToMap(resultMap, item, cardMap, player)

                # 将表内card加入map
                for i in range(3):
                    cardListStr = cardNumList[r][player * 3 + i]
                    if cardListStr is not None:
                        addCardToMap(resultMap, cardListStr, cardMap, player)

                # 统计
                for i in range(2, 5):
                    writeStr = ""
                    for card in resultMap[player][i].keys():
                        cardNum = resultMap[player][i][card]
                        if cardNum > 4:
                            resultMap[player][i][card] = 4
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
                            turnReward = cardRareMap[i][0]
                            resType, resId, resValue = str.split(turnReward, "=")
                            resKey = resType + "=" + resId
                            if resKey not in resMap[player]:
                                resMap[player][resKey] = int(resValue) * turnNum
                            else:
                                resMap[player][resKey] += int(resValue) * turnNum
                    offset = player * 3 + 4 - i
                    cardNumList[r][offset] = writeStr

                resStr = ""
                for key in resMap[player]:
                    if resStr == "":
                        resStr = key + "=" + str(resMap[player][key])
                    else:
                        resStr = resStr + "|" + key + "=" + str(resMap[player][key])
                resList[r][player] = resStr
    calcDrawSht.range(
        cells(3, cD_cardNum_col + 1), cells(cD_maxRow, cD_maxCol)
    ).value = cardNumList
    calcDrawSht.range(
        cells(3, cD_hangNum_col + 1), cells(cD_maxRow, cD_hangNum_col + 12)
    ).value = hangNumList
    calcDrawSht.range(
        cells(3, cD_res_col + 1), cells(cD_maxRow, cD_res_col + 6)
    ).value = resList


def addCardToMap(cardNumMap, itemStr, cardMap, player):
    itemList = str.split(itemStr, "|")
    for item in itemList:
        itemList = str.split(item, "=")
        if len(itemList) == 2:
            itemId = int(itemList[0])
            itemNum = int(itemList[1])
        else:
            itemId = int(itemList[1])
            itemNum = int(itemList[2])

        if itemId < 300000:
            cardId = itemId
        else:
            cardId = itemId - 200000

        cardQuality = int(cardMap[cardId])
        if itemId in cardNumMap[player][cardQuality]:
            cardNumMap[player][cardQuality][itemId] += itemNum
        else:
            cardNumMap[player][cardQuality][itemId] = itemNum


def getRandomResult(totalWeight, weightList, dropIdList):
    rndValue = random.randint(1, totalWeight)
    for i in range(len(weightList)):
        if rndValue <= weightList[i]:
            return dropIdList[i]
        else:
            rndValue -= weightList[i]


def getDropList(dropStr):
    totalWeight = 0
    weightList = []
    dropList = []
    strList = str.split(dropStr, "|")
    for dStr in strList:
        drop, weight = str.split(dStr, "=")
        weight = int(weight)
        totalWeight += weight
        weightList.append(weight)
        dropList.append(int(drop))

    return totalWeight, dropList, weightList
