import xlwings as xw
from common.printer import Printer
from common import common

getDataOrder = common.getDataOrder
toStr = common.toStr
toInt = common.toInt


# 脚本说明：
# 用于牵绊度等级奖励配置
#
def main():
    isClear = True  # 是否清除旧数据

    printer = Printer()
    tablePath = common.getTablePath()
    wbChangeMap = {}  # 保存改变了来源的表格

    app = xw.App(visible=False, add_book=False)
    try:
        itemMap = {}
        rewardTypeData = xw.books.active.sheets['奖励描述'].used_range.value
        for r in range(1, len(rewardTypeData)):
            if rewardTypeData[r][0] is not None:
                d = {}
                for c in range(1, len(rewardTypeData[0])):
                    d[rewardTypeData[0][c]] = rewardTypeData[r][c]
                itemMap[rewardTypeData[r][0]] = d

        # region 数据准备
        app.screen_updating = False
        app.display_alerts = False
        dataSht = xw.sheets.active
        dataRng = dataSht.used_range
        dv_columnData = dataRng.rows[0].value
        d_lv_col = getDataOrder(dv_columnData, '等级')
        d_itemType_col = getDataOrder(dv_columnData, '道具奖励Type')
        d_itemID_col = getDataOrder(dv_columnData, '道具奖励ID')
        d_itemNum_col = getDataOrder(dv_columnData, '道具奖励数量')
        d_ResType1_col = getDataOrder(dv_columnData, '解锁内容1')
        d_ResType2_col = getDataOrder(dv_columnData, '解锁内容2')
        d_ResType3_col = getDataOrder(dv_columnData, '解锁内容3')
        d_STResID_col = getDataOrder(dv_columnData, 'ST解锁内容ID')
        d_YSResID_col = getDataOrder(dv_columnData, 'YS解锁内容ID')
        d_3ResID_col = getDataOrder(dv_columnData, '3解锁内容ID')
        d_4ResID_col = getDataOrder(dv_columnData, '4解锁内容ID')
        d_RYResID_col = getDataOrder(dv_columnData, 'RY解锁内容ID')

        printer.setStartTime("开始打开相关表格...", 'green')

        for k in itemMap:
            wbName = itemMap[k]['奖励表']
            if wbName is not None and not isBookOpen(app, wbName):
                app.books.open(tablePath + "\\" + wbName, 0)
        loveWb = xw.books.open(tablePath + "\\" + "LovePointLevel.xlsx", 0)
        printer.setCompareTime(printer.printGapTime("表格打开完毕，耗时:"))
        loveSht = loveWb.sheets['LovePointReward']
        loveRange = loveSht.used_range

        lv_columnData = loveSht.used_range.rows[2].value
        lv_ID = loveRange.columns[getDataOrder(lv_columnData, 'ID')].value
        l_drop_col = getDataOrder(lv_columnData, 'Drop')
        l_reward_col = getDataOrder(lv_columnData, 'Reward')
        l_desc1_col = getDataOrder(lv_columnData, 'ExtraRewardDes1')
        l_desc2_col = getDataOrder(lv_columnData, 'ExtraRewardDes2')
        l_desc3_col = getDataOrder(lv_columnData, 'ExtraRewardDes3')
        l_unknow_col = getDataOrder(lv_columnData, 'Unknown1')
        l_unknow2_col = getDataOrder(lv_columnData, 'Unknown2')
        l_unknow3_col = getDataOrder(lv_columnData, 'Unknown3')
        lv_reward = loveRange.columns[l_reward_col].value
        lv_drop = loveRange.columns[l_drop_col].value
        lv_desc1 = loveRange.columns[l_desc1_col].value
        lv_desc2 = loveRange.columns[l_desc2_col].value
        lv_desc3 = loveRange.columns[l_desc3_col].value
        lv_unknow = loveRange.columns[l_unknow_col].value
        lv_unknow2 = loveRange.columns[l_unknow2_col].value
        lv_unknow3 = loveRange.columns[l_unknow3_col].value
        if isClear:
            for i in range(3, len(lv_reward)):
                lv_drop[i] = None
                lv_reward[i] = None
                lv_desc1[i] = None
                lv_unknow[i] = None
                lv_desc2[i] = None
                lv_unknow2[i] = None
                lv_desc3[i] = None
                lv_unknow3[i] = None
        # endregion

        def updateList(id, rewards):
            """更新等级奖励，rewards内包含多个reward字典

            Args:
                id (int): ID
                rewards (list): [{reward,desc,type,order}]
            """

            idRow = common.getListRow(lv_ID, id)
            for i in range(len(rewards)):
                r = rewards[i]
                if r['type'] != -1:  # 普通掉落
                    if lv_reward[idRow] is None or lv_reward[idRow] == '':
                        lv_reward[idRow] = r['reward']
                    elif r['reward'] is not None and r['reward'] != '':
                        lv_reward[idRow] = lv_reward[idRow] + '|' + r['reward']
                    if r['order'] == 0:
                        lv_desc1[idRow] = r['desc']
                    elif r['order'] == 1:
                        lv_desc2[idRow] = r['desc']
                    else:
                        lv_desc3[idRow] = r['desc']
                else:  # 随机掉落
                    if lv_drop[idRow] is None or lv_drop[idRow] == '':
                        lv_drop[idRow] = r['reward']
                    elif r['reward'] is not None and r['reward'] != '':
                        lv_drop[idRow] = lv_drop[idRow] + '|' + r['reward']
                    if r['order'] == 0:
                        lv_desc1[idRow] = r['desc'] + '：{0}'
                        lv_unknow[idRow] = r['desc']
                    elif r['order'] == 1:
                        lv_desc2[idRow] = r['desc'] + '：{0}'
                        lv_unknow2[idRow] = r['desc']
                    else:
                        lv_desc3[idRow] = r['desc'] + '：{0}'
                        lv_unknow3[idRow] = r['desc']

        printer.setStartTime("正在处理数据...")
        rowCount = -1
        for row in dataRng.value:
            rowCount += 1
            printer.printProgress(rowCount, len(dataRng.value) - 1)
            if not common.isNumber(row[d_lv_col]):
                continue
            for i in range(1, 6):
                lev = toInt(row[d_lv_col])
                id = i * 1000 + lev
                rewards = []
                # 资源奖励
                if i == 1:
                    rewards = getRewardData(itemMap, wbChangeMap, row[d_ResType1_col], row[d_ResType2_col], row[d_ResType3_col],
                                            row[d_STResID_col], row[d_itemType_col], row[d_itemID_col], row[d_itemNum_col], i,
                                            lev)
                elif i == 2:
                    rewards = getRewardData(itemMap, wbChangeMap, row[d_ResType1_col], row[d_ResType2_col], row[d_ResType3_col],
                                            row[d_YSResID_col], row[d_itemType_col], row[d_itemID_col], row[d_itemNum_col], i,
                                            lev)
                elif i == 3:
                    rewards = getRewardData(itemMap, wbChangeMap, row[d_ResType1_col], row[d_ResType2_col], row[d_ResType3_col],
                                            row[d_3ResID_col], row[d_itemType_col], row[d_itemID_col], row[d_itemNum_col], i,
                                            lev)
                elif i == 4:
                    rewards = getRewardData(itemMap, wbChangeMap, row[d_ResType1_col], row[d_ResType2_col], row[d_ResType3_col],
                                            row[d_4ResID_col], row[d_itemType_col], row[d_itemID_col], row[d_itemNum_col], i,
                                            lev)
                else:
                    rewards = getRewardData(itemMap, wbChangeMap, row[d_ResType1_col], row[d_ResType2_col], row[d_ResType3_col],
                                            row[d_RYResID_col], row[d_itemType_col], row[d_itemID_col], row[d_itemNum_col], i,
                                            lev)
                updateList(id, rewards)

        loveSht.cells(1, l_reward_col + 1).options(transpose=True).value = lv_reward
        loveSht.cells(1, l_drop_col + 1).options(transpose=True).value = lv_drop
        loveSht.cells(1, l_desc1_col + 1).options(transpose=True).value = lv_desc1
        loveSht.cells(1, l_unknow_col + 1).options(transpose=True).value = lv_unknow
        loveSht.cells(1, l_desc2_col + 1).options(transpose=True).value = lv_desc2
        loveSht.cells(1, l_unknow2_col + 1).options(transpose=True).value = lv_unknow2
        loveSht.cells(1, l_desc3_col + 1).options(transpose=True).value = lv_desc3
        loveSht.cells(1, l_unknow3_col + 1).options(transpose=True).value = lv_unknow3
        loveWb.save()
        loveWb.close()

    finally:
        for b in app.books:
            for i in itemMap:
                if b.name == itemMap[i]['奖励表'] and itemMap[i]['来源字段'] is not None and b.name in wbChangeMap:
                    b.save()
                    break
            b.close()
        app.screen_updating = True
        app.display_alerts = True
        app.quit()
        print("")
        printer.setCompareTime(printer.printGapTime("数据处理完毕，耗时:"))


def getRewardData(itemMap, wbChangeMap, resType, resType2, resType3, resId, itemType, itemId, itemNum, role, lev):
    """获取牵绊度等级对应奖励

    Args:
        itemMap (map)
        wbChangeMap (map): 保存变更的表格
        resType (str): 资源类型1
        resType2 (str): 资源类型2
        resType3 (str): 资源类型3
        resId (str): 资源id
        itemType (int): 道具类型
        itemId (int): 道具id
        itemNum (int): 道具数量
        role (int): 男主id
        lev (int): 牵绊度等级

    Returns:
        rewards: [{'reward','desc','type','order'}]
    """
    rewards = []
    idList = splitResId(resId)
    for i in range(len(idList)):
        if i == 2:
            resType = resType3
            order = 2
        elif i == 1:
            resType = resType2
            order = 1
        else:
            order = 0
        if resType in itemMap.keys():
            if role == 1:
                itemMap[resType]['奖励计数'] += 1
            for id in idList[i]:
                reward = {}
                reward['reward'] = ''
                reward['order'] = order
                if itemMap[resType]['无需掉落'] == 1:  # 纯文本描述
                    reward['desc'] = getRewardDesc(itemMap, wbChangeMap, resType, id, role=role, lev=lev)
                elif common.isNumberValid(id):  # >0的id有效
                    reward['reward'] = getReward(itemMap[resType]['奖励类型'], id)
                    if reward['reward'] != '':
                        reward['desc'] = getRewardDesc(itemMap, wbChangeMap, resType, id, role=role, lev=lev)
                elif common.isNumberValid(itemId):  # id无效时取道具id填充
                    resType = '道具'
                    reward['reward'] = toStr(itemType) + '=' + toStr(itemId) + '=' + toStr(itemNum)
                    reward['desc'] = getRewardDesc(itemMap, wbChangeMap, resType, itemId, itemNum)
                reward['type'] = itemMap[resType]['奖励类型']
                rewards.append(reward)
    return rewards


def getRewardDesc(itemMap, wbChangeMap, itemType, itemID, role=0, lev=0):
    """获取奖励描述

    Args:
        itemMap (map)
        wbChangeMap (map): _description_
        itemType (str): 奖励类型
        itemID (int): 奖励id
        itemNum (int, optional): 奖励数量. Defaults to 0.
        role (int, optional): 男主id. Defaults to 0.
        lev (int, optional): 牵绊度等级. Defaults to 0.

    Returns:
        str: 奖励描述
    """
    chatId = [0, 101, 201, 1, 1, 501]  # 各男主闲聊组id
    if itemType == '闲聊':
        itemID = chatId[role]

    if itemMap[itemType]['奖励表'] is None:
        desc = toStr(itemMap[itemType]['文字描述'])
    else:
        rewardWb = xw.books[itemMap[itemType]['奖励表']]
        rewardSht = rewardWb.sheets[itemMap[itemType]['奖励sheet']]
        r_reward = rewardSht.used_range
        rv_column = r_reward.rows[2].value
        rv_id = r_reward.columns[0].value
        r_row = common.getListRow(rv_id, itemID)
        if r_row == -1 or itemMap[itemType]['描述字段'] is None:
            desc = toStr(itemMap[itemType]['文字描述'])
        else:
            r_name_col = getDataOrder(rv_column, itemMap[itemType]['描述字段'])
            if itemType == '闲聊':
                desc = itemMap[itemType]['文字描述'].format(r_reward.columns[r_name_col][r_row].value,
                                                        int(1 + itemMap[itemType]['奖励计数']))
            elif '{' in itemMap[itemType]['文字描述']:
                desc = itemMap[itemType]['文字描述'].format(r_reward.columns[r_name_col][r_row].value)
            else:
                desc = toStr(itemMap[itemType]['文字描述'])

        # 修改来源描述
        if r_row != -1 and itemMap[itemType]['来源字段'] is not None:
            r_desc_col = getDataOrder(rv_column, itemMap[itemType]['来源字段'])
            setRewardScource(wbChangeMap, itemMap[itemType]['奖励表'], rewardSht, itemMap[itemType]['来源类型'], r_row + 1,
                             r_desc_col + 1, role, lev)
    return desc


def setRewardScource(wbChangeMap, wbName, rewardSht, scourceType, row, desc_col, role, lev):
    """更新奖励道具的来源说明

    Args:
        wbChangeMap (map): 保存修改过的表格
        wbName (str): 奖励表格名
        rewardSht (sht): 奖励道具表
        scourceType(int): 来源类型：0=带描述，1=只有等级
        r_row (int): 奖励道具id对应行
        r_desc_col (int): 来源对应列
        role (int): 男主id
        lev (int): 牵绊度等级
    """
    # 修改来源描述
    if scourceType == 0:
        desc = '{Role' + toStr(role) + '}牵绊度达到' + toStr(lev) + '级'
    elif scourceType == 1:
        desc = int(lev)
    if rewardSht.cells(row, desc_col).value != desc:
        wbChangeMap[wbName] = 1
        rewardSht.cells(row, desc_col).value = desc


def getReward(itemType, itemID, itemNum=1):
    """获取奖励内容

    Args:
        itemType (str): 奖励类型，-1=随机掉落，0=无掉落
        itemID (int): 奖励id
        itemNum (int, optional): 奖励数量. Defaults to 1.

    Returns:
        普通奖励返回三元组，随机掉落返回掉落id
    """
    reward = ''
    if itemType > 0:
        reward = toStr(itemType) + '=' + toStr(itemID) + '=' + toStr(itemNum)
    elif itemType == -1:  # 随机掉落池
        reward = itemID
    return reward


def splitResId(resId: str):
    """解析资源id

    Args:
        resId (str): 资源id，不同类型用;分隔，相同类型用|分隔

    Returns:
        idList: [[],[]]资源id列表，不存在的id返回-1
    """
    sep1 = ';'
    returnList = []
    resId = toStr(resId)
    if sep1 in resId:  # 存在2种不同的资源类型
        idTypes = str.split(resId, sep1)
        for t in idTypes:
            returnList.append(splitId(t))  # 同类型资源存在多个
    else:  # 只存在一种资源类型
        returnList.append(splitId(resId))
    return returnList


def splitId(idStr: str):
    """获取id列表

    Args:
        idStr (str): 待分隔的字符串

    Returns:
        list: id组成的列表，空id返回[-1]
    """
    sep2 = '|'
    if sep2 in idStr:  # 同类型资源存在多个
        ids = str.split(idStr, sep2)
        for i in range(len(ids)):
            ids[i] = toInt(ids[i])
        return ids
    else:  # 同类型资源只有1个
        return [toInt(idStr)]


def isBookOpen(app, bookName):
    """检测表格是否打开

    Args:
        app (_type_): 表格所属app
        bookName (str): 表格名

    Returns:
        result: True/False
    """
    for b in app.books:
        if b.name == bookName:
            return True
    else:
        return False
