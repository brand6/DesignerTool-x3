import xlwings as xw
import numpy as np
from common import common
from xlwings import Sheet
from common.common import getDataColOrder
from common.common import getRowData
from common.common import split
from common.common import toStr

# 脚本说明：
# 用于更新关卡首通时给的资源类奖励
#


def main():
    rewardMap = {'短信': '20', '电话': '21', '朋友圈': '22', '公众号': '25'}

    rewardWb = xw.books.active
    rewardSht: Sheet = rewardWb.sheets['主线奖励']
    rewardData = rewardSht.used_range.value

    tablePath = common.getTablePath()
    stageWb = xw.books.open(tablePath + "\\" + 'CommonStageEntry.xlsx', 0)
    stageSht: Sheet = stageWb.sheets['CommonStageEntry']
    stageData = np.array(stageSht.used_range.value)

    rewardIdCol = getDataColOrder(rewardData, '资源id')
    rewardTypeCol = getDataColOrder(rewardData, '资源类型')
    rewardStageCol = getDataColOrder(rewardData, '资源投放方式')
    rewardStageIdCol = getDataColOrder(rewardData, '关卡id')
    rewardStageOrderCol = getDataColOrder(rewardData, '关卡序号')

    stageIdCol = getDataColOrder(stageData, 'ID', 2)
    stageOrderCol = getDataColOrder(stageData, 'NumTab', 2)
    stageRewardCol = getDataColOrder(stageData, 'FirstReward', 2)
    stageIds = stageData[:, stageIdCol]
    stageOrders = stageData[:, stageOrderCol]
    stageRewards = stageData[:, stageRewardCol]

    for i in range(1, len(rewardData)):
        if rewardData[i][rewardIdCol] is not None:
            rewardId = toStr(rewardData[i][rewardIdCol])
            rewardType = rewardMap[rewardData[i][rewardTypeCol]]
            rewardStr = rewardType + '=' + rewardId + '=1'

            rewardStageOrder = split(rewardData[i][rewardStageCol], '：')[1]
            rewardStageInfos = getRowData(rewardStageOrder, rewardStageOrderCol, [rewardStageIdCol, rewardStageOrderCol],
                                          rewardData)
            updateStageReward(stageIds, stageOrders, stageRewards, rewardStageInfos, rewardStr)

    for i in range(len(stageRewards)):
        stageRewards[i] = checkStageReward(stageRewards[i])
    stageSht.cells(1, stageRewardCol + 1).options(transpose=True).value = stageRewards
    input('程序执行结束，按回车保存数据，叉掉本窗口不保存数据')
    stageWb.save()
    stageWb.close()


def updateStageReward(stageIds, stageOrders, stageRewards, rewardStageInfos, rewardStr):
    """更新关卡资源奖励

    Args:
        stageIds (_type_): 关卡id列表
        stageOrders (_type_): 关卡order列表
        stageRewards (_type_): 关卡奖励列表
        rewardStageInfos (_type_): 奖励关卡信息[id,order]
        rewardStr (_type_): 奖励内容
    """
    for i in range(len(stageIds)):
        if toStr(stageIds[i]) == toStr(rewardStageInfos[0]) and stageOrders[i] == rewardStageInfos[1]:  # 添加资源到关卡掉落
            if stageRewards[i] is None or stageRewards[i] == '':
                stageRewards[i] = rewardStr
                print(stageIds[i], '：', stageOrders[i], '添加奖励', rewardStr)
            elif rewardStr not in stageRewards[i]:
                stageRewards[i] = stageRewards[i] + '|' + rewardStr
                print(stageIds[i], '：', stageOrders[i], '添加奖励', rewardStr)
            else:
                print(stageIds[i], '：', stageOrders[i], '已有奖励', rewardStr)
        elif stageRewards[i] is not None and rewardStr in stageRewards[i]:  # 移除资源在其他主线的掉落
            print(stageIds[i], '：', stageOrders[i], '移除奖励', rewardStr)
            stageRewards[i] = stageRewards[i].replace(rewardStr, '')
            stageRewards[i] = checkStageReward(stageRewards[i])


def checkStageReward(stageReward):
    """清除奖励字段里的||和尾部的|

    Args:
        stageRewards (_type_): 奖励列表
    """
    if stageReward is not None:
        if '||' in stageReward:
            stageReward = stageReward.replace('||', '|')
        if stageReward[-1] == '|':
            stageReward = stageReward[:-1]
        if stageReward[0] == '|':
            stageReward = stageReward[1:]
    return stageReward
