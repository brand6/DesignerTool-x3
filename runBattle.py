import sys
import traceback

from battle import copyData, hitParam, monsterIdSync, propertyCalc, propIdClear, skillSync, stagePropertyCalc, stageReward
from common import common

programMap = {
    "stagePropertyCalc": stagePropertyCalc,
    "stageReward": stageReward,
    "copyData": copyData,
    "propertyCalc": propertyCalc,
    "monsterIdSync": monsterIdSync,
    "hitParam": hitParam,
    "skillSync": skillSync,
    "propIdClear": propIdClear,
}
common.quitHideApp()
if len(sys.argv) == 1:
    stagePropertyCalc.main()
else:
    try:
        for p in sys.argv[1:]:
            print("执行脚本：" + p)
            if p in programMap:
                programMap[p].main()
            else:
                input("未找到可执行脚本:" + p)

    except BaseException:
        traceback.print_exc()
        input("Error...")
