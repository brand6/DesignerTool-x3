import sys
import traceback
from common import common

from battle import copyData
from battle import stagePropertyCalc
from battle import stageReward
from battle import propertyCalc
from battle import monsterHurtCalc
from battle import monsterIdSync
from battle import monsterIdAdd
from battle import hitParam
from battle import skillSync
from battle import propIdClear

programMap = {
    'stagePropertyCalc': stagePropertyCalc,
    'stageReward': stageReward,
    'copyData': copyData,
    'propertyCalc': propertyCalc,
    'monsterHurtCalc': monsterHurtCalc,
    'monsterIdSync': monsterIdSync,
    'monsterIdAdd': monsterIdAdd,
    'hitParam': hitParam,
    'skillSync': skillSync,
    'propIdClear': propIdClear
}
common.quitHideApp()
if len(sys.argv) == 1:
    stagePropertyCalc.main()
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
