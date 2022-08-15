import sys
import traceback
from common import common
from battle import copyData
from battle import stageTimeCalc
from battle import propertyCalc

try:
    common.quitHideApp()
    if len(sys.argv) == 1:
        propertyCalc.main()
    for p in sys.argv[1:]:
        print("执行脚本：" + p)
        if p == 'stageTimeCalc':
            stageTimeCalc.main()
        if p == 'copyData':
            copyData.main()
        if p == 'propertyCalc':
            propertyCalc.main()
        else:
            print("未找到可执行脚本")
    input("Press Enter To Exit...")
except BaseException:
    traceback.print_exc()
    input("Error...")
