import sys
import traceback
from common import common
from battle import copyData
from battle import stageTimeCalc
from battle import propertyCalc

try:
    common.quitHideApp()
    if len(sys.argv) == 1:
        stageTimeCalc.main()
    for p in sys.argv[1:]:
        print("执行脚本：" + p)
        if p == 'stageTimeCalc':
            stageTimeCalc.main()
        if p == 'copyData':
            copyData.main()
        if p == 'propertyCalc':
            propertyCalc.main()
        else:
            input("未找到可执行脚本:" + p)

except BaseException:
    traceback.print_exc()
    input("Error...")
