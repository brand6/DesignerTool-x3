import sys
import traceback
from common import common
from activity import resourceStatistic, resourceSync

try:
    programMap = {'resourceStatistic': resourceStatistic, 'resourceSync': resourceSync}
    common.quitHideApp()
    if len(sys.argv) == 1:
        resourceStatistic.main()
    else:
        program = sys.argv[1]
        print("执行脚本：" + program)
        programMap[program].main()
except FileExistsError:  # BaseException:
    traceback.print_exc()
    input("Error...")
