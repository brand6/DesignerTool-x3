import sys
import traceback
from common import common
from activity import activityResStatistic

try:
    common.quitHideApp()
    if len(sys.argv) == 1:
        activityResStatistic.main()
    for p in sys.argv[1:]:
        print("执行脚本：" + p)
        if p == 'activityResStatistic':
            activityResStatistic.main()
        else:
            print("未找到可执行脚本")
    input("Press Enter To Exit...")
except BaseException:
    traceback.print_exc()
    input("Error...")
