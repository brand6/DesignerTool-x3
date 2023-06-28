import sys
import traceback
from common import common
from activity import activityResStatistic
from activity import resouceStatistic
from activity import resourceSync

try:
    programMap = {
        'activityResStatistic': activityResStatistic,
        'resouceStatistic': resouceStatistic,
        'resourceSync': resourceSync
    }
    common.quitHideApp()
    if len(sys.argv) == 1:
        resouceStatistic.main()
        resouceStatistic.main(True)
    elif sys.argv[1] == 'resouceStatistic':
        print("执行脚本-免费资源统计：" + sys.argv[1])
        resouceStatistic.main(False)
        print("执行脚本-付费资源统计：" + sys.argv[1])
        resouceStatistic.main(True)
    else:
        program = sys.argv[1]
        print("执行脚本：" + program)
        programMap[program].main()
except FileExistsError:  # BaseException:
    traceback.print_exc()
    input("Error...")
