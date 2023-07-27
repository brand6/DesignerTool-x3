import sys
import traceback
from common import common
from tlog import tlog, tlogAll, weiDu

try:
    programMap = {'tlog': tlog, 'tlogAll': tlogAll, 'weiDu': weiDu}
    common.quitHideApp()
    if len(sys.argv) == 1:
        tlog.main()
    for p in sys.argv[1:]:
        print("执行脚本：" + p)
        if p in programMap:
            programMap[p].main()
        else:

            input("未找到可执行脚本:" + p)
except BaseException:
    traceback.print_exc()
    input("Error...")
