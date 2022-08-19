import sys
import traceback
from common import common
from tlog import tlog
from tlog import tlogAll

try:
    common.quitHideApp()
    if len(sys.argv) == 1:
        tlogAll.main()
    for p in sys.argv[1:]:
        print("执行脚本：" + p)
        if p == 'tlog':
            tlog.main()
        elif p == 'tlogAll':
            tlogAll.main()
        else:
            input("未找到可执行脚本:" + p)
except BaseException:
    traceback.print_exc()
    input("Error...")
