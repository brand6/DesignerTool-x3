import sys
import traceback
from common import common
from love import lovePoint, lovePointRandomDrop, lovePointTask

try:
    programMap = {
        "lovePoint": lovePoint,
        "lovePointRandomDrop": lovePointRandomDrop,
        "lovePointTask": lovePointTask,
    }
    common.quitHideApp()
    if len(sys.argv) == 1:
        lovePointTask.main()
    for p in sys.argv[1:]:
        print("执行脚本：" + p)
        if p in programMap:
            programMap[p].main()
        else:
            input("未找到可执行脚本:" + p)
except BaseException:
    traceback.print_exc()
    input("Error...")
