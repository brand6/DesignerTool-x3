import sys
import traceback
from common import common
from lovePoint import lovePoint
from lovePoint import lovePointRandomDrop
from lovePoint import lovePointTask
from lovePoint import lovePointDraw
from lovePoint import lovePointLev

try:
    programMap = {
        'lovePoint': lovePoint,
        'lovePointRandomDrop': lovePointRandomDrop,
        'lovePointTask': lovePointTask,
        'lovePointDraw': lovePointDraw,
        'lovePointLev': lovePointLev
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
