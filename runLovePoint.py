import sys
import traceback
from common import common
from lovePoint import lovePoint
from lovePoint import lovePointRandomDrop
from lovePoint import lovePointTask
from lovePoint import lovePointDraw
from lovePoint import lovePointLev

try:
    common.quitHideApp()
    if len(sys.argv) == 1:
        lovePoint.main()
    for p in sys.argv[1:]:
        print("执行脚本：" + p)
        if p == 'lovePoint':
            lovePoint.main()
        elif p == 'lovePointRandomDrop':
            lovePointRandomDrop.main()
        elif p == 'lovePointTask':
            lovePointTask.main()
        elif p == 'lovePointDraw':
            lovePointDraw.main()
        elif p == 'lovePointLev':
            lovePointLev.main()
        else:
            input("未找到可执行脚本:" + p)
except BaseException:
    traceback.print_exc()
    input("Error...")
