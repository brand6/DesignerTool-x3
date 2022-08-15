import sys
import traceback
from common import common
from gacha import draw

try:
    common.quitHideApp()
    if len(sys.argv) == 1:
        draw.main()
    for p in sys.argv[1:]:
        print("执行脚本：" + p)
        if p == 'draw':
            draw.main()
        else:
            print("未找到可执行脚本")
    input("Press Enter To Exit...")
except BaseException:
    traceback.print_exc()
    input("Error...")
