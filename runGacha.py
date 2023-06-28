import sys
import traceback

from common import common
from gacha import draw, newDraw, gemInit, gemUpgrade, drawCard, drawRole

try:
    programMap = {
        'draw': draw,
        'newDraw': newDraw,
        'gemInit': gemInit,
        'gemUpgrade': gemUpgrade,
        'drawCard': drawCard,
        'drawRole': drawRole
    }
    common.quitHideApp()
    if len(sys.argv) == 1:
        drawCard.main()
    for p in sys.argv[1:]:
        print("执行脚本：" + p)
        if p in programMap:
            programMap[p].main()
        else:
            input("未找到可执行脚本:" + p)
except BaseException:
    traceback.print_exc()
    input("Error...")
