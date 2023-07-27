import sys
import traceback

from common import common
from gacha import gemInit, gemUpgrade, drawResult, hungUpResult

try:
    programMap = {'gemInit': gemInit, 'gemUpgrade': gemUpgrade, 'drawResult': drawResult, 'hungUpResult': hungUpResult}
    common.quitHideApp()
    if len(sys.argv) == 1:
        hungUpResult.main()
    for p in sys.argv[1:]:
        print("执行脚本：" + p)
        if p in programMap:
            programMap[p].main()
        else:
            input("未找到可执行脚本:" + p)

except BaseException:
    traceback.print_exc()
    input("Error...")
