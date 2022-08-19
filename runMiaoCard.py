import sys
import traceback
from common import common
from miaoCard import miaoCardAIAction
from miaoCard import miaoCardInit
from miaoCard import miaoCardWinRate

try:
    common.quitHideApp()
    if len(sys.argv) == 1:
        miaoCardAIAction.main()
    for p in sys.argv[1:]:
        print("执行脚本：" + p)
        if p == 'miaoCardAIAction':
            miaoCardAIAction.main()
        elif p == 'miaoCardInit':
            miaoCardInit.main()
        elif p == 'miaoCardWinRate':
            miaoCardWinRate.main()
        else:
            input("未找到可执行脚本:" + p)
except BaseException:
    traceback.print_exc()
    input("Error...")
