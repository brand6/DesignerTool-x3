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
        if p == 'miaoCardAIAction':
            miaoCardAIAction.main()
        elif p == 'miaoCardInit':
            miaoCardInit.main()
        elif p == 'miaoCardWinRate':
            print("执行脚本：" + p)
            miaoCardWinRate.main()
            input("Press Enter To Exit...")
        else:
            print("未找到可执行脚本")
except BaseException:
    traceback.print_exc()
    input("Error...")
