from common import common
import sys
import traceback

try:
    if len(sys.argv) == 1:
        pass
    elif sys.argv[1] == 'getProgramPath':
        tablePath = common.getTablePath(True)
except FileExistsError:  # BaseException:
    traceback.print_exc()
    input("Error...")
