import xlwings as xw
from xlwings import Range
from xlwings import Sheet

# 脚本说明：
# 用于统计新手活动的资源总量
#


def main():
    book = xw.apps.active.books.active
    statisSht: Sheet = book.sheets['#资源汇总']
    statisSht.range('B3').expand('table').value = None

    statisList = []

    for sht in book.sheets:
        if sht.name[0] != '#':
            beginTag = False
            dataRng: Range = sht.used_range
            for r in range(dataRng.rows.count):
                rng: Range = dataRng[r, 1]
                if rng.value == '天数':
                    beginTag = True
                elif beginTag is True:
                    if rng.value is not None:
                        rowList = [sht.name]
                        rowList.extend(rng.expand('right').value)
                        statisList.append(rowList)
                    else:
                        break
    statisList.sort(key=lambda x: x[1])
    statisSht.range('B3').value = statisList
