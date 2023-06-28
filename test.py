import xlwings as xw

sht = xw.sheets.active
data = sht.used_range.value

sht['4:5'].delete('up')