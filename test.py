import calendar
import datetime
from mmap import mmap, ACCESS_READ
from xlrd import open_workbook, cellname
from tempfile import TemporaryFile
from xlwt import Workbook
import time
from xlutils.copy import copy

# sd = open_workbook('sd.xls')
# # ssd = copy(sd)
# # VV1 = ssd.get_sheet(0)
# VV1 = sd.sheet_by_index(0)
# print(VV1.cell(0, 0))
s=2
if s<10:
    dd='0%s'%(s,)

print(dd)