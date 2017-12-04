import calendar
import datetime
from mmap import mmap, ACCESS_READ
from xlrd import open_workbook, cellname
from tempfile import TemporaryFile
from xlwt import Workbook
import time
from xlutils.copy import copy

# 开始日期
YEAR = 2017
MONTH = 12
DAY = 4

# teambition导出的表格
TB_EXCEL = 'ss.xlsx'

# 模板表格
MODEL_EXCEL = 'sd.xls'

# 任务数量
MAX_TASK_NUM=34

ss = open_workbook(TB_EXCEL)
V1 = ss.sheet_by_index(0)
i = 0

sd = open_workbook(MODEL_EXCEL, formatting_info=True)
ssd = copy(sd)
VV1 = ssd.get_sheet(0)

# 写入全部任务
for row_index in range(V1.nrows):
    if row_index == 0:
        continue
    if V1.cell(row_index, 8).value != '':
        if V1.cell(row_index, 5).value != '':
            i = i + 1
            VV1.write(i + 3, 1, V1.cell(row_index, 0).value)

# 写入日期
for index in range(15):
    tom = datetime.datetime(YEAR, MONTH, DAY) + datetime.timedelta(days=index)
    text = tom.strftime('%m-%d')
    VV1.write(MAX_TASK_NUM, index + 2, text)

ssd.save('sd.xls')

# 写入功能截至时间
j = 0
sd = open_workbook(MODEL_EXCEL, formatting_info=True)
VVV1 = sd.sheet_by_index(0)
for row_index in range(V1.nrows):
    if row_index == 0:
        continue
    if V1.cell(row_index, 8).value != '':
        if V1.cell(row_index, 5).value != '':
            print(V1.cell(row_index, 5).value)
            stime = time.strptime(V1.cell(row_index, 5).value, '%Y-%m-%d %H:%M:%S')
            dd = ''
            if stime[2] < 10:
                dd = '0%s' % (stime[2],)
            else:
                dd = stime[2]
            str = "%s-%s" % (stime[1], dd)
            print(str)
            j = j + 1
            for date_index in range(15):
                val = VVV1.cell(MAX_TASK_NUM, date_index + 2).value
                # print(val)
                if str == val:
                    VV1.write(j + 3, date_index + 2, '√')

ssd.save('sd.xls')
