import datetime
import shutil
import time

from xlrd import open_workbook
from xlutils.copy import copy

# 开始日期
from xlwt import XFStyle, Borders, Pattern

YEAR = 2017
MONTH = 12
DAY = 4

# teambition导出的表格
TB_EXCEL = 'comein.xlsx'

# 模板表格
MODEL_EXCEL = 'output.xls'

shutil.copy('model_excel.xls', 'output.xls')

# 任务数量
MAX_TASK_NUM = 34

# 任务时间跨度
TIME_SPAN = 20

# 整理那个页面的数据
SHEET = 0

ss = open_workbook(TB_EXCEL)
V1 = ss.sheet_by_index(SHEET)
i = 0

sd = open_workbook(MODEL_EXCEL, formatting_info=True)
ssd = copy(sd)
VV1 = ssd.get_sheet(0)

# 写入全部任务
style = XFStyle()
pattern = Pattern()
borders = Borders()
borders.right = Borders.THIN
borders.left = Borders.THIN
style.borders = borders

for row_index in range(V1.nrows):
    if row_index == 0:
        continue
    if V1.cell(row_index, 8).value != '':
        if V1.cell(row_index, 5).value != '':
            i = i + 1
            pre = ''
            if V1.cell(row_index, 1).value != '':
                pre = V1.cell(row_index, 1).value + '-'

            VV1.write(i + 3, 1, pre + V1.cell(row_index, 0).value, style)

# 写入日期
style = XFStyle()
borders = Borders()
borders.top = Borders.THIN
borders.bottom = Borders.THIN
style.borders = borders

for index in range(TIME_SPAN):
    tom = datetime.datetime(YEAR, MONTH, DAY) + datetime.timedelta(days=index)
    text = tom.strftime('%m-%d')

    if index == TIME_SPAN - 1:
        borders.right = Borders.THIN
        style.borders = borders

    VV1.write(MAX_TASK_NUM, index + 2, text, style)

ssd.save(MODEL_EXCEL)

# 写入功能截至时间
j = 0
sd = open_workbook(MODEL_EXCEL, formatting_info=True)
VVV1 = sd.sheet_by_index(0)
for row_index in range(V1.nrows):
    if row_index == 0:
        continue
    if V1.cell(row_index, 8).value != '':
        if V1.cell(row_index, 5).value != '':
            stime = time.strptime(V1.cell(row_index, 5).value, '%Y-%m-%d %H:%M:%S')
            dd = ''
            if stime[2] < 10:
                dd = '0%s' % (stime[2],)
            else:
                dd = stime[2]
            str = "%s-%s" % (stime[1], dd)
            j = j + 1
            for date_index in range(TIME_SPAN):
                val = VVV1.cell(MAX_TASK_NUM, date_index + 2).value
                if str == val:
                    style = XFStyle()
                    pattern = Pattern()
                    pattern.pattern = Pattern.SOLID_PATTERN

                    if V1.cell(row_index, 12).value == 'Y':  # 已完成
                        if V1.cell(row_index, 17).value == 'Y':  # 已超期
                            pattern.pattern_fore_colour = 0x0B
                            style.pattern = pattern
                            VV1.write(j + 3, date_index + 2, '√', style)
                        else:
                            pattern.pattern_fore_colour = 0x0F
                            style.pattern = pattern
                            VV1.write(j + 3, date_index + 2, '√', style)
                    else:  # 没完成
                        if V1.cell(row_index, 17).value == 'Y':  # 已超期
                            pattern.pattern_fore_colour = 0x0A
                            style.pattern = pattern
                            VV1.write(j + 3, date_index + 2, '√', style)
                        else:
                            pattern.pattern_fore_colour = 0x0FFB
                            VV1.write(j + 3, date_index + 2, '√')

# style = XFStyle()
# borders = Borders()
# borders.right = Borders.THIN
# style.borders = borders
# val = VVV1.cell(0, 1).value
# print(val)
# VV1.write(0, 1, 'fds', style)
ssd.save(MODEL_EXCEL)
