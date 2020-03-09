# 导入地理编码模块
import compare_location as CL
# 导入日期时间处理模块
import datetime
import time
import re

# 导入OpenPyXL处理EXCEL的xlsx文件
from openpyxl import Workbook
from openpyxl import load_workbook

# 获取当前年份
current_year = datetime.datetime.now().year
current_month = datetime.datetime.now().month

# #提示输入考勤年份和月份
# prompt = "请输入考勤年份（默认为" + str(current_year) + "年)："
# year = input(prompt)
# prompt = "请输入考勤月份（默认为" + str(current_month) + "月)："
# month = input(prompt)
# if year == '':
#     year = current_year
# if month == '':
#     month = current_month
# print(year,month)

# （一）处理坐班打卡记录
# 数据源文件：坐班打卡记录.xlsx
# 关键字段：1.员工编号 2.签卡时间 3.数据来源
# 构建坐班列表OCR_list (office checking recode)，字段有：id, check_time, source
# 读入坐班打卡记录关键字段至OCR_list中
OCR_list = []

wb_OCR = load_workbook(filename="data/坐班打卡记录.xlsx")
ws_OCR = wb_OCR.active
row_index = 1
for per_row in ws_OCR.iter_rows():
    if row_index > 1:
        department = per_row[0].value
        id = per_row[1].value
        name = per_row[2].value
        check_time = time.strptime(per_row[3].value[:16], "%Y-%m-%d %H:%M")
        source = re.findall(r'[(](.*?)[)]', per_row[4].value)[0]
        OCR_list.append([department, id, name, check_time, source])
    row_index += 1
# print(OCR_list)

# （二）处理外勤打卡记录
# 数据源文件：外勤打卡记录.xlsx
# 关键字段：1.员工编号 2.签卡时间 3.地点
# 构建外勤列表MCR_list (mobile checking recode)，字段有：id, check_time, location
# 读入外勤打卡记录关键字段至MCR_list中
MCR_list = []

wb_MCR = load_workbook(filename="data/外勤打卡记录.xlsx")
ws_MCR = wb_MCR.active
row_index = 1
for per_row in ws_MCR.iter_rows():
    if row_index > 1:
        id = per_row[1].value
        check_time = time.strptime(per_row[4].value, "%Y-%m-%d %H:%M")
        location = per_row[3].value
        MCR_list.append([id, check_time, location])
    row_index += 1
# print(MCR_list)

# （三）生成当月考勤记录列表
check_list = []
# 处理坐班考勤记录，将同一人、同一天的数据聚合成一条记录
# 关键字段:1.部门 2.ID 3.姓名 4.类型（坐班/外勤） 5.签卡(check_time)[列表]

for i in range(0, len(OCR_list)):
    department = OCR_list[i][0]
    id = OCR_list[i][1]
    name = OCR_list[i][2]
    date = "%4d-%02d-%02d" % (OCR_list[i][3].tm_year, OCR_list[i][3].tm_mon, OCR_list[i][3].tm_mday)
    # 若列表为空，直接添加记录
    if len(check_list) == 0:
        check_list.append([department, id, name, date, "坐班", []])
        continue
    # 先扫描一遍签卡记录看看有没有同一人、同一天
    found = False
    for j in range(0, len(check_list)):
        if check_list[j][1] == id and check_list[j][3] == date:
            found = True
    if found == False:
        check_list.append([department, id, name, date, "坐班", []])
# 此时，check_list中已是聚合后的列表
# print(check_list)
# 再扫描打卡记录，将时间写入聚合后的列表中

for i in range(0, len(OCR_list)):
    for j in range(0, len(check_list)):
        date = "%4d-%02d-%02d" % (OCR_list[i][3].tm_year, OCR_list[i][3].tm_mon, OCR_list[i][3].tm_mday)
        if check_list[j][1] == OCR_list[i][1] and check_list[j][3] == date:
            check_list[j][5].append("%02d:%02d" % (OCR_list[i][3].tm_hour, OCR_list[i][3].tm_min))

print(check_list)
#
# address1 = '北京市海淀区美林花园'
# address2 = '北京市海淀区北理工附中小学部'
#
#
# result = CL.compare_location(address1, address2, 200)
#
# print(result)
