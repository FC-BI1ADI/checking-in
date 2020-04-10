# 导入统计汇总相关包
import numpy as np
import pandas as pd
import matplotlib as mpl
# 导入地理编码模块
import compare_location as CL
# 导入OpenPyXL处理EXCEL的xlsx文件
import openpyxl


# out_check(id,out_time,out_address)
# 功能：判断外出地址是否符合要求
# 参数：员工编号，外出时间，外出地址
# 返回：若同一天能查到2条同一位置的签卡记录则返回True，否则返回False
def out_check(oc_df, id, out_time, out_address):
    time_list = []
    for i in oc_df.index:
        oc_id = oc_df.loc[i, "员工编号"]
        oc_day = oc_df.loc[i, "签卡时间"][:10]
        oc_address = oc_df.loc[i, "地点"]
        if oc_id == id and oc_day == out_time and CL.compare_location(out_address, oc_address, 400) == 1:
            time_list.append(oc_df.loc[i, "签卡时间"][11:])

    if len(time_list) >= 2:
        return True
    else:
        return False


# 调整显示格式
pd.set_option('display.max_columns', 10)
pd.set_option('display.max_rows', 1000)
pd.set_option('display.width', 200)
# 读入外出记录单文件
df = pd.read_excel("DATA/IN外出记录单.xlsx", header=2, usecols=[1, 2, 3, 4, 5, 6, 9, 11])
print(df.head())
# 调整各列的顺序
order = ["部门", "员工编号", "姓名", "人员类别", "外出时间", "外出类型", "外出地址", "相关项目编号"]
df = df[order]
# 增加外出校核列
df["项目类型"] = None
df["外出校核"] = None

# 扫描项目编号列，判断是否为pipeline项目、非pipeline项目、无项目编号
for i in df.index:
    project_type = ""

    if pd.isnull(df.loc[i, "相关项目编号"]):
        project_type = "无项目编号"
    else:
        if df.loc[i, "相关项目编号"][0:1] == 'P':
            project_type = "Pipeline项目"
        else:
            project_type = "非Pipeline项目"
    df.loc[i, "项目类型"] = project_type

# 读入外勤签卡记录
df_check = pd.read_excel("DATA/IN外勤签卡记录.xlsx", header=0, usecols=[1, 3, 4])
# print(df_check)

# 较验外出合规性
for i in df.index:
    if out_check(df_check, df.loc[i, "员工编号"], df.loc[i, "外出时间"], df.loc[i, "外出地址"]) == True:
        df.loc[i, "外出校核"] = True
    else:
        df.loc[i, "外出校核"] = False
# df中是完全的外勤记录信息
print(df)
df.to_excel("data/OUT外勤汇总.xlsx",columns=["部门","员工编号","姓名","人员类别","外出类型","项目类型","外出校核"])


# sales_list = []
# # 统计外勤信息分类导出
# for i in df.index:
#     if df.loc[i,]


# 销售类的导出字段
# 周/月/季度
# 姓名
# 分支机构
# 职能
# 考勤打卡异常次数
# 上下班打卡异常次数
# 外出打卡异常次数
# 外出拜访次数
# Pipeline项目拜访次数
# 非Pipeline项目拜访次数
# 无项目编号外出拜访次数
# 商务非正式交流次数
# Pipeline项目商务非正式交流次数
# 非Pipeline项目商务非正式交流次数
# 无项目编号商务非正式交流次数    “其他”类型外出次数
# 无项目编号的“其他”类型外出次数

# 技术类的导出字段
# 周/月/季度
# 姓名
# 分支机构
# 职能
# 考勤打卡异常次数
# 上下班打卡异常次数
# 外出打卡异常次数
# 外出次数
# Pipeline项目外出次数
# 非Pipeline项目外出次数
# 无项目编号外出次数
# 商务非正式交流次数
# Pipeline项目商务非正式交流次数
# 非Pipeline项目商务非正式交流次数
# 无项目编号商务非正式交流次数
# 客户交流次数
# Pipeline项目客户交流次数
# 非Pipeline项目客户交流次数
# 无项目编号客户交流次数
# 投标活动次数
# Pipeline项目投标次数
# 非Pipeline项目投标次数
# 无项目编号投标次数
# 培训次数
# Pipeline项目培训次数
# 非Pipeline项目培训次数
# 无项目编号培训次数
# 安装实施次数
# 故障排除次数
# 巡检次数
# “其他”类型外出次数
# 无项目编号的“其他”类型外出次数

