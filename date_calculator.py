
from datetime import date, datetime, timedelta
import time

# is_leap_year 用于判断年份是否是闰年
def is_leap_year(year):
    if year % 4 == 0 and year % 100 != 0 or year % 400 == 0:
        return True
    else:
        return False
    
def str_to_date(str):
    # str为格式为XXXX-XX-XX的日期字符串
    str += " 00:00:00"
    return time.strptime(str,"%Y-%m-%d %H:%M:%S")

def date_to_str(date):
    return time.strftime("%Y-%m-%d",date)

# date_calculator 用于计算 time结构类日期的计算
def calculate_n_day(date,n):
# 将date转化为时间戳
    timeStamp = int(time.mktime(date))
    timeStamp += n*86400
    return time.localtime(timeStamp)

# interval_days 用于计算两个time结构类日期间的日期差
def interval_day(date1,date2):
    # 将两个日期都转化为时间戳
    timeStamp1 = int(time.mktime(date1))
    timeStamp2 = int(time.mktime(date2))
    delta = timeStamp2 - timeStamp1
    return int(delta/86400)

