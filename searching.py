from get_data import *
from xls_functions import *
import os
# datatime = '2021-06-01'
datatime = input("请输入需要导出的日期（日期格式为：2020-06-01）：")
local_time = time_trans_stamp(datatime) # 将输入日期转换成时间戳
stop_time = local_time + 86400000*2 # 截止日期为输入日期加2
stop_local_time = stamp_trans_time(stop_time)[1]
author_name = '王玉国'
print(f"正在查找{datatime}到{stop_local_time}记者《{author_name}》的新闻，请稍等.....")
print('-----------------------------------')

filename = f'{datatime}至{stop_local_time}.xls'
sheet_name = f'{author_name}'
data = get_data(local_time,author_name)  # 接口获取的数据

if os.path.exists(filename):
    write_excel_xls_append(filename, data[0])  # 写入表格内容
else:
    title = [['Datetime', 'view', 'video', 'title', 'name', 'url']]
    write_excel_xls(filename, sheet_name, title)  # 创建表格并写入表头
    write_excel_xls_append(filename, data[0])  # 写入表格内容
    # write_excel_xls_append(filename,data[1])