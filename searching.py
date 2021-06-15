from get_data import *
from xls_functions import *

# datetime = '2021-06-13'
datetime = input("请输入需要导出的日期（日期格式为：2021-06-01）：")
date_format = is_valid_date(datetime)
while True:
    if date_format:
        local_time = time_trans_stamp(datetime)  # 将输入日期转换成时间戳
        break
    else:
        datetime = input("日期输入不正确，请重新输入需要导出的日期（日期格式为：2020-06-01）：")
        date_format = is_valid_date(datetime)

stop_time = local_time + 86400000 * 2  # 截止日期为输入日期加2
stop_local_time = stamp_trans_time(stop_time)[2]

# author_name = input("请输入需要导出的记者名字：")

author_names = ['王玉国','张颖','白桦','张赫','康玉斌'] # 多个记者名字

filename = f'{datetime}至{stop_local_time}.xls'
# filename = '2021-06-01至06-03.xls'
# print(f"正在查找{datetime}到{stop_local_time}记者《{author_name}》的新闻，请稍等.....")
print(f"你需要查询的时间为{datetime}到{stop_local_time}的新闻，请稍等.....")
print('-----------------------------------')

data = get_data(local_time, author_names)  # 获取稿件数据
for author_name in author_names:
    title = [['Datetime', 'view', 'video', 'title', 'name', 'url'],]
    write_excel_xls(filename, author_name, title)  # 创建表格并写入表头
    write_excel_xls_append(filename, author_name, data[author_name])  # 写入表格内容

# 查找完成后弹窗提示
root = tkinter.Tk()
root.withdraw()
messagebox.showinfo('提示','查找完成！')