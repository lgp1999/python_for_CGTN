from get_data import *
from xls_functions import *

# datetime = '2021-06-01'
datetime = input("请输入需要导出的日期（日期格式为：2021-06-01）：")
date_format = is_valid_date(datetime)
while True:
    if date_format:
        local_time = time_trans_stamp(datetime)  # 将输入日期转换成时间戳
        break
    else:
        datetime = input("日期输入不正确，请重新输入需要导出的日期（日期格式为：2020-06-01）：")
        date_format = is_valid_date(datetime)
date = int(input("请输入要查询的天数（不超过30天）："))
if date > 30:
    date = int(input("查询日期不可超过30天，请重新输入："))
stop_time = local_time + 86400000 * (date - 1)  # 截止日期
stop_local_time = stamp_trans_time(stop_time)[2]

author_names = []
# 自定义输入记者名字
while True:
    author_name = input("请输入需要导出的记者名字(输入‘q’则跳过)：")
    if author_name != 'q':
        author_names.append(author_name)
    if author_name == 'q':
        break
# 如果没有输入任何记者名字，则为默认
if len(author_names) == 0:
    author_names = ['邓宗宇', '江华', '李冠男', '李耀洋', '王玉国', '魏帆', '徐公正', '邹合义', '张颖', '李春元', '徐明', '马瑾瑾', '郝晓丽', '薛婧萌', '李长皓',
                    '余鹏', '朱赫', '杨春', '白桦', '侯茂华', '孔琳琳', '田晓春', '王楠', '王璇', '张赫', '雷昊', '康玉斌', '陈明磊', '林晖', '梁弢',
                    '康炘冬', '贾延宁', '阮佳闻', '宋承杰', '张婧昊', '郑治', '时光', '赵洪超', '殷欣', '易歆']

filename = f'{datetime}至{stop_local_time}.xls'
# filename = '2021-06-01至06-03.xls'
# print(f"正在查找{datetime}到{stop_local_time}记者《{author_name}》的新闻，请稍等.....")
print(f"你需要查询的时间为{datetime}到{stop_local_time}的新闻，请稍等.....")
print(f"查询的记者名为{author_names}")
print('-----------------------------------')

data = get_data(local_time, author_names, date)  # 获取稿件数据
# for author_name in author_names:
title = [['Datetime', 'view', 'video', 'title', 'name', 'url'], ]
write_excel_xls(filename, title)  # 创建表格并写入表头
write_excel_xls_append(filename, data)  # 写入表格内容

# 查找完成后弹窗提示
root = tkinter.Tk()
root.withdraw()
messagebox.showinfo('提示', '查找完成！')
