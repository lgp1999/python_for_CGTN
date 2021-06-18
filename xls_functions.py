import os
import sys
import tkinter
from tkinter import messagebox
import xlrd
import xlwt
from xlutils.copy import copy


def get_style(font1=False, font2=False, alignment=False, bg_color=False, border=True, bold=False):
    '''
    设置表格样式
    :param font1: 默认为普通单元格，否则选择此字体格式
    :param font2: 默认为普通单元格,否则选择此字体
    :param alignment: 是否设置单元格格式，默认不设置
    :param bg_color: 是否设置单元格背景色，默认不设置
    :param border:是否需要加边框，默认需要
    :param bold: 是否需要加粗，默认不加粗
    :return:
    '''
    style = xlwt.XFStyle()
    if font1:
        # 字体
        font = xlwt.Font()
        font.name = "宋体"  # 字体名字
        font.bold = bold  # 加粗，false为不加粗
        font.underline = False  # 下划线
        font.italic = False
        font.colour_index = 0
        font.height = 280  # 200为10号字体
        style.font = font

    if font2:
        font = xlwt.Font()
        font.height = 240
        font.name = '宋体'
        font.bold = bold  # 加粗，false为不加粗
        font.underline = False  # 下划线
        style.font = font
    # 单元格居中
    if alignment:
        align = xlwt.Alignment()
        # 水平方向居中,0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        align.horz = 0x02
        # 竖直方向居中，# 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        align.vert = 0x01
        style.alignment = align

    # 背景色
    if bg_color:
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['green']  # 设置单元格背景色为黄色
        style.pattern = pattern

    # 边框
    if border:
        border = xlwt.Borders()  # 给单元格加框线
        border.left = xlwt.Borders.THIN  # 左
        border.top = xlwt.Borders.THIN  # 上
        border.right = xlwt.Borders.THIN  # 右
        border.bottom = xlwt.Borders.THIN  # 下
        border.left_colour = 0x40  # 边框线颜色
        border.right_colour = 0x40
        border.top_colour = 0x40
        border.bottom_colour = 0x40
        style.borders = border

    return style


def create_excel_xls(path, sheet_name='sheet1'):
    '''
    如果该路径下工作簿不存在则创建工作簿
    :param path:工作簿路径
    :param sheet_name:创建工作簿时必须要同时创建一个sheet1表格
    :return:修改后的工作簿
    '''
    if os.path.exists(path):
        print(f'{path}已存在')
    else:
        workbook = xlwt.Workbook(encoding='utf-8')  # 新建一个工作簿
        workbook.add_sheet(sheet_name)
        workbook.save(path)  # 保存工作簿
        print(f'工作簿《{path}》表格《{sheet_name}》创建成功')


def create_sheet(path, sheet_name):
    '''
    创建表格sheet
    :param path: xls工作簿名称
    :param sheet_name: 表格名称
    :return: 修改后的工作簿
    '''
    # 打开工作簿,参数formatting_info为True时指打开后保持已有表格的格式
    workbook = xlrd.open_workbook(path, formatting_info=True)
    new_workbook = copy(workbook)  # 将xlrd对象转换为xlwt对象
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    if sheet_name not in sheets:
        new_workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
        print(f'表格《{sheet_name}》创建成功')
        print(f'开始向表格《{sheet_name}》添加内容')
    else:
        print(f'开始向表格《{sheet_name}》添加内容')
    new_workbook.save(path)
    return new_workbook

def write_excel_xls(path, value, sheet_name='sheet1'):
    '''
    打开一张表格，在表格最顶行插入value数据并保存
    :param path: 文件路径及名称
    :param value: 写入表格中的内容
    :param sheet_name: 表格名称,默认为sheet1
    :return:修改后的工作簿
    '''

    index = len(value)  # 获取需要写入数据的行数
    try:
        create_excel_xls(path, sheet_name)
        workbook = create_sheet(path, sheet_name)
        sheet = workbook.get_sheet(sheet_name)  # 根据名字打开表格
        for i in range(0, index):
            for j in range(0, len(value[i])):
                if i == 0:
                    sheet.write(i, j, value[i][j],
                                    get_style(font1=True, bg_color=True, alignment=True, border=True))
                else:
                    sheet.write(i, j, value[i][j], get_style(font2=True))
        workbook.save(path)  # 保存工作簿
        print(f"写入数据成功！")
    except Exception as e:
        # 异常弹窗提示
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo("ERROR", f'{e}')
        sys.exit()

def write_excel_xls_append(path, value, sheet_name='sheet1'):
    '''
    向xls表格中追加value内容并保存（此方法会导致表格中原来数据的样式被初始化）
    :param path: 文件路径
    :param sheet_name: 表格名称
    :param value: 追加内容
    :return:
    '''

    index = len(value)  # 获取需要写入数据的行数
    try:
        # 打开工作簿,参数formatting_info为True时指打开后保持已有表格的格式
        workbook = xlrd.open_workbook(path, formatting_info=True)
        # sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheet_name)  # 获取工作簿中名叫{sheet_name}的表格
        rows_old = worksheet.nrows  # 获取{sheet_name}表格中已存在的数据的行数
        new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
        new_worksheet  = new_workbook.get_sheet(sheet_name)  # 根据名字打开表格
        for i in range(0, index):
            for j in range(0, len(value[i])):
                # 追加写入数据，注意是从i+rows_old行开始写入
                new_worksheet.write(i + rows_old, j, value[i][j], get_style(font2=True))
        new_workbook.save(path)  # 保存工作簿
        print(f"写入数据成功！")
    except Exception as e:
        # 异常弹窗提示
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo("ERROR", f'{e}')

def read_excel_xls(path):
    '''
    读取xls表格内容并显示
    :param path:
    :return:
    '''
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    for i in range(0, worksheet.nrows):
        for j in range(0, worksheet.ncols):
            print(worksheet.cell_value(i, j), "\t", end="")  # 逐行逐列读取数据
        print()
