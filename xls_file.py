import os

import xlrd
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy  # 将xlrd.Book对象复制到xlwt.Workbook对象的工具。

file_dir = "G:\\school\\"


def nsfile(s):
    # 判断文件夹是否存在，如果不存在则创建
    b = os.path.exists(file_dir)
    if not b:
        os.mkdir(file_dir)

    # 创建xls文件对象
    wb = xlwt.Workbook()
    # 新建表单
    sh = wb.add_sheet('A new sheet')
    # 按位置添加数据，前面两个参数是位置，后面一个是单元格内容
    sh.write(0, 0, '名称')
    sh.write(0, 1, '地区')
    sh.write(0, 2, '电话')
    sh.write(0, 3, '地址')
    sh.write(0, 4, '网址')
    # 保存文件
    wb.save(file_dir + s + '.xls')


# 获得xls当前已写入的行数
def get_lines(fn):
    file_d = open_workbook(fn)
    # 获得第一个页签对象
    select_sheet = file_d.sheets()[0]
    row_list = []
    # 获取总共的行数
    rows_num = select_sheet.nrows
    return rows_num


# 修改学校信息
def wtxls(p, n, city, tel, adr, www):
    file_name = file_dir + p + '.xls'
    # 打开想要更改的excel文件
    old_excel = xlrd.open_workbook(file_name, formatting_info=True)

    cur_line = get_lines(file_name)
    # 将操作文件对象拷贝，变成可写的workbook对象
    new_excel = copy(old_excel)
    # 获得第一个sheet的对象
    ws = new_excel.get_sheet(0)
    # 写入数据
    ws.write(cur_line, 0, n)
    ws.write(cur_line, 1, city)
    ws.write(cur_line, 2, tel)
    ws.write(cur_line, 3, adr)
    ws.write(cur_line, 4, www)
    # 另存为excel文件，并将文件命名
    new_excel.save(file_name)
    # print("省份：" + city + "，当前数量" + cur_line + "，总数量：" + "")
    # 省份 当前页数 总页数 当前数量 总数量


if __name__ == "__main__":
    print("无法运行")
else:
    print("正在运行。。。")
