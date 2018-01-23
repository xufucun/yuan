import urllib.request

import os
import xlrd  # read xls
import xlwt  # write xls
from xlrd import open_workbook
from bs4 import BeautifulSoup  # 解析xml
from xlutils.copy import copy  # 将xlrd.Book对象复制到xlwt.Workbook对象的工具。

base_url = "http://www.ruyile.com/xuexiao/?t=1"
file_dir = "G:\\school\\"

# 模拟浏览器
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) '
                         'Chrome/51.0.2704.63 Safari/537.36'}

# TODO 改进
# TODO 1. 错误信息，跳过的信息，
# TODO 1. 显示当前的省份，总数量，当前数量
# TODO 2. 进度条百分比
# TODO 3. 多线程同时抓取
# def log_msg(province, all_school):
#     print("当前省份：" + province + "共搜索到" + all_school + "个幼儿园")
#
#
# def log_school(count_page, all_page):
#     print("第" + count_page + "条，共" + all_page + "条")


def nsfile(s):
    # TODO 改进 此处应该创建临时文件，当前省份全部爬完后，移动到用户目录
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


# 获得行数
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


# 使用BeautifulSoup解析
def soupa(url):
    try:
        # 请求
        request = urllib.request.Request(url=url, headers=headers)
        # 爬取结果
        response = urllib.request.urlopen(request, timeout=8)
        # 读取网页数据
        data = response.read()
        # # 设置解码方式 当前使用BeautifulSoup无需解码
        # data = data.decode('utf-8')
        soup = BeautifulSoup(data, "lxml")
        return soup
    except Exception as e:
        print("出现异常-->" + str(e))


# 单个学校信息
def school_count(pro, url, name):
    # 声明全局变量
    global school_city, school_nature, school_tel, school_www, school_adr
    # 抓取学校信息
    school_info = soupa(url).find(class_='xxsx')
    # print(name + school_info.get_text())
    # 开始循环
    for div in school_info:
        # print(div.get_text())
        # if div.get_text.
        # wtxls(pro,0,0,name)
        if "所属地区" in div.get_text():
            school_city = div.get_text().replace('所属地区：', '')
            # print(div.get_text())
        if "学校性质" in div.get_text():
            school_nature = div.get_text().replace('学校性质：', '')
            # print(div.get_text())
        if "招生电话" in div.get_text():
            school_tel = div.get_text().replace('招生电话：', '')
            # print(div.get_text())
        if "学校网址" in div.get_text():
            school_www = div.get_text().replace('学校网址：', '')
            # print(div.get_text())
        if "学校地址" in div.get_text():
            school_adr = div.get_text().replace('学校地址：', '')
            # print(div.get_text())
        # wtxls(pro,0,1,)
    # 写入表格
    wtxls(pro, name, school_city, school_tel, school_adr, school_www)


# 获取单个省份的信息
# pro 省份名称
# url 省份链接
def province_school(pro, url):
    # print("当前省份:" + pro + "省份链接:" + url)
    # 使用BeautifulSoup解析
    soupd = soupa(url)
    # 获取总页数
    zys = soupd.find(class_='zys')
    # print("总页数" + zys.get_text())
    # 获取当前页数（可以从1开始）
    cur = soupd.find(class_='fy').strong.get_text()
    # 转换成int类型，方便循环
    current_page = int(cur)
    all_pages = int(zys.get_text())
    # 开始循环
    while current_page < all_pages:
        # print("当前页数:" + str(current_page))
        cur_url = (url + '&p=' + str(current_page)).replace(' ', '')
        # print(cur_url)
        soupd = soupa(cur_url)
        # 查找所有sk属性
        divs = soupd.find_all(class_='sk')
        # 遍历divs里面的div
        for div in divs:
            # 获取单个学校名称和链接
            school_con = div.h4.a
            school_url = school_con.get('href')
            school_name = school_con.get_text()
            school_count(pro, school_url, school_name)

        current_page += 1


# 读取首页全部省份名称和省份链接
def main():
    # 获取全部省份
    province = soupa(base_url).find(class_='qylb')
    # 遍历所有省份
    for a in province:
        # 获取省份名称
        pro = a.get_text()
        # 获取省份链接
        link = a.get('href')
        # 创建文件
        nsfile(str(pro))
        # 使用province_school处理
        province_school(pro, link)


main()
