import urllib.request

import os

import xlrd
import xlwt
from bs4 import BeautifulSoup

base_url = "http://www.ruyile.com/xuexiao/?t=1"
file_dir = "G:\\school\\"

# 模拟浏览器
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) '
                         'Chrome/51.0.2704.63 Safari/537.36'}

# def log_msg(province, all_school):
#     print("当前省份：" + province + "共搜索到" + all_school + "个幼儿园")
#
#
# def log_school(count_page, all_page):
#     print("第" + count_page + "条，共" + all_page + "条")


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
    sh.write(0, 1, '性质')
    sh.write(0, 2, '地区')
    sh.write(0, 3, '电话')
    sh.write(0, 4, '地址')
    # 保存文件
    wb.save(file_dir+s+'.xls')


# 修改xls文件(不确定)
def wtxls(p, x, y, st):
    wb = xlwt.Workbook()
    sh = wb.add_sheet('A new sheet')
    sh.write(x, y, st)
    wb.save(file_dir+p+'.xls')


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
def school_count(pro,url, name):
    school_info = soupa(url).find(class_='xxsx')
    print(name + school_info.get_text())
    for div in school_info:
        # print(div.get_text())
        # if div.get_text.
        if "所属地区" in div.get_text():
            school_city = div.get_text()
            print(div.get_text())
        if "学校性质" in div.get_text():
            print(div.get_text())
        if "招生电话" in div.get_text():
            print(div.get_text())
        if "学校网址" in div.get_text():
            print(div.get_text())
        if "学校地址" in div.get_text():
            print(div.get_text())

    # TODO 解析全部信息
    # TODO 写入表格


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
            school_count(pro,school_url, school_name)

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