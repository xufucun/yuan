# 抓取数据
import urllib.request
from bs4 import BeautifulSoup
import xlrd
import xlwt

# 网址（不包含页数）
base_url = "http://www.ruyile.com/xuexiao/?t=1&p="
# 模拟浏览器
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) '
                         'Chrome/51.0.2704.63 Safari/537.36'}

page_hover_int = 1
page_all_int = 105981


# 创建表格

# 写入方法

# 判断省份表格是否存在 存在 写入 ，不存在 创建 写入

# workbook = xlrd.open_workbook(r'F:\demo.xlsx')


# 单个学校信息
def school(url, name):
    # print(url)
    request_scholl = urllib.request.Request(url=url, headers=headers)
    response_school = urllib.request.urlopen(request_scholl)
    data_school = response_school.read()
    soup_school = BeautifulSoup(data_school, 'lxml')
    school_info = soup_school.find(class_='xxsx')
    print(name + school_info.get_text())
    # TODO 读取省份
    # TODO 批判文件名省份文件名是否存在， 不存在则创建
    # TODO 写入表格



while page_hover_int < page_all_int:
    print(page_hover_int)

    y_url = (base_url + str(page_hover_int)).replace(' ', '')  # 拼接URL.去除空格
    print(y_url)

    request = urllib.request.Request(url=y_url, headers=headers)
    response = urllib.request.urlopen(request)
    data = response.read()
    soup = BeautifulSoup(data, 'lxml')

    divs = soup.find_all(class_='sk')
    for div in divs:
        school_con = div.h4.a
        school_url = school_con.get('href')
        school_name = school_con.get_text()
        school(school_url, school_name)

    page_hover_int += 1
