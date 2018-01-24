import urllib.request

from bs4 import BeautifulSoup  # 解析xml

from xls_file import nsfile, wtxls

base_url = "http://www.ruyile.com/xuexiao/?t=1"


# 模拟浏览器
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) '
                         'Chrome/51.0.2704.63 Safari/537.36'}


# TODO 改进
# TODO 1. 错误信息，跳过的信息，
# TODO 1. 显示当前的省份，总数量，当前数量
# TODO 2. 进度条百分比
# TODO 3. 多线程同时抓取
# TODO 4. UI界面
# TODO 5 .创建临时文件，当前省份全部爬完后，移动到用户目录

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
    print(name)
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
    # print('----------------------------------------------------')


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
    while current_page <= all_pages:
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


if __name__ == "__main__":
    main()
