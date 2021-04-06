# -*- coding: utf-8

import os
import re
import sys
import xlrd2
import time
import random
import pickle
import requests
import urllib.request
from requests_html import HTML
from requests_html import HTMLSession

ROOT_SHEET = "藏书目录"
ROOT_WEB = 'https://drive.my-elibrary.com'

ROOT_DIR = 'C:\\backup\\data'
ROOT_EXCEL = '1.xlsx'
ROOT_OBJECT = 'all.object'
ROOT_ERROR_OBJECT = 'err.object'
ROOT_SUCCESS_OBJECT = 'suc.object'
ROOT_ERROR_file = 'error.txt'
ROOT_ERR200_file = 'err200.txt'
GLOBAL_DOWN_LIST = []
GLOBAL_DOWN_ERROR = []
GLOBAL_DOWN_ERR200 = []
GLOBAL_DOWN_SUCCE = []
GLOBAL_DOWN_SIZE = 0

# 下载第三方站点文件
class GetOtherFile:
    def __init__(self, file, url):
        self.url = url
        self.file = file
        self.status = 0
        self.size = 0

    def __get_header(self):
        USER_AGENTS = [
            "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; AcooBrowser; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
            "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; Acoo Browser; SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 3.0.04506)",
            "Mozilla/4.0 (compatible; MSIE 7.0; AOL 9.5; AOLBuild 4337.35; Windows NT 5.1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
            "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)",
            "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 2.0.50727; Media Center PC 6.0)",
            "Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 1.0.3705; .NET CLR 1.1.4322)",
            "Mozilla/4.0 (compatible; MSIE 7.0b; Windows NT 5.2; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.2; .NET CLR 3.0.04506.30)",
            "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN) AppleWebKit/523.15 (KHTML, like Gecko, Safari/419.3) Arora/0.3 (Change: 287 c9dfb30)",
            "Mozilla/5.0 (X11; U; Linux; en-US) AppleWebKit/527+ (KHTML, like Gecko, Safari/419.3) Arora/0.6",
            "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.1.2pre) Gecko/20070215 K-Ninja/2.1.1",
            "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN; rv:1.9) Gecko/20080705 Firefox/3.0 Kapiko/3.0",
            "Mozilla/5.0 (X11; Linux i686; U;) Gecko/20070322 Kazehakase/0.4.5",
            "Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.8) Gecko Fedora/1.9.0.8-1.fc10 Kazehakase/0.5.6",
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) AppleWebKit/535.20 (KHTML, like Gecko) Chrome/19.0.1036.7 Safari/535.20",
            "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; fr) Presto/2.9.168 Version/11.52",
        ]
        random_agent = USER_AGENTS[random.randint(0, len(USER_AGENTS) - 1)]
        self.headers = {
            'User-Agent': random_agent,
        }

    # 获取包含下载链接的页面
    def __get_container_page(self):
        self.__get_header()
        try:
            session = HTMLSession()
            self.r = session.get(self.url, headers=self.headers)
            self.r.raise_for_status()
        except requests.exceptions.RequestException as e:
            # 获取容器页面出错
            print(e)
            self.r = None
            self.status = -1 # 获取容器页面出错
        return self.r

    # 获取外链的值
    def __get_3rd_url(self):
        self.__get_container_page()
        str_xpath = "/html/body/div[1]/div[3]/a"
        other_url_set = self.r.html.xpath(str_xpath, first=True).links
        url_list = list(other_url_set)
        if len(url_list) > 0:
            print("第三方链接：" + url_list[0])
            self.url_3rd = url_list[0]
        else:
            self.url_3rd = ""
            self.status = -2 # 解析容器页面出错

    def get_3rd_file(self):
        self.__get_3rd_url()
        try:
            r = requests.get(self.url_3rd, stream=True)
            r.raise_for_status()
            f = open(self.file, "wb")
            for chunk in r.iter_content(chunk_size=1024):
                if chunk:
                    f.write(chunk)

            f.flush()
            print("第三方站点下载完成(%s)" % (self.url))
        except requests.exceptions.RequestException as e:
            # 打印出错误信息
            print(e)
            self.status = -3 # 下载目标文件出错

        self.size = os.path.getsize(self.file)
        global GLOBAL_DOWN_ERR200
        GLOBAL_DOWN_ERR200.append(self)
        print("第三方站点下载完成(%d)" % (self.size))

class FileSave:
    # 传入文件名及URL都是已经处理过的，可以直接使用
    def __init__(self, file, url):
        self.name = file# 文件绝对路径
        self.dir = ""   # 保存绝对路径
        self.url = url  # URL绝对路径
        self.status = 0 # 状态 0(未下载) 1(成功下载)
        self.size = 0	# 文件大小

    def __get_header(self):
        USER_AGENTS = [
            "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; AcooBrowser; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
            "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; Acoo Browser; SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 3.0.04506)",
            "Mozilla/4.0 (compatible; MSIE 7.0; AOL 9.5; AOLBuild 4337.35; Windows NT 5.1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
            "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)",
            "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 2.0.50727; Media Center PC 6.0)",
            "Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 1.0.3705; .NET CLR 1.1.4322)",
            "Mozilla/4.0 (compatible; MSIE 7.0b; Windows NT 5.2; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.2; .NET CLR 3.0.04506.30)",
            "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN) AppleWebKit/523.15 (KHTML, like Gecko, Safari/419.3) Arora/0.3 (Change: 287 c9dfb30)",
            "Mozilla/5.0 (X11; U; Linux; en-US) AppleWebKit/527+ (KHTML, like Gecko, Safari/419.3) Arora/0.6",
            "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.1.2pre) Gecko/20070215 K-Ninja/2.1.1",
            "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN; rv:1.9) Gecko/20080705 Firefox/3.0 Kapiko/3.0",
            "Mozilla/5.0 (X11; Linux i686; U;) Gecko/20070322 Kazehakase/0.4.5",
            "Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.8) Gecko Fedora/1.9.0.8-1.fc10 Kazehakase/0.5.6",
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) AppleWebKit/535.20 (KHTML, like Gecko) Chrome/19.0.1036.7 Safari/535.20",
            "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; fr) Presto/2.9.168 Version/11.52",
        ]

        random_agent = USER_AGENTS[random.randint(0, len(USER_AGENTS) - 1)]
        self.headers = {
            'User-Agent': random_agent,
        }

    # 记录错误文件
    def __write_err_file(self):
        with open(ROOT_ERROR_file, 'a', encoding='utf-8') as f:
            f.write(self.url + '\n')
        return

    # 记录第三方网站信息
    def __write_200_file(self):
        with open(ROOT_ERR200_file, 'a', encoding='utf-8') as f:
            f.write(self.url + '\n')
        return

    def __ready_dir(self):
        end_pos = self.name.rfind("\\")
        if end_pos != -1:
            self.dir = self.name[0:end_pos]
        else:
            self.dir = self.name

        if not os.path.isdir(self.dir):
            #print("创建目录：" + self.dir)
            os.makedirs(self.dir)

    def __dow_error(self):
        print("下载出现异常。")

        if self.status_code == 200:
            # 外部链接
            Other = GetOtherFile(self.name, self.url)
            self.url = Other.get_save()
            self.down_file()

            global GLOBAL_DOWN_ERR200
            GLOBAL_DOWN_ERR200.append(self)
            self.__write_200_file()
        GLOBAL_DOWN_ERROR.append(self)
        self.write_err_file()

    def __get_length(self):
        is_chunked = self.response.headers.get('transfer-encoding', '') == 'chunked'
        content_length_s = self.response.headers.get('content-length')
        if not is_chunked and content_length_s.isdigit():
            content_length = int(content_length_s)
        else:
            content_length = 0

        return content_length

    def __print_info(self):
        str_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        print("下载总数据量：%.2f(MBit) 当前时间：%s" % (GLOBAL_DOWN_SIZE / (1024 * 1024), str_time))

    def down_file(self):
        # 判断路径是否存在
        self.__ready_dir()
        self.__get_header()
        # 判断文件是否存在
        global GLOBAL_DOWN_SIZE
        global GLOBAL_DOWN_SUCCE
        global GLOBAL_DOWN_ERROR
        fsize = 0
        if os.path.exists(self.name):
            fsize = os.path.getsize(self.name)

        start = time.time()  # 下载开始时间
        try:
            self.response = requests.get(self.url, headers=self.headers, stream=True)
            size = 0  # 初始化已下载大小
            chunk_size = 1024  # 每次下载的数据大小
            self.response.raise_for_status()
            # 判断是否响应成功
            if self.response.status_code == 200:
                self.size = self.__get_length()  # 下载文件总大小
                if fsize != 0 and fsize == self.size:
                    print("文件已存在且文件的大小(%s)检查正确，无需下载！ {%s}" % (size2human(fsize), self.name))
                    # 更新下载成功列表
                    GLOBAL_DOWN_SUCCE.append(self)
                    return
                if self.size == 0:
                    # 未成功获取到下载文件的大小，可能是文件保存在第三方站点，需要另外下载
                    print('不能正确获取到将下载文件大小(%s)' % (self.url))
                    other_3rd_file = GetOtherFile(self.name, self.url)
                    # 检查文件是否存在，如果存在，则默认是正确下载了
                    if os.path.exists(self.name):
                        print("可能是第三方站点下载，文件已存在默认为已成功下载：%s" % (self.name))
                        GLOBAL_DOWN_SUCCE.append(self)
                    else:
                        other_3rd_file.get_3rd_file()
                        if other_3rd_file.status != 0:
                            print("第三方站点下载失败，文件信息：%s" %(other_3rd_file.file))
                            print("第三方站点下载失败，连接信息：%s" % (other_3rd_file.url))
                            GLOBAL_DOWN_ERROR.append(self)
                            self.__write_err_file(self)
                        else:
                            print("第三方站点下载成功，连接信息：%s" % (other_3rd_file.file))
                            GLOBAL_DOWN_SUCCE.append(self)
                            GLOBAL_DOWN_SIZE = GLOBAL_DOWN_SIZE + other_3rd_file.size
                            self.__print_info()
                    return
                # 当前站点下载且没有下载过此文件
                if self.size != 0 and fsize == 0:
                    print('开始下载：%s' % (self.url))
                    print('开始下载，文件大小:[{size:.2f}] MB'.format(size=self.size / chunk_size / 1024))  # 开始下载，显示下载文件大小
                    with open(self.name, 'wb') as file:  # 显示进度条
                        for data in self.response.iter_content(chunk_size=chunk_size):
                            file.write(data)
                            size += len(data)
                            now_t = time.time()
                            print('\r' + '[下载进度]：%s%.2f%%(实时速度：%.2fKB/秒)' % ('>' * int(size * 50 / self.size), float(size / self.size * 100), float(size / ((now_t - start) * 1024))), end=' ')
                        file.flush()

                    end = time.time()  # 下载结束时间
                    total = end - start
                    speed = self.size / (total * 1024)
                    print('下载完成！用时： %.2f秒，占用带宽：%.2fKBit/秒' % (total, speed))  # 输出下载用时时间
                    fsize = os.path.getsize(self.name)
                    if fsize == self.size:
                        self.status = 0
                        # 更新已经下载总量
                        GLOBAL_DOWN_SIZE = GLOBAL_DOWN_SIZE + self.size
                        # 更新下载成功列表
                        GLOBAL_DOWN_SUCCE.append(self)
                        print('文件下载成功！(%s)' % (self.name))
                    else:
                        print("文件下载失败因大小不正确(%d/%d)(%s)。" % (fsize, self.size, self.name))
                        GLOBAL_DOWN_ERROR.append(self)
                        self.__write_err_file()

            str_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
            print("下载总数据量：%.2f(MBit) 当前时间：%s" % (GLOBAL_DOWN_SIZE/(1024 * 1024), str_time))
        except requests.exceptions.ConnectionError as error:
            self.status = -2 # 连接出错，需要重新连接
            print("\033[1;31m文件下载失败：" + self.url + "\033[0m")
        else:
            self.status = -1 # 发生异常
            print("\033[1;31m文件下载失败：" + self.url + "，错误码：" + str(self.response.status_code) + "\033[0m")
            GLOBAL_DOWN_ERROR.append(self)
            self.__write_err_file()


def get_down_object():
    if os.path.exists(ROOT_OBJECT):
        print("对象文件已存在，可以直接读取到对象列表中。")
        f = open(ROOT_OBJECT, 'rb')
        global GLOBAL_DOWN_LIST
        GLOBAL_DOWN_LIST = pickle.load(f)
    else:
        get_books_list()

def save_error_obj():
    global GLOBAL_DOWN_ERROR
    l = len(GLOBAL_DOWN_ERROR)
    if l % 100 == 0:
        update_status_object_file

def download_file():
    number = len(GLOBAL_DOWN_LIST)
    i = 1
    for o in GLOBAL_DOWN_LIST:
        str_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        print("开始处理：当前下载/总数[%d/%d]，状态：已成功/别处下载/错误[%d/%d/%d] (%s)" % (i, number, len(GLOBAL_DOWN_SUCCE), len(GLOBAL_DOWN_ERR200), len(GLOBAL_DOWN_ERROR), str_time))
        i = i + 1
        o.down_file()
        save_error_obj()

# 大小单位转换
SUFFIXES = {1000:['KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'],
            1024:['KiB', 'MiB', 'GiB', 'TiB', 'PiB', 'EiB', 'ZiB', 'YiB']}
def size2human(size,is_1024_byte=False):
    #mutiple默认是1024
    mutiple=1000 if is_1024_byte else 1024
    #与for遍历结合起来，这样来进行递级的转换
    for suffix in SUFFIXES[mutiple]:
        size/=mutiple
        #直到Size小于能往下一个单位变的数值
        if size<mutiple:
            return '{0:.1f}{1}'.format(size,suffix)
    raise ValueError('number too large') #抛出异常

def get_books_list():
    book = xlrd2.open_workbook(ROOT_EXCEL)
    sheet = book.sheet_by_name(ROOT_SHEET)

    for i in range(sheet.nrows):
        url_list = sheet.row_values(i)      # 简体文件名 | 繁体文件名 | 网站目录
        save_dir = ROOT_DIR + url_list[2]   # 保存路径
        # 特殊字符替换
        save_dir = save_dir.replace('?', '.')
        down_url = get_down_url(url_list[2])# 下载URL

        file_obj = FileSave(save_dir, down_url)
        global GLOBAL_DOWN_LIST
        GLOBAL_DOWN_LIST.append(file_obj)

    update_object_file()

def update_object_file():
    global GLOBAL_DOWN_LIST
    f = open(ROOT_OBJECT, 'wb')
    pickle.dump(GLOBAL_DOWN_LIST, f)

def update_status_object_file():
    global GLOBAL_DOWN_ERROR
    f = open(ROOT_ERROR_OBJECT, 'wb')
    pickle.dump(GLOBAL_DOWN_ERROR, f)
    print("下载成功/错误文件个数：[%d/%d]" % (len(GLOBAL_DOWN_SUCCE), len(GLOBAL_DOWN_ERROR)))

def get_down_url(url):
    url = url.replace("\\", "/")
    ret = ROOT_WEB + url
    return ret

def init_log():
    # 初始化环境
    str_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    with open(ROOT_ERROR_file, 'w') as f:
        f.write(str_time + '\n')
    with open(ROOT_ERR200_file, 'a') as f:
        f.write(str_time + '\n')
# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    init_log()
    get_down_object()
    download_file()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
