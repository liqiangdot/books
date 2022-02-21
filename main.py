# -*- coding: utf-8

import os
import gc
import re
import sys
import xlrd2
import time
import math
import random
import pickle
import requests
import threading
import urllib.request
from requests_html import HTML
from requests_html import HTMLSession
import socket

socket.setdefaulttimeout(30)

ROOT_SHEET = "藏书目录"
ROOT_WEB = 'https://drive.my-elibrary.com'

ROOT_DIR = 'Z:\\backup_other'
ROOT_EXCEL = '1.xlsx'
# 对象列表
ROOT_OBJECT = 'all.object'
# 记录所有错误列表
ROOT_ERROR_file = 'error.txt'
# 记录异常事件
ROOT_ERROR_EXECPT_file = 'except.txt'
# 记录404事件列表
ROOT_404ERR_file = '404err.txt'
# 记录非404事件列表
ROOT_NO404ERR_file = 'no404err.txt'
# 记录写文件失败事件列表
ROOT_WRITE_ERR_file = 'writeerr.txt'

ROOT_NUMBER_SUCC_file = 'number.txt'
ROOT_NUMBER_SUCC_file2 = 'number2.txt'

GLOBAL_DOWN_LIST = []

# 总下载个数
GLOBAL_DOWN_NUMBER = 0
# 下载出错个数
GLOBAL_DOWN_ERROR_NUMBER = 0
# 第三方网站下载成功个数
GLOBAL_DOWN_3RDLIST_NUMBER = 0
# 下载成功个数
GLOBAL_DOWN_SUCCE_NUMBER = 0
# 总下载数据字节数
GLOBAL_DOWN_SIZE = 0
# 每次执行要下载成功的文件数
GLOBAL_DOWN_FILE_NUMBER = 2048

# 当前下载数据字节数
GLOBAL_NOW_DOWN_SIZE = 0

# 下载片大小
GLOBAL_DOWN_MAX = (1024 * 1024)
# 开始时间
GLOBAL_START_TIME = 0
# 超时时间
GLOBAL_TIMEOUT = 16
# 退出标志
GLOBAL_READY_EXIT_1 = 0
GLOBAL_READY_EXIT_2 = 0


# 生成随机的UA
def get_ua():
    user_agents = [
        "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/14.0.835.163 Safari/535.1",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_0) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11",
        "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:6.0) Gecko/20100101 Firefox/6.0",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.6; rv:2.0.1) Gecko/20100101 Firefox/4.0.1",
        "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_8; en-us) AppleWebKit/534.50 (KHTML, like Gecko) Version/5.1 Safari/534.50",
        "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-us) AppleWebKit/534.50 (KHTML, like Gecko) Version/5.1 Safari/534.50",
        "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; en) Presto/2.8.131 Version/11.11",
        "Opera/9.80 (Windows NT 6.1; U; en) Presto/2.8.131 Version/11.11",
        "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko",
        "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0;",
        "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
        "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0)",
        "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)",
        "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; Trident/4.0; GTB7.0)",
        "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1)",
        "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1)",
        "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; Maxthon 2.0)",
        "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; Trident/4.0; SE 2.X MetaSr 1.0; SE 2.X MetaSr 1.0; .NET CLR 2.0.50727; SE 2.X MetaSr 1.0)",
        "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; 360SE)",
        "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; TencentTraveler 4.0)",
        "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/534.50 (KHTML, like Gecko) Version/5.1 Safari/534.50",
        "Opera/9.80 (Windows NT 6.1; U; zh-cn) Presto/2.9.168 Version/11.50",
        "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 2.0.50727; SLCC2; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; Tablet PC 2.0; .NET4.0E)",
        "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; InfoPath.3)",
        "Mozilla/5.0 (Windows; U; Windows NT 6.1; ) AppleWebKit/534.12 (KHTML, like Gecko) Maxthon/3.0 Safari/534.12",
        "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E)",
        "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E; SE 2.X MetaSr 1.0)",
        "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/534.3 (KHTML, like Gecko) Chrome/6.0.472.33 Safari/534.3 SE 2.X MetaSr 1.0",
        "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E)",
        "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/13.0.782.41 Safari/535.1 QQBrowser/6.9.11079.201",
        "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E) QQBrowser/6.9.11079.20a1",
        "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)",
        "Mozilla/5.0 (Linux; U; Android 2.3.6; en-us; Nexus S Build/GRK39F) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1",
        "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/532.5 (KHTML, like Gecko) Chrome/4.0.249.0 Safari/532.5",
        "Mozilla/5.0 (Windows; U; Windows NT 5.2; en-US) AppleWebKit/532.9 (KHTML, like Gecko) Chrome/5.0.310.0 Safari/532.9",
        "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US) AppleWebKit/534.7 (KHTML, like Gecko) Chrome/7.0.514.0 Safari/534.7",
        "Mozilla/5.0 (Windows; U; Windows NT 6.0; en-US) AppleWebKit/534.14 (KHTML, like Gecko) Chrome/9.0.601.0 Safari/534.14",
        "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/534.14 (KHTML, like Gecko) Chrome/10.0.601.0 Safari/534.14",
        "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/534.20 (KHTML, like Gecko) Chrome/11.0.672.2 Safari/534.20",
        "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/534.27 (KHTML, like Gecko) Chrome/12.0.712.0 Safari/534.27",
        "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/13.0.782.24 Safari/535.1",
        "Mozilla/5.0 (Windows NT 6.0) AppleWebKit/535.2 (KHTML, like Gecko) Chrome/15.0.874.120 Safari/535.2",
        "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.36 Safari/535.7",
        "Mozilla/5.0 (Windows; U; Windows NT 6.0 x64; en-US; rv:1.9pre) Gecko/2008072421 Minefield/3.0.2pre",
        "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.10) Gecko/2009042316 Firefox/3.0.10",
        "Mozilla/5.0 (Windows; U; Windows NT 6.0; en-GB; rv:1.9.0.11) Gecko/2009060215 Firefox/3.0.11 (.NET CLR 3.5.30729)",
        "Mozilla/5.0 (Windows; U; Windows NT 6.0; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 GTB5",
        "Mozilla/5.0 (Windows; U; Windows NT 5.1; tr; rv:1.9.2.8) Gecko/20100722 Firefox/3.6.8 ( .NET CLR 3.5.30729; .NET4.0E)",
        "Mozilla/5.0 (Windows NT 6.1; rv:2.0.1) Gecko/20100101 Firefox/4.0.1",
        "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:2.0.1) Gecko/20100101 Firefox/4.0.1",
        "Mozilla/5.0 (Windows NT 5.1; rv:5.0) Gecko/20100101 Firefox/5.0",
        "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:6.0a2) Gecko/20110622 Firefox/6.0a2",
        "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:7.0.1) Gecko/20100101 Firefox/7.0.1",
        "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:2.0b4pre) Gecko/20100815 Minefield/4.0b4pre",
        "Mozilla/4.0 (compatible; MSIE 5.5; Windows NT 5.0 )",
        "Mozilla/4.0 (compatible; MSIE 5.5; Windows 98; Win 9x 4.90)",
        "Mozilla/5.0 (Windows; U; Windows XP) Gecko MultiZilla/1.6.1.0a",
        # "Mozilla/2.02E (Win95; U)",
        # "Mozilla/3.01Gold (Win95; I)",
        # "Mozilla/4.8 [en] (Windows NT 5.1; U)",
        # "Mozilla/5.0 (Windows; U; Win98; en-US; rv:1.4) Gecko Netscape/7.1 (ax)",
        "HTC_Dream Mozilla/5.0 (Linux; U; Android 1.5; en-ca; Build/CUPCAKE) AppleWebKit/528.5  (KHTML, like Gecko) Version/3.1.2 Mobile Safari/525.20.1",
        "Mozilla/5.0 (hp-tablet; Linux; hpwOS/3.0.2; U; de-DE) AppleWebKit/534.6 (KHTML, like Gecko) wOSBrowser/234.40.1 Safari/534.6 TouchPad/1.0",
        "Mozilla/5.0 (Linux; U; Android 1.5; en-us; sdk Build/CUPCAKE) AppleWebkit/528.5  (KHTML, like Gecko) Version/3.1.2 Mobile Safari/525.20.1",
        "Mozilla/5.0 (Linux; U; Android 2.1; en-us; Nexus One Build/ERD62) AppleWebKit/530.17 (KHTML, like Gecko) Version/4.0 Mobile Safari/530.17",
        "Mozilla/5.0 (Linux; U; Android 2.2; en-us; Nexus One Build/FRF91) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1",
        "Mozilla/5.0 (Linux; U; Android 1.5; en-us; htc_bahamas Build/CRB17) AppleWebKit/528.5  (KHTML, like Gecko) Version/3.1.2 Mobile Safari/525.20.1",
        "Mozilla/5.0 (Linux; U; Android 2.1-update1; de-de; HTC Desire 1.19.161.5 Build/ERE27) AppleWebKit/530.17 (KHTML, like Gecko) Version/4.0 Mobile Safari/530.17",
        "Mozilla/5.0 (Linux; U; Android 2.2; en-us; Sprint APA9292KT Build/FRF91) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1",
        "Mozilla/5.0 (Linux; U; Android 1.5; de-ch; HTC Hero Build/CUPCAKE) AppleWebKit/528.5  (KHTML, like Gecko) Version/3.1.2 Mobile Safari/525.20.1",
        "Mozilla/5.0 (Linux; U; Android 2.2; en-us; ADR6300 Build/FRF91) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1",
        "Mozilla/5.0 (Linux; U; Android 2.1; en-us; HTC Legend Build/cupcake) AppleWebKit/530.17 (KHTML, like Gecko) Version/4.0 Mobile Safari/530.17",
        "Mozilla/5.0 (Linux; U; Android 1.5; de-de; HTC Magic Build/PLAT-RC33) AppleWebKit/528.5  (KHTML, like Gecko) Version/3.1.2 Mobile Safari/525.20.1 FirePHP/0.3",
        "Mozilla/5.0 (Linux; U; Android 1.6; en-us; HTC_TATTOO_A3288 Build/DRC79) AppleWebKit/528.5  (KHTML, like Gecko) Version/3.1.2 Mobile Safari/525.20.1",
        "Mozilla/5.0 (Linux; U; Android 1.0; en-us; dream) AppleWebKit/525.10  (KHTML, like Gecko) Version/3.0.4 Mobile Safari/523.12.2",
        "Mozilla/5.0 (Linux; U; Android 1.5; en-us; T-Mobile G1 Build/CRB43) AppleWebKit/528.5  (KHTML, like Gecko) Version/3.1.2 Mobile Safari 525.20.1",
        "Mozilla/5.0 (Linux; U; Android 1.5; en-gb; T-Mobile_G2_Touch Build/CUPCAKE) AppleWebKit/528.5  (KHTML, like Gecko) Version/3.1.2 Mobile Safari/525.20.1",
        "Mozilla/5.0 (Linux; U; Android 2.0; en-us; Droid Build/ESD20) AppleWebKit/530.17 (KHTML, like Gecko) Version/4.0 Mobile Safari/530.17",
        "Mozilla/5.0 (Linux; U; Android 2.2; en-us; Droid Build/FRG22D) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1",
        "Mozilla/5.0 (Linux; U; Android 2.0; en-us; Milestone Build/ SHOLS_U2_01.03.1) AppleWebKit/530.17 (KHTML, like Gecko) Version/4.0 Mobile Safari/530.17",
        "Mozilla/5.0 (Linux; U; Android 2.0.1; de-de; Milestone Build/SHOLS_U2_01.14.0) AppleWebKit/530.17 (KHTML, like Gecko) Version/4.0 Mobile Safari/530.17",
        "Mozilla/5.0 (Linux; U; Android 3.0; en-us; Xoom Build/HRI39) AppleWebKit/525.10  (KHTML, like Gecko) Version/3.0.4 Mobile Safari/523.12.2",
        "Mozilla/5.0 (Linux; U; Android 0.5; en-us) AppleWebKit/522  (KHTML, like Gecko) Safari/419.3",
        "Mozilla/5.0 (Linux; U; Android 1.1; en-gb; dream) AppleWebKit/525.10  (KHTML, like Gecko) Version/3.0.4 Mobile Safari/523.12.2",
        "Mozilla/5.0 (Linux; U; Android 2.0; en-us; Droid Build/ESD20) AppleWebKit/530.17 (KHTML, like Gecko) Version/4.0 Mobile Safari/530.17",
        "Mozilla/5.0 (Linux; U; Android 2.1; en-us; Nexus One Build/ERD62) AppleWebKit/530.17 (KHTML, like Gecko) Version/4.0 Mobile Safari/530.17",
        "Mozilla/5.0 (Linux; U; Android 2.2; en-us; Sprint APA9292KT Build/FRF91) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1",
        "Mozilla/5.0 (Linux; U; Android 2.2; en-us; ADR6300 Build/FRF91) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1",
        "Mozilla/5.0 (Linux; U; Android 2.2; en-ca; GT-P1000M Build/FROYO) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1",
        "Mozilla/5.0 (Linux; U; Android 3.0.1; fr-fr; A500 Build/HRI66) AppleWebKit/534.13 (KHTML, like Gecko) Version/4.0 Safari/534.13",
        "Mozilla/5.0 (Linux; U; Android 3.0; en-us; Xoom Build/HRI39) AppleWebKit/525.10  (KHTML, like Gecko) Version/3.0.4 Mobile Safari/523.12.2",
        "Mozilla/5.0 (Linux; U; Android 1.6; es-es; SonyEricssonX10i Build/R1FA016) AppleWebKit/528.5  (KHTML, like Gecko) Version/3.1.2 Mobile Safari/525.20.1",
        "Mozilla/5.0 (Linux; U; Android 1.6; en-us; SonyEricssonX10i Build/R1AA056) AppleWebKit/528.5  (KHTML, like Gecko) Version/3.1.2 Mobile Safari/525.20.1"
    ]
    random_agent = user_agents[random.randint(0, len(user_agents) - 1)]
    headers = {
        'User-Agent': random_agent,
    }
    return headers


# 输入下载信息
def print_down_info():
    str_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    print("总数据量：[%s]，成功/错误/三方/总数：[%d/%d/%d/%d]，当前时间：[%s]" % (
        size2human(GLOBAL_DOWN_SIZE), GLOBAL_DOWN_SUCCE_NUMBER, GLOBAL_DOWN_ERROR_NUMBER, GLOBAL_DOWN_3RDLIST_NUMBER,
        GLOBAL_DOWN_NUMBER, str_time))


def write_number_file(e, file):
    with open(file, 'w+', encoding='utf-8') as f:
        f.write(str(e))


def read_number_file(file):
    if os.path.exists(file):
        with open(file, 'r', encoding='utf-8') as f:
            num = f.read()
            if num == '':
                return 0
            else:
                return int(num)
    else:
        return 0


# 记录异常事件
def write_except_file(e):
    with open(ROOT_ERROR_EXECPT_file, 'a', encoding='utf-8') as f:
        f.write(repr(e) + '\n')


# 读写文件错误
def file_write_err_file(file):
    with open(ROOT_WRITE_ERR_file, 'a', encoding='utf-8') as f:
        f.write(file + '\n')


# 默认是写错误日志和非404日志。
def write_err_file(url, status=0):
    with open(ROOT_ERROR_file, 'a', encoding='utf-8') as f:
        f.write(url + '\n')
    # 非404事件
    if status == 0:
        with open(ROOT_NO404ERR_file, 'a', encoding='utf-8') as f:
            f.write(url + '\n')
    else:
        with open(ROOT_404ERR_file, 'a', encoding='utf-8') as f:
            f.write(url + '\n')


# 大小单位转换
SUFFIXES = {1000: ['KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'],
            1024: ['KiB', 'MiB', 'GiB', 'TiB', 'PiB', 'EiB', 'ZiB', 'YiB']}


def size2human(size, is_1024_byte=False):
    # mutiple默认是1024
    mutiple = 1000 if is_1024_byte else 1024
    # 与for遍历结合起来，这样来进行递级的转换
    for suffix in SUFFIXES[mutiple]:
        size /= mutiple
        # 直到Size小于能往下一个单位变的数值
        if size < mutiple:
            return '{0:.1f}{1}'.format(size, suffix)
    raise ValueError('number too large')  # 抛出异常


# 秒转换成分、小时和天
def size2Time(allTime):
    day = 24 * 60 * 60
    hour = 60 * 60
    min = 60
    if allTime < 60:
        return "%d sec" % math.ceil(allTime)
    elif allTime > day:
        days = divmod(allTime, day)
        return ("%d 天, %s" % (int(days[0]), size2Time(days[1])))
    elif allTime > hour:
        hours = divmod(allTime, hour)
        return ('%d 小时, %s' % (int(hours[0]), size2Time(hours[1])))
    else:
        mins = divmod(allTime, min)
        return ("%d 分, %d 秒" % (int(mins[0]), math.ceil(mins[1])))


# 下载第三方站点文件
class GetOtherFile:
    def __init__(self, file, url, dir):
        self.url = url
        self.file = file
        self.status = 0
        self.size = 0
        self.url_lists = []
        self.dir = dir  # 绝对路径
        self.url_3rd = ""

    def __get_header(self):
        self.headers = get_ua()

    # 获取包含下载链接的页面
    def __get_container_page(self):
        self.__get_header()
        try:
            session = HTMLSession()
            self.r = session.get(self.url, headers=self.headers, timeout=GLOBAL_TIMEOUT)
            self.r.raise_for_status()
        except (requests.exceptions.HTTPError, requests.exceptions.ConnectionError,
                requests.exceptions.Timeout, requests.exceptions.RequestException, Exception) as err:
            print("Error:", err)
            self.r = None
            self.status = -1  # 获取容器页面出错
            print("获取包含第三方站点链接页面失败，URL[%s]，文件名[%s]" % (self.url, self.file))
            write_except_file(err)

    # 获取外链的值
    def __get_3rd_url(self):
        self.__get_container_page()
        if self.status < 0:
            return  # 获取第三方站点链接页面出错，返回。

        # 增加第三方链接判断
        try:
            ret_find = self.r.url.find("cdmission-my.sharepoint.com")
            if ret_find != -1:
                self.url_lists.append(self.r.url)
                print("第三方链接：" + self.r.url)
                return
        except Exception as err:
            write_except_file(err)
            return

        str_xpath = "/html/body/div[1]/div[3]/a"
        url_list = []
        try:
            other_url_set = self.r.html.xpath(str_xpath, first=True)  # .links
            if other_url_set == None:
                # 获取外网第三方链接失败，尝试另外一种方法
                self.url_lists.clear()
                items = self.r.html.find('.item')
                for item in items:
                    size = item.find('.size', first=True)
                    if size.text == 'NaN KB' or size.text == '':
                        continue
                    # 需要下载的文件链接
                    url_list = list(item.absolute_links)  # set转list
            else:
                url_list = list(other_url_set.links)  # set转list

            if len(url_list) == 0:
                self.status = -1
            else:
                for l in url_list:
                    self.url_lists.append(l)
                    print("第三方链接：" + l)

        except Exception as err:
            write_except_file(err)

    # 下载多个文件
    def get_3rd_files(self):
        self.__get_3rd_url()
        if self.status < 0:
            return  # 获取第三方站点链接出错，返回。

        if len(self.url_3rd.strip()) != 0:
            try:
                r = requests.get(self.url_3rd.split('\n')[0], headers=self.headers, stream=True, timeout=GLOBAL_TIMEOUT)
                r.raise_for_status()
                i = 0
                size = 0
                start_t = time.time()
                with open(self.file, 'wb') as f:  # 显示进度条
                    for data in r.iter_content(chunk_size=GLOBAL_DOWN_MAX):
                        if data:
                            f.write(data)
                            size += len(data)
                            i = i + 1
                            now_t = time.time()
                            print(
                                '\r' + '[下载进度]：%s(实时速度：%.2fKB/秒)' % ('>' * i, float(size / ((now_t - start_t) * 1024))),
                                end=' ')
                            f.flush()

            except (requests.exceptions.HTTPError, requests.exceptions.ConnectionError,
                    requests.exceptions.Timeout, requests.exceptions.RequestException, Exception) as err:
                print("Error:", err)
                self.r = None
                self.status = -3  # 第三方站点下载目标文件出错
                print("下载第三方站点文件失败，URL[%s]，文件名[%s]" % (self.url, self.file))
                write_except_file(err)
                write_err_file(self.url)
        else:
            i = 1
            for url in self.url_lists:
                print(url)
                filename = os.path.basename(url)
                if len(filename) > 64:
                    filename = os.path.basename(self.file)
                self.file = self.dir + "\\" + str(i) + filename
                try:
                    r = requests.get(url.split('\n')[0], headers=self.headers, stream=True,
                                     timeout=GLOBAL_TIMEOUT)
                    r.raise_for_status()
                    i = 0
                    size = 0
                    start_t = time.time()
                    with open(self.file, 'wb') as f:  # 显示进度条
                        for data in r.iter_content(chunk_size=GLOBAL_DOWN_MAX):
                            if data:
                                f.write(data)
                                size += len(data)
                                i = i + 1
                                now_t = time.time()
                                # print('\r' + '[下载进度]：%s(实时速度：%.2fKB/秒)' % (
                                # '>' * i, float(size / ((now_t - start_t) * 1024))), end=' ')
                                f.flush()

                    # 获取文件大小
                    self.size = os.path.getsize(self.file)
                    print("第三方站点下载完成[%d][%s]" % (self.size, self.file))
                    self.status = 0

                except (requests.exceptions.HTTPError, requests.exceptions.ConnectionError,
                        requests.exceptions.Timeout, requests.exceptions.RequestException, Exception) as err:
                    print("Error:", err)
                    self.r = None
                    self.status = -3  # 第三方站点下载目标文件出错
                    print("\033[1;31m下载第三方站点文件失败，URL[%s]，文件名[%s]\033[0m" % (self.url, self.file))
                    write_except_file(err)
                    write_err_file(url)


class C_DownFiles:
    # 传入文件名及URL都是已经处理过的，可以直接使用
    def __init__(self, file, url):
        self.name = file  # 文件绝对路径
        self.dir = ""  # 保存绝对路径
        self.url = url  # URL绝对路径
        self.status = 0  # 状态 0(未下载) 1(成功下载)
        self.size = 0  # 文件大小
        self.redirect = False

    def __get_header(self):
        self.headers = get_ua()

    # 记录错误文件
    def __write_err_file(self):
        with open(ROOT_ERROR_file, 'a', encoding='utf-8') as f:
            f.write(self.url + '\n')

    def __write_404err_file(self):
        with open(ROOT_404ERR_file, 'a', encoding='utf-8') as f:
            f.write(self.url + '\n')

    # 预备目录环境
    def __ready_dir(self):
        end_pos = self.name.rfind("\\")
        if end_pos != -1:
            self.dir = self.name[0:end_pos]
        else:
            self.dir = self.name

        self.dir = self.dir.replace("?", "")

        if not os.path.isdir(self.dir):
            # print("创建目录：" + self.dir)
            os.makedirs(self.dir)

    # 获取远程文件大小
    def __get_length(self):
        is_chunked = self.response.headers.get('transfer-encoding', '') == 'chunked'
        content_length_s = self.response.headers.get('content-length')
        if not is_chunked and content_length_s.isdigit():
            content_length = int(content_length_s)
        else:
            content_length = 0
        return content_length

    def get_error_list(self):
        self.err_file = ROOT_ERROR_file
        for line in open(self.err_file, encoding='utf-8'):
            self.url = line
            f_p = self.url.find(ROOT_WEB)
            if f_p != -1:
                file_name = self.url[f_p + len(ROOT_WEB):]
                file_name = file_name.replace("/", "\\")
                file_name = ROOT_DIR + file_name
                self.name = file_name

                print(self.name)
                print(self.url)

                self.down_file()

    # 默认为仅检查文件是否存在且大小不为零则认为是已经成功下载
    def down_file(self, chk=True):
        # 判断路径是否存在
        self.__ready_dir()
        self.__get_header()
        # 判断文件是否存在
        global GLOBAL_DOWN_SIZE
        global GLOBAL_NOW_DOWN_SIZE
        global GLOBAL_DOWN_SUCCE_NUMBER
        global GLOBAL_DOWN_ERROR_NUMBER
        global GLOBAL_DOWN_3RDLIST_NUMBER

        if os.path.exists(self.name):
            fsize = os.path.getsize(self.name)
        else:
            fsize = 0

        start = time.time()  # 下载开始时间
        try:
            self.response = requests.get(self.url.split('\n')[0], headers=self.headers, stream=True,
                                         timeout=GLOBAL_TIMEOUT, allow_redirects=True)
            try:
                size = 0  # 初始化已下载大小
                chunk_size = GLOBAL_DOWN_MAX  # 每次下载的数据大小
                self.response.raise_for_status()
                if self.response.status_code == 302:
                    self.redirect = True
                    self.url = "https:" + self.response.headers['Location']
                    print('下载转向到(%s)' % (self.url))
                    self.response = requests.get(self.url.split('\n')[0], headers=self.headers, stream=True,
                                                 timeout=GLOBAL_TIMEOUT, allow_redirects=False)
                    self.response.raise_for_status()
                # 判断是否响应成功
                if self.response.status_code == 200:
                    self.size = self.__get_length()  # 下载文件总大小
                    # 当前站点下载且下载成功
                    if fsize != 0 and fsize == self.size:
                        # 更新下载成功列表
                        GLOBAL_DOWN_SUCCE_NUMBER += 1
                        # 更新已下载数据
                        GLOBAL_DOWN_SIZE = GLOBAL_DOWN_SIZE + fsize
                        print("文件之前已下载成功，大小(%s)名称(%s)" % (size2human(fsize), self.name))
                        self.response.close()
                        print_down_info()
                        return
                    # 未成功获取到下载文件的大小，可能是文件保存在第三方站点，需要另外下载
                    if self.size == 0 and self.redirect == False:
                        # reditList = self.response.history  # 可以看出获取的是一个地址序列
                        # print(f'获取重定向最终的url：{reditList[len(reditList) - 1].headers["Location"]}')
                        print('在本站点不能获取到将下载文件大小(%s)' % (self.url))
                        if os.path.exists(self.name):
                            fsize = os.path.getsize(self.name)
                            if fsize != 0:
                                print("可能是第三方站点下载，文件已存在且大小不为零：%s" % (self.name))

                        # 需要下载第三方站点文件
                        other_3rd_file = GetOtherFile(self.name, self.url, self.dir)
                        # 检查文件是否存在，如果存在，则默认是正确下载了
                        other_3rd_file.get_3rd_files()
                        if other_3rd_file.status != 0:
                            print("第三方站点下载失败，文件信息：%s \n连接信息：%s" % (other_3rd_file.file, other_3rd_file.url))
                            GLOBAL_DOWN_ERROR_NUMBER += 1
                            self.status = -1
                            self.__write_err_file()
                        else:
                            print("第三方站点下载成功，连接信息：%s" % (other_3rd_file.file))
                            GLOBAL_DOWN_SUCCE_NUMBER += 1
                            GLOBAL_DOWN_3RDLIST_NUMBER += 1
                            GLOBAL_DOWN_SIZE = GLOBAL_DOWN_SIZE + other_3rd_file.size
                            print_down_info()
                        return
                    # 当前站点下载且没有下载过此文件
                    if ((self.size != 0 and fsize == 0) or
                            (self.size != 0 and fsize != 0 and self.size != fsize) or
                            (self.redirect == True)):
                        print('开始下载：%s' % (self.url))
                        with open(self.name, 'wb') as file:  # 显示进度条
                            for data in self.response.iter_content(chunk_size=GLOBAL_DOWN_MAX):
                                file.write(data)
                                size += len(data)
                                now_t = time.time()
                                if self.size == 0:
                                    print('\r' + '[文件下载实时速度：%.2fKB/秒)' % (float(size / ((now_t - start) * 1024))),
                                          end=' ')
                                else:
                                    print('\r' + '[文件大小(%s)]：%s%.2f%%(实时速度：%.2fKB/秒)' % ((size2human(size=self.size)),
                                                                                         '>' * int(
                                                                                             size * 50 / self.size),
                                                                                         float(size / self.size * 100),
                                                                                         float(size / ((
                                                                                                                   now_t - start) * 1024))),
                                          end=' ')
                                    GLOBAL_NOW_DOWN_SIZE += len(data)
                            file.flush()

                        end = time.time()  # 下载结束时间
                        total = end - start
                        speed = self.size / (total * 1024)
                        print('下载完成！用时： %.2f秒，占用带宽：%.2fKBit/秒' % (total, speed))  # 输出下载用时时间
                        self.response.close()
                        fsize = os.path.getsize(self.name)
                        if fsize == self.size or self.size == 0:
                            self.status = 0
                            # 更新已经下载总量
                            GLOBAL_DOWN_SIZE = GLOBAL_DOWN_SIZE + self.size
                            # 更新下载成功列表
                            GLOBAL_DOWN_SUCCE_NUMBER += 1
                            print('文件下载成功！(%s)' % (self.name))
                        else:
                            print("文件下载失败因大小不正确(%d/%d)(%s)。" % (fsize, self.size, self.name))
                            GLOBAL_DOWN_ERROR_NUMBER += 1
                            self.status = -1
                            self.__write_err_file()
            except Exception as err:
                self.status = -1
                # 成功获取到网站内容处理
                if self.response.status_code == 404:
                    self.__write_404err_file()
                else:
                    self.__write_err_file()
                    write_except_file(err)
                    #sys.exit()

            print_down_info()
        except Exception as err:
            print(err)
            self.status = -1  # 发生异常
            print("\033[1;31m文件下载失败：" + self.url + "\033[0m")
            GLOBAL_DOWN_ERROR_NUMBER += 1
            self.__write_err_file()
            write_except_file(err)
            #sys.exit()


def get_down_objects():
    if os.path.exists(ROOT_OBJECT):
        print("\n对象文件已存在，可以直接读取到对象列表中。")
        f = open(ROOT_OBJECT, 'rb')
        global GLOBAL_DOWN_LIST
        GLOBAL_DOWN_LIST = pickle.load(f)
        f.close()
    else:
        print("\n对象文件不存在，直接读取Excel到对象列表中。")
        get_books_list()
    global GLOBAL_DOWN_NUMBER
    GLOBAL_DOWN_NUMBER = len(GLOBAL_DOWN_LIST)
    print("下载对象总数[%d]" % (GLOBAL_DOWN_NUMBER))


def updata_all_objects():
    global GLOBAL_DOWN_LIST
    global GLOBAL_DOWN_NUMBER
    Now_Down_Number = len(GLOBAL_DOWN_LIST)
    GLOBAL_DOWN_NUMBER = Now_Down_Number
    if (Now_Down_Number % 256 == 0):
        save_all_objects()


class OptGlobalObjList:
    def __init__(self):
        self._value_lock = threading.Lock()

    def download_file(self):
        i = 0
        j = 0
        global GLOBAL_DOWN_LIST
        global GLOBAL_DOWN_FILE_NUMBER
        global GLOBAL_NOW_DOWN_SIZE
        global GLOBAL_READY_EXIT_1
        global GLOBAL_READY_EXIT_2
        num = read_number_file(ROOT_NUMBER_SUCC_file)
        list_num = len(GLOBAL_DOWN_LIST)
        if num == 0:
            num = 1
        if num == list_num:
            num = 1
            print("下载从新开始！！！")

        i = num

        # for o in GLOBAL_DOWN_LIST:
        for o in GLOBAL_DOWN_LIST[num - 1:]:
            try:
                # 准备退出
                if GLOBAL_READY_EXIT_2 == 1:
                    print("当前下载<1>：退出")
                    os._exit(0)
                # 等待退出
                if GLOBAL_READY_EXIT_1 == 1:
                    print("当前下载<1>：等待<2>完成下载，退出")
                    time.sleep(1)
                    continue

                total_sec = time.time() - GLOBAL_START_TIME
                str_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
                print("当前状态<1>：下载/总数[%d/%d]，状态：已成功/第三方/已错误[%d/%d/%d]，时间：%s，总用时(%s)" % (
                i, GLOBAL_DOWN_NUMBER, GLOBAL_DOWN_SUCCE_NUMBER, GLOBAL_DOWN_3RDLIST_NUMBER, GLOBAL_DOWN_ERROR_NUMBER,
                str_time, size2Time(total_sec)))
                i = i + 1
                j = j + 1

                print("当前下载<1>：%s" % (o.name))
                o.down_file()

                write_number_file((i - 1), ROOT_NUMBER_SUCC_file)
                #print("当前列表序号<1>：%d，已下载个数<%d>" % (i - 1, j))
                # 检查返回值，成功从对象列表中删除。
                # if o.status == 0:
                #    with self._value_lock:
                #        GLOBAL_DOWN_LIST.remove(o)
            # finally :
            except Exception as e:
                print("当前列表序号<2>：发生异常")
                print(e.args)
                print(str(e))
                print(repr(e))
                time.sleep(1)
            finally:
                print("\033[1;45m 当前列表序号<1>：%d，已下载个数<%d>，序号<2>的状态[%d], 文件下载实时大小：[%s] \033[0m" % (i - 1, j, GLOBAL_READY_EXIT_2, size2human(GLOBAL_NOW_DOWN_SIZE)))
                if j > GLOBAL_DOWN_FILE_NUMBER:
                    GLOBAL_READY_EXIT_1 = 1

            # 删除100个之后，则保存一次对象文件
            # updata_all_objects()

    def download_file2(self):
        global GLOBAL_DOWN_LIST
        global GLOBAL_NOW_DOWN_SIZE
        global GLOBAL_DOWN_FILE_NUMBER
        global GLOBAL_READY_EXIT_1
        global GLOBAL_READY_EXIT_2

        i = 0
        j = 1
        num = read_number_file(ROOT_NUMBER_SUCC_file2)
        list_num = len(GLOBAL_DOWN_LIST)
        if num == 0:
            num = 1
        if num == list_num:
            num = 1
            print("下载从新开始！！！")

        i = num

        rev_list = []
        rev_list = list(reversed(GLOBAL_DOWN_LIST))
        for o in rev_list[num - 1:]:
            # for o in reversed(GLOBAL_DOWN_LIST):
            try:
                # 准备退出
                if GLOBAL_READY_EXIT_1 == 1:
                    print("当前下载<2>：退出")
                    os._exit(0)

                # 等待退出
                if GLOBAL_READY_EXIT_2 == 1:
                    print("当前下载<2>：等待<1>完成下载，退出")
                    time.sleep(1)
                    continue

                total_sec = time.time() - GLOBAL_START_TIME
                str_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
                print("当前状态<2>：下载/总数[%d/%d]，状态：已成功/第三方/已错误[%d/%d/%d]，时间：%s，总用时(%s)" %
                      (i, GLOBAL_DOWN_NUMBER, GLOBAL_DOWN_SUCCE_NUMBER, GLOBAL_DOWN_3RDLIST_NUMBER,
                       GLOBAL_DOWN_ERROR_NUMBER, str_time, size2Time(total_sec)))
                i = i + 1
                j = j + 1

                print("当前下载<2>：%s" % (o.name))

                o.down_file()
                write_number_file((i - 1), ROOT_NUMBER_SUCC_file2)
                #print("当前列表序号<2>：%d，已下载个数<%d>" % (i - 1, j))
                # 检查返回值，成功从对象列表中删除。
                # if o.status == 0:
                #    with self._value_lock:
                #        GLOBAL_DOWN_LIST.remove(o)
                # 删除100个之后，则保存一次对象文件
                # updata_all_objects()
            except Exception as e:
                print("当前列表序号<2>：发生异常")
                print(e.args)
                print(str(e))
                print(repr(e))
                time.sleep(1)
            finally:
                print("\033[1;45m 当前列表序号<2>：%d，已下载个数<%d>，序号<1>的状态[%d], 文件下载实时大小：[%s] \033[0m" % (i - 1, j, GLOBAL_READY_EXIT_1, size2human(GLOBAL_NOW_DOWN_SIZE)))
                if j > GLOBAL_DOWN_FILE_NUMBER:
                    GLOBAL_READY_EXIT_2 = 1


def get_down_url(url):
    url = url.replace("\\", "/")
    ret = ROOT_WEB + url
    return ret


def get_books_list():
    book = xlrd2.open_workbook(ROOT_EXCEL)
    sheet = book.sheet_by_name(ROOT_SHEET)

    for i in range(sheet.nrows):
        url_list = sheet.row_values(i)  # 简体文件名 | 繁体文件名 | 网站目录
        save_dir = ROOT_DIR + url_list[2]  # 保存路径
        # 特殊字符替换
        save_dir = save_dir.replace('?', '.')
        down_url = get_down_url(url_list[2])  # 下载URL
        # print("从Excel文件读取：%s" % (save_dir))
        file_save = C_DownFiles(save_dir, down_url)
        GLOBAL_DOWN_LIST.append(file_save)

    # 更新下载对象到本地文件中
    save_all_objects()


class SharedCounter:
    '''
    A counter object that can be shared by multiple threads.
    '''

    def __init__(self):
        self._value_lock = threading.Lock()

    def save_all_objects(self):
        with self._value_lock:
            global GLOBAL_DOWN_LIST
            f = open(ROOT_OBJECT, 'wb')
            pickle.dump(GLOBAL_DOWN_LIST, f)
            f.flush()
            f.close()


def save_all_objects():
    sc = SharedCounter()
    sc.save_all_objects()

    # global GLOBAL_DOWN_LIST
    # f = open(ROOT_OBJECT, 'wb')
    # pickle.dump(GLOBAL_DOWN_LIST, f)
    # f.flush()
    # f.close()


def init_log():
    # 初始化环境
    str_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    # 错误文件
    with open(ROOT_ERROR_file, 'a') as f:
        f.write(str_time + '\n')
    # 异常文件
    with open(ROOT_ERROR_EXECPT_file, 'a') as f:
        f.write(str_time + '\n')
    # 404文件
    with open(ROOT_404ERR_file, 'a') as f:
        f.write(str_time + '\n')
    # 非404文件
    with open(ROOT_NO404ERR_file, 'a') as f:
        f.write(str_time + '\n')


def get_all_threat():
    optObj = OptGlobalObjList()
    t = threading.Thread(target=optObj.download_file, name='LoopThread1')
    t.start()
    s = threading.Thread(target=optObj.download_file2, name='LoopThread2')
    s.start()

    t.join()
    s.join()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    GLOBAL_START_TIME = time.time()
    init_log()
    get_down_objects()
    get_all_threat()
    # download_file()
    # dw = FileSave("", "")
    # dw.get_error_list()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
