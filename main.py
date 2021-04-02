# -*- coding: utf-8

import os
import sys
import xlrd2
import time
import pickle
import requests
import urllib.request
from requests_html import HTML
from requests_html import HTMLSession

ROOT_SHEET = "藏书目录"
ROOT_WEB = 'https://drive.my-elibrary.com'

ROOT_DIR = 'Z:\\bacpup_other\\data'
ROOT_EXCEL = '1.xlsx'
ROOT_OBJECT = 'all.object'
ROOT_ERROR_OBJECT = 'err.object'
ROOT_SUCCESS_OBJECT = 'suc.object'
GLOBAL_DOWN_LIST = []
GLOBAL_DOWN_ERROR = []
GLOBAL_DOWN_SUCCE = []
GLOBAL_DOWN_SIZE = 0

class FileSave:
    # 传入文件名及URL都是已经处理过的，可以直接使用
    def __init__(self, file, url):
        self.name = file# 文件绝对路径
        self.dir = ""   # 保存绝对路径
        self.url = url  # URL绝对路径
        self.status = 0 # 状态 0(未下载) 1(成功下载)
        self.size = 0	# 文件大小
        self.status_code = 0

    def ready_dir(self):
        end_pos = self.name.rfind("\\")
        if end_pos != -1:
            self.dir = self.name[0:end_pos]
        else:
            self.dir = self.name

        if not os.path.isdir(self.dir):
            #print("创建目录：" + self.dir)
            os.makedirs(self.dir)

    def down_file(self):
        # 判断路径是否存在
        self.ready_dir()

        # 判断文件是否存在
        global GLOBAL_DOWN_SUCCE
        global GLOBAL_DOWN_ERROR
        fsize = 0
        if os.path.exists(self.name):
            fsize = os.path.getsize(self.name)
            print("文件已存在且文件的大小：%.2fKBit (%s)" % ((fsize/1024), self.name))

        start = time.time()  # 下载开始时间
        try:
            response = requests.get(self.url, stream=True)
            self.status_code = response.status_code
            size = 0  # 初始化已下载大小
            chunk_size = 1024  # 每次下载的数据大小
            self.size = content_size = int(response.headers['content-length'])  # 下载文件总大小
            if response.status_code == 200:  # 判断是否响应成功
                if fsize != 0 and fsize == self.size:
                    print("文件已成功下载，无需再下载。")
                    # 更新下载成功列表
                    GLOBAL_DOWN_SUCCE.append(self)
                    return
                print('开始下载，文件大小:[{size:.2f}] MB'.format(size = content_size / chunk_size / 1024))  # 开始下载，显示下载文件大小
                with open(self.name, 'wb') as file:  # 显示进度条
                    for data in response.iter_content(chunk_size=chunk_size):
                        file.write(data)
                        size += len(data)
                        now_t = time.time()
                        print('\r' + '[下载进度]：%s%.2f%%(实时速度：%.2fKB/秒)' % ('>' * int(size * 50 / content_size), float(size / content_size * 100), float(size/((now_t - start) * 1024))), end=' ')
                    file.flush()

            end = time.time()  # 下载结束时间
            total = end - start
            speed = content_size/(total * 1024)
            print('下载完成！用时： %.2f秒，占用带宽：%.2fKBit/秒' % (total, speed))  # 输出下载用时时间
            fsize = os.path.getsize(self.name)
            if fsize == self.size:
                self.status = 1
                # 更新下载总量
                global  GLOBAL_DOWN_SIZE
                GLOBAL_DOWN_SIZE = GLOBAL_DOWN_SIZE + content_size
                # 更新下载成功列表
                GLOBAL_DOWN_SUCCE.append(self)
            else:
                print("下载完成但文件大小不正确(%d/%d)。"%(fsize, self.size))
                GLOBAL_DOWN_ERROR.append(self)

            str_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
            print("下载总数据量：%.2f(MBit) 当前时间：%s" % (GLOBAL_DOWN_SIZE/(1024 * 1024), str_time))
        except:
            self.status = 0
            print("文件下载失败：" + self.url + "，错误码：" + str(response.status_code))
            GLOBAL_DOWN_ERROR.append(self)

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
    global GLOBAL_DOWN_LIST
    number = len(GLOBAL_DOWN_LIST)
    i = 1
    for o in GLOBAL_DOWN_LIST:
        print("开始处理：下载/总数[%d/%d]，当前成功/错误[%d/%d]" % (i, number, len(GLOBAL_DOWN_SUCCE), len(GLOBAL_DOWN_ERROR)))
        i = i + 1
        o.down_file()
        save_error_obj()

def tosize(size):
    def strofsize(integer, remainder, level):
        if integer >= 1024:
            remainder = integer % 1024
            integer //= 1024
            level += 1
            return strofsize(integer, remainder, level)
        else:
            return integer, remainder, level

    units = ['B', 'KB', 'MB', 'GB', 'TB', 'PB']
    integer, remainder, level = strofsize(size, 0, 0)
    if level+1 > len(units):
        level = -1
    return ( '{}.{:>03d} {}'.format(integer, remainder, units[level]) )

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

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    get_down_object()
    download_file()
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
