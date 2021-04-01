# -*- coding: utf-8
# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
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
        self.size = 0
        self.status_code = 0

    def ready_dir(self):
        end_pos = self.name.rfind("\\")
        if end_pos != -1:
            self.dir = self.name[0:end_pos]
        else:
            self.dir = self.name

        if not os.path.isdir(self.dir):
            print("创建目录：" + self.dir)
            os.makedirs(self.dir)
        else:
            print("目录存在：" + self.dir)

    def down_file(self):
        # 判断路径是否存在
        self.ready_dir()
        # 判断文件正确与否
        if os.path.exists(self.name):
            fsize = os.path.getsize(self.name)

            if self.size == fsize:
                print("文件已存在且文件大小一致，忽略：%s，文件的大小：%.2fKBit" % (self.name, (fsize / 1024)))
            else:
                os.remove(self.name)
                print("文件已存在但文件大小不一致，删除：%s，文件的大小：%.2f/%.2fKBit" % (self.name, (fsize / 1024), (self.size/1024)))
        # 判断文件是否存在
        if os.path.exists(self.name):
            fsize = os.path.getsize(self.name)
            print("文件已存在，忽略：%s，文件的大小：%.2fKBit" % ( self.name ,(fsize/1024)))
        else:
            start = time.time()  # 下载开始时间
            try:
                response = requests.get(self.url, stream=True)
                self.status_code = response.status_code
                size = 0  # 初始化已下载大小
                chunk_size = 1024  # 每次下载的数据大小
                content_size = int(response.headers['content-length'])  # 下载文件总大小
                self.size = content_size
                if response.status_code == 200:  # 判断是否响应成功
                    print('开始下载，文件大小:[{size:.2f}] MB'.format(size = content_size / chunk_size / 1024))  # 开始下载，显示下载文件大小
                    with open(self.name, 'wb') as file:  # 显示进度条
                        for data in response.iter_content(chunk_size=chunk_size):
                            file.write(data)
                            size += len(data)
                            print('\r' + '[下载进度]：%s%.2f%%' % ('>' * int(size * 50 / content_size), float(size / content_size * 100)), end=' ')
                        file.flush()

                end = time.time()  # 下载结束时间
                total = end - start
                speed = content_size/(total * 1024)
                print('下载完成！用时： %.2f秒，占用带宽：%.2fKBit/秒' % (total, speed))  # 输出下载用时时间
                fsize = os.path.getsize(self.name)
                if fsize == self.size:
                    self.status = 1
                    global  GLOBAL_DOWN_SIZE
                    GLOBAL_DOWN_SIZE = GLOBAL_DOWN_SIZE + content_size
                    global GLOBAL_DOWN_SUCCE
                    GLOBAL_DOWN_SUCCE.append(self)
                else:
                    print("下载完成，文件大小不正确。")
                str_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
                print("下载总数据量：%.2f(MBit) 当前时间：%s" % (GLOBAL_DOWN_SIZE/(1024 * 1024), str_time))
            except:
                self.status = 0
                print("文件下载失败：" + self.url + "，错误码：" + str(response.status_code))
                global GLOBAL_DOWN_ERROR
                GLOBAL_DOWN_ERROR.append(self)

def get_down_object():
    if os.path.exists(ROOT_OBJECT):
        print("对象文件已存在，可以直接读取到对象列表中。")
        f = open(ROOT_OBJECT, 'rb')
        global GLOBAL_DOWN_LIST
        GLOBAL_DOWN_LIST = pickle.load(f)
    else:
        get_books_list()

def download_file():
    global GLOBAL_DOWN_LIST
    number = len(GLOBAL_DOWN_LIST)
    i = 1
    for o in GLOBAL_DOWN_LIST:
        print("开始处理[%d/%d]" % (i, number))
        i = i + 1
        if o.status == 1:
            # 判断文件大小
            if os.path.exists(o.name):
                fsize = os.path.getsize(o.name)
                if fsize == o.size:
                    print("文件已成功下载：" + o.name)
                else:
                    print("文件下载已失败：" + o.name)
                    o.status = 0
                    o.down_file()
        else:
            print("文件将下载：" + o.name)
            o.down_file()

def get_books_list():
    book = xlrd2.open_workbook(ROOT_EXCEL)
    sheet = book.sheet_by_name(ROOT_SHEET)

    for i in range(sheet.nrows):
        url_list = sheet.row_values(i)      # 简体文件名 | 繁体文件名 | 网站目录
        save_dir = ROOT_DIR + url_list[2]   # 保存路径
        # 特殊字符替换
        save_dir = save_dir.replace('?', '.')
        down_url = get_down_url(url_list[2])# 下载URL
        #print("下载链接：" + down_url)
        #print("保存文件：" + save_dir)

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

    global GLOBAL_DOWN_SUCCE
    f = open(ROOT_SUCCESS_OBJECT, 'wb')
    pickle.dump(GLOBAL_DOWN_SUCCE, f)

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
