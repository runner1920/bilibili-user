# -*-coding:utf8-*-

import requests
import json
import random
import pymysql
import sys
import datetime
import time
import re
import random
import openpyxl
from imp import reload
from multiprocessing.dummy import Pool as ThreadPool


def datetime_to_timestamp_in_milliseconds(d):
    def current_milli_time(): return int(round(time.time() * 1000))

    return current_milli_time()


reload(sys)


def LoadUserAgents(uafile):
    uas = []
    with open(uafile, 'rb') as uaf:
        for ua in uaf.readlines():
            if ua:
                uas.append(ua.strip()[:-1])
    random.shuffle(uas)
    return uas


uas = LoadUserAgents("user_agents.txt")
head = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36',
    'X-Requested-With': 'XMLHttpRequest',
    'Origin': 'http://www.cccf.com.cn',
    'Host': 'www.cccf.com.cn',
    'Referer': 'http://www.cccf.com.cn/certSearch/search',
    'AlexaToolbar-ALX_NS_PH': 'AlexaToolbar/alx-4.0',
    'Accept-Language': 'zh-CN,zh;q=0.8,en;q=0.6,ja;q=0.4',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
}

# Please replace your own proxies.
proxies = {
    'http': 'http://120.26.110.59:8080',
    'http': 'http://120.52.32.46:80',
    'http': 'http://218.85.133.62:80',
}
time1 = time.time()

urls = []
maxPage = 2100

companyList = set()


def getsource(page):
    payload = {
        'page': page
    }
    ua = random.choice(uas)
    head = {
        'User-Agent': ua
    }
    jscontent = requests \
        .session() \
        .post('http://www.cccf.com.cn/certSearch/search', headers=head, data=payload) \
        .text

    if '认证委托人' in jscontent:
        pass
        patter = re.compile('认证委托人 ：[\u4e00-\u9fa5]{5,30}')
        companys = patter.findall(jscontent)
        print(companys, page)
        for companyName in companys:
            companyList.add(companyName.replace('认证委托人 ：', ''))
    else:
        print('数据错误')
        return


if __name__ == "__main__":
    for pageNum in range(7, 17):
        t = random.randint(3, 5)
        print('休眠时间', str(t) + '秒', '数据页数', pageNum)
        time.sleep(t)
        getsource(pageNum)
        # 每10页写入excel
        if pageNum % 3 == 0:
            # 写入excel
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.append(['公司名称'])
            for companyName in companyList:
                ws.append([companyName])
            wb.save('D:/公司数据第' + str(pageNum) + '条.xlsx')
            companyList = set()
