# -*-coding:utf8-*-

import requests
import pymysql
import sys
import time
import re
import random
from imp import reload

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

def getsource():
    ua = random.choice(uas)
    head = {
        'User-Agent': ua
    }
    jscontent = requests \
        .session() \
        .get('https://tiyu.baidu.com/tokyoly/home/tab/%E5%A5%96%E7%89%8C%E6%A6%9C/from/pc', headers=head) \
        .text

    # print(jscontent)
    jscontent = jscontent.replace(" china", "")
    patter = re.compile('style="width:0.82rem;" data-a-7e0f0a6e>[\u4e00-\u9fa5]{1,10}')
    companys = patter.findall(jscontent)
    # print(companys)
    goldpatter = re.compile('class="item-gold" data-a-7e0f0a6e>[0-9]{1,2}')
    golds = goldpatter.findall(jscontent)
    # print(golds)
    silverpatter = re.compile('item-silver" data-a-7e0f0a6e>[0-9]{1,2}')
    silvers = silverpatter.findall(jscontent)
    copperpatter = re.compile('item-copper" data-a-7e0f0a6e>[0-9]{1,2}')
    coppers = copperpatter.findall(jscontent)
    addpatter = re.compile('item-all" data-a-7e0f0a6e>[0-9]{1,2}')
    alls = addpatter.findall(jscontent)
    # print(alls)
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(['国家名称', "金牌", "银牌", "铜牌", "全部"])
    # for companyName in companyList:
    #     ws.append([companyName])
    for index in range(len(companys)):
        company = companys[index]
        gold = golds[index]
        silver = silvers[index]
        copper = coppers[index]
        all = alls[index]
        ws.append([company[company.find(">")+1:], gold[gold.find(">")+1:], silver[silver.find(">")+1:]
                      , copper[copper.find(">")+1:], all[all.find(">")+1:]])
    wb.save('D:/奥运数据'+time.strftime("%Y-%m-%d", time.localtime()) +'.xlsx')
    print("输出完成")

getsource()
