# -*-coding:utf8-*-

import pymysql

companyList = []
conn = pymysql.connect(
    host='192.168.60.200', user='root', passwd='admin123', db='bilibili', charset='utf8')
cur = conn.cursor()
cur.execute('select DISTINCT company_name from company ORDER BY page_num')
results = cur.fetchall()
for row in results:
    companyList.append(row[0])
conn.close()

# 写入excel
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.append(['公司名称'])
for companyName in companyList:
    ws.append([companyName])
wb.save('D:/消防数据.xlsx')
