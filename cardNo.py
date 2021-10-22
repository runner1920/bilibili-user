# -*-coding:utf8-*-

import pymysql
import random


def getRandom():
    str = ""
    for i in range(6):
        ch = chr(random.randrange(ord('0'), ord('9') + 1))
        str += ch
    return str

conn = pymysql.connect(
    host='192.168.1.61', user='root', passwd='admin123', db='bilibili', charset='utf8')
cur = conn.cursor()

startNum = 100123
for pageNum in range(0, 2000):
    while (len(set(map(int, str(startNum))))!=4):
        startNum = startNum + 1
    cardNo = startNum
    cardPwd = getRandom()
    while (len(set(map(int, str(cardPwd))))!=4):
        cardPwd = getRandom()

    cur.execute('INSERT INTO t_card_info(card_no, card_pwd) \
                        VALUES ("%s","%s")'
                %
                ("6100630686"+str(cardNo), cardPwd))
    startNum = startNum + random.randint(10,20)

conn.commit()
