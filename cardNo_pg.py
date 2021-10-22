# -*-coding:utf8-*-

import psycopg2
import random
import pymysql


def getRandom():
    str2 = ""
    for i in range(6):
        ch = str(random.randint(0, 9))
        str2 += ch
    return str2

# conn = psycopg2.connect(database="python", user="postgres", password="postgres", host="localhost", port="5433")
conn = pymysql.connect(
    host='192.168.1.61', user='root', passwd='admin123', db='python', charset='utf8')
cur = conn.cursor()

startNum = 13973
for pageNum in range(0, 10000):
    startNum = startNum + 1
    cardNo = startNum
    cardPwd = getRandom()
    while (len(set(map(int, str(cardPwd))))!=4):
        cardPwd = getRandom()
    fullCard = "6100630686"+str(startNum)+str(random.randint(0, 9))
    # cur.execute('INSERT INTO t_card_info VALUES ('+fullCard+','+cardPwd+')')
    cur.execute('INSERT INTO t_card_info(card_no, card_pwd) \
                        VALUES ("%s","%s")'
                %
                (fullCard, cardPwd))

conn.commit()
