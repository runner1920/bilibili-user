# -*-coding:utf8-*-
from pyhive import hive


conn = hive.Connection(host='192.168.10.25', port=10000, database='db_test', username='root', password='root')
cursor = conn.cursor()
cursor.execute('show tables')
for result in cursor.fetchall():
    print(result)

import pip
print(pip.pep425tags.get_supported())