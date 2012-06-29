#coding:utf-8
#batch import csv to access

import csv
import pyodbc

# 定义数据库连接
cnxn = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\\Users\\Lee\\Desktop\\DistrictCode.mdb', autocommit=True)
cursor = cnxn.cursor()

# 指定csv文件列表
citylistfile = r'D:\Private\Py\pyodbc\batchAccess\citylist.csv'

# 创建表
tablename = 'DingCan_MatchOK'  #"订餐匹配结果"

if not cursor.tables(table = tablename).fetchone():
    sql_crtb = """create table [%s](
      ID1 INTEGER,
      OrgName varchar(140),
      prov varchar(20),
      city varchar(20),
      OrgAddr varchar(255),
      MatchName varchar(140),
      MatchAddress varchar(255),
      POIID varchar(20),
      Coordinate varchar(40),
      MatchType varchar(20)
      )""" % tablename

    cursor.execute(sql_crtb)
    cursor.commit()


# 顺序读取csv文件并导入到表中
city_files = csv.reader(open(citylistfile, 'rb'))

for city in city_files:
    print city[0],

    reader = csv.reader(open(city[0].decode('utf8'), 'rb'))
    reader.next()  # 跳过标题行

    for line in reader:
        sql_insert = "insert into [%s] values (%s)" % (tablename,
            ','.join(["'%s'"% i.decode('utf8') for i in line]))

        #print sql_insert
        cursor.execute(sql_insert)
    cursor.commit()

    print 'OK'
