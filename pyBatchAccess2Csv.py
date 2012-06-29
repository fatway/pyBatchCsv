#coding:utf-8
#batch export csv from access database

#import sys, csv, ConfigParser
import csv
import pyodbc

#pyodbc不支持数据库中文路径？
##cnnstr = """'DRIVER={Microsoft Access Driver (*.mdb)};
##DBQ=%s',
##autocommit=True
##""" % u'E:\\上海号百匹配0626\\HaobaiData0626.mdb'


# 定义数据库连接
cnxn = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb)};DBQ=E:\\HaobaiData0626.mdb',autocommit=True)

# 指定导出csv路径
out_csv_path = u'E:\\上海号百匹配0626\\分城市导出\\'
has_header = True


# 根据城市列表分城市读取table
citys = [u"北京市", u"长沙市", u"常州市", u"成都市", u"东莞市", u"佛山市", u"福州市", u"广州市", u"杭州市", u"济南市", u"南京市", u"宁波市", u"青岛市", u"汕头市", u"上海市", u"深圳市", u"苏州市", u"天津市", u"温州市", u"无锡市", u"武汉市", u"西安市", u"厦门市", u"扬州市", u"重庆市", u"珠海市"]

for city in citys:
    cursor = cnxn.cursor()
    #tablename = u'订餐-' + city
    tablename = u'对标-' + city

    # 指定输出的csv文件
    wtr = csv.writer(open(out_csv_path + tablename + '.csv', 'wb'))

    # 文件标题行
    if has_header:
        #wtr.writerow(cursor.columns(table=tablename).column_name)  #err
        wtr.writerow(['SERIAL','CHINAME','PROVINCE','CITY','ADDRESS'])
        
    # 按行写内容
    #cursor.execute(u"select * from [订餐数据格式OK] where [市] = ?", city)
    cursor.execute(u"select * from [对标数据格式OK] where [市] = ?", city)
    for row in cursor.fetchall():
        #wtr.writerow([s.encode('utf8') for s in row])
        wtr.writerow([row[0], row[1].encode('utf8'), row[2].encode('utf8'), row[3].encode('utf8'), row[4].encode('utf8')])

