# -*- coding: utf-8 -*-
#                       _oo0oo_
#                      o8888888o
#                      88" . "88
#                      (| -_- |)
#                      0\  =  /0
#                    ___/`---'\___
#                  .' \\|     |// '.
#                 / \\|||  :  |||// \
#                / _||||| -:- |||||- \
#               |   | \\\  -  /// |   |
#               | \_|  ''\---/''  |_/ |
#               \  .-\__  '-'  ___/-. /
#             ___'. .'  /--.--\  `. .'___
#          ."" '<  `.___\_<|>_/___.' >' "".
#         | | :  `- \`.;`\ _ /`;.`/ - ` : | |
#         \  \ `_.   \_ __\ /__ _/   .-` /  /
#     =====`-.____`.___ \_____/___.-`___.-'=====
#                       `=---='
#
#
#     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
#               佛祖保佑         永無BUG
#####################################################################
#########################   天祐   ##############################
#####################################################################
import csv
import pyodbc
co = False
count = 1
conn = pyodbc.connect(DRIVER='{SQL Server Native Client 10.0}',SERVER='JHOUGO-PC',DATABASE='Top',Trusted_Connection='yes')
cur = conn.cursor()
with open(u'D:\\共用\\小的監理所資料\\people\\typolice\\20150812 dbo.police.csv', 'rb') as csvfile:#'D:\\共用\\小的監理所資料\\people\\typolice\\20150812 dbo.police.csv'
    spamreader = csv.reader(csvfile, dialect='excel')
    for row in spamreader:
        if co :
	        cur.execute("INSERT INTO dbo.People_6 (vil_dt,vil_time,d_id_no,d_id_no1,owner,o_addr,n_ownerno1,n_owner,n_addr,driver,driver1,d_addr)  VALUES(?,?,?,?,?,?,?,?,?,?,?,?) ",row[0].decode('UTF-8'),row[1].decode('UTF-8'),row[2].decode('UTF-8'),row[3].decode('UTF-8'),row[4].decode('UTF-8'),row[5].decode('UTF-8'),row[6].decode('UTF-8'),row[7].decode('UTF-8'),row[8].decode('UTF-8'),row[9].decode('UTF-8'),row[10].decode('UTF-8'),row[11].decode('UTF-8'))
	        conn.commit()
        #print count 
        count +=1
	co = True
print count
cur.close()
conn.close() 
