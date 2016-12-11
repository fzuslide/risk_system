# -*- coding: utf-8 -*-
"""
Created on Sun Dec 04 14:29:58 2016

@author: IORI
"""

import MySQLdb
import MySQLdb.cursors

class mysql_conn:
    def __init__(self,isdic=True):
        if isdic:
            conn=MySQLdb.Connect(host='10.100.7.52',port=3306,user='fxgl',passwd='fxgl',db='ht_risk', charset='utf8',cursorclass = MySQLdb.cursors.DictCursor)
        else:
            conn=MySQLdb.Connect(host='10.100.7.52',port=3306,user='fxgl',passwd='fxgl',db='ht_risk', charset='utf8')
        cursor=conn.cursor()
        self.conn=conn
        self.cursor=cursor
        self.isdic=isdic
    def close(self):
        self.cursor.close()
        self.conn.close()
    def start(self):
        if self.isdic:
            conn=MySQLdb.Connect(host='10.100.7.52',port=3306,user='fxgl',passwd='fxgl',db='ht_risk', charset='utf8',cursorclass = MySQLdb.cursors.DictCursor)
        else:
            conn=MySQLdb.Connect(host='10.100.7.52',port=3306,user='fxgl',passwd='fxgl',db='ht_risk', charset='utf8')
        cursor=conn.cursor()
        self.conn=conn
        self.cursor=cursor
    def commit(self):
        self.conn.commit()
    def fetchall(self):
        return self.cursor.fetchall()      
    def fetchone(self):
        return self.cursor.fetchone()
    def execute(self,sql):
        try:
            self.cursor.execute(sql)
        except Exception as e:
            print sql
            self.conn.rollback()
            self.close()
            raise e
    def insert_table(self,tbname,colname,data):   
        #tbname表名,colname英文列名，list，data数据，list
        conn,cursor=self.conn,self.cursor
        sql='desc %s' %tbname
        cursor.execute(sql)
        rs=cursor.fetchall()
        col_type={}
        for x in rs:
            col_type[x['Field']]=x['Type'].split('(')[0]
        n=len(colname)
        for x in data:
            sql='insert into %s set ' %tbname
            for i in range(n):
                if col_type[colname[i]] in ['float','int','double']:            
                    sql+='%s=%s,' %(colname[i],x[i])
                else:
                    sql+='%s=\'%s\',' %(colname[i],x[i])
            sql=sql[:-1] 
            self.execute(sql)
        print u'数据库插入成功',tbname