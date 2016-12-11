# -*- coding: utf-8 -*-
"""
Created on Wed Nov 09 00:21:00 2016

@author: IORI
"""

from __future__ import division
import math
import datetime
import calendar
import MySQLdb
import MySQLdb.cursors
import tkMessageBox
import tkSimpleDialog
import functions as fn
import re
def conn_mysql():
    conn=MySQLdb.Connect(host='10.100.7.52',port=3306,user='fxgl',passwd='fxgl',db='ht_risk', charset='utf8',cursorclass = MySQLdb.cursors.DictCursor)
    return conn
def monitor_cor_position(invclass,codeclass):
    '''公司持仓监控计算'''
    conn=conn_mysql()
    cursor=conn.cursor()
    sql='select * from corpos_monitor_coefficient'
    cursor.execute(sql)
    rs=cursor.fetchall()
    vdata={}
    for x in rs:
        vdata[x['vari']]=x        
    vari_longpos,vari_shortpos={},{}
    for x in invclass:
        for pos in invclass[x].Position:
            if pos['house'] in ['SHFE','CFFEX']:
                if vari_longpos.has_key(pos['code']):
                    vari_longpos[pos['code']]+=pos['longnums']
                    vari_shortpos[pos['code']]+=pos['shortnums']                    
                else:
                    vari_longpos[pos['code']],vari_shortpos[pos['code']]=pos['longnums'],pos['shortnums']                    
    market_pos={}
    for x in codeclass:
        if codeclass[x].House in ['SHF','CFE']:
            code=x#codeclass[x].Vari
            if market_pos.has_key(code):
                market_pos[code]+=codeclass[x].Inf['open_interest']
            else:
                market_pos[code]=codeclass[x].Inf['open_interest']
    outdata=[]
    colname=[u'交易日',u'持仓合约',u'会员持仓限额',u'会员多头持仓',u'会员空头持仓',u'多头持仓/限额',u'空头持仓/限额',u'多头风险等级',u'空头风险等级']
    for x in vari_longpos:
        data=[]
        if market_pos[x]>=vdata[codeclass[x].Vari]['limit']:
            vl=math.floor(market_pos[x]/2*0.25*(1+vdata[codeclass[x].Vari]['bus_cof']+vdata[codeclass[x].Vari]['cre_cof']))
            if x in ['IF','IH','IC','TF','T']:
                vl=math.floor(market_pos[x]*0.25*(1+vdata[codeclass[x].Vari]['bus_cof']+vdata[codeclass[x].Vari]['cre_cof']))
            r1,r2=vari_longpos[x]/vl,vari_shortpos[x]/vl
            now=datetime.datetime.now().strftime('%Y-%m-%d')
            data=[now,x,vl,vari_longpos[x],vari_shortpos[x],round(r1,2),round(r2,2)]
            for r in [r1,r2]:
                if r>0.8:
                    data.append(u'极度风险')
                elif r>0.6:
                    data.append(u'高度风险')
                elif r>0.4:
                    data.append(u'中度风险')
                elif r>0.2:
                    data.append(u'轻度风险')
                else:
                    data.append(u'没有风险')
        if data:
            outdata.append(data)
    return colname,outdata
def update_posmonitor(outdata):
    '''公司持仓监控操作记录至数据库'''
    name=tkSimpleDialog.askstring(u'华泰期货',u'请输入监控员姓名',initialvalue ='')
    if name==None:
        return
    rmarks=tkSimpleDialog.askstring(u'华泰期货',u'请输入备注信息',initialvalue ='')
    if rmarks==None:
        rmarks=''
    conn=conn_mysql()
    cursor=conn.cursor()
    print u'数据库连接成功'
    hr=0
    er=0
    error=False
    try:
        for x in outdata:
            if x[-1]==u'高度风险' or x[-2]==u'高度风险':
                hr+=1
            if x[-1]==u'极度风险' or x[-2]==u'极度风险':
                er+=1
            tpstr=list_tuple_to_sqlstr(x)
            sql='INSERT INTO corpos_monitor_detail VALUES%s'%tpstr
            cursor.execute(sql)
    except Exception as e:
        error=True
        print 1      
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'数据库更新失败')
        conn.rollback()
        cursor.close()
        conn.close()
        raise e
    if not error:
        totalinf=[outdata[0][0],hr,er,name,rmarks]
        try:
            tpstr=list_tuple_to_sqlstr(totalinf)
            sql='INSERT INTO corpos_monitor_inf VALUES%s'%tpstr
            print sql
            cursor.execute(sql) 
        except Exception as e:
            error=True
            print 2
            tkMessageBox.showinfo(title=u'温馨提醒',message=u'数据库更新失败')
            conn.rollback()
            cursor.close()
            conn.close()
            raise e
    if error:
        return
    tkMessageBox.showinfo(title=u'温馨提醒',message=u'数据库更新成功')
    conn.commit()
    cursor.close()
    conn.close()
    print u'关闭数据库连接'
    return        
def list_tuple_to_sqlstr(tp):
    '''将list或tuple转化成str格式的tuple以插入数据库'''
    a='('
    for x in tp:
        if x is None:
            a+='null,'
        elif type(x)==float or type(x)==int:
            a+='%s,' %x
        else:
            a+='\'%s\',' %x           
    a=a[:-1]+')'
    return a   
def out_excel(colname,outdata,filename,shname='Sheet1'):
    import xlwt
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet(shname)     
    nrow=len(outdata)
    ncol=len(colname)    
    for j in range(ncol):
        worksheet.write(0, j, label =fn.english_to_ch(colname[j]))
    for i in range(nrow):
        for j in range(ncol):
            worksheet.write(i+1, j, label =outdata[i][j])
    workbook.save(filename)
def colname_to_sql(tbname,colname_en,colname_ch):
    '''把表英文列名对应的中文名写入至数据库'''
    conn=conn_mysql()
    cursor=conn.cursor()
    if len(colname_en)<>len(colname_ch):
        print '表里列名数据库更新失败，数据长度不一样'
        return
    for i in range(len(colname_en)):
        sql='insert into table_column_name set table_name=\'%s\',col_name_en=\'%s\',col_name_ch=\'%s\''%(tbname,colname_en[i],colname_ch[i])       
        try:
            cursor.execute(sql)
        except Exception as e:
            print sql
            conn.rollback()
            cursor.close()
            conn.close()           
    conn.commit()    
    cursor.close()
    conn.close()
rcolname=['invid','inv_name','invdepartment','spmark','riskdegree','cor_riskdegree','leftcapital','cor_leftcapital','seat','margin','capital','lastriskstate','maxcode','posstrut','forcedbound','phone']        
def get_monitor_pos():
    '''获取营业部持仓监控数据'''
    res={u'质押配比监控':'mon_ple_client',u'临近交割月监控':'delivery_mon_pos',u'客户超仓监控':'client_pos_mon'
    ,u'重大持仓监控':'major_pos_monitor',u'不活跃持仓监控':'unactive_pos_monitor'}
    tbname={}
    for x in res:
        tbname[res[x]]=x
    now=datetime.datetime.now().strftime('%Y-%m-%d')
    conn=MySQLdb.Connect(host='10.100.7.52',port=3306,user='fxgl',passwd='fxgl',db='ht_risk', charset='utf8')
    cursor=conn.cursor()
    res={}
    for x in tbname:
        sql='desc %s'%x
        cursor.execute(sql)
        rs=cursor.fetchall()
        colname=[]
        for y in rs:
            colname.append(y[0])
        sql='select * from %s where tr_date=\'%s\'' %(x,now)
        cursor.execute(sql)
        rs=cursor.fetchall()
        if rs:
            colname_ch=get_colname_ch_from_sql({x:colname})
            res[tbname[x]]={}
            idx=colname.index('invdepartment')
            for y in rs:
                if not res[tbname[x]].has_key(y[idx]):
                    res[tbname[x]][y[idx]]=[colname_ch[x],[]]
                res[tbname[x]][y[idx]][1].append(y)
    if not res:
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'没有找到当天持仓监控数据')
    return res
def get_colname_ch_from_sql(tbdic):
    '''给定英文名字获取中文名字,输入变量为字典'''
    conn=conn_mysql()
    cursor=conn.cursor()
    res={}
    for tbname in tbdic:
        sql='select * from table_column_name where table_name=\'%s\''%tbname
        cursor.execute(sql)
        rs=cursor.fetchall()
        temp={}
        for x in rs:
            temp[x['col_name_en']]=x['col_name_ch']
        cl=[]
        for x in tbdic[tbname]:
            cl.append(temp[x])
        res[tbname]=cl
    cursor.close()
    conn.close()
    return res    
    
        