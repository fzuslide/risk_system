# -*- coding: utf-8 -*-
'''
盘后风险试算文件,压力测试文件
'''
from __future__ import division
import os
import os.path
import csv
import sys
import copy
import operator
import tkSimpleDialog
import math
import xlrd
import stresstest as stress
def return_ndays(r,n,t):
    """
       以各品种制定的收益r，最接近的行情天数n,已经具有的历史行情数据t作为输入参数，
       返回每个品种对应的n天最接近指定收益率r的行情数据
    """
    rootdir = r"\\10.100.6.20\fkfile\python_risksystem\data\vdata"
    csv_name = [rootdir + '\\' + filepath for filepath in os.listdir(rootdir)]
    file_pathname = [filepath for filepath in os.listdir(rootdir)]
    my_result = {} 
    # 保留原始输入的参数值,由于t跟n都为基本类型int
    t_temp = t 
    n_temp = n
    for i in range(len(csv_name)):
        t = t_temp
        n = n_temp
        pingzhong_name = file_pathname[i][0:file_pathname[i].find('_')] 
        if pingzhong_name not in r.keys():
            continue
        f = open(csv_name[i],"rb") 
        f1 = open(csv_name[i],"rb") 
        future_code = []
        trade_date = []
        pre_return = []
        max_bodong = []
        next_return_jiesuan = []
        next_price_kaipan = []
        next_return = []
        trade_date1 = []
        abs_chaju = []
        my_dict = {}
        my_new_dict = {}
        contentList = [row for row in csv.reader(f) if row]
        if (len(contentList)-1) < t:
            t = len(contentList)-1
        if n > (len(contentList)-1):
            n = len(contentList)-1
        if n > t:
            n = t
        contentList = contentList[-t:] #改变 -（t+1） 250
         
        contentList1 = [row for row in csv.reader(f1) if row]
        contentList1 = contentList1[-t:] #改变 -（t+1） 250
        for g, row in enumerate(contentList):
        # 记录下交易日期
            trade_date.append(row[1])
        # 当天收益率，计算差距
            abs_chaju.append(abs(float(row[2])-r[pingzhong_name])) # r -0.00005
        # 字典查询
            dicd = {trade_date[-1]:abs_chaju[-1]}
            my_dict.update(dicd)
        my_new_dict = sorted(my_dict.iteritems(), key=lambda my_dict:my_dict[1])
        my_new_dict = my_new_dict[:n] # n 20
        for j,row in enumerate(contentList1):
            my_new_len = len(my_new_dict)
            for k in range(my_new_len):
                if row[1] == my_new_dict[k][0]:
                     future_code.append(row[0])
                     trade_date1.append(row[1])
                     pre_return.append(row[2])
                     max_bodong.append(row[3])
                     next_return_jiesuan.append(row[4])
                     next_return.append(row[5])
                     next_price_kaipan.append(row[6])
        for k in range(len(future_code)):
            if k == 0:
                temp={}
                temp[pingzhong_name]= [[future_code[0],trade_date1[0],\
                                          pre_return[0],max_bodong[0],next_return_jiesuan[0],next_return[0],next_price_kaipan[0]]]
                my_result.update(temp)
            else:
                my_result[pingzhong_name].append([future_code[k],trade_date1[k],\
                                          pre_return[k],max_bodong[k],next_return_jiesuan[k],next_return[k],next_price_kaipan[k]])                            
        f1.close()
        f.close()
    return my_result   
def update():
    rootdir = r"\\10.100.6.20\fkfile\VariData\MainCode"
    resultdir = r"\\10.100.6.20\fkfile\python_risksystem\data\vdata"
    csv_name = []
    
    future_code = []
    trade_date = []
    pre_return = []
    max_bodong = []
    next_return_jiesuan = []
    next_return = []
    next_price_clo = 0.
    next_price_jiesuan = 0.
    
    csv_name = [rootdir + '\\' + filepath for filepath in os.listdir(rootdir)]
    #print csv_name
    for i in range(len(csv_name)):
    
        g = open(resultdir + "\\" + csv_name[i].split('\\')[-1] ,'wb')
        write = csv.writer(g)
        f = open(csv_name[i],"rb")
    
        future_code = []
        trade_date = []
        pre_return = []
        max_bodong = []
        next_return_jiesuan = []
        next_return = []
        next_price_clo = []
        next_price_jiesuan = []
        next_price_kaipan = []
        
        contentList = [row for row in csv.reader(f) if row]
        contentList = contentList[1:]
    
        for i, row in enumerate(contentList):
        # 记录下期货代码
            future_code.append(row[1])
            #print '代码长度为 %s' % len(future_code)
        # 记录下交易日期
            trade_date.append(row[2])
        # 当天收益率，以昨天为标准计算
            pre_return.append((float(row[6])-float(row[15]))/float(row[15]))
            if i+1 < len(contentList):
                a1 = (float(contentList[i+1][3])-float(contentList[i+1][4]))/float(row[7])
                a2 = (float(contentList[i+1][3])-float(row[6]))/float(row[7])
                a3 = (float(row[6])-float(contentList[i+1][4]))/float(row[7])
                max_bodong.append(max(a1,a2,a3))
                next_price_clo.append(float(contentList[i+1][6]))
                next_price_jiesuan.append(float(contentList[i+1][7]))
                next_return_jiesuan.append((next_price_jiesuan[len(next_price_clo)-1]-float(row[7]))/float(row[7]))
                next_return.append((next_price_clo[len(next_price_clo)-1]-float(row[6]))/float(row[6]))
                next_price_kaipan.append((float(contentList[i+1][5])-float(row[6]))/float(row[6]))
                
            else:
                max_bodong.append('')
                next_price_clo.append('')
                next_price_jiesuan.append('')
                next_return_jiesuan.append('')
                next_return.append('')
                next_price_kaipan.append('')
    
        write.writerow(["ext_code","date","pre_return","max_bodong","next_return_jiesuan","next_return","next_price_kaipan"])
        len_total = len(future_code)
        for j in range(len(future_code)):
            if max_bodong[j] !='':
                write.writerow([future_code[j],trade_date[j],str(pre_return[j]),str(max_bodong[j]),str(next_return_jiesuan[j]),str(next_return[j]),str(next_price_kaipan[j])])
        g.close()
def pricevar(variclass):
    '''行情var'''
    sdate=tkSimpleDialog.askstring(u'华泰期货',u'历史数据从哪天开始？',initialvalue ='2015-09-01')
    delta=tkSimpleDialog.askinteger(u'华泰期货',u'连续几天行情？',initialvalue =6)
    #variclass=data['variclass']
    data2,allpch,adt=stress.getfu(sdate,delta,variclass)
    td={}
    r=[0.9,0.95,0.99]
    colname=[u'涨90%',u'涨95%',u'涨99%',u'跌90%',u'跌95%',u'跌99%']
    for i in range(len(colname)):
        colname[i]=str(delta)+u'天连续行情'+colname[i]
    colname=[u'品种']+colname
    for x in variclass:
        td[x]=[]
        ylist=[]        
        for y in allpch:
            ylist.append(y[x])
        templist=copy.copy(ylist)
        ylist.sort(reverse=False)#从小到大
        for y in r:
            k=int(math.ceil(y*len(ylist)))
            v=round(ylist[k],4)
            td[x].append(v)
        ylist.sort(reverse=True)#从大到小
        for y in r:
            k=int(math.ceil(y*len(ylist)))
            v=round(ylist[k],4)
            td[x].append(v)
    return td,colname
def pricelimt(td,colname,variclass,fname):
    #dp=u'\\\\10.100.6.20\\fkfile\\IORI-陈志荣\\三表\\2016年国庆调保.xls'
    book = xlrd.open_workbook(fname)
    sh=book.sheets()[0]
    n=sh.nrows   
    colname+=[u'涨1个板',u'涨1.5个板',u'涨2个板',u'跌1个板',u'跌1.5个板',u'跌2个板']
    for i in range(4,n-1):
        vari=sh.cell_value(i,1).upper()
        l1=float(sh.cell_value(i,10))
        ul2=(1+float(sh.cell_value(i,10)))*(1+0.03+float(sh.cell_value(i,10)))-1
        dl2=(1-float(sh.cell_value(i,10)))*(1-0.03-float(sh.cell_value(i,10)))-1
        if variclass[vari].House=='CZC':
            ul2=(1+float(sh.cell_value(i,10)))*(1+float(sh.cell_value(i,10)))-1
            dl2=(1-float(sh.cell_value(i,10)))*(1-float(sh.cell_value(i,10)))-1            
        td[vari].append(l1)
        td[vari].append(1.5*l1)
        td[vari].append(ul2)
        td[vari].append(-l1)
        td[vari].append(-l1*1.5)
        td[vari].append(dl2)
    return td,colname        
              
#update()    