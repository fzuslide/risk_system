# -*- coding: utf-8 -*-
"""
Created on Tue Jun 14 11:11:55 2016

@author: Administrator
"""
from __future__ import division
import Tkinter as tk
import xlrd
import copy
import codecs
import re
import os
import datetime
import math 
import pickle
import tkMessageBox
import tkSimpleDialog
import functions_2 as fn2
from operator import itemgetter, attrgetter
from classes import *
from cal_margin import *
vacations=['20160101','20160208','20160209','20160210','20160211','20160212','20160404','20160502'
,'20160609','20160610','20160915','20160916','20161003','20161004','20161005','20161006','20161007']
def read_baicsh(gui='',pak=0):
    ''' 函数说明：读取能用得上得表'''
    rule=re.compile(r'[^a-zA-z]')
    ntime=datetime.datetime.now()-datetime.timedelta(days=6)
    tt=[]
    dt=[]
    dp=u'\\\\10.100.6.20\\fkfile\\开仓保证金量\\new保证金率.xls'
    #dp=u'\\\\10.100.6.20\\fkfile\\IORI-陈志荣\\三表\\new保证金率.xls'
    book = xlrd.open_workbook(dp)
    sh=book.sheets()[0]
    n=sh.nrows     
    colname=sh.row_values(1)
    code_house={}
    codeclass={}
    colk={'Mrate':colname.index(u'交易所保证金率'),'Limt':colname.index(u'涨跌停板幅度%'),'Units':colname.index(u'合约数量乘数')
    ,'Minprice':colname.index(u'最小变动价位'),'house':colname.index(u'交易所')}
    isdelta=False
    if u'保证金率调整' in colname:
        colk['delta_rate']=colname.index(u'保证金率调整')
        isdelta=True
    for i in range(2,n):
        code=sh.cell_value(i,colname.index(u'合约')).upper()
        if len(code)<=6:
            codeclass[code]=FutureClass(rule.sub('',code),code)  
            codeclass[code].Inf['Mrate']=sh.cell_value(i,colk['Mrate'])
            codeclass[code].Inf['Limt']=sh.cell_value(i,colk['Limt'])
            codeclass[code].Units=sh.cell_value(i,colk['Units'])
            codeclass[code].Inf['Minprice']=sh.cell_value(i,colk['Minprice'])
            code_house[code]=sh.cell_value(i,colk['house'])[0:3]
            if code_house[code]==u'CFF':
                code_house[code]=u'CFE' 
            if isdelta:
                codeclass[code].Inf['delta_rate']=float(sh.cell_value(i,colk['delta_rate']))
            codeclass[code].House=code_house[code]
    ini_codeclass=codeclass
    date=datetime.datetime.fromtimestamp(os.path.getmtime(dp))
    tt.append('保证金率 '+date.strftime('%Y-%m-%d %H:%M:%S'))
    dt.append(date)
    if gui<>'':
        gui.labels[pak].config(text=tt[-1])
        if ntime>dt[-1]:
            gui.labels[pak].config(text='（该更新了）'+tt[-1],fg='red')
        pak+=1
        gui.root.update()
    

    dp=u'\\\\10.100.6.20\\fkfile\\大连交易参数表'
    files = os.listdir(dp) 
    temp=0
    for x in files:
        if os.path.getmtime(dp+'\\'+x)>temp:
            temp=os.path.getmtime(dp+'\\'+x)
            xx=x
    dp=dp+'\\'+xx
    book = xlrd.open_workbook(dp)
    sh=book.sheets()[0]
    n=sh.nrows
    sp_order=sh.col_values(2,2,n)
    for i in range(len(sp_order)):
        sp_order[i]=sp_order[i].upper()
    ini_sporder=sp_order    
    date=datetime.datetime.fromtimestamp(os.path.getmtime(dp))
    tt.append('大连套利表 '+date.strftime('%Y-%m-%d %H:%M:%S'))
    dt.append(date)
    if gui<>'':
        gui.labels[pak].config(text=tt[-1])
        if ntime>dt[-1]:
            gui.labels[pak].config(text='（该更新了）'+tt[-1],fg='red')
        pak+=1
        gui.root.update()
    
    dp=u'\\\\10.100.6.20\\fkfile\\风险试算\\对应表\\风控投资者信息查询.xls'
    book = xlrd.open_workbook(dp)
    sh=book.sheets()[0]
    n=sh.nrows   
    colname=sh.row_values(1)
    ivid=sh.col_values(colname.index(u'客户号'),2,n)
    ivid=change_id_type(ivid)
    ivde=sh.col_values(colname.index(u'对应的营业部'),2,n)
    ini_belong=dict(zip(ivid,ivde))
    date=datetime.datetime.fromtimestamp(os.path.getmtime(dp))
    tt.append('客户对应表 '+date.strftime('%Y-%m-%d %H:%M:%S'))
    dt.append(date)
    if gui<>'':
        gui.labels[pak].config(text=tt[-1])
        if ntime>dt[-1]:
            gui.labels[pak].config(text='（该更新了）'+tt[-1],fg='red')
        pak+=1
        gui.root.update()
    
    dp=u'\\\\10.100.6.20\\fkfile\\风险试算\\对应表\\用户权限分配表.csv'    
    f=codecs.open(dp,'r','utf-8')
    colname=f.readline().replace('"','').replace('\r\n','').split(',')
    k1=colname.index(u'用户代码')
    k2=colname.index(u'交易中心标识号')
    line=f.readline().replace('"','').replace('\r\n','').split(',')
    ini_seat={}
    ivid=[]
    ivseat=[]
    while line!=['']:
        ivid.append(line[k1])
        ivseat.append(line[k2])
        #ini_seat[line[k1]]=line[k2]
        line=f.readline().replace('"','').replace('\r\n','').split(',')
    ivid=change_id_type(ivid)
    ini_seat=dict(zip(ivid,ivseat))
    f.close()
    date=datetime.datetime.fromtimestamp(os.path.getmtime(dp))
    tt.append('席位表 '+date.strftime('%Y-%m-%d %H:%M:%S'))   
    dt.append(date)
    if gui<>'':
        gui.labels[pak].config(text=tt[-1])
        if ntime>dt[-1]:
            gui.labels[pak].config(text='（该更新了）'+tt[-1],fg='red')
        pak+=1
        gui.root.update()    
    
    
    dp=u'\\\\10.100.6.20\\fkfile\\风险试算\\对应表\\投资者保证金率属性.csv'
    #dp=u'\\\\10.100.6.20\\fkfile\\IORI-陈志荣\\三表\\投资者保证金率属性.csv'
    try :
        f=codecs.open(dp,'r','utf-8')
        ini_sprate=f.readlines()
    except:
        f=codecs.open(dp,'r','gbk')#过节启用
        ini_sprate=f.readlines()           
    f.close()
    date=datetime.datetime.fromtimestamp(os.path.getmtime(dp))
    tt.append('优惠客户表 '+date.strftime('%Y-%m-%d %H:%M:%S'))   
    dt.append(date)
    if gui<>'':
        gui.labels[pak].config(text=tt[-1])
        if ntime>dt[-1]:
            gui.labels[pak].config(text='（该更新了）'+tt[-1],fg='red')
        pak+=1    
        gui.root.update()
    
    dp=u'\\\\10.100.6.20\\fkfile\\工具箱\\投资者电话导出和短信生成\\投资者信息查询.xls'
    book = xlrd.open_workbook(dp)
    sh=book.sheets()[0]
    n=sh.nrows   
    colname=sh.row_values(1)
    ivid=sh.col_values(colname.index(u'投资者代码'),1,n)
    ivde=sh.col_values(colname.index(u'手机'),1,n)
    ivtype=sh.col_values(colname.index(u'投资者类型'),1,n)
    sh=book.sheets()[1]
    n=sh.nrows   
    colname=sh.row_values(0)
    ivid=ivid+sh.col_values(colname.index(u'投资者代码'),1,n)
    ivde=ivde+sh.col_values(colname.index(u'手机'),1,n)
    ivtype=ivtype+sh.col_values(colname.index(u'投资者类型'),1,n)
    ivid=change_id_type(ivid)
    ini_phone=dict(zip(ivid,ivde)) 
    ini_invtype=dict(zip(ivid,ivtype))      
    date=datetime.datetime.fromtimestamp(os.path.getmtime(dp))
    tt.append('基本信息表 '+date.strftime('%Y-%m-%d %H:%M:%S'))   
    dt.append(date)
    if gui<>'':
        gui.labels[pak].config(text=tt[-1])
        if ntime>dt[-1]:
            gui.labels[pak].config(text='（该更新了）'+tt[-1],fg='red')
        pak+=1  
        gui.root.update()
        gui.data['sheet_info']=tt
        gui.data['ini_codeclass']=ini_codeclass
        gui.data['ini_sporder']=ini_sporder
        gui.data['ini_belong']=ini_belong
        gui.data['ini_seat']=ini_seat
        gui.data['ini_sprate']=ini_sprate
        gui.data['ini_phone']=ini_phone
        gui.data['ini_invtype']=ini_invtype        
        gui.data['code_house']=code_house
        gui.data['shfe_unsp'],gui.data['cfe_unsp']=cal_shfe_unsp(ini_codeclass)
    return ini_codeclass,ini_sporder,ini_belong,ini_seat,ini_sprate,ini_phone,code_house,ini_invtype,tt

def read_sh3(invfname=False,posfname=False,riskfname=False,gui=''):
    ''' 函数说明：读取三表'''
    #dp='E:\\HTRM\\风险指标计算系统\\历史夜盘盘前投资者资金信息\\'
    #filename=dp+'投资者资金信息20160201.xls'
    allinv={}
    allpos=[]
    allcode={}
    if invfname<>False:
        xlsname=invfname
        #xlsname=unicode(filename , "utf8")
        book = xlrd.open_workbook(xlsname)
        sh=book.sheets()[0]
        n=sh.nrows    
        colname=sh.row_values(1)
        colk={'invid':colname.index(u'投资者代码'),'inv_name':colname.index(u'投资者名称'),'invdepartment':colname.index(u'组织架构名称')
        ,'riskstate':colname.index(u'风险状态'),'lastriskstate':colname.index(u'昨风险状态'),'riskdegree':colname.index(u'交易所风险度')
        ,'capital':colname.index(u'权益'),'lastcapital':colname.index(u'昨权益'),'margin':colname.index(u'交易所保证金'),'spmark':colname.index(u'特殊标志')
        ,'closeprofit':colname.index(u'平仓盈亏'),'holdprofit':colname.index(u'持仓盈亏'),'cashmove':colname.index(u'出入金')
        ,'fee':colname.index(u'手续费'),'frozenmoney':colname.index(u'资金冻结'),'deliverymargin':colname.index(u'交割保证金')
        ,'mortgagemoney':colname.index(u'质押金额'),'leftcapital':colname.index(u'交易所可用资金'),'cor_leftcapital':colname.index(u'可用资金')
        ,'cor_riskdegree':colname.index(u'风险度'),'leastmoney':colname.index(u'保底资金'),'cor_margin':colname.index(u'保证金'),'rr_bankmoney':colname.index(u'可取资金')}
        for i in range(2,n-2):
            invinf={}
            for x in colk:
                invinf[x]=sh.cell_value(i,colk[x])
            #for x in invinf:
                #invinf[x]=invinf[x].encode('gbk')
            allinv[invinf['invid']]=invinf
    
    #dp='E:\\HTRM\\风险指标计算系统\\历史夜盘盘前持仓信息\\'
    #filename=dp+'持仓查询20160201.xls'
    if posfname<>False:
        xlsname=posfname
        #xlsname=unicode(filename , "utf8")      
        book = xlrd.open_workbook(xlsname)
        sh=book.sheets()[0]
        n=sh.nrows      
        colname=sh.row_values(1)
        rule=re.compile(r'[^a-zA-z]')
        colk={'invid':colname.index(u'投资者代码'),'code':colname.index(u'合约'),'longnums':colname.index(u'总买持')
        ,'longopenprice':colname.index(u'买开仓均价'),'longholdprice':colname.index(u'买均价'),'longmar':colname.index(u'买保证金')
        ,'shortnums':colname.index(u'总卖持'),'shortopenprice':colname.index(u'卖开仓均价'),'shortholdprice':colname.index(u'卖均价')
        ,'shortmar':colname.index(u'卖保证金'),'posident':colname.index(u'投保'),'realmargin':colname.index(u'实收交易所保证金'),'house':colname.index(u'交易所')}
        for i in range(2,n-2):
            posinf={}
            for x in colk:
                posinf[x]=sh.cell_value(i,colk[x])
            posinf['code']=posinf['code'].upper()
            posinf['vari']=rule.sub('',posinf['code'])     
            allpos.append(posinf)
            
    #dp='E:\\HTRM\\风险指标计算系统\\历史夜盘前实时风控行情\\'
    #filename=dp+'实时风控行情20160201.xls'
    if riskfname<>False:
        xlsname=riskfname
        #xlsname=unicode(filename , "utf8")
        book = xlrd.open_workbook(xlsname)
        sh=book.sheets()[0]
        n=sh.nrows     
        colname=sh.row_values(1)
        rule=re.compile(r'[^a-zA-z]')
        colk={'code':colname.index(u'合约'),'settlement':colname.index(u'今结算'),'meanprice':colname.index(u'均价'),'high':colname.index(u'最高价')
        ,'low':colname.index(u'最低价'),'opn':colname.index(u'今开盘'),'cls':colname.index(u'今收盘'),'open_interest':colname.index(u'持仓量')
        ,'volume':colname.index(u'成交量'),'price_now':colname.index(u'最新价'),'lastcls':colname.index(u'昨收盘')
        ,'lastsettlement':colname.index(u'昨结算'),'lastopen_interest':colname.index(u'昨持仓'),'bid':colname.index(u'买价'),'ask':colname.index(u'卖价')
        ,'limtup':colname.index(u'涨停价'),'limtdown':colname.index(u'跌停价')}
        for i in range(2,n):
            codeinf={}
            if u'合约' in colname:
                for x in colk:
                    codeinf[x]=sh.cell_value(i,colk[x])
                codeinf['code']=codeinf['code'].upper()
                codeinf['vari']=rule.sub('',codeinf['code'])
            else:                
                codeinf['code']=sh.cell_value(i,colname.index(u'合约代码')).upper()
                codeinf['vari']=rule.sub('',codeinf['code'])
                codeinf['settlement']=sh.cell_value(i,colname.index(u'今结算'))
                codeinf['meanprice']=sh.cell_value(i,colname.index(u'平均价'))
                codeinf['high']=string.atof(sh.cell_value(i,colname.index(u'最高价')))
                codeinf['low']=sh.cell_value(i,colname.index(u'最低价'))  
                codeinf['opn']=sh.cell_value(i,colname.index(u'开盘价'))
                codeinf['cls']=sh.cell_value(i,colname.index(u'今收盘'))
                codeinf['open_interest']=sh.cell_value(i,colname.index(u'持仓量'))
                codeinf['volume']=sh.cell_value(i,colname.index(u'成交量'))
                codeinf['price_now']=string.atof(sh.cell_value(i,colname.index(u'最新价')))
                codeinf['lastcls']=sh.cell_value(i,colname.index(u'昨收盘'))
                codeinf['lastsettlement']=sh.cell_value(i,colname.index(u'昨结算'))
                codeinf['lastopen_interest']=string.atof(sh.cell_value(i,colname.index(u'昨持仓')))
                codeinf['volume']=sh.cell_value(i,colname.index(u'成交量'))
                codeinf['bid']=sh.cell_value(i,colname.index(u'买价'))
                codeinf['ask']=sh.cell_value(i,colname.index(u'卖价'))
                for x in codeinf:
                    if x<>'code' and x<>'vari':
                        codeinf[x]=string.atof(codeinf[x])
            allcode[codeinf['code']]=codeinf
    tt=[]
    date=datetime.datetime.fromtimestamp(os.path.getmtime(invfname))
    tt.append(u'资金表 '+date.strftime('%Y-%m-%d %H:%M:%S')) 
    date=datetime.datetime.fromtimestamp(os.path.getmtime(posfname))
    tt.append(u'持仓表 '+date.strftime('%Y-%m-%d %H:%M:%S'))
    date=datetime.datetime.fromtimestamp(os.path.getmtime(riskfname))
    tt.append(u'行情表 '+date.strftime('%Y-%m-%d %H:%M:%S'))
    if gui<>'':
        for i in range(3):        
            gui.labels[i].config(text=tt[i],fg='orangered')
        gui.root.update()
        if len(gui.data['sheet_info'])<=6:
            gui.data['sheet_info']=tt+gui.data['sheet_info']
        else:
            gui.data['sheet_info'][:3]=tt
        gui.data['allinv']=allinv
        gui.data['allpos']=allpos
        gui.data['allcode']=allcode       
    return allinv,allpos,allcode,tt
def creat_class(data):
    ''' 函数说明：把三表内容转换为预设计的类'''
    allinv,allpos,allrisk=data['allinv'],data['allpos'],data['allcode']
    ini_codeclass=data['ini_codeclass']
    investor={}
    code={}
    vclass={}
    for x in allrisk:
        if not vclass.has_key(allrisk[x]['vari']):        
            vclass[allrisk[x]['vari']]=VariClass(allrisk[x]['vari'])
        if ini_codeclass.has_key(x):
            ini_codeclass[x].intvalue(allrisk[x])
    t=''
    uninv=''   
    allposcode=[]         
    for x in allinv:
        investor[x]=InvestorClass(allinv[x])
    for posinf in allpos:
        invid=posinf['invid']
        if investor.has_key(invid):
            investor[invid].addpos(posinf)
        else:
            uninv=u'有持仓找不到投资者' 
            print invid
        if len(posinf['code'])<=6:
            allposcode.append(posinf['code'])

        
    for x in investor:
        investor[x].getpurepos()
    data['invclass']=investor
    data['variclass']=vclass
    cal_meanprice(ini_codeclass)
    #data['codeclass']=ini_codeclass
    
    if len(uninv)>0:
        t+=uninv+'\n'
    unexit_code=[x for x in ini_codeclass.keys() if x not in allrisk.keys()]
    if len(unexit_code)>0:
        t+=u'风控行情表缺少合约'  
        for x in unexit_code:
            t+=x
        t+='\n'    
    unexit_ini=[x for x in set(allposcode) if x not in ini_codeclass.keys()]
    if len(unexit_ini)>0:
        t+=u'保证金率表缺少合约，大表计算将出错：'  
        for x in unexit_code:
            t+=x
    if len(t)>0:
        import tkMessageBox
        print t    
def allclient_cal(data,k1=0,k2=0,ps='settlement',isaddrate=True):
    ''' 函数说明：计算所有客户风险度，并做相应处理'''  
    #k1-ptag最小变动价位选取方向，k2-ctag是否引入价格变动
    [data['ini_codeclass'][x].adj_price(k1,k2,ps,isaddrate) for x in data['ini_codeclass']]
    invclass=allcal(data['invclass'],data['ini_codeclass'],data['ini_sporder'],data['code_house'],data['variclass'],data['shfe_unsp'],data['cfe_unsp'])    
    alldeparment,allvari=cal_otherinf(invclass,data['invclass'],data['ini_codeclass'])
    ivid=invclass.keys()
    rk=[(i,invclass[ivid[i]].InvInf['riskdegree']) for i in range(len(ivid))]
    rk.sort(key=itemgetter(1),reverse=True)
    ivid_sorted=[ivid[x[0]] for x in rk]
    i=len(ivid)-1
    locb=[]
    while float(rk[i][1])<0:
        locb.append(ivid_sorted[i])
        ivid_sorted.pop()
        i-=1
    ivid_sorted=locb+ivid_sorted
    data['ivid_sorted']=ivid_sorted
    i=1
    tabdata={}
    for iid in ivid_sorted:
        tabdata[i]={}
        tabdata[i]=invclass[iid].InvInf
        i+=1        
    return invclass,tabdata,alldeparment,allvari  
def cal_otherinf(newinvc,oldinvc,codeclass):
    '''函数说明：计算所有客户其他信息'''
    alldeparment=[]
    allvari=[]
    for x in newinvc:
        tags=[]
        ninf=newinvc[x].InvInf
        oinf=oldinvc[x].InvInf
        ninf['mortgagestatus']=''
        if ninf['mortgagemoney']>0:
            if ninf['capital']<ninf['mortgagemoney']:
                ninf['mortgagestatus']=u'质押严重异常'
            if ninf['capital']<1.25*ninf['mortgagemoney']:
                ninf['mortgagestatus']=u'质押异常'
        m1=ninf['capital']-1.25*ninf['mortgagemoney']#质押配比
        m2=ninf['leftcapital']#结算可用       
        m3=oinf['cor_leftcapital']-max(oinf['holdprofit'],0)-max(oinf['closeprofit'],0)-oinf['leastmoney']-oinf['mortgagemoney']
        #公司可用-平仓盈利-盘面持仓盈利-保底资金-质押金额
        ninf['bankmoney']=min([m1,m2,m3])
        m2=ninf['cor_leftcapital']-max(ninf['holdprofit'],0)-max(oinf['closeprofit'],0)#结算后可用-结算价计算的浮动盈利-平仓盈利
        m3=oinf['capital']-max(oinf['holdprofit'],0)-max(oinf['closeprofit'],0)-oinf['frozenmoney']-oinf['cor_margin']-oinf['deliverymargin']
        #盘面权益-平仓盈利-盘面持仓盈利-资金冻结-公司盘面保证金-交割保证金
        ninf['oa_bankmoney']=min([m1,m2,m3])
        ninf['maxcode']=''
        ninf['posstrut']=''
        ninf['mcoderate']=''
        ninf['codedir']=''
        ninf['outrange']=''
        ninf['forcedrange']=''
        ninf['forcednums']=''
        ninf['forcedbound']=''
        if ninf['leftcapital']<0:
            tags.append('forced_limited')
        if ninf['leftcapital']<0:
            tags.append('margin_call')
        tags.append(ninf['lastriskstate'])
        tags.append(ninf['invdepartment'])
        if ninf['invdepartment'] not in alldeparment:
            alldeparment.append(ninf['invdepartment'])
        ninf['tags']=tags
        
        codevalue=0  
        allvalue=0
        codenums=0
        pl=0
        ps=0
        for posinf in oldinvc[x].Position:
            if len(posinf['code'])>6:
                continue
            posnums=posinf['longnums']-posinf['shortnums']
            code=posinf['code']            
            ttv=abs(posnums*codeclass[code].Units*codeclass[code].Price)
            if ttv>codevalue:
                ninf['maxcode']=code
                codenums=posnums
                pl=posinf['longnums']
                ps=posinf['shortnums']
                codevalue=ttv
            allvalue+=ttv
            tags.append(codeclass[code].Vari)
            if codeclass[code].Vari not in allvari:
                allvari.append(codeclass[code].Vari)
        ninf['tags']=tags
        if ninf['maxcode']<>'':
            ninf['mcoderate']=codevalue/allvalue
            if ninf['mcoderate']<1:
                ninf['posstrut']=u'复合持仓'
            else:
                ninf['posstrut']=u'单一持仓'
            if codenums>0:
                ninf['codedir']=u'多仓'                
            else:
                ninf['codedir']=u'空仓'
            ninf['posstrut']+=ninf['codedir']
            cclass=codeclass[ninf['maxcode']]
            ninf['outrange']=abs((codeclass[ninf['maxcode']].Price-ninf['capital']/(codeclass[ninf['maxcode']].Units*codenums))/codeclass[ninf['maxcode']].Price-1)*100
            if codenums-cclass.Rate*max(pl,ps)<>0:
                ninf['forcedbound']=cclass.Price-ninf['leftcapital']/(cclass.Units*(codenums-cclass.Rate*max(pl,ps)))
            else:
                ninf['forcedbound']=0                    
            k=abs(codenums)/codenums
            ninf['forcednums']=ninf['leftcapital']/(-k*(cclass.Inf['price_now']-cclass.Price)*cclass.Units-cclass.Rate*cclass.Price*cclass.Units)
    return alldeparment,allvari
def change_id_type(ll):
    ''' 用于投资者代码统一转换为字符格式'''
    if type(ll[0])==float:
        rls=[str(int(x)) for x in ll]
        return rls
    elif  type(ll[0])==unicode:
        return ll
    else:
        print u'客户号为未知类型'
        print type(ll[0])
        return ll
def float_to_str(f):
    ''' 用于数字格式转化为银行数字格式'''
    s=str(f)
    if not re.match('[0-9-]',s):
        return s
    k=len(s)
    if '.' in s:
        k=s.index('.')
    k=k-3
    while k>0:
        s=s[:k]+','+s[k:]
        k-=3
    if s[0]=='-' and s[1]==',':
        s=s[0]+s[2:]
    return s
def english_to_ch(s):
    ''' 函数说明：列名转换为中文名字'''
    tp=[('invid',u'投资者代码'),('inv_name',u'投资者名称'),('invdepartment',u'组织架构名称'),('riskstate',u'昨风险状态'),('lastriskstate',u'昨风险状态')
    ,('riskdegree',u'交易所风险度'),('capital',u'权益'),('lastcapital',u'昨权益'),('spmark',u'特殊标志'),('margin',u'交易所保证金')
    ,('closeprofit',u'平仓盈亏'),('holdprofit',u'持仓盈亏'),('cashmove',u'出入金'),('fee',u'手续费'),('frozenmoney',u'资金冻结')
    ,('deliverymargin',u'交割保证金'),('mortgagemoney',u'质押金额'),('leftcapital',u'交易所可用资金'),('code',u'合约代码'),('vari',u'品种')
    ,('longnums',u'总买持'),('longopenprice',u'买开仓均价'),('longholdprice',u'买均价'),('longmar',u'买保证金'),('shortnums',u'总卖持')
    ,('shortopenprice',u'卖开仓均价'),('shortholdprice',u'卖均价'),('shortmar',u'卖保证金'),('posident',u'投保'),('settlement',u'今结算')
    ,('meanprice',u'均价'),('high',u'最高价'),('low',u'最低价'),('opn',u'今开盘'),('cls',u'今收盘'),('open_interest',u'持仓量')
    ,('volume',u'成交量'),('price_now',u'最新价'),('lastcls',u'昨收盘'),('lastsettlement',u'昨结算'),('lastopen_interest',u'昨持仓')
    ,('volume',u'成交量'),('seat',u'席位'),('phone',u'手机号'),('cor_margin',u'公司保证金'),('cor_riskdegree',u'公司风险度'),('cor_leftcapital',u'公司可用资金')
    ,('bid',u'买价'),('ask',u'卖价'),('delta_price',u'结算价调整'),('delta_rate',u'保证金调整'),('limtup',u'涨停价'),('limtdown',u'跌停价')
    ,('r_mrate',u'调整后保证金率'),('Mrate',u'交易所保证金率'),('r_price',u'试算价格'),('mortgagestatus',u'质押异常'),('bankmoney',u'银期可取资金')
    ,('leastmoney',u'保底资金'),('maxcode',u'价值最大合约'),('posstrut',u'持仓结构'),('mcoderate',u'持仓比例'),('codedir',u'持仓方向')
    ,('outrange',u'穿仓幅度'),('forcedbound',u'强平边界'),('forcednums',u'强平手数'),('cor_longmar',u'公司买保证金'),('cor_shortmar',u'公司卖保证金')
    ,('house',u'交易所'),('oa_bankmoney',u'OA可取'),('phone',u'手机号'),('rr_bankmoney',u'盘面可取资金')]
    en=[x[0] for x in tp]
    ch=[x[1] for x in tp]
    if s in en:
        return ch[en.index(s)]
    else:
        #print u'找不到对应的中文名：'+s
        return s
def cal_tradeday(sdate='2016-01-11',edate='2016-01-9'):
    ''' 函数说明：计算两个日期的交易日'''
    if type(sdate)==datetime.datetime:
        sdate=datetime.datetime.strftime(sdate, "%Y-%m-%d")
    if type(edate)==datetime.datetime:
        edate=datetime.datetime.strftime(edate, "%Y-%m-%d")
    v=[]    
    for x in vacations:
        x=x[0:4]+'-'+x[4:6]+'-'+x[6:8]
        v.append(x)
    oneday = datetime.timedelta(days=1)
    tempdate=datetime.datetime.strptime(sdate, "%Y-%m-%d")#字符串转时间
    n=0
    while tempdate<=datetime.datetime.strptime(edate, "%Y-%m-%d"):
        if tempdate.strftime("%Y-%m-%d") not in v:
            if tempdate.weekday()<5:   
                n+=1
        tempdate+=oneday
    return n    
def cal_shfe_unsp(vari_codeclass):
    ''' 函数说明：计算上期所和中金所无套利优惠合约'''
    unspcode=[]
    edate1=datetime.datetime.now()
    sdate=datetime.datetime(edate1.year,edate1.month,1)
    edate2=datetime.datetime(edate1.year,edate1.month,15)    
    while edate2.strftime("%Y%m%d") in vacations or edate2.weekday()>=5:
        edate2=edate2+datetime.timedelta(days=1)
    n=cal_tradeday(sdate,edate2)-cal_tradeday(sdate,edate1)
    if n>=0 and n<=5:
        s=datetime.datetime.strftime(edate1,"%Y-%m-%d")
        scode=s[2:4]+s[5:7]
        for x in vari_codeclass:
            if vari_codeclass[x].House=='SHF':
                code=vari_codeclass[x].Vari+scode
                if code not in unspcode:
                    unspcode.append(code)
    scode=datetime.datetime.strftime(edate1,"%Y%m")[2:]
    cfe_un=['T'+scode,'TF'+scode]
    return unspcode,cfe_un
def match_corp_marr(invclass,variclass,ini_sprate,isforce=False):
    ''' 函数说明：公司保证金率匹配'''
    sdate=datetime.datetime.now()
    edate=sdate-datetime.timedelta(days=sdate.day-1)+datetime.timedelta(days=31)
    edate2=datetime.datetime(edate.year,edate.month,1)-datetime.timedelta(days=1)
    n=cal_tradeday(sdate,edate2)
    mark=0
    #if True:
    if n>=2 and n<=3:
        mark=1
        last_code=[]
        belast_code=[]
        lcd=datetime.datetime.strftime(edate2+datetime.timedelta(days=1),"%Y%m%d")[2:6]
        llcd=lcd[1:]
        blcd=datetime.datetime.strftime(edate2+datetime.timedelta(days=32),"%Y%m%d")[2:6]
        for x in variclass:
            if variclass[x].House=='SHF':
                last_code.append(x+lcd)
                belast_code.append(x+blcd)
            if variclass[x].House=='DCE':
                last_code.append(x+lcd)
            if variclass[x].House=='CZC':
                last_code.append(x+llcd)
    if isforce:
        mark=0                     
    colname=ini_sprate[0].replace('"','').replace('\r\n','').split(',')
    sp_clain={}
    com_clain={}
    last_cl={}
    belast_cl={}
    colk={'vari':colname.index(u'合约代码'),'ivid':colname.index(u'投资者代码')
    ,'mname':colname.index(u'保证金分段名称'),'rate':colname.index(u'投机多头保证金率')}
    spcial_code={}
    for i in range(1,len(ini_sprate)):
        line=ini_sprate[i].replace('"','').replace('\r\n','').split(',')
        if line[colk['ivid']]==u'所有':
            if line[colk['mname']]==u'上市月后含1个交易日':
                com_clain[line[colk['vari']].upper()]=line[colk['rate']]
            elif line[colk['mname']]==u'交割月后含1个交易日前2个交易日' or line[colk['mname']]==u'交割月后含1个公历日前2个交易日':
                last_cl[line[colk['vari']].upper()]=line[colk['rate']]
            elif line[colk['mname']]==u'交割月前1个月后含1个交易日前2个交易日':
                belast_cl[line[colk['vari']].upper()]=line[colk['rate']]
            if len(line[colk['vari']])>2:
                spcial_code[line[colk['vari']].upper()]=line[colk['rate']]
        else:
            t=line[colk['vari']].upper()+'_'+line[colk['ivid']]
            sp_clain[t]=line[colk['rate']]
    com_clain['IF']=copy.copy(com_clain['IC'])
    for x in invclass:
        pos=invclass[x].Position
        for posinf in pos:
            if len(posinf['code'])>6:
                continue
            nf=True
            ky=posinf['vari']+'_'+x
            if sp_clain.has_key(ky):
                posinf['cor_rate']=float(sp_clain[ky])
                nf=False
            if nf:
                nf2=True
                if mark:
                    if posinf['code'] in last_code:
                        posinf['cor_rate']=float(last_cl[posinf['vari']])
                        nf2=False
                    elif posinf['code'] in belast_code:
                        posinf['cor_rate']=float(belast_cl[posinf['vari']])
                        nf2=False                        
                if nf2:
                    posinf['cor_rate']=float(com_clain[posinf['vari']])
                if spcial_code.has_key(posinf['code']):
                    posinf['cor_rate']=float(spcial_code[posinf['code']])
def cal_meanprice(codeclass2):
    ''' 函数说明：预估冷门合约计算价 '''
    vdic={}    
    codeclass=copy.deepcopy(codeclass2)
    for x in codeclass:
        if vdic.has_key(codeclass[x].Vari):
            vdic[codeclass[x].Vari].append(x)
        else:
            vdic[codeclass[x].Vari]=[x]
    for x in vdic:
        vdic[x].sort()
        clist=vdic[x]
        for i in range(len(vdic[x])):
            cinf=codeclass[clist[i]]
            if cinf.Inf['meanprice']<>0:
                continue
            iset=False
            #print clist[i]
            if cinf.Inf['bid']==0 or cinf.Inf['bid']=='':
                if cinf.Inf['ask']==0 or cinf.Inf['ask']=='':  
                    pass
                elif abs(cinf.Inf['ask']-cinf.Inf['lastsettlement']*(1-cinf.Inf['Limt']))<=cinf.Inf['Minprice']:
                    cinf.Inf['meanprice']=cinf.Inf['ask']
                    iset=True
            elif abs(cinf.Inf['bid']-cinf.Inf['lastsettlement']*(1+cinf.Inf['Limt']))<=cinf.Inf['Minprice']:
                cinf.Inf['meanprice']=cinf.Inf['bid']
                iset=True
            else:
                p3=[cinf.Inf['bid'],cinf.Inf['ask'],cinf.Inf['lastsettlement']]
                p3.sort()
                cinf.Inf['meanprice']=p3[1]
                iset=True                
            if not iset:
                k=i
                iset2=False
                while k-1>=0:
                    if codeclass2[clist[k]].Inf['meanprice']<>0:
                        cinf.Inf['meanprice']=cinf.Inf['lastsettlement']*codeclass[clist[k]].Inf['meanprice']/codeclass[clist[k]].Inf['lastsettlement']
                        cinf.Inf['meanprice']=math.floor(cinf.Inf['meanprice']/cinf.Inf['Minprice'])*cinf.Inf['Minprice']
                        iset2=True
                        break
                    k-=1
                if not iset2:
                    cinf.Inf['meanprice']=cinf.Inf['lastsettlement']
    for x in codeclass:
        codeclass2[x]=codeclass[x]     
def upperlist(isadd=0):
    lt=[['a','m','y','p','oi','cf','rm','sr','c','cs','jd'],['cu','zn','ni','al','sn','pb'],
    ['rb','i','j','jm','hc','zc'],['ag','au'],['ta','ma','bu','l','pp','v','ru','fg'],
    ['if','ih','ic'],['t','tf']]
    for i in range(len(lt)):
        for j in range(len(lt[i])):
            lt[i][j]=lt[i][j].upper()
    if isadd==0:
        return lt
    if isadd==1:
        lt2=[]
        for x in lt:
            lt2+=x
        return lt2  
def get_CFE_setl(codeclass,path):
    ''' 函数说明：预估中金所结算价 '''
    from WindPy import w
    now=datetime.datetime.now()
    strtoday=now.strftime('%Y-%m-%d')
    today=datetime.datetime.strptime(strtoday, '%Y-%m-%d')
    s1=strtoday+' 14:00:00'
    s2=strtoday+' 14:15:00'
    e1=strtoday+' 15:00:00'
    e2=strtoday+' 15:15:00'
    s1time=datetime.datetime.strptime(s1, '%Y-%m-%d %H:%M:%S')#字符串转时间   
    if now<s1time:
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'请14:00后再获取模拟结算价') 
        return
    sercode=[]
    for x in codeclass:
        if codeclass[x].House=='CFE':
            sercode.append(x)
    with open(path+'cfe_vol.pickle', 'rb') as f:cfe_vol = pickle.load(f)
    tc=cfe_vol.keys()[0]
    tempstr=cfe_vol[tc][-1][0]
    temptime=datetime.datetime.strptime(tempstr, '%Y-%m-%d %H:%M:%S')
    if temptime<today:
        cfe_vol={}
        for x in sercode:
            cfe_vol[x]=[]
        sdate=strtoday+' 09:00:00'
    else:
        sdate=tempstr
    w.start()
    for x in sercode:
        if sdate<>strtoday+' 09:00:00':
            sdate=cfe_vol[x][-1][0]
        windreulst=w.wst(x+'.CFE','volume,last',sdate,e2)
        if len(cfe_vol[x])>0:
            while True:
                if cfe_vol[x][-1][0]==sdate:
                    cfe_vol[x].pop()
                else:
                    break
        for i in range(len(windreulst.Times)):
            v=[windreulst.Times[i].strftime('%Y-%m-%d %H:%M:%S')]
            v.append(windreulst.Data[0][i])
            v.append(windreulst.Data[1][i])
            cfe_vol[x].append(v)
        tf=['T','TF']
        ss=s1
        if codeclass[x].Vari in tf:
            ss=s2
        i=0
        while i<len(cfe_vol[x])-1:
            if cfe_vol[x][i][0]<ss and cfe_vol[x][i+1][0]>=ss:
                break
            i+=1
        i+=1
        vol=0
        volpri=0
        if i<len(cfe_vol[x]):
            for k in range(i,len(cfe_vol[x])):
                deltavol=cfe_vol[x][k][1]-cfe_vol[x][k-1][1]
                vol+=deltavol
                volpri+=deltavol*cfe_vol[x][k][2]
            if vol>0:
                t=volpri/vol
                t=math.floor(t/codeclass[x].Inf['Minprice'])*codeclass[x].Inf['Minprice']
                codeclass[x].Inf['settlement']=t    
                codeclass[x].Inf['meanprice']=t
    with open(path+'cfe_vol.pickle', 'wb') as f:pickle.dump(cfe_vol, f)   
def get_rtdata(codeclass):
    ''' 函数说明：获取最新行情数据 '''
    from WindPy import w
    w.start()
    windname={'rt_high':'high','rt_low':'low','rt_latest':'price_now','rt_vwap':'meanprice','rt_oi':'open_interest','rt_vol':'volume','rt_bid1':'bid','rt_ask1':'ask'}
    colname=windname.keys()
    for x in codeclass:    
        windreulst=w.wsq(x+'.'+codeclass[x].House,colname)
        for i in range(len(colname)):
            codeclass[x].Inf[windname[colname[i]]]=windreulst.Data[i][0]
def get_leastdata(codeclass,warmr=0.01):
    ''' 函数说明：获取最新价格 '''
    from WindPy import w
    w.start()
    code=[]
    unvari=['WR','FU','B','FB','BB','WH','RI','LR','PM','RS','JR']
    ngvari='p;j;a;b;m;y;jm;i;rm;sr;ta;ma;oi;cf;fg;tc;zc;cu;al;zn;pb;ag;au;ru;rb;hc;bu;ni;sn'
    ngvari=ngvari.upper().split(';')    
    for x in codeclass:
        code.append(x+'.'+codeclass[x].House)
    windreulst=w.wsq(code,'rt_latest')
    res=[]
    tags=[]
    for i,x in enumerate(code):
        cc=x.split('.')[0]
        if codeclass[cc].Vari in unvari or windreulst.Data[0][i]==0 or codeclass[cc].Inf['volume']==0:
            continue
        ul=codeclass[cc].Inf['limtup']
        dl=codeclass[cc].Inf['limtdown']
        ls=codeclass[cc].Inf['lastsettlement']
        if ls==0:
            continue
        r=windreulst.Data[0][i]/ls-1
        if (ul-windreulst.Data[0][i])/ls<warmr or (windreulst.Data[0][i]-dl)/ls<warmr:
            od=[cc]
            rlast=windreulst.Data[0][i]
            od.append(rlast)
            if r>0:
                tags.append('red')
                od.append(ul)
            else:
                tags.append('green')
                od.append(dl)
            od.append(round(r*100,2))
            od.append(round(ul/ls-1,2)*100)
            res.append(od)
    return res,tags
def forced_client_inf(invclass,codeclass):
    ''' 函数说明：第二天强平客户信息'''  
    ivid=[]    
    for x in invclass:
        if invclass[x].InvInf['lastriskstate']==u'强平':
            ivid.append(x)
    warminf=[]
    for x in ivid:
        for pos in invclass[x].Position:
            if codeclass[pos['code']].House=='SHF':
                text=pos['code']+u'——目前持仓量:'+str(codeclass[pos['code']].Inf['open_interest'])+'。'
                text+=x+'-'+invclass[x].Name+':'+str(pos['longnums']-pos['shortnums'])+u'手'
                warminf.append(text)
    return warminf
def cal_invfound(filename,codeclass):
    ''' 函数说明：导入成交并结算投资者保障基金'''  
    #filename=self.path+'\\'+u'成交查询.xls'
    fun_rate=6/10000000
    book = xlrd.open_workbook(filename)
    sh=book.sheets()[0]
    n=sh.nrows      
    colname=sh.row_values(1)
    colk={'code':colname.index(u'合约'),'ivid':colname.index(u'投资者代码'),'nums':colname.index(u'成交手数'),'price':colname.index(u'成交价')}
    invfound={}
    name={}
    print sh.cell_value(n-2,0)
    for i in range(2,n-3):
        code=sh.cell_value(i,colk['code']).upper()
        ivid=sh.cell_value(i,colk['ivid'])
        if invfound.has_key(ivid):
            invfound[ivid]+=float(sh.cell_value(i,colk['nums']))*float(sh.cell_value(i,colk['price']))*codeclass[code].Units*fun_rate
        else:
            invfound[ivid]=float(sh.cell_value(i,colk['nums']))*float(sh.cell_value(i,colk['price']))*codeclass[code].Units*fun_rate
            name[ivid]=sh.cell_value(i,colname.index(u'投资者名称'))
    outdata=[]
    for x in invfound:
        outdata.append([x,name[x],round(invfound[x],2)])
    outcol=[u'投资者代码',u'投资者名称','投资者保障基金']    
    return outcol,outdata
def out_inv(invclass,tree,path):
    ''' 函数说明：输出表格'''
    import xlwt
    res=tkMessageBox.askquestion(u"请选择", u"强平客户（yes)还是全部客户(no)？")
    ck=0
    if res=='yes':
        ck=1    
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet(u'客户资金信息')
    i=0
    colname=tree['columns']
    for j in range(len(colname)):
        worksheet.write(i, j, label =english_to_ch(colname[j]))
    i+=1
    l1=[]
    l2=[]     
    for x in invclass:
        if invclass[x].InvInf['capital']<0:
            l1.append((invclass[x].InvInf['capital'],x))
        else:
            if ck:
                if invclass[x].InvInf['riskdegree']>=100:
                    l2.append((invclass[x].InvInf['riskdegree'],x))
            else:
                l2.append((invclass[x].InvInf['riskdegree'],x))
    l1.sort()#从小到大排序
    l2.sort(reverse=True)#从大到小排序         
    for x in l1:
        ivid=x[1]
        for j in range(len(colname)):
            worksheet.write(i, j, label = invclass[ivid].InvInf[colname[j]])
        i+=1
    for x in l2:
        ivid=x[1]
        for j in range(len(colname)):
            worksheet.write(i, j, label = invclass[ivid].InvInf[colname[j]])
        i+=1        
    now=datetime.datetime.now()
    nowf=now.strftime('%Y%m%d')
    outfilename=path+'\\'+nowf+u'试算结果.xls'
    workbook.save(outfilename)    
def splitout_inv(invclass,tree,path,datapath):
    ''' 函数说明：输出表格'''
    import xlwt  
    import codecs
    import sys
    whitelists=['13500886']
    res=tkMessageBox.askquestion(u"请选择", u"是否添加持仓监控数据？")
    posmon={}
    if res=='yes':
        posmon=fn2.get_monitor_pos()
    res=tkMessageBox.askquestion(u"请选择", u"导出哪个可用小于0，交易所线（yes)还是公司线(no)？")
    lcp='leftcapital'
    rk='riskdegree'
    if res=='no':
        lcp='cor_leftcapital'   
        rk='cor_riskdegree'
    risk_f=tkSimpleDialog.askfloat(u'华泰期货',u'交易所风险度高于多少进行标识',initialvalue =180)
    cp_f=tkSimpleDialog.askfloat(u'华泰期货',u'交易所强平金额高于多少进行标识',initialvalue =500000)    
    cor_risk_f=tkSimpleDialog.askfloat(u'华泰期货',u'公司风险度高于多少进行标识',initialvalue =180)
    cor_cp_f=tkSimpleDialog.askfloat(u'华泰期货',u'公司强平金额高于多少进行标识',initialvalue =500000)     
    phonelist=[]
    unphlist=[]
    whtext=u'以下为试算中含有的白名单：\r\n'
    now=datetime.datetime.now()
    nowf=now.strftime('%Y%m%d')    
    path=path+'\\'+nowf+u'试算结果'
    if not os.path.exists(path):
        os.makedirs(path)
    odic={}
    fnums=0
    for x in invclass:
        if invclass[x].InvInf[lcp]<0:
            fnums+=1
            if odic.has_key(invclass[x].InvInf['invdepartment']):
                odic[invclass[x].InvInf['invdepartment']].append((invclass[x].InvInf[rk],x))
            else:
                odic[invclass[x].InvInf['invdepartment']]=[(invclass[x].InvInf[rk],x)]
    alldep=odic.keys()
    for x in posmon:
        alldep+=posmon[x].keys()
    alldep=set(alldep)
    colname=tree['columns']   
    font = xlwt.Font()
    font.colour_index=2  
    style2 = xlwt.XFStyle()
    style2.font=font
    style = xlwt.XFStyle()
    for x in alldep:
        i=0
        fcp=0
        fn=0
        workbook = xlwt.Workbook(encoding = 'ascii')
        worksheet = workbook.add_sheet(u'客户资金信息')
        if odic.has_key(x):
            #插入试算数据
            fn=len(odic[x])
            for j in range(len(colname)):
                worksheet.write(i, j, label =english_to_ch(colname[j]))
            i+=1
            odic[x].sort(reverse=True)#从大到小排序         
            for y in odic[x]:
                ivid=y[1]
                fcp+=invclass[ivid].InvInf[lcp]
                ost=style
                if cp_f<>None and invclass[ivid].InvInf['leftcapital']<=-cp_f:
                    ost=style2
                if risk_f<>None and invclass[ivid].InvInf['riskdegree']>=risk_f:
                    ost=style2
                if cor_cp_f<>None and invclass[ivid].InvInf['cor_leftcapital']<=-cor_cp_f:
                    ost=style2
                if cor_risk_f<>None and invclass[ivid].InvInf['cor_riskdegree']>=cor_risk_f:
                    ost=style2                
                phone=invclass[ivid].InvInf['phone']            
                if len(phone)<>11 or phone[0]<>'1':
                    ost=style2
                    unphlist.append(ivid+'-'+invclass[ivid].Name+': '+phone+' ,'+invclass[ivid].InvInf['invdepartment'])               
                elif ivid in whitelists:
                    whtext+=ivid+'-'+invclass[ivid].Name+u'-手机-'+phone+'\r\n'
                else: 
                    phonelist.append(phone)
                for j in range(len(colname)):
                    worksheet.write(i, j, label = invclass[ivid].InvInf[colname[j]],style=ost)
                i+=1
            
        #插入持仓监控数据
        for mon_tbname in posmon:
            if posmon[mon_tbname].has_key(x):
                worksheet.write(i, 0, label = mon_tbname)
                i+=1
                n=len(posmon[mon_tbname][x][0])
                for j in range(n):
                    worksheet.write(i, j, label = posmon[mon_tbname][x][0][j])
                i+=1
                for mon_data in posmon[mon_tbname][x][1]:
                    for j in range(n):
                        worksheet.write(i, j, label = mon_data[j])
                    i+=1
        #插入持仓监控数据
                    
        outfilename=path+'\\'+x+nowf+u'试算结果 '+u'强平金额'+float_to_str(round(fcp,2))+u' 人数'+str(fn)+'.xls '
        workbook.save(outfilename)
    tphone=''
    for x in phonelist:    
        tphone+=','+x
    tphone=tphone[1:]
    f=codecs.open(path+'\\information.txt','w','utf-8')
    f.write(tphone+'\r\n\r\n')
    f.write(u'强平人数:'+str(fnums)+u'   其中可发出人数：'+str(len(phonelist))+u'    手机号不正确人数：'+str(len(unphlist))+'\r\n\r\n')
    f.write(u'以为手机错误客户:'+'\r\n')
    for x in unphlist:
        f.write(x+'\r\n')
    f.write(whtext)
    f.close()
    reload(sys)
    sys.setdefaultencoding('utf8')
    txtname='notepad '+datapath+'\\information.txt'        
    #os.system(txtname.encode('cp936'))    
    a=os.popen(txtname.encode('cp936'))    
    return     
def get_lastmonth():
    ''' 函数说明：得出交割月与临近交割月，例子本月1608，返回1608,1609,1610'''
    date1=datetime.datetime.now()
    date2=date1-datetime.timedelta(days=date1.day-1)+datetime.timedelta(days=31)
    date3=datetime.datetime(date2.year,date2.month,1)+datetime.timedelta(days=31)
    s1=datetime.datetime.strftime(date1,"%Y%m%d")[2:6]
    s2=datetime.datetime.strftime(date2,"%Y%m%d")[2:6]
    s3=datetime.datetime.strftime(date3,"%Y%m%d")[2:6]
    return s1,s2,s3
def recover_data(ini_codeclass,invclass):
    ''' 函数说明：行情数据修复，缺失合约补全以及没有价格合约补全'''
    from WindPy import w
    poscode=[]
    for x in invclass:
        for y in invclass[x].Position:
            if len(y['code'])<=6 and y['code'] not in poscode:
                poscode.append(y['code'])
    w.start()
    windname={'rt_high':'high','rt_low':'low','rt_latest':'price_now','rt_vwap':'meanprice','rt_oi':'open_interest','rt_vol':'volume','rt_bid1':'bid','rt_ask1':'ask'}
    colname=windname.keys()    
    rule=re.compile(r'[^a-zA-z]')
    qcode=''
    for code in poscode:
        iswind=False
        if code not in ini_codeclass.keys():
            iswind=True
            vari=rule.sub('',code)
            ini_codeclass[code]=FutureClass(vari,code)
        elif ini_codeclass[code].Inf['meanprice']==0:
            iswind=True
        if iswind:
            extcode=code+'.'+ini_codeclass[code].House
            windreulst=w.wsq(extcode,colname)
            for i,x in enumerate(colname):
                ini_codeclass[code].Inf[windname[x]]=windreulst.Data[i][0]
            qcode=qcode+code+','
    for x in ini_codeclass.keys():
        if ini_codeclass[x].House=='CFE' and x not in poscode:
            ini_codeclass.pop(x)
    if qcode<>'':
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'修复了'+qcode+u'合约行情\n注意保证金率表是否没更新，或者实时风控行情表没有添加新上市合约\n每导一次三表都要按此按钮')
    else:
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'没有合约需要修复')
def mon_ple_client(invclass):
    ''' 函数说明：质押客户监控'''
    now=datetime.datetime.now().strftime('%Y-%m-%d')        
    colname=['invid','inv_name','invdepartment','lastriskstate','mortgagemoney','capital']   
    outdata=[]
    sigama=0.035
    lever=15
    for x in invclass:
        if invclass[x].InvInf['mortgagemoney']>0:
            data=[now]
            for y in colname:
                data.append(invclass[x].InvInf[y])
            data.append(invclass[x].InvInf['capital']-1.25*invclass[x].InvInf['mortgagemoney'])
            isshfe=False
            for pos in invclass[x].Position:
                if pos['house']=='SHFE':
                    isshfe=True
                    break
            if isshfe:
                data.append(u'是')
            else:
                data.append(u'否')
            dy=invclass[x].InvInf['capital']-invclass[x].InvInf['mortgagemoney']-(invclass[x].InvInf['margin']*lever*sigama)
            data.append(dy)            
            if invclass[x].InvInf['capital']-invclass[x].InvInf['mortgagemoney']<=0:
                data.append(u'黑色：权益少于质押金额！！！')
                outdata.append(data)
            elif invclass[x].InvInf['capital']-1.25*invclass[x].InvInf['mortgagemoney']<=0:
                data.append(u'追加：权益少于1.25倍质押金额')
                outdata.append(data)                
            elif dy<=0:
                data.append(u'预警：现金承压能力不足')
                outdata.append(data)
            #else:
            #    data.append(u'正常')
    colname=['tr_date']+colname
    colname1=colname+['cp_minus_125mor','SHFE_pos','pro_loss','degree']        
    colname+=[u'权益-1.25倍质押',u'是否含上海持仓',u'现金承压提示',u'质押提示']
    for i in range(len(colname)):
        colname[i]=english_to_ch(colname[i])
    return colname,outdata,colname1
def out_left_client(invclass,invclass2):
    ''' 函数说明：可出大于可用客户监控'''
    colname=['invid','inv_name','invdepartment','lastriskstate','riskdegree','leftcapital']   
    outdata=[]    
    for x in invclass2:
        if invclass[x].InvInf['rr_bankmoney']>0 and invclass[x].InvInf['rr_bankmoney']>invclass2[x].InvInf['leftcapital']:
            data=[]
            for y in colname:
                data.append(invclass2[x].InvInf[y])
            data.append(invclass[x].InvInf['rr_bankmoney'])
            data.append(invclass[x].InvInf['rr_bankmoney']-invclass2[x].InvInf['leftcapital'])
            r=round(invclass2[x].InvInf['margin']/(invclass2[x].InvInf['capital']-invclass[x].InvInf['rr_bankmoney'])*100,2)
            data.append(r)
            outdata.append(data)
    colname+=['rr_bankmoney',u'出金后强平金额',u'出金后风险度']
    return colname,outdata
                
    
    
    
    
    
    
    
    