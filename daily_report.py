# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 23:52:29 2016

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
import functions_2 as fn2
import re
import xlrd
from operator import itemgetter, attrgetter
from win32com import client
import pickle
def delivery_mon(invclass):
    '''临近交割月持仓监控'''
    s1,s2,s3=fn.get_lastmonth()
    colname=[u'交易日',u'投资者代码',u'投资者名称',u'营业部名称',u'交易所代码',u'合约代码',u'多头持仓',u'空头持仓']
    colname1=['tr_date','invid','inv_name','invdepartment','house','code','longnums','shortnums']
    outdata=[]    
    now=datetime.datetime.now()
    isremind=(fn.cal_tradeday(now,get_day(s2,0,-1))<=10)
    for x in invclass:
        if invclass[x].InvInf['type']<>u'自然人':
            continue
        for pos in invclass[x].Position:
            data=[]
            if len(pos['code'])<=6:
                if pos['house'] not in ['CFFEX','SHFE'] and pos['code'][-3:] in [s2[1:]] and isremind:
                    data=[now.strftime('%Y-%m-%d'),x,invclass[x].Name,invclass[x].InvInf['invdepartment'],pos['house'],pos['code'],pos['longnums'],pos['shortnums']]
                    outdata.append(data)
                elif pos['house']=='SHFE' and pos['code'][-3:] in [s1[1:]]:
                    data=[now.strftime('%Y-%m-%d'),x,invclass[x].Name,invclass[x].InvInf['invdepartment'],pos['house'],pos['code'],pos['longnums'],pos['shortnums']]
                    outdata.append(data)                    
    outdata.sort(key=itemgetter(2))#,reverse=True)
    return colname,outdata,colname1
def client_pos_mon(invclass,ini_codeclass,datapath):
    '''公司客户超仓监控'''
    now=datetime.datetime.now()
    now_m=now.month
    ft=datetime.datetime(now.year,now.month,1)
    now_d=fn.cal_tradeday(ft,now)
    dp=datapath+u'客户超仓参数.xls'
    book = xlrd.open_workbook(dp)
    sh=book.sheets()[0]
    n=sh.nrows
    vcol={}
    for i in range(n):
        if sh.cell_value(i,0)<>u'品种':
            vcol[sh.cell_value(i,0).upper()]=i
    outdata=[]
    colname=[u'交易日',u'投资者代码',u'投资者名称',u'营业部名称',u'交易所代码',u'合约代码',u'投保',u'多头持仓',u'空头持仓',u'持仓上限']
    colname1=['tr_date','invid','inv_name','invdepartment','house','code','posident','longnums','shortnums','pos_limit']
    dtpar,gppar=client_pos_mon_par()
    for x in invclass:
        for pos in invclass[x].Position:
            if len(pos['code'])<=6 and pos['posident']<>u'套保':
                is_exceed=False
                if pos['vari'] in dtpar['gp1']:
                    par=[pos['vari'],pos['code'],pos['longnums'],pos['shortnums'],ini_codeclass[pos['code']].Inf['open_interest']]
                    par.append(sh.cell_value(vcol[pos['vari']],1))
                    par.append(sh.cell_value(vcol[pos['vari']],2))
                    par.append(sh.cell_value(vcol[pos['vari']],4))
                    par.append(sh.cell_value(vcol[pos['vari']],7))
                    is_exceed,limit=client_pos_1(par)
                if pos['vari'] in dtpar['gp2']:
                    par=[pos['vari'],pos['code'],pos['longnums'],pos['shortnums'],ini_codeclass[pos['code']].House]
                    par.append(sh.cell_value(vcol[pos['vari']],1))
                    par.append(sh.cell_value(vcol[pos['vari']],4))
                    par.append(sh.cell_value(vcol[pos['vari']],7))
                    is_exceed,limit=client_pos_2(par)
                if pos['vari'] in dtpar['gp3']:
                    par=[pos['vari'],pos['code'],pos['longnums'],pos['shortnums'],ini_codeclass[pos['code']].Inf['open_interest']]
                    par.append(sh.cell_value(vcol[pos['vari']],1))
                    par.append(sh.cell_value(vcol[pos['vari']],2))
                    par.append(sh.cell_value(vcol[pos['vari']],3))
                    par.append(sh.cell_value(vcol[pos['vari']],4))
                    par.append(sh.cell_value(vcol[pos['vari']],7))
                    is_exceed,limit=client_pos_3(par)     
                if pos['vari'] in dtpar['gp4']:
                    par=[pos['vari'],pos['code'],pos['longnums'],pos['shortnums'],ini_codeclass[pos['code']].House]
                    par.append(sh.cell_value(vcol[pos['vari']],1))
                    par.append(sh.cell_value(vcol[pos['vari']],4))
                    par.append(sh.cell_value(vcol[pos['vari']],5))
                    par.append(sh.cell_value(vcol[pos['vari']],7))
                    is_exceed,limit=client_pos_4(par) 
                if is_exceed:
                    data=[now.strftime('%Y-%m-%d'),x,invclass[x].Name,invclass[x].InvInf['invdepartment'],pos['house'],pos['code'],pos['posident'],pos['longnums'],pos['shortnums'],limit]
                    outdata.append(data)
    outdata.sort(key=itemgetter(2))
    return colname,outdata,colname1        
def client_pos_1(par):
    '''gp1的超仓监控'''
    #品种，合约，多头，空头，市场持仓量，限额标准，限额百分比，限额1，限额2
    rule=re.compile(r'[^0-9]')
    dmon=rule.sub('',par[1])
    now=datetime.datetime.now()
    l1,tt=client_pos_mon_par()
    l1=l1['l1']
    nT=8#提前多少个交易日
    dl1=get_day(dmon,l1[0][0],l1[0][1])#限额日期1
    dl2=get_day(dmon,l1[1][0],l1[1][1])#限额日期2
    is_exceed=False   
    if fn.cal_tradeday(now,dl2)<=nT:
        if par[2]>par[-1] or par[3]>par[-1]:
            is_exceed=True
            return is_exceed,par[-1]
    if fn.cal_tradeday(now,dl1)<=nT:
        if par[2]>par[-2] or par[3]>par[-2]:    
            is_exceed=True
            return is_exceed,par[-2]
    if par[4]>par[5] and (par[2]>par[4]*par[6] or par[3]>par[4]*par[6]):
        is_exceed=True
        return is_exceed,par[5]*par[6]
    return is_exceed,0                           
def client_pos_2(par):
    '''gp2的超仓监控'''
    #品种，合约，多头，空头，交易所，限额1，限额2，限额3
    nT=8#提前多少个交易日
    rule=re.compile(r'[^0-9]')
    dmon=rule.sub('',par[1])
    now=datetime.datetime.now()
    dtpar,gppar=client_pos_mon_par()
    tl=False
    if par[0] in gppar['gp2_1']:
        tl=dtpar['l1']
    elif par[0] in gppar['gp2_2']:
        tl=dtpar['l2']
    elif par[0] in gppar['gp2_3']:
        tl=dtpar['l3']
    if tl:
        dl1=get_day(dmon,tl[0][0],tl[0][1])#限额日期1
        dl2=get_day(dmon,tl[1][0],tl[1][1])#限额日期2
    elif par[4] in gppar['gp2_4']:
        dl1=get_day(dmon,-1,1)
        dl1=datetime.datetime(dl1.year,dl1.month,16)
        dl2=get_day(dmon,0,1)
    else:
        dl1=get_day(dmon,-1,1)
        dl1=datetime.datetime(dl1.year,dl1.month,21)
        dl2=get_day(dmon,0,1)        
    is_exceed=False   
    if fn.cal_tradeday(now,dl2)<=nT:
        if par[2]>par[-1] or par[3]>par[-1]:
            is_exceed=True
            return is_exceed,par[-1]
    if fn.cal_tradeday(now,dl1)<=nT:
        if par[2]>par[-2] or par[3]>par[-2]:    
            is_exceed=True
            return is_exceed,par[-2]
    if par[2]>par[-3] or par[3]>par[-3]:    
        is_exceed=True
        return is_exceed,par[-3]
    return is_exceed,0   
def client_pos_3(par):
    '''gp3的超仓监控'''
    #品种，合约，多头，空头，市场持仓量，限额标准，限额百分比，限额1，限额2，限额3
    rule=re.compile(r'[^0-9]')
    dmon=rule.sub('',par[1])
    now=datetime.datetime.now()
    l1,tt=client_pos_mon_par()
    l1=l1['l3']
    nT=8#提前多少个交易日
    dl1=get_day(dmon,l1[0][0],l1[0][1])#限额日期1
    dl2=get_day(dmon,l1[1][0],l1[1][1])#限额日期2
    is_exceed=False   
    if fn.cal_tradeday(now,dl2)<=nT:
        if par[2]>par[-1] or par[3]>par[-1]:
            is_exceed=True
            return is_exceed,par[-1]
    if fn.cal_tradeday(now,dl1)<=nT:
        if par[2]>par[-2] or par[3]>par[-2]:    
            is_exceed=True
            return is_exceed,par[-2]
    if par[4]>par[5]:
        if par[2]>par[4]*par[6] or par[3]>par[4]*par[6]:
            is_exceed=True
            return is_exceed,par[4]*par[6]
    elif par[2]>par[-3] or par[3]>par[-3]:
        is_exceed=True
        return is_exceed,par[-3]        
    return is_exceed,0   
def client_pos_4(par):
    '''gp4的超仓监控'''
    #品种，合约，多头，空头，交易所，限额1，限额2，限额3，限额4
    rule=re.compile(r'[^0-9]')
    dmon=rule.sub('',par[1])
    now=datetime.datetime.now()
    l1,tt=client_pos_mon_par()
    l1=l1['l4']
    nT=8#提前多少个交易日
    dl1=get_day(dmon,l1[0][0],l1[0][1])#限额日期1
    dl2=get_day(dmon,l1[1][0],l1[1][1])#限额日期2
    dl3=get_day(dmon,l1[2][0],l1[2][1])#限额日期2
    is_exceed=False   
    if fn.cal_tradeday(now,dl3)<=nT:
        if par[2]>par[-1] or par[3]>par[-1]:
            is_exceed=True
            return is_exceed,par[-1]
    if fn.cal_tradeday(now,dl2)<=nT:
        if par[2]>par[-2] or par[3]>par[-2]:    
            is_exceed=True
            return is_exceed,par[-2]
    if fn.cal_tradeday(now,dl1)<=nT:
        if par[2]>par[-3] or par[3]>par[-3]:    
            is_exceed=True
            return is_exceed,par[-3]
    if par[2]>par[-4] or par[3]>par[-4]:    
        is_exceed=True
        return is_exceed,par[-4]        
    return is_exceed,0     
def client_pos_mon_par():
    '''公司客户超仓监控日期划分参数'''
    gp1=['CU','ZN','AL','RB','WR']
    gp2=['RU','PB','AU','AG','BU','HC','NI','SN','FU','J','JM','CZCE','CFFEX']
    gp3=['Y','L','B','A','M','V','C','I','P','FB','BB','PP','CS']
    gp4=['JD']
    gp2_1=['RU','PB','AU','AG','BU','HC','NI','SN']
    gp2_2=['FU']
    gp2_3=['J','JM']
    gp2_4=['CZCE']
    gp2_5=['CFFEX']
    l1=[(-1,1),(0,1)]
    l2=[(-2,1),(-1,1)]
    l3=[(-1,10),(0,1)]
    l4=[(-1,1),(-1,10),(0,1)]
    l5=[16,(0,1)]
    l6=[21,(0,1)]
    res={'gp1':gp1,'gp2':gp2,'gp3':gp3,'gp4':gp4,'l1':l1,'l2':l2,'l3':l3,'l4':l4,'l5':l5,'l6':l6}
    gp2res={'gp2_1':gp2_1,'gp2_2':gp2_2,'gp2_3':gp2_3,'gp2_4':gp2_4,'gp2_5':gp2_5}
    return res,gp2res
def get_day(dmon,emon,day):
    '''根据交割月前emon个月，第day交易日返回日期'''
    #0代表第一个交易日，-1代表最后一个交易日
    mon=int(dmon[-2:])
    y=int(dmon[:-2])
    oneday=datetime.timedelta(days=1)
    if y>10:
        year=2000+y
    else:
        year=2010+y        
    if mon+emon<=0:
        year=year-1
        rmon=12+mon+emon
    elif mon+emon>12:
        year+=1
        rmon=mon+emon-12      
    else:
        rmon=mon+emon      
    if day==-1:
        temp=datetime.datetime(year,rmon,1)
        temp2=datetime.datetime(year,rmon,calendar.monthrange(year,rmon)[1])
        n=fn.cal_tradeday(temp,temp2)
        res=get_day(dmon,emon,n)
    else:
        res=datetime.datetime(year,rmon,1)
        temp=datetime.datetime(year,rmon,1)
        while fn.cal_tradeday(temp,res)<>day:
            res=res+oneday
    return res
def corpos_ratio(invclass,ini_codeclass):
    '''公司持仓占市场比'''
    corpos={}    
    for x in invclass:
        for pos in invclass[x].Position:
            if len(pos['code'])<=6:
                if corpos.has_key(pos['vari']):
                    corpos[pos['vari']][0]+=pos['longnums']
                    corpos[pos['vari']][1]+=pos['shortnums']
                else:
                    corpos[pos['vari']]=[pos['longnums'],pos['shortnums']]
    vkp={}
    for x in ini_codeclass:
        if vkp.has_key(ini_codeclass[x].Vari):
            if ini_codeclass[x].House=='CFE':
                vkp[ini_codeclass[x].Vari]+=ini_codeclass[x].Inf['open_interest']
            else:
                vkp[ini_codeclass[x].Vari]+=ini_codeclass[x].Inf['open_interest']/2
        else:
            if ini_codeclass[x].House=='CFE':
                vkp[ini_codeclass[x].Vari]=ini_codeclass[x].Inf['open_interest']
            else:
                vkp[ini_codeclass[x].Vari]=ini_codeclass[x].Inf['open_interest']/2
    lratio=[]
    sratio=[]                    
    for x in corpos:
        lratio.append((corpos[x][0]/vkp[x],x,corpos[x][0]))
        sratio.append((corpos[x][1]/vkp[x],x,corpos[x][1]))
    lratio.sort(reverse=True)
    sratio.sort(reverse=True)
    colname=[u'交易日',u'多头品种',u'多头持仓占比',u'多头持仓',u'空头品种',u'空头持仓占比',u'空头持仓']
    colname1=['tr_date','longvari','longratio','longnums','shortvari','shortratio','shortnums']
    outdata=[]
    now=datetime.datetime.now().strftime('%Y-%m-%d')
    for i in range(len(lratio)):
        lr=round(lratio[i][0],4)#*100)+'%'
        sr=round(sratio[i][0],4)#*100)+'%'
        data=[now,lratio[i][1],lr,lratio[i][2],sratio[i][1],sr,sratio[i][2]] 
        outdata.append(data)
    return colname,outdata,colname1  
def monitor_cor_position(invclass,codeclass,datapath):
    '''公司持仓监控计算'''
    '''
    conn=conn_mysql()
    cursor=conn.cursor()
    sql='select * from corpos_monitor_coefficient'
    cursor.execute(sql)
    rs=cursor.fetchall()
    '''
    rs=[]
    dp=datapath+u'公司超仓参数.xls'
    book = xlrd.open_workbook(dp)
    sh=book.sheets()[0]
    n=sh.nrows 
    shcname=sh.row_values(0)
    for i in range(1,n):
        tt={}
        for j in range(len(shcname)):
            tt[shcname[j]]=sh.cell_value(i,j)
        rs.append(tt)
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
    now=datetime.datetime.now().strftime('%Y-%m-%d')
    colname=[u'交易日',u'持仓合约',u'会员持仓限额',u'会员多头持仓',u'会员空头持仓',u'多头持仓/限额',u'空头持仓/限额',u'多头风险等级',u'空头风险等级']
    colname1=['tr_date','code','pos_limit','longnums','shortnums','longratio','shortratio','longdegree','shortdegree']
    for x in vari_longpos:
        data=[]
        if market_pos[x]>=vdata[codeclass[x].Vari]['limit']:
            vl=math.floor(market_pos[x]/2*0.25*(1+vdata[codeclass[x].Vari]['bus_cof']+vdata[codeclass[x].Vari]['cre_cof']))
            if x in ['IF','IH','IC','TF','T']:
                vl=math.floor(market_pos[x]*0.25*(1+vdata[codeclass[x].Vari]['bus_cof']+vdata[codeclass[x].Vari]['cre_cof']))
            r1,r2=vari_longpos[x]/vl,vari_shortpos[x]/vl           
            sr1=round(r1,4)
            sr2=round(r2,4)
            data=[now,x,vl,vari_longpos[x],vari_shortpos[x],sr1,sr2]
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
            #if data[-2]<>u'没有风险' or data[-1]<>u'没有风险':
            outdata.append(data)
    return colname,outdata,colname1
def unactive_pos_monitor(invclass,ini_codeclass,invclass2):
    '''不活跃持仓监控'''
    tn=500#成交量
    oi=500#持仓量
    uncode={}
    colname=[u'交易日',u'投资者代码',u'投资者名称',u'营业部名称',u'合约代码',u'投保',u'多头',u'空头',u'占市场持仓比例',u'风险等级']
    colname1=['tr_date','invid','inv_name','invdepartment','code','posident','longnums','shortnums','posratio','degree']    
    for x in ini_codeclass:
        if  ini_codeclass[x].Inf['open_interest']<=oi or  ini_codeclass[x].Inf['volume']<=tn:
            toi=0
            oinp=0            
            for y in ini_codeclass:
                if ini_codeclass[y].Vari==ini_codeclass[x].Vari:
                    toi+=ini_codeclass[y].Inf['volume']
                    oinp+=ini_codeclass[y].Inf['volume']*ini_codeclass[y].Inf['settlement']
            if toi>0:
                uncode[x]=oinp/toi
            else:
                uncode[x]=ini_codeclass[x].Inf['settlement']
    allixda,alldata=[],[]
    now=datetime.datetime.now().strftime('%Y-%m-%d')
    for x in invclass:
        for pos in invclass[x].Position:
            if pos['code'] in uncode:
                data=[now,x,invclass[x].Name,invclass[x].InvInf['invdepartment'],pos['code'],pos['posident'],pos['longnums'],pos['shortnums']]
                ixda={'ivid':x,'code':pos['code']}
                ixda['mknums']=min(ini_codeclass[pos['code']].Inf['open_interest'],ini_codeclass[pos['code']].Inf['volume'])#成家量和持仓量取小
                codenums=pos['longnums']-pos['shortnums']
                setl=ini_codeclass[pos['code']].Inf['settlement']
                if codenums==0:
                    ixda['outrange']=100
                else:
                    obound=setl-invclass2[x].InvInf['capital']/(codenums*ini_codeclass[pos['code']].Units)
                    ixda['outrange']=obound/setl-1#穿仓幅度
                ixda['riskdegree']=invclass2[x].InvInf['riskdegree']#交易所风险度
                y=(pos['longnums']*(setl-pos['longopenprice'])+pos['shortnums']*(pos['shortopenprice']-setl))*ini_codeclass[pos['code']].Units
                ixda['lossratio']=y/invclass2[x].InvInf['margin']#亏损占保证金比例
                ixda['deviation']=setl/uncode[pos['code']]-1
                if ini_codeclass[pos['code']].Inf['volume']==0:
                    tvol=1
                else:
                    tvol=ini_codeclass[pos['code']].Inf['volume']
                if ini_codeclass[pos['code']].House=='CFE':
                    ixda['oiratio']=max(pos['longnums'],pos['shortnums'])/ini_codeclass[pos['code']].Inf['open_interest']
                    ixda['tnratio']=abs(codenums)/tvol
                else:
                    ixda['oiratio']=max(pos['longnums'],pos['shortnums'])/(ini_codeclass[pos['code']].Inf['open_interest']/2)#占市场持仓比例
                    ixda['tnratio']=abs(codenums)/(tvol/2)#占市场成交比例
                data.append(ixda['oiratio'])
                allixda.append(ixda)
                alldata.append(data)
    outdata=[]
    for i,ixda in enumerate(allixda):
        points=0
        if ixda['mknums']<=100:
            points+=20
        if ixda['outrange']<=0.05:
            points+=15
        elif ixda['outrange']<=0.1:
            points+=10
        if ixda['riskdegree']>=75:
            points+=15
        elif ixda['riskdegree']>=50:
            points+=10
        if -ixda['lossratio']>=1:
            points+=5
        if abs(ixda['deviation'])>=0.2:
            points+=20
        elif abs(ixda['deviation'])>=0.1:
            points+=10
        if ixda['oiratio']>=0.1:
            points+=15
        elif ixda['oiratio']>=0.05:
            points+=10
        if ixda['tnratio']>=1:
            points+=15
        elif ixda['tnratio']>=0.5:
            points+=10
        if points>=50:
            alldata[i].append(u'极度')
        elif points>=35:
            alldata[i].append(u'高度')
        elif points>=25:
            alldata[i].append(u'中度')
        elif points>=15:
            alldata[i].append(u'低度')
        else:
            alldata[i].append(u'轻度')
        if alldata[i][-1] in [u'极度',u'高度',u'中度'] and alldata[i][1] not in ['20023766','10701018','10701012','10701015','10701206']:
            outdata.append(alldata[i])
    return colname,outdata,colname1  
def major_pos_monitor(invclass,ini_codeclass,invclass2,datapath,variclass): 
    '''重大持仓监控'''
    vidx={}#连续三天涨跌幅
    for x in ini_codeclass:
        if not vidx.has_key(ini_codeclass[x].Vari):
            vidx[ini_codeclass[x].Vari]=0
    conn=MySQLdb.Connect(host='10.100.7.52',port=3306,user='fxgl',passwd='fxgl',db='ht_risk', charset='utf8',cursorclass = MySQLdb.cursors.DictCursor)
    cursor=conn.cursor()
    for x in vidx:
        sql='select * from maincode where vari=\'%s\' order by tr_date desc limit 3' %x
        cursor.execute(sql)
        rs=cursor.fetchall()
        '''
        if rs[0]['tr_date'].strftime('%Y-%m-%d')<>datetime.datetime.now().strftime('%Y-%m-%d'):
            print x
            print u'重大持仓制表失败，请检查是否具有最新行情'
            cursor.close()
            conn.close()
            return [],[],[]
        '''
        for mf in rs:
            #print x,mf['cls'],mf['pre_cls']
            if not(mf['pre_cls']==None or  mf['pre_cls']==0 or mf['cls']==None):
                vidx[x]+=mf['cls']/mf['pre_cls']-1
    cursor.close()
    conn.close()
    
    
    bnums={}#重仓手数
    fname=datapath+u'重大持仓参数.xls'
    book = xlrd.open_workbook(fname)
    sh=book.sheets()[0]
    n=sh.nrows 
    for i in range(1,n):
        bnums[sh.cell_value(i,0).upper()]=sh.cell_value(i,1)   
        
    colname=[u'交易日',u'投资者代码',u'投资者名称',u'营业部名称',u'交易所',u'产品代码',u'投机套保标志',u'净持仓'
    ,u'多头持仓',u'空头持仓',u'浮动盈亏',u'当日权益',u'交易所保证金',u'交易所风险度',u'公司风险度'
    ,u'重仓占比',u'重大亏损指标',u'品种指数涨跌幅',u'风险级别']    
    colname1=[u'tr_date','invid','inv_name','invdepartment','house','vari','posident','posnums','longnums','shortnums'
    ,'pnl','capital','margin','riskdegree','cor_riskdegree','mratio','cpratio','variyield','degree']
    alldata=[]
    now=datetime.datetime.now().strftime('%Y-%m-%d')
    for x in invclass:
        varipos={}
        for pos in invclass[x].Position:
            if len(pos['code'])>=6:
                continue
            ishaskey(varipos,pos['posident'],pos['vari'])
            varipos[pos['posident']][pos['vari']][0]+=pos['longnums']
            varipos[pos['posident']][pos['vari']][1]+=pos['shortnums']
        for posi in varipos:
            for vari in varipos[posi]:
                if abs(varipos[posi][vari][0]-varipos[posi][vari][1])>=bnums[vari]:
                    data={}
                    ml,ms=0,0
                    cp=0
                    for pos in invclass[x].Position:
                        if pos['vari']==vari and pos['posident']==posi:
                            tcode=ini_codeclass[pos['code']]
                            ml+=pos['longnums']*(tcode.Inf['Mrate']+tcode.Inf['delta_rate'])*tcode.Units*tcode.Inf['settlement']
                            ms+=pos['shortnums']*(tcode.Inf['Mrate']+tcode.Inf['delta_rate'])*tcode.Units*tcode.Inf['settlement']
                            cp+=(pos['longnums']*(tcode.Inf['settlement']-pos['longopenprice'])+pos['shortnums']*(pos['shortopenprice']-tcode.Inf['settlement']))*tcode.Units
                    data['tr_date']=now
                    data['invid']=x
                    data['inv_name']=invclass[x].Name
                    data['invdepartment']=invclass[x].InvInf['invdepartment']
                    data['house']=variclass[vari].House
                    data['vari']=vari
                    data['posident']=posi
                    data['posnums']=varipos[posi][vari][0]-varipos[posi][vari][1]
                    data['longnums']=varipos[posi][vari][0]
                    data['shortnums']=varipos[posi][vari][1]
                    data['pnl']=cp
                    data['capital']=invclass2[x].InvInf['capital']
                    data['margin']=invclass2[x].InvInf['margin']
                    data['riskdegree']=invclass2[x].InvInf['riskdegree']
                    data['cor_riskdegree']=invclass2[x].InvInf['cor_riskdegree']
                    data['mratio']=max(ml,ms)/data['margin']
                    data['cpratio']=cp/data['margin']
                    data['variyield']=vidx[vari]
                    alldata.append(data)
    effdata=[]
    for data in alldata:
        judge_major(data,variclass)
        if data['degree'] in [u'极度',u'高度']:
            tpdata=[]
            for cn in colname1:
                tpdata.append(data[cn])
            effdata.append(tpdata)
    return colname,effdata,colname1
def judge_major(data,variclass):
    idx=[0,0,0]
    r1,r2,r3=50,30,70
    if data['mratio']>=0.5:
        idx[0]=1
    if data['cpratio']<=-0.5:
        idx[1]=1
    if abs(data['variyield'])>=1.5*variclass[data['vari']].Inf['Limt'] and data['variyield']*data['posnums']<0:
        idx[2]=1
    f1=(idx[0] and (idx[1] or idx[2]) and data['riskdegree']>=r1)
    f2=(idx[0] and not idx[1] and not idx[2] and data['riskdegree']>=r3)
    f3=(not idx[0] and idx[1] and idx[2] and data['riskdegree']>=r1)
    if f1 or f2 or f3:
        data['degree']=u'极度'
        return
    f4=(idx[0] and idx[1] and idx[2] and data['riskdegree']<r1 and data['riskdegree']>=r2)
    f5=(idx[0] and not idx[1] and not idx[2] and data['riskdegree']>=r1)
    f6=(not idx[0] and idx[1] and not idx[2] and data['riskdegree']>=r1)
    f7=(not idx[0] and not idx[1] and idx[2] and data['riskdegree']>=r1)
    if f4 or f5 or f6 or f7:
        data['degree']=u'高度'
        return
    f8=(idx[0] and idx[1] and not idx[2] and data['riskdegree']<r1 and data['riskdegree']>=r2)
    f9=(idx[0] and not idx[1] and idx[2] and data['riskdegree']<r1 and data['riskdegree']>=r2)  
    f10=(not idx[0] and idx[1] and idx[2] and data['riskdegree']<r1 and data['riskdegree']>=r2) 
    f11=(not idx[0] and not idx[1] and not idx[2] and data['riskdegree']>=r1)
    if f8 or f9 or f10 or f11:
        data['degree']=u'中度'
        return     
    f12=(idx[0] and not idx[1] and not idx[2] and data['riskdegree']<r1 and data['riskdegree']>=r2)
    f13=(idx[0] and (not idx[1] or not idx[2]) and data['riskdegree']<r1 and data['riskdegree']>=r2) 
    if f12 or f13:
        data['degree']=u'低度'
        return
    data['degree']=u'轻度'
    return
def ishaskey(varipos,posident,vari):
    if varipos.has_key(posident):
        if varipos[posident].has_key(vari):
            return
        else:
            varipos[posident][vari]=[0,0]
            return
    else:
        varipos[posident]={}
        varipos[posident][vari]=[0,0]
        return
def cTable(doc,docrange,data,colname):
    '''创建word表格'''
    rows=len(data)
    cols=len(colname)
    tab=doc.Tables.Add(doc.Range(docrange.End,docrange.End),rows+1,cols)
    tab.Style=u'网格型'
    for i in range(cols):
        tab.Cell(1,i+1).Range.Text=colname[i]
    for i in range(rows):
        for j in range(cols):
            tab.Cell(i+2,j+1).Range.Text=data[i][j]
    return tab.Range.End,tab.Range.End
def out_PDF(alldata,filepath):
    '''输出至pdf'''
    fname=datetime.datetime.now().strftime('%Y%m%d')+u'风险监控报告'
    word = client.Dispatch('Word.Application')
    newdoc = word.Documents.Add()
    r = newdoc.Range(0,0)
    r.Style.Font.Size='8'
    r.InsertAfter(fname+'\n')
    for colname,outdata,tbname,unname in alldata:
        r.InsertAfter('\n'+tbname+'\n')
        px,py=cTable(newdoc,r,outdata,colname)
        r = newdoc.Range(px,py)
    par1 = newdoc.Paragraphs(1).Range
    par1.Font.Size='18'
    par1.Font.Bold=True
    par1.ParagraphFormat.Alignment = 1
    #newdoc.Paragraphs(2).Range.Font.Bold=True
       
    ffname=filepath+'\\'+fname+u'.pdf'
    ffname2=filepath+'\\'+fname+u'.docx'
    #newdoc.SaveAs(ffname2)
    client.gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
    newdoc.ExportAsFixedFormat(ffname, client.constants.wdExportFormatPDF,   
      Item = client.constants.wdExportDocumentWithMarkup, 
      CreateBookmarks = client.constants.wdExportCreateHeadingBookmarks) 
    newdoc.Close(False)
    word.Quit()    
def out_excel(alldata,filepath):
    import xlwt
    now=datetime.datetime.now().strftime('%Y-%m-%d')
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet(now+u'持仓风险监控')
    i=0
    for data in alldata:
        worksheet.write(i, 0, label =data[2])
        print data[2]
        i+=1
        n=len(data[1])
        if n==0:
            continue
        cols=len(data[0])
        #print n,len(data[0])
        #print data[1]
        for j in range(cols):
            worksheet.write(i, j, label =data[0][j])
        i+=1
        for x in data[1]:
            for j in range(cols):
                worksheet.write(i, j, label =x[j])
            i+=1
        i+=1
    outfilename=filepath+'\\'+now+u'持仓风险监控.xls'
    workbook.save(outfilename) 
def create_report(data,datapath,filepath):
    '''生成每日监控报告'''
    alldata=[]
    
    ple_colname,ple_data,colname1=fn.mon_ple_client(data['invclass2'])#质押配比数据
    alldata.append((ple_colname,ple_data,u'质押配比监控',colname1))
    print u'质押配比监控','done',len(ple_data)

    delpos_colname,delpos_data,colname1=delivery_mon(data['invclass'])#临近交割月监控
    alldata.append((delpos_colname,delpos_data,u'临近交割月监控',colname1))
    print u'临近交割月监控','done',len(delpos_data)
    
    clipos_colname,clipos_data,colname1=client_pos_mon(data['invclass'],data['ini_codeclass'],datapath)#客户超仓监控
    alldata.append((clipos_colname,clipos_data,u'客户超仓监控',colname1))
    print u'客户超仓监控','done',len(clipos_data)
    
    cpr_colname,cpr_data,colname1=corpos_ratio(data['invclass'],data['ini_codeclass'])#公司持仓占市场
    alldata.append((cpr_colname,cpr_data,u'公司持仓占比市场',colname1))
    print u'公司持仓占比市场','done',len(cpr_data)

    monitor_cor_colname,monitor_cor_data,colname1=monitor_cor_position(data['invclass'],data['ini_codeclass'],datapath)#公司超仓监控
    alldata.append((monitor_cor_colname,monitor_cor_data,u'公司超仓监控',colname1))
    print u'公司超仓监控','done',len(monitor_cor_data)
    
    unactive_colname,unactive_data,colname1=unactive_pos_monitor(data['invclass'],data['ini_codeclass'],data['invclass2'])#不活跃持仓监控
    alldata.append((unactive_colname,unactive_data,u'不活跃持仓监控--不包含蔡雪萍、陈雪玲、郝全永、郝志红、关鸿伟',colname1))
    print u'不活跃持仓监控','done',len(unactive_data)
    
    maj_colname,maj_data,colname1=major_pos_monitor(data['invclass'],data['ini_codeclass'],data['invclass2'],datapath,data['variclass'])#重大持仓监控
    alldata.append((maj_colname,maj_data,u'重大持仓监控',colname1))
    print u'重大持仓监控','done',len(maj_data)   
    
    out_excel(alldata,filepath)
    
    #with open(datapath+'alldata.pickle', 'wb') as f:pickle.dump(alldata, f)
    
    #out_excel(alldata,filepath)
    #with open(datapath+'alldata.pickle', 'rb') as f:alldata = pickle.load(f)
    res=tkMessageBox.askquestion(u"各类持仓监控计算完成", u"是否把数据插入至数据库？")
    if res=='yes':
        storing_data(alldata)
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'数据库更新完成')
def storing_data(alldata):
    tr_day=datetime.datetime.now().strftime('%Y-%m-%d')    
    res={u'质押配比监控':'mon_ple_client',u'临近交割月监控':'delivery_mon_pos',u'客户超仓监控':'client_pos_mon'
    ,u'公司持仓占比市场':'corpos_ratio',u'公司超仓监控':'corpos_monitor_detail',u'重大持仓监控':'major_pos_monitor'
    ,u'不活跃持仓监控--不包含蔡雪萍、陈雪玲、郝全永、郝志红、关鸿伟':'unactive_pos_monitor'}    
    conn=MySQLdb.Connect(host='10.100.7.52',port=3306,user='fxgl',passwd='fxgl',db='ht_risk', charset='utf8',cursorclass = MySQLdb.cursors.DictCursor)
    cursor=conn.cursor() 
    for data in alldata:
        od={}
        od['colname1']=data[-1]
        od['colname2']=data[0]
        od['data']=data[1]
        od['tbname']=data[2]
        insert_into_sql(res[od['tbname']],od['colname1'],od['data'],conn,cursor)
    conn.commit()    
    cursor.close()
    conn.close()
def insert_colname(alldata):
    res={u'质押配比监控':'mon_ple_client',u'临近交割月监控':'delivery_mon_pos',u'客户超仓监控':'client_pos_mon'
    ,u'公司持仓占比市场':'corpos_ratio',u'公司超仓监控':'corpos_monitor_detail',u'重大持仓监控':'major_pos_monitor'
    ,u'不活跃持仓监控--不包含蔡雪萍、陈雪玲、郝全永、郝志红、关鸿伟':'unactive_pos_monitor'}
    for data in alldata:
        fn2.colname_to_sql(res[data[2]],data[-1],data[0])
    print u'插入列名至数据库成功'
def insert_into_sql(tbname,colname,data,conn,cursor):
   
    sql='desc %s' %tbname
    print sql
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
        try:
            cursor.execute(sql)
        except Exception as e:
            print tbname,'ERROR IN SQL execute'
            print sql
            conn.rollback()
            cursor.close()
            conn.close()
            raise e
    print u'数据库插入成功',tbname