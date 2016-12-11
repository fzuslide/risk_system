# -*- coding: utf-8 -*-
"""
Created on Wed Aug 10 21:57:59 2016

@author: IORI
"""
from __future__ import division
import functions as fn
import math
import datetime
import xlrd
import re
import tkSimpleDialog
from scipy import stats
import matplotlib.pyplot as plt
from classes import *
from option_class import *
from WindPy import w
import Tkinter as tk
tt=tk.Tk()
trade_date=tkSimpleDialog.askstring(u'华泰期货',u'请输入获取交易日期的收盘价',initialvalue ='20160921')
tk.Tk.destroy(tt)
def input_position():
    path=u'C:\\Users\\Administrator\\Desktop\\场外期权\\交易数据08.09\\'
    fname=path+u'PnL Explained Vanilla20160921.xlsm'
    book = xlrd.open_workbook(fname)
    sh=book.sheet_by_name('Calc')
    n=sh.nrows   
    for i in range(n):
        if sh.cell_value(i,2)==u'金牛数据':
            i+=2
            break
    fu_colname=sh.row_values(i)
    fu_pos=[]
    i+=1
    while sh.cell_value(i,2)<>'INSERT ROW':
        fu_pos.append(sh.row_values(i))
        i+=1
    while sh.cell_value(i,2)<>'NEW TRADES':
        i+=1
    i+=3
    while sh.cell_value(i,2)<>'INSERT ROW':
        fu_pos.append(sh.row_values(i))
        i+=1           
    fucol={}
    for x in fu_colname:
        fucol[x]=fu_colname.index(x)
        
    sh=book.sheet_by_name('Data')
    n=sh.nrows
    op_colname=sh.row_values(3)
    k=op_colname.index(u'CLOSED /\nEXPIRED')
    op_pos=[]
    for i in range(4,n):
        if sh.cell_value(i,k)=='N':
            op_pos.append(sh.row_values(i))
    opcol={}
    for x in op_colname:
        opcol[x]=op_colname.index(x)  
    
    sh=book.sheet_by_name('S')
    n=sh.nrows    
    codeprice={}
    lastcodeprice={}
    for i in range(6,n):
        codeprice[sh.cell_value(i,1)]=float(sh.cell_value(i,11))
        lastcodeprice[sh.cell_value(i,1)]=float(sh.cell_value(i,10))
    
    book = xlrd.open_workbook(path+u'Exotics_Pnl 20160921.xlsx')
    sh=book.sheet_by_name(u'奇异期权')
    n=sh.nrows  

    baop_colname=sh.row_values(1)
    baop_pos=[]
    k=baop_colname.index(u'EXPIRED')
    for i in range(2,n):
        if sh.cell_value(i,k)=='N':
            baop_pos.append(sh.row_values(i))

    baopcol={}
    for x in baop_colname:
        baopcol[x]=baop_colname.index(x)  
    AllData={}
    AllData['codeprice'],AllData['lastcodeprice']=codeprice,lastcodeprice
    AllData['opcol'],AllData['op_pos'],AllData['fucol'],AllData['fu_pos'],AllData['baopcol'],AllData['baop_pos']=opcol,op_pos,fucol,fu_pos,baopcol,baop_pos
    return AllData
def create_class():
    AllData=input_position()
    opcol,op_pos,fucol,fu_pos,baopcol,baop_pos=AllData['opcol'],AllData['op_pos'],AllData['fucol'],AllData['fu_pos'],AllData['baopcol'],AllData['baop_pos']
    alloption={}
    deltat=0#日期调整
    allcode=[]
    for x in op_pos:
        vari=x[opcol['GROUP']]
        if vari=='':
            continue        
        allcode.append(x[opcol['UNDERLYING']])
        if not alloption.has_key(vari):
            alloption[vari]=[]
        values={'r':0.03,'q':0.03}
        values['underlying']=x[opcol['UNDERLYING']]
        values['S']=x[opcol['PRICE\nTODAY']]
        values['K']=x[opcol['K (CNY)\nTODAY']]
        values['vol']=x[opcol['SIGMA\nTODAY']]
        values['type']=x[opcol['C / P']]
        expiry=xlrd.xldate.xldate_as_datetime(int(x[opcol['EXPIRY']]),0)
        #now=datetime.datetime.now()
        #date=datetime.datetime(now.year,now.month,now.day,0,0)
        date=datetime.datetime.strptime(trade_date,"%Y%m%d")
        '''*******************交易日请注意*********************
        '''
        T=((expiry-date).days+deltat)/365
        values['T']=T
        values['LOTx']=x[opcol['LOTx']]
        values['LOTS']=x[opcol[u'LOTS\n(CNY)']]
        temp=EuropeanOption()
        temp.setvalue(values)
        alloption[vari].append(temp)
    rule=re.compile(r'[^a-zA-z]')
    for x in baop_pos:
        if x[baopcol['EXPIRED']]=='Y':
            continue
        vari=rule.sub('',x[baopcol['UNDERLYING']])
        if vari=='':
            continue         
        allcode.append(x[baopcol['UNDERLYING']])
        if not alloption.has_key(vari):
            alloption[vari]=[]
        values={'r':0.03,'q':0.03}
        values['underlying']=x[baopcol['UNDERLYING']]
        values['type']=x[baopcol['C / P']]
        values['S']=x[baopcol['CLOSE\n(TODAY)']]
        values['K']=x[baopcol['K']]
        values['H']=x[baopcol['BARRIER H']]
        values['k']=x[baopcol['REBATE']]
        values['T']=x[baopcol['T\n(TODAY)']]
        values['vol']=x[baopcol['sT']] 
        values['LOTx']=x[baopcol['LOTx']] 
        values['LOTS']=x[baopcol[u'LOTS']]        
        temp=StandardBarrier()
        temp.setvalue(values)
        alloption[vari].append(temp)
    allfu={}
    for x in fu_pos:
        vari=x[fucol['GROUP']]
        if vari=='':
            continue
        if not allfu.has_key(vari):
            allfu[vari]=[]
        temp={}
        temp['code']=x[fucol['UNDERLYING']].upper()
        allcode.append(temp['code'])
        temp['buy']=x[fucol['BUY']]
        temp['sell']=x[fucol['BUY']]
        temp['avgbuy']=x[fucol['AVG\nBUY']]
        temp['avgsell']=x[fucol['AVG\nSELL']]
        temp['LOTx']=x[fucol['LOTx']]
        temp['LOTS']=x[fucol['OPEN \nPOSN']]
        if temp['LOTS']==0:
            continue
        allfu[vari].append(temp)
    codedata,maincode=getcodedata(set(allcode),alloption)
    for x in alloption:
        for y in alloption[x]:
            y.Greeks()
    return allfu,alloption,codedata,maincode
def getcodedata(allcode,alloption):
    acode=list(allcode)
    wcode=[]
    vspl={}
    rule=re.compile(r'[^a-zA-z]')
    for x in acode:    
        vari=rule.sub('',x)
        if not vspl.has_key(vari):
            vspl[vari]=[]
        temp=FutureClass(vari,x)
        wcode.append(x+'.'+temp.House)
        vspl[vari].append(temp)
    w.start()
    date=trade_date[:4]+'-'+trade_date[4:6]+'-'+trade_date[6:]
    windreulst=w.wsd(wcode,'close',date,date,'Fill=Previous')
    codedata=dict(zip(acode,windreulst.Data[0]))
    windreulst=w.wsd(wcode,'oi',date,date,'Fill=Previous')
    openint=dict(zip(acode,windreulst.Data[0]))
    maincode={}
    for x in vspl:
        tc=''
        to=0        
        for y in vspl[x]:
            if openint[y.Code]>to:                
                maincode[x]=y
                t0=openint[y.Code]
                y.Inf['open_interest']=t0
                y.Inf['cls']=codedata[y.Code]     
    return codedata,maincode
def cal_greeks(allfu,alloption,ds=0,dv=0):
    gks={}
    for x in allfu:
        if not gks.has_key(x):
            gks[x]={'delta':0,'gamma':0,'theta':0,'vega':0,'rho':0}
        for y in allfu[x]:            
            gks[x]['delta']+=y['LOTS']
    #op_pos2=[op_pos[-1]]
    for x in alloption:
        if not gks.has_key(x):
            gks[x]={'delta':0,'gamma':0,'theta':0,'vega':0,'rho':0}
        for y in alloption[x]:
            S0=y.S
            vol0=y.vol
            y.S=S0*(1+ds)
            y.vol=vol0*(1+dv)
            ygreek=y.Greeks()
            for nn in gks[x]:
                gks[x][nn]+=ygreek[nn]
            y.S=S0
            y.vol=vol0
    return gks
def greek_simulation(vari='AU'):
    dl=-0.05
    ul=0.05
    n=100    
    rlist=[dl+(ul-dl)/n*i for i in range(n+1)]    
    allfu,alloption,allcode,maincode=create_class() 
    reslist=[]
    for x in rlist:
        tgreeks=cal_greeks(allfu,alloption,ds=x)
        reslist.append(tgreeks)
    yname=['delta','gamma','vega','theta']
    ydic={}
    for y in yname:
        ydic[y]=[x[vari][y] for x in reslist]
    #fig=plt.figure(1)
    f,axarr=plt.subplots(2,2)
    f.suptitle('Greeks for '+vari)
    fy=[axarr[0,0],axarr[0,1],axarr[1,0],axarr[1,1]]
    i=0
    for x in ydic:
        plt.sca(fy[i])
        plt.plot(rlist,ydic[x])
        fy[i].set_title(x)
        plt.axis([dl, ul,min(ydic[x]),max(ydic[x])])
        i+=1
    plt.show()
    return reslist[int(n/2)]
def cal_pnl(allfu,alloption,codedata,ds=0,dv=0):
    pnl={}
    for x in alloption:
        if not pnl.has_key(x):
            pnl[x]=0
        for opclass in alloption[x]:
            if opclass.S<>codedata[opclass.underlying]:
                print opclass.underlying,'wind收盘价和excel收盘价对不上',opclass.S,codedata[opclass.underlying]
                print opclass.type,opclass.K,opclass.LOTS
            S0=opclass.S
            vol0=opclass.vol
            t0=opclass.T
            v0=opclass.OptionValue()
            opclass.S=S0*(1+ds)
            opclass.vol=vol0*(1+dv)
            opclass.T=t0-1/365
            v1=opclass.OptionValue()
            oppnl=(v1-v0)*opclass.LOTx*opclass.LOTS
            if math.isnan(oppnl):
                print opclass.type,opclass.K,opclass.LOTS
            pnl[x]+=oppnl
            opclass.S=S0
            opclass.vol=vol0
            opclass.T=t0
    for x in allfu:
        if not pnl.has_key(x):
            pnl[x]=0
        for y in allfu[x]:
            fupnl=y['LOTx']*y['LOTS']*codedata[y['code']]*ds
            if math.isnan(fupnl):
                print y
            pnl[x]+=fupnl
    return pnl
def pnl_simulation(vari='AU'):
    dl=-0.05
    ul=0.05
    n=100
    allfu,alloption,codedata,maincode=create_class()    
    rlist=[dl+(ul-dl)/n*i for i in range(n+1)]    
    #gks=cal_greeks(allfu,alloption,ds=0,dv=0)
    reslist=[]
    splist=[]#分解的盈亏
    for x in rlist:
        tpnl=cal_pnl(allfu,alloption,codedata,ds=x)
        reslist.append(tpnl)
        sppnl={}
        for y in alloption:
            sppnl[y]=0
            for op in alloption[y]:
                sppnl[y]+=op.S*x*op.greeks['delta']*op.LOTx+op.greeks['vega']*(x**2-op.vol**2*1/365)/(2*op.vol*op.T)
        for y in allfu:
            if not sppnl.has_key(y):
                sppnl[y]=0
            for fu in allfu[y]:
                sppnl[y]+=codedata[fu['code']]*x*fu['LOTx']*fu['LOTS']
        splist.append(sppnl)
    ylist1=[]
    ylist2=[]
    ylist3=[]
    zl=[]
    for i in range(len(rlist)):
        ylist1.append(reslist[i][vari])
        ylist2.append(splist[i][vari])
        ylist3.append(reslist[i][vari]-splist[i][vari])
        zl.append(0)
    f,axarr=plt.subplots(1,3)
    #plt.plot(rlist,ylist)
    f.suptitle('PnL for '+vari)
    fy=[axarr[0],axarr[1],axarr[2]]
    plt.sca(fy[0])
    plt.plot(rlist,ylist1,rlist,zl)
    plt.axis([dl, ul,min(ylist1),max(ylist1)])
    plt.sca(fy[1])
    plt.plot(rlist,ylist2,rlist,zl)
    plt.axis([dl, ul,min(ylist2),max(ylist2)])    
    plt.sca(fy[2])
    plt.plot(rlist,ylist3,rlist,zl)  
    plt.axis([dl, ul,min(ylist3),max(ylist3)])
    #plt.axis([dl, ul,min(ylist),max(ylist)])
    #plt.title(vari+' PnL')
    '''
    yname=['delta','gamma','vega','theta']
    ydic={}
    for y in yname:
        ydic[y]=[x[vari][y] for x in reslist]
    #fig=plt.figure(1)
    f,axarr=plt.subplots(2,2)
    f.suptitle('Greeks for '+vari)
    fy=[axarr[0,0],axarr[0,1],axarr[1,0],axarr[1,1]]
    i=0
    for x in ydic:
        plt.sca(fy[i])
        plt.plot(rlist,ydic[x])
        fy[i].set_title(x)
        plt.axis([dl, ul,min(ydic[x]),max(ydic[x])])
        i+=1
    plt.show()
    '''
    return reslist[int(n/2)]  
#allfu,alloption,codedata=create_class()
#gks=cal_greeks(allfu,alloption,ds=0,dv=0)
#greek_simulation()
#pnl=cal_pnl(allfu,alloption,codedata,ds=0,dv=0)
pnl_simulation('AU')