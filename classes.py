# -*- coding: utf-8 -*-
"""
Created on Tue Jun 14 11:20:00 2016

@author: Administrator
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Mar 10 11:04:25 2016

@author: DELL
"""
import pandas as pd                    
import copy
import math
#import os
#import datetime
#import numpy as np

DataPath='\\\\10.100.6.20\\fkfile\VariData'
#DataPath='C:\\Users\\IORI\\Desktop\\python_risksystem'
filename=DataPath+'\\VariUnits.csv'

Variunits=pd.read_csv(filename)


class VariClass:
    def __init__(self,vari):
        self.Vari=vari.upper()
        self.Inf={}
        self.Units,self.House=self.getunits()
        self.Inf['Mrate_change']=0
        
    def getunits(self):
        runits=0
        rhouse=0
        self.Inf['Mrate']=0
        self.Inf['Limt']=0
        self.Inf['Mrate_all']=[0,0,0]       
        self.Inf['Limt_all']=[0,0,0]
        self.Inf['Dif_rate']=0
        self.Inf['Minprice']=0
        for i in range(len(Variunits)):
            vari=Variunits.vari.at[i]
            if vari==self.Vari:
                #lia=Variunits.vari.isin([self.Vari])
                runits=Variunits.units.at[i]
                rhouse=Variunits.house.at[i]
                self.Inf['Mrate']=Variunits.EX_margin_rate1.at[i]
                self.Inf['Limt']=Variunits.limit1.at[i]
                self.Inf['Minprice']=Variunits.minprice.at[i]
                self.Inf['Mrate_all']=Variunits.iloc[i,7:10]
                self.Inf['Limt_all']=Variunits.iloc[i,4:7]
                self.Inf['Dif_rate']=Variunits.com_dif.at[i]
                return runits,rhouse
                break
class FutureClass(VariClass):
    def __init__(self,vari,code):
        VariClass.__init__(self,vari)
        self.Code=code.upper()
        #self.Inf['Lbound']是否停板
        info=['Lbound','settlement','meanprice','high','low','opn','cls','open_interest','volume','price_now',
              'lastsettlement','lastcls','delta_price','lastopen_interest','delta_rate','ask','bid','limtup',
              'limtdown','Price','Rate']
        for x in info:
            self.Inf[x]=0
        self.Price=self.Inf['Price']
        self.Rate=self.Inf['Mrate']
    def setvalue(self,name,value):
        if self.Inf.has_key(name):
            self.Inf[name]=value
        else:
            if name<>'vari' and name<>'code':
                print self.Code+'      can not set value:'+name
    def adj_price(self,ptag=0,ctag=0,ps='settlement',isaddrate=True):
        #ptag最小变动价位选取方向，ctag是否引入价格变动
        if ps=='settlement':
            if self.Inf['settlement']<>'' and self.Inf['settlement']<>0:
                price=self.Inf['settlement']
            else:
                price=self.Inf['meanprice']
        if ps=='meanprice':
            price=self.Inf['meanprice']
        if ps=='price_now':
            price=self.Inf['price_now']
        if ps=='lastsettlement':
            price=self.Inf['lastsettlement']
        if ctag:
            price=price*(1+self.Inf['delta_price'])
            if not ptag:
                price=math.floor(price/self.Inf['Minprice'])*self.Inf['Minprice']
            else:
                if self.House=='SHF':
                    price=math.floor(price/self.Inf['Minprice'])*self.Inf['Minprice']
                elif self.House=='CZC':
                    if self.Inf['delta_price']>0:
                        price=math.ceil(price/self.Inf['Minprice'])*self.Inf['Minprice']
                    else:
                        price=math.floor(price/self.Inf['Minprice'])*self.Inf['Minprice']
                elif self.House=='DCE' or self.House=='CFE':
                    if self.Inf['delta_price']<0:
                        price=math.ceil(price/self.Inf['Minprice'])*self.Inf['Minprice']
                    else:
                        price=math.floor(price/self.Inf['Minprice'])*self.Inf['Minprice']                   
        self.Price=math.floor(price/self.Inf['Minprice'])*self.Inf['Minprice']
        if isaddrate:
            self.Rate=self.Inf['Mrate']+self.Inf['delta_rate']
        else:
            self.Rate=self.Inf['Mrate']
        self.Inf['Rate']=self.Rate
        self.Inf['Price']=self.Price
    def adj_rate(self):
        self.Rate=self.Inf['Mrate']+self.Inf['delta_rate']
    def intvalue(self,data):
        for x in data:
            self.setvalue(x,data[x])
class InvestorClass:
    def __init__(self,invinf):
        #InvInf is a dictionary for
        #invid,name,invdepartment,riskstate,lastriskstate,riskdegree,capital,lastcapital,margin,spmark,closeprofit,holdprofit,cashmove,fee
        self.InvID=invinf['invid']
        self.Name=invinf['inv_name']
        self.InvInf=invinf        
        self.Position=[]
        self.PurePos=[]  
    def addpos(self,data):
        #data is a Dictionary for:
        #invid,vari,code,longnums,longopenprice,longholdprice,longmar,shortnums,shortopenprice,shortholdprice,shortmar,posident
        self.Position.append(data)
    def getpurepos(self):        
        self.PurePos=copy.deepcopy(self.Position)
        n=len(self.Position)
        for i in range(n-1):
            p1=self.PurePos[i]
            for j in range(i+1,n):
                p2=self.PurePos[j]
                if p1['vari']==p2['vari']:
                    lnums=min(p1['longnums'],p2['shortnums'])
                    snums=min(p1['shortnums'],p2['longnums'])
                    p1['longnums']=p1['longnums']-lnums
                    p1['shortnums']=p1['shortnums']-snums
                    p2['shortnums']=p2['shortnums']-lnums
                    p2['longnums']=p2['longnums']-snums
                    self.PurePos[i]=p1
                    self.PurePos[j]=p2
        for x in self.PurePos:
            if x['longnums']==0 and x['shortnums']==0:
                self.PurePos.remove(x)

    