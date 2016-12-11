# -*- coding: utf-8 -*-
"""
Created on Mon Sep 19 21:28:54 2016

@author: IORI
"""

from __future__ import division
import math
from scipy import stats
class EuropeanOption:
    def __init__(self):
        self.S,self.K,self.T,self.vol,self.r,self.q='','','','','',''
        #S标的价格，K执行价，T时间，vol波动率，r无风险利率，q利息利率
        self.P=''#期权价格
        self.LOTx=1#期权乘数，一般等于期货乘数
        self.LOTS=1#期权手数
        self.type,self.underlying='',''
        self.greeks={}
    def setvalue(self,values):
        for x in values:
            if x=='S':
                self.S=float(values[x])
            if x=='K':
                self.K=float(values[x])
            if x=='T':
                self.T=float(values[x])
            if x=='vol':
                self.vol=float(values[x])
            if x=='r':
                self.r=float(values[x]) 
            if x=='q':
                self.q=float(values[x])
            if x=='type':
                self.type=values[x]
            if x=='underlying':
                self.underlying=values[x]
            if x=='LOTx':
                self.LOTx=float(values[x])
            if x=='LOTS':
                self.LOTS=float(values[x])                
    def OptionValue(self,CP=0,S=0,K=0,vol=0,r=0,q=0,T=0,isget=False):
        if not isget:               
            CP,S,K,vol,r,q,T=self.type,self.S,self.K,self.vol,self.r,self.q,self.T
        if T==0:
            if CP=='C':
                return max(S-K,0)
            else:
                return max(K-S,0)
        d1=(math.log(S/K)+(r-q+0.5*vol**2)*T)/(vol*math.sqrt(T))
        d2=(math.log(S/K)+(r-q-0.5*vol**2)*T)/(vol*math.sqrt(T))
        nd1=stats.norm.cdf(d1)
        nd2=stats.norm.cdf(d2)
        nnd1=stats.norm.cdf(-d1)
        nnd2=stats.norm.cdf(-d2)
        if CP=='C':
            value=S*math.exp(-q*T)*nd1-K*math.exp(-r*T)*nd2
        else:
            value=-S*math.exp(-q*T)*nnd1+K*math.exp(-r*T)*nnd2
        return value   
    def IV(self,guess=0.15):
        CP,S,K,r,q,T=self.type,self.S,self.K,self.r,self.q,self.T
        if self.P=='':
            print 'can not find option price'
            return 0
        option_value=self.P
        dv=0.005
        e=0.00001
        maxIter=100
        vol=guess
        rrvol=self.vol
        i=1
        while True:
            self.vol=vol
            value=self.OptionValue()
            dx=value-option_value
            if abs(dx)<e or i>=maxIter:
                break
            vega=greeks(CP,S,K,vol,r,q,T)['vega']
            vol=vol-dx/vega
            #if vol<=0:
            #    vol=dv
            i+=1
        self.vol=rrvol
        return vol
    def Greeks(self):
        S,vol,r,T=self.S,self.vol,self.r,self.T
        
        self.S=S+0.5
        uP=self.OptionValue()
        self.S=S-0.5
        dP=self.OptionValue()
        self.greeks['delta']=uP-dP
        
        self.S=S+1
        uP=self.OptionValue()
        self.S=S-1
        dP=self.OptionValue()
        self.S=S
        mP=self.OptionValue()        
        self.greeks['gamma']=uP+dP-2*mP
        
        self.vol=vol+0.005
        uP=self.OptionValue()
        self.vol=vol-0.005
        dP=self.OptionValue()
        self.greeks['vega']=uP-dP
        self.vol=vol
        
        if T<1/365:
            self.T=0.00001
        else:
            self.T=T-1/365
        uP=self.OptionValue()
        self.T=T
        dP=self.OptionValue()
        self.greeks['theta']=uP-dP
        
        self.r=r+0.00005
        uP=self.OptionValue()
        self.r=r-0.00005
        dP=self.OptionValue()
        self.greeks['rho']=uP-dP
        self.r=r
        
        self.greeks['delta']=self.greeks['delta']*self.LOTS
        self.greeks['gamma']=self.greeks['gamma']*self.LOTS
        nn=['theta','vega','rho']
        for x in nn:
            self.greeks[x]=self.greeks[x]*self.LOTx*self.LOTS
        return self.greeks
class StandardBarrier(EuropeanOption):
    def __init__(self):  
        EuropeanOption.__init__(self)    
        self.H=''#障碍水平
        self.k=''#期权发生敲入或敲出后的支付
    def setvalue(self,values):
        EuropeanOption.setvalue(self,values)
        for x in values:
            if x=='H':
                self.H=float(values[x])
            if x=='k':
                self.k=float(values[x])
    def OptionValue(self):
        TypeFlag, S, X, H=self.type,self.S,self.K,self.H
        k, T, r , v=self.k,self.T,self.r,self.vol
        b=self.r-self.q
        dt=1/365
        if H>S:
            H=H*math.exp(0.5826*v*math.sqrt(dt))
        elif H<S:
            H=H*math.exp(-0.5826*v*math.sqrt(dt))
        mu=(b-v**2/2)/(v**2)
        lbda=math.sqrt(mu**2+2*r/v**2)
        X1=math.log(S/X)/(v*math.sqrt(T))+(1+mu)*v*math.sqrt(T)
        X2=math.log(S/H)/(v*math.sqrt(T))+(1+mu)*v*math.sqrt(T)
        y1=math.log(H**2/(S*X))/(v*math.sqrt(T))+(1+mu)*v*math.sqrt(T)
        y2=math.log(H/S)/(v*math.sqrt(T))+(1+mu)*v*math.sqrt(T)
        z=math.log(H/S)/(v*math.sqrt(T))+lbda*v*math.sqrt(T)
        
        if TypeFlag=='cdi' or TypeFlag=='cdo':
            eta,phi=1,1
        elif TypeFlag=='cui' or TypeFlag=='cuo':
            eta,phi=-1,1
        elif TypeFlag=='pdi' or TypeFlag=='pdo':
            eta,phi=1,-1
        elif TypeFlag=='pui' or TypeFlag=='puo':
            eta,phi=-1,-1
            
        f1=phi*S*math.exp((b - r) * T) * stats.norm.cdf(phi * X1) - phi * X * math.exp(-r * T) * stats.norm.cdf(phi * X1 - phi * v * math.sqrt(T))
        f2=phi*S*math.exp((b - r) * T) * stats.norm.cdf(phi * X2) - phi * X * math.exp(-r * T) * stats.norm.cdf(phi * X2 - phi * v * math.sqrt(T))
        f3 = phi * S * math.exp((b - r) * T) * (H / S)**(2 * (mu + 1)) * stats.norm.cdf(eta * y1) - phi * X * math.exp(-r * T) * (H / S)**(2 * mu) * stats.norm.cdf(eta * y1 - eta * v * math.sqrt(T))  
        f4 = phi * S * math.exp((b - r) * T) * (H / S)**(2 * (mu + 1)) * stats.norm.cdf(eta * y2) - phi * X * math.exp(-r * T) * (H / S)**(2 * mu) * stats.norm.cdf(eta * y2 - eta * v * math.sqrt(T)) 
        f5 = k * math.exp(-r * T) * (stats.norm.cdf(eta * X2 - eta * v * math.sqrt(T)) - (H / S)**(2 * mu) * stats.norm.cdf(eta * y2 - eta * v * math.sqrt(T))) 
        f6 = k * ((H / S)** (mu + lbda) * stats.norm.cdf(eta * z) + (H / S)**(mu - lbda) * stats.norm.cdf(eta * z - 2 * eta * lbda * v * math.sqrt(T)))
        
        if X>H:
            if TypeFlag=='cdi':
                optionvalue=f3+f5
            if TypeFlag=='cui':
                optionvalue=f1+f5
            if TypeFlag=='pdi':
                optionvalue=f2-f3+f4+f5
            if TypeFlag=='pui':
                optionvalue=f1-f2+f4+f5
            if TypeFlag=='cdo':
                optionvalue=f1-f3+f6
            if TypeFlag=='cuo':
                optionvalue=f6
            if TypeFlag=='pdo':
                optionvalue=f1-f2+f3-f4+f6
            if TypeFlag=='puo':
                optionvalue=f2-f4+f6
        elif X<H:
            if TypeFlag=='cdi':
                optionvalue=f1-f2+f4+f5
            if TypeFlag=='cui':
                optionvalue=f2-f3+f4+f5
            if TypeFlag=='pdi':
                optionvalue=f1+f5
            if TypeFlag=='pui':
                optionvalue=f3+f5
            if TypeFlag=='cdo':
                optionvalue=f2+f6-f4
            if TypeFlag=='cuo':
                optionvalue=f1-f2+f3-f4+f6
            if TypeFlag=='pdo':
                optionvalue=f6
            if TypeFlag=='puo':
                optionvalue=f1-f3+f6
        return optionvalue         