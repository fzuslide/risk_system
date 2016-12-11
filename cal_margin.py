# -*- coding: utf-8 -*-
"""
Created on Wed Mar 16 13:56:14 2016

@author: DELL
"""
from __future__ import division
import pandas as pd
import xlrd
import re
import copy
from operator import itemgetter, attrgetter
import functions as fn
from classes import *
def DCE_sp(dce_pos,codeclass,sp_order):
    ''' 函数说明：大商所保证金计算'''
    m=0
    cor_m=0
    pos={}
    pos2=[]
    pos3=[]
    for i in range(len(dce_pos)):
        code=dce_pos[i]['code']
        if len(code)<=6:
            if dce_pos[i]['posident']==u'套保':
                pos2.append(dce_pos[i])
            else:
                if pos.has_key(code):
                    pos[code]['longnums']=pos[code]['longnums']+dce_pos[i]['longnums']
                    pos[code]['shortnums']=pos[code]['shortnums']+dce_pos[i]['shortnums']
                else:
                    pos[code]=dce_pos[i]
    spcode=[]
    for x in pos:
        for y in pos:
            s=x+'&'+y
            if s in sp_order:
                spcode.append([s,sp_order.index(s)])
    spcode=sorted(spcode,key=lambda asd:asd[1])
    for sp in spcode:
        tppos={}
        code1=sp[0].split('&')[0]
        code2=sp[0].split('&')[1]
        for x in pos.keys():
            posinf1=pos[x]
            if posinf1['code']==code1:
                l1=posinf1['longnums']
                s1=posinf1['shortnums']                                               
                for y in pos.keys():
                    posinf2=pos[y]
                    if posinf2['code']==code2:
                        l2=posinf2['longnums']
                        s2=posinf2['shortnums']
                        ll=min(l1,s2)
                        ss=min(s1,l2)
                        if ll==0 and ss==0:
                            break
                        else:
                            tppos['code']='SP '+sp[0]
                            spmargin=[(ll*codeclass[code1].Units*codeclass[code1].Price*codeclass[code1].Rate,ll*codeclass[code1].Units*codeclass[code1].Price*(codeclass[code1].Rate+posinf1['cor_rate']))
                            ,(ll*codeclass[code2].Units*codeclass[code2].Price*codeclass[code2].Rate,ll*codeclass[code2].Units*codeclass[code2].Price*(codeclass[code2].Rate+posinf2['cor_rate']))]
                            spmargin.sort(key=itemgetter(1),reverse=True)#从大到小排序
                            m+=spmargin[0][0]
                            cor_m+=spmargin[0][1]
                            tppos['longnums']=ll
                            tppos['longmar']=spmargin[0][0]
                            tppos['cor_longmar']=spmargin[0][1]
                            spmargin=[(ss*codeclass[code1].Units*codeclass[code1].Price*codeclass[code1].Rate,ss*codeclass[code1].Units*codeclass[code1].Price*(codeclass[code1].Rate+posinf1['cor_rate']))
                            ,(ss*codeclass[code2].Units*codeclass[code2].Price*codeclass[code2].Rate,ss*codeclass[code2].Units*codeclass[code2].Price*(codeclass[code2].Rate+posinf2['cor_rate']))]
                            spmargin.sort(key=itemgetter(1),reverse=True)#从大到小排序
                            m+=spmargin[0][0]
                            cor_m+=spmargin[0][1]                            
                            tppos['shortnums']=ss
                            tppos['shortmar']=spmargin[0][0]
                            tppos['cor_shortmar']=spmargin[0][1]                            
                            #m=m+max(ll*codeclass[code1].Units*codeclass[code1].Price*codeclass[code1].Rate,ll*codeclass[code2].Units*codeclass[code2].Price*codeclass[code2].Rate)
                            #m=m+max(ss*codeclass[code1].Units*codeclass[code1].Price*codeclass[code1].Rate,ss*codeclass[code2].Units*codeclass[code2].Price*codeclass[code2].Rate)
                            posinf1['longnums']=posinf1['longnums']-ll
                            posinf1['shortnums']=posinf1['shortnums']-ss
                            posinf2['longnums']=posinf2['longnums']-ss
                            posinf2['shortnums']=posinf2['shortnums']-ll
                            pos[y]=posinf2
                            pos[x]=posinf1
                            tppos['longholdprice']=''
                            tppos['shortholdprice']=''
                            tppos['house']='DCE'
                            pos3.append(tppos)                                                             
                            break
                break
    for x in pos:
        posinf=pos[x]
        code=posinf['code']
        mm=max(posinf['longnums'],posinf['shortnums'])*codeclass[code].Units*codeclass[code].Price*codeclass[code].Rate
        m+=mm
        cor_mm=max(posinf['longnums'],posinf['shortnums'])*codeclass[code].Units*codeclass[code].Price*(codeclass[code].Rate+posinf['cor_rate'])
        cor_m+=cor_mm
        if posinf['longnums']>posinf['shortnums']:
            posinf['longmar']=mm
            posinf['shortmar']=0
            posinf['cor_longmar']=cor_mm
            posinf['cor_shortmar']=0
        else:
            posinf['longmar']=0
            posinf['shortmar']=mm
            posinf['cor_longmar']=0
            posinf['cor_shortmar']=cor_mm   
        pos3.append(posinf)
            
    for posinf in pos2:
        code=posinf['code']
        m=m+(posinf['longnums']+posinf['shortnums'])*codeclass[code].Units*codeclass[code].Price*codeclass[code].Rate 
        cor_m+=(posinf['longnums']+posinf['shortnums'])*codeclass[code].Units*codeclass[code].Price*(codeclass[code].Rate+posinf['cor_rate']) 
        posinf['longmar']=posinf['longnums']*codeclass[code].Units*codeclass[code].Price*codeclass[code].Rate 
        posinf['shortmar']=+posinf['shortnums']*codeclass[code].Units*codeclass[code].Price*codeclass[code].Rate
        posinf['cor_longmar']=posinf['longnums']*codeclass[code].Units*codeclass[code].Price*(codeclass[code].Rate+posinf['cor_rate']) 
        posinf['cor_shortmar']=+posinf['shortnums']*codeclass[code].Units*codeclass[code].Price*(codeclass[code].Rate+posinf['cor_rate']) 
        pos3.append(posinf)
    return m,cor_m,pos3
def SHF_sp(shf_pos,codeclass,unspcode):
    ''' 函数说明：上期所保证金计算'''
    m=0
    cor_m=0
    pos={}
    vari={}
    for i in range(len(shf_pos)):
        code=shf_pos[i]['code']        
        if code in unspcode:
            n1=shf_pos[i]['longnums']*codeclass[code].Units*codeclass[code].Price*codeclass[code].Rate
            n2=shf_pos[i]['shortnums']*codeclass[code].Units*codeclass[code].Price*codeclass[code].Rate
            cn1=shf_pos[i]['longnums']*codeclass[code].Units*codeclass[code].Price*(codeclass[code].Rate+shf_pos[i]['cor_rate'])
            cn2=shf_pos[i]['shortnums']*codeclass[code].Units*codeclass[code].Price*(codeclass[code].Rate+shf_pos[i]['cor_rate'])
            m+=n1+n2
            cor_m+=cn1+cn2                        
            shf_pos[i]['longmar']=n1
            shf_pos[i]['shortmar']=n2
            shf_pos[i]['cor_longmar']=cn1
            shf_pos[i]['cor_shortmar']=cn2            
        else:
            pos[i]=shf_pos[i]
            tmv=codeclass[pos[i]['code']].Vari
            if vari.has_key(tmv):
                vari[tmv].append(i)
            else:
                vari[tmv]=[i]
    for x in vari:
        ml=0
        ms=0
        cml=0
        cms=0
        for i in vari[x]:           
            posinf=pos[i]
            code=posinf['code']
            n1=posinf['longnums']*codeclass[code].Units*codeclass[code].Price*codeclass[code].Rate
            n2=posinf['shortnums']*codeclass[code].Units*codeclass[code].Price*codeclass[code].Rate
            cn1=posinf['longnums']*codeclass[code].Units*codeclass[code].Price*(codeclass[code].Rate+posinf['cor_rate'])
            cn2=posinf['shortnums']*codeclass[code].Units*codeclass[code].Price*(codeclass[code].Rate+posinf['cor_rate'])
            ml+=n1
            ms+=n2
            cml+=cn1
            cms+=cn2
            posinf['longmar']=n1
            posinf['shortmar']=n2
            posinf['cor_longmar']=cn1
            posinf['cor_shortmar']=cn2
        if ml>ms:
            m+=ml
            cor_m+=cml
            for i in vari[x]:
                pos[i]['shortmar'],pos[i]['cor_shortmar']=0,0
        else:
            m+=ms
            cor_m+=cms
            for i in vari[x]:
                pos[i]['longmar'],pos[i]['cor_longmar']=0,0            
    return m,cor_m,shf_pos
def CZC_sp(czc_pos,codeclass):
    ''' 函数说明：郑商所保证金计算'''
    m=0
    cor_m=0
    spos={}
    lpos={}
    for i in range(len(czc_pos)):
        code=czc_pos[i]['code']
        if len(code)>6:
            lpos[i]=czc_pos[i]
        else:
            if spos.has_key(code):
                spos[code]['longnums']=spos[code]['longnums']+czc_pos[i]['longnums']
                spos[code]['shortnums']=spos[code]['shortnums']+czc_pos[i]['shortnums']
            else:
                spos[code]=czc_pos[i]
    for i in lpos:
        sp=lpos[i]['code']
        code1=sp.split(' ')[1].split('&')[0]
        code2=sp.split(' ')[1].split('&')[1]
        posinf1=spos[code1]
        posinf2=spos[code2]
        lnums=lpos[i]['longnums']
        snums=lpos[i]['shortnums']
        spmargin=[(lnums*codeclass[code1].Units*codeclass[code1].Price*codeclass[code1].Rate,lnums*codeclass[code1].Units*codeclass[code1].Price*(codeclass[code1].Rate+posinf1['cor_rate']))
        ,(lnums*codeclass[code2].Units*codeclass[code2].Price*codeclass[code2].Rate,lnums*codeclass[code2].Units*codeclass[code2].Price*(codeclass[code2].Rate+posinf2['cor_rate']))]
        spmargin.sort(key=itemgetter(1),reverse=True)        
        m+=spmargin[0][0]
        cor_m+=spmargin[0][1]
        lpos[i]['longmar']=spmargin[0][0]
        lpos[i]['cor_longmar']=spmargin[0][1]
        #m=m+lnums*max(codeclass[code1].Units*codeclass[code1].Price*codeclass[code1].Rate,codeclass[code2].Units*codeclass[code2].Price*codeclass[code2].Rate)
        spmargin=[(snums*codeclass[code1].Units*codeclass[code1].Price*codeclass[code1].Rate,snums*codeclass[code1].Units*codeclass[code1].Price*(codeclass[code1].Rate+posinf1['cor_rate']))
        ,(snums*codeclass[code2].Units*codeclass[code2].Price*codeclass[code2].Rate,snums*codeclass[code2].Units*codeclass[code2].Price*(codeclass[code2].Rate+posinf2['cor_rate']))]
        spmargin.sort(key=itemgetter(1),reverse=True)        
        m+=spmargin[0][0]
        cor_m+=spmargin[0][1]
        lpos[i]['shortmar']=spmargin[0][0]
        lpos[i]['cor_shortmar']=spmargin[0][1]        
        #m=m+snums*max(codeclass[code1].Units*codeclass[code1].Price*codeclass[code1].Rate,codeclass[code2].Units*codeclass[code2].Price*codeclass[code2].Rate)     
        spos[code1]['longnums']=spos[code1]['longnums']-lnums
        spos[code2]['shortnums']=spos[code2]['shortnums']-lnums
        spos[code1]['shortnums']=spos[code1]['shortnums']-snums
        spos[code2]['longnums']=spos[code2]['longnums']-snums
    for x in spos:
        n1=spos[x]['longnums']*codeclass[x].Units*codeclass[x].Price*codeclass[x].Rate
        n2=spos[x]['shortnums']*codeclass[x].Units*codeclass[x].Price*codeclass[x].Rate
        cn1=spos[x]['longnums']*codeclass[x].Units*codeclass[x].Price*(codeclass[x].Rate+spos[x]['cor_rate'])
        cn2=spos[x]['shortnums']*codeclass[x].Units*codeclass[x].Price*(codeclass[x].Rate+spos[x]['cor_rate'])
        if n1>n2:
            m+=n1
            cor_m+=cn1
            spos[x]['longmar'],spos[x]['shortmar'],spos[x]['cor_longmar'],spos[x]['cor_shortmar']=n1,0,cn1,0
        else:
            m+=n2
            cor_m+=cn2
            spos[x]['longmar'],spos[x]['shortmar'],spos[x]['cor_longmar'],spos[x]['cor_shortmar']=0,n2,0,cn2
    return m,cor_m,spos.values()+lpos.values()       
def CFE_sp(cfe_pos,codeclass,cfe_unsp):
    ''' 函数说明：中金所保证金计算'''
    m=0
    cor_m=0
    pos2=[]
    idpos={}
    for posinf in cfe_pos:
        if idpos.has_key(posinf['posident']):
            idpos[posinf['posident']].append(posinf)
        else:
            idpos[posinf['posident']]=[posinf]
    for x in idpos:
        pos={}
        for posinf in idpos[x]:
            vari=codeclass[posinf['code']].Vari
            if vari==u'TF':
                vari=u'T'
            if pos.has_key(vari):
                pos[vari].append(posinf)
            else:
                pos[vari]=[posinf]
        for y in pos:
            ml=0
            ms=0
            cml=0
            cms=0
            for posinf in pos[y]:
                code=posinf['code']
                if x==u'套保' and y in ['IF','IH','IC']:
                    mr=0.2
                else:
                    mr=codeclass[code].Rate                
                n1=posinf['longnums']*codeclass[code].Units*codeclass[code].Price*mr                
                n2=posinf['shortnums']*codeclass[code].Units*codeclass[code].Price*mr                 
                cn1=posinf['longnums']*codeclass[code].Units*codeclass[code].Price*(mr+posinf['cor_rate'])               
                cn2=posinf['shortnums']*codeclass[code].Units*codeclass[code].Price*(mr+posinf['cor_rate'])
                if code in cfe_unsp:
                    m+=n1+n2
                    cor_m+=cn1+cn2
                    pos[y].remove(posinf)
                else:
                    ml=ml+n1
                    ms=ms+n2
                    cml+=cn1
                    cms+=cn2
                posinf['longmar'],posinf['shortmar'],posinf['cor_longmar'],posinf['cor_shortmar']=n1,n2,cn1,cn2
                pos2.append(posinf)
            if ml>ms:
                m+=ml
                cor_m+=cml
                for posinf in pos[y]:
                    posinf['shortmar'],posinf['cor_shortmar']=0,0                     
            else:
                m+=ms
                cor_m+=cms  
                for posinf in pos[y]:
                    posinf['longmar'],posinf['cor_longmar']=0,0                 
    return m,cor_m,pos2
def allcal(invclass,codeclass,sp_order,code_house,variclass,shfe_unsp,cfe_unsp=[]):
    ''' 函数说明：计算所有客户盈亏和保证金'''
    invclass2=copy.deepcopy(invclass)
    for x in invclass2:
        hc={'holdprofit_DCE':0,'holdprofit_CZC':0,'holdprofit_SHF':0,'holdprofit_CFE':0}
        hm={'margin_DCE':0,'margin_CZC':0,'margin_SHF':0,'margin_CFE':0}        
        if len(invclass2[x].Position)==0:
            invclass2[x].InvInf['cor_margin']=0
            invclass2[x].InvInf['cor_riskdegree']=0
            invclass2[x].InvInf['cor_leftcapital']=0
            for y in hc.keys():
                invclass2[x].InvInf[y]=hc[y]
            for y in hm.keys():
                invclass2[x].InvInf[y]=hm[y]            
            continue
        invinf=invclass2[x].InvInf
        pos=invclass2[x].Position
        cap=invinf['lastcapital']+invinf['closeprofit']+invinf['cashmove']-invinf['fee']+invinf['frozenmoney']-invinf['deliverymargin']
        dce_pos=[]
        shf_pos=[]
        czc_pos=[]
        cfe_pos=[]
        holdcap=0
        for posinf in pos:
            code=posinf['code']
            if len(code)<=6:
                house=variclass[posinf['vari']].House
            else:                
                rule=re.compile(r'[^a-zA-z]')
                vr=rule.sub('',code.split('&')[1])
                house=variclass[vr].House
            if house==u'DCE':
               dce_pos.append(posinf)
            if house==u'SHF':
               shf_pos.append(posinf)
            if house==u'CZC':
               czc_pos.append(posinf)
            if house==u'CFE':
               cfe_pos.append(posinf)                              
            if len(posinf['code'])<=6:         
                tcap=posinf['longnums']*codeclass[code].Units*(codeclass[code].Price-posinf['longholdprice'])+posinf['shortnums']*codeclass[code].Units*(posinf['shortholdprice']-codeclass[code].Price)                    
                holdcap+=tcap
                ss='holdprofit_'+codeclass[posinf['code']].House
                hc[ss]+=tcap
        unspcode=shfe_unsp
        m=0
        cor_m=0
        tm,tcm,dce_pos=DCE_sp(dce_pos,codeclass,sp_order)
        hm['margin_DCE']=tm
        m+=tm
        cor_m+=tcm
        tm,tcm,shf_pos=SHF_sp(shf_pos,codeclass,unspcode)
        hm['margin_SHF']=tm
        m+=tm
        cor_m+=tcm
        tm,tcm,czc_pos=CZC_sp(czc_pos,codeclass)
        hm['margin_CZC']=tm
        m+=tm
        cor_m+=tcm
        tm,tcm,cfe_pos=CFE_sp(cfe_pos,codeclass,cfe_unsp)
        hm['margin_CFE']=tm
        m+=tm
        cor_m+=tcm        
        invinf['capital']=cap+holdcap
        invinf['margin']=m
        invinf['cor_margin']=cor_m
        invinf['leftcapital']=invinf['capital']-invinf['margin']
        invinf['cor_leftcapital']=invinf['capital']-invinf['cor_margin']
        if invinf['capital']==0:
            invinf['riskdegree']=0
            invinf['cor_riskdegree']=0
        else:
            invinf['riskdegree']=m/invinf['capital']*100
            invinf['cor_riskdegree']=cor_m/invinf['capital']*100
        invinf['holdprofit']=holdcap  
        for y in hc.keys():
            invinf[y]=hc[y]
        for y in hm.keys():
            invinf[y]=hm[y]          
        #invclass2[x]=InvestorClass(invinf)
        #invclass2[x].Position=[]
        invclass2[x].Position=cfe_pos+shf_pos+czc_pos+dce_pos
    return invclass2
def out_inv(invclass):
    ''' 函数说明：输入表格（弃用）'''
    import xlwt
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet('My Worksheet')
    i=0
    colname=['invid','margin','cor_margin','capital']
    for j in range(4):
        worksheet.write(i, j, label =colname[j])
    i+=1
    for x in invclass:
        for j in range(4):
            worksheet.write(i, j, label = invclass[x].InvInf[colname[j]])
    workbook.save('Excel_Workbook.xls')
def check(invclass):
    ''' 函数说明：结果验证'''
    dp=u'\\\\10.100.6.20\\fkfile\\IORI-陈志荣\\三表\\0317试算结果.xls'
    book = xlrd.open_workbook(dp)
    sh=book.sheets()[0]
    n=sh.nrows    
    colname=sh.row_values(0)
    mlab={}
    res=[]
    for i in range(2,n):
        invinf={}
        invinf['invid']=str(int(sh.cell_value(i,colname.index(u'投资者代码'))))
        invinf['capital']=sh.cell_value(i,colname.index(u'权益'))
        invinf['margin']=sh.cell_value(i,colname.index(u'保证金'))       
        mlab[invinf['invid']]=invinf
        py=invclass[invinf['invid']].InvInf
        if abs(py['capital']-invinf['capital'])>0.1 or abs(py['margin']-invinf['margin'])>0.1:
            res.append(py)
    print len(res)
    return res       
