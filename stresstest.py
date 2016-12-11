# -*- coding: utf-8 -*-
"""
Created on Thu Jun 30 16:06:26 2016

@author: Administrator
"""
from __future__ import division
import functions as fun
import cal_margin
import pandas as pd
import math
import datetime
import pickle
import xlrd
from WindPy import w
import os
def historytest():
    starttime = datetime.datetime.now()
    data,allpch,adt=getfu()
    codeclass=data['codeclass']
    mnres=[]
    yares=[]#80088412
    i=0
    for j in range(len(allpch)):
        t=allpch[j]
        if max(map(abs,t.values()))>=0:
            m1=0
            n1=0
            m2=0
            n2=0
            for x in codeclass:
                codeclass[x].Inf['delta_price']=t[codeclass[x].Vari]
            [data['codeclass'][x].adj_price(0,1) for x in data['codeclass']]    
            invclass=cal_margin.allcal(data['invclass'],data['codeclass'],data['ini_sporder'],data['code_house'],data['variclass'],data['shfe_unsp'])
            for x in invclass:
                if invclass[x].InvInf['leftcapital']<0:
                    m2+=invclass[x].InvInf['leftcapital']
                    n2+=1
                    if invclass[x].InvInf['capital']<0:
                        m1+=invclass[x].InvInf['capital']
                        n1+=1
            mnres.append([adt[j],m1,n1,m2,n2])
            yares.append([adt[j],invclass[u'80088412'].InvInf['capital'],invclass[u'80088412'].InvInf['leftcapital']])
            i+=1
            if i%10==0:
                endtime = datetime.datetime.now()
                print i
                print (endtime - starttime).seconds
        else:
            mnres.append([adt[j],0,0,0,0])    
    df=pd.DataFrame(mnres,columns=[u'日期',u'穿仓金额',u'穿仓人数',u'强平金额',u'强平人数'])
    df.to_excel('testresult20160630.xls')
    df=pd.DataFrame(yares,columns=[u'日期',u'权益',u'可用资金'])
    df.to_excel('yaresult20160630.xls')
    endtime = datetime.datetime.now()
    print (endtime - starttime).seconds 
def screntest():
    with open('data.pickle', 'rb') as f:data = pickle.load(f)    
    scenres=[]
    yascenres=[]
    codeclass=data['codeclass']
    DataPath='\\\\10.100.6.20\\fkfile\VariData'
    filename=DataPath+'\VariUnits.csv'
    Variunits=pd.read_csv(filename)
    for x in codeclass:
        for i in range(len(Variunits)):
            vari=Variunits.vari.at[i]
            if vari==codeclass[x].Vari:
                codeclass[x].Inf['Mrate_all']=Variunits.iloc[i,7:10]
                codeclass[x].Inf['Limt_all']=Variunits.iloc[i,4:7]
                break
    for ud in [1,-1]:
        for k in range(3):
            for x in codeclass:
                if codeclass[x].Vari in ['T','TF']:
                    continue
                t=1
                for j in range(k+1):
                    t=t*(1+ud*codeclass[x].Inf['Limt_all'][j])
                codeclass[x].Inf['delta_price']=t-1
            m1=0
            n1=0
            m2=0
            n2=0
            [codeclass[x].adj_price(1,1) for x in codeclass] 
            invclass2=cal_margin.allcal(data['invclass'],codeclass,data['ini_sporder'],data['code_house'],data['variclass'],data['shfe_unsp'])
            for x in invclass2:
                if invclass2[x].InvInf['leftcapital']<0:
                    m2+=invclass2[x].InvInf['leftcapital']
                    n2+=1
                    if invclass2[x].InvInf['capital']<0:
                        m1+=invclass2[x].InvInf['capital']
                        n1+=1
            scenres.append([str(ud*(k+1))+u'个板',m1,n1,m2,n2])
            yascenres.append([str(ud*(k+1))+u'个板',invclass2[u'80088412'].InvInf['capital'],invclass2[u'80088412'].InvInf['leftcapital']])
    df=pd.DataFrame(scenres,columns=[u'压力情景',u'穿仓金额',u'穿仓人数',u'强平金额',u'强平人数'])
    df.to_excel('scenresult.xls')
    df=pd.DataFrame(yascenres,columns=[u'压力情景',u'权益',u'可用资金'])
    df.to_excel('yascenresult.xls')
def getfu(sdate='2015-09-01 00:00:00.000',delta=6,variclass=''):
    dp='\\\\10.100.6.20\\fkfile\\VariData\\MainCode\\'
    data=''
    if variclass=='':
        path=os.getcwd()
        with open(os.path.dirname(path)+'\\data\\'+'data.pickle', 'rb') as f:data = pickle.load(f)    
        variclass=data['variclass']
    idvari=fun.upperlist()
    effvari=fun.upperlist(1)
    
    maincode={}
    for x in variclass.keys():
        fname=dp+x+'_Maincode.csv'
        maincode[x]=pd.read_csv(fname)
    indexres={}
    for x in idvari:
        fname='\\\\10.100.6.20\\fkfile\\VariData\\IndexResult\\'
        for y in x:
            fname+=y+'_'
        fname+='index.csv'
        indexres[idvari.index(x)]=pd.read_csv(fname)
        
    isdt=maincode['IF'].date>=sdate
    allpch=[]#品种历史收益率
    i=0
    for dt in maincode['IF'].date[isdt]:
        vpch={}#品种收益率
        for x in variclass.keys():
            yd=0
            if max(maincode[x].date.isin([dt])):
                lia=maincode[x].date.isin([dt])
                k=lia.tolist().index(True)   
                spoint=max(0,k-delta+1)
                for i in range(spoint,k+1):
                    yd+=maincode[x].cls.at[i]/maincode[x].pre_cls.at[i]-1
                #vpch[x]=maincode[x].settlement[lia]/maincode[x].pre_settlement[lia]-1
                #vpch[x]=vpch[x].tolist()[0]
                vpch[x]=yd
            else:
                isf=False
                for i in range(len(idvari)):
                    if x in idvari[i]:
                        if max(indexres[i].date.isin([dt])):
                            lia=indexres[i].date.isin([dt])
                            k=lia.tolist().index(True)
                            spoint=max(0,k-delta+1)
                            for i in range(spoint,k+1):
                                yd+=ath.exp(indexres[i].at['yield',i])-1
                            #vpch[x]=math.exp(indexres[i]['yield',lia])-1
                            #vpch[x]=vpch[x].tolist()[0]
                            vpch[x]=yd
                        else:
                            vpch[x]=0
                        isf=True
                        break
                if not isf:
                    vpch[x]=0
        allpch.append(vpch)
    adt=maincode['IF'].date[isdt].tolist()
    return data,allpch,adt
def getselfdata():
    dp=u'C:\\Users\\Administrator\\Desktop\\stresstest\\持仓 20160630_to 志荣.xlsx'
    book = xlrd.open_workbook(dp)
    sh=book.sheet_by_name(u'期货')
    n=sh.nrows     
    colname=sh.row_values(0)
    fulist=[]
    for i in range(1,n):
        fulist.append([sh.cell_value(i,colname.index(u'合约')).upper()])
        if sh.cell_value(i,colname.index(u'买卖'))==u'卖':
            fulist[-1].append(-1*sh.cell_value(i,colname.index(u'数量')))
        else:
            fulist[-1].append(1*sh.cell_value(i,colname.index(u'数量')))
        fulist[-1].append(sh.cell_value(i,colname.index(u'品种')).upper())
        fulist[-1].append(sh.cell_value(i,colname.index(u'公司')))
    sh=book.sheet_by_name(u'证券-研究所')
    n=sh.nrows     
    colname=sh.row_values(0)        
    stlist=[]
    for i in range(1,n):
        stlist.append([sh.cell_value(i,colname.index(u'证券代码'))])
        stlist[-1].append(sh.cell_value(i,colname.index(u'股票余额')))
        stlist[-1].append(sh.cell_value(i,colname.index(u'市价')))
        stlist[-1].append(sh.cell_value(i,colname.index(u'交易市场')))
        stlist[-1].append(sh.cell_value(i,colname.index('tag')))
        stlist[-1].append(sh.cell_value(i,colname.index(u'公司')))
    sh=book.sheet_by_name(u'期权-华泰账户')
    n=sh.nrows     
    colname=sh.row_values(0)        
    oplist=[]
    for i in range(1,n):
        oplist.append([sh.cell_value(i,colname.index(u'代码'))])
        oplist[-1].append(sh.cell_value(i,colname.index(u'名称')))
        oplist[-1].append(sh.cell_value(i,colname.index(u'类别')))
        oplist[-1].append(sh.cell_value(i,colname.index(u'买卖')))
        oplist[-1].append(sh.cell_value(i,colname.index(u'持仓')))
        oplist[-1].append(sh.cell_value(i,colname.index(u'最新价')))
        oplist[-1].append(sh.cell_value(i,colname.index(u'Delta')))
        oplist[-1].append(sh.cell_value(i,colname.index(u'Gamma')))
        oplist[-1].append(sh.cell_value(i,colname.index(u'Rho')))
        oplist[-1].append(sh.cell_value(i,colname.index(u'Theta')))
        oplist[-1].append(sh.cell_value(i,colname.index(u'Vega')))
        oplist[-1].append(sh.cell_value(i,colname.index(u'公司')))
    return fulist,stlist,oplist
def getst(stlist):
    w.start()
    sdate='2015-06-30'
    edate='2016-06-30'
    code=[]
    industrylist=[]
    for x in stlist:
        if x[4]==0:
            x.append('A')
            continue
        if x[3]==u'深圳Ａ':     
            ss='.SZ'
        elif x[3]==u'上海Ａ':
            ss='.SH'
        else:
            print x[0]
        wcode=str6(str(int(x[0])))+ss
        x[0]=wcode
        code.append(wcode)
        windresult=w.wsd(wcode,'indexcode_sw',edate,edate,'industryType=2')
        if windresult.Data[0][0]==None:
            windresult=w.wsd(wcode,'indexcode_sw',edate,edate,'industryType=1')            
        industrylist.append(windresult.Data[0][0])
        x.append(windresult.Data[0][0])
    code.append('510050.SH')
    sercode=list(set(code+industrylist))
    #windresult=w.wsd(sercode,'pct_chg',sdate,edate,'Fill=Previous')
    windresult=w.wsd(sercode,'pct_chg',sdate,edate,'Fill=Previous')
    pct_ch=windresult.Data
    windresult=w.wsd(sercode,'trade_status',sdate,edate,'Fill=Previous')
    status=windresult.Data
    return stlist,sercode,pct_ch,status
def str6(s):
    while len(s)<6:
        s='0'+s
    return s
def crdic(team,value=0):
    dic={}
    for x in team:
        if value=='list':
            dic[x]=[]
        else:
            dic[x]=value
    return dic
def calself():
    data,allpch,adt=getfu()
    fulist,stlist,oplist=getselfdata()
    stlist,sercode,pct_ch,status=getst(stlist)
    codeclass=data['codeclass']
    fucp=[]
    stcp=[]
    opcp=[]
    team=[]
    for x in fulist:
        team.append(x[-1])
    for x in allpch:
        dic=crdic(set(team))
        for y in fulist:
            dic[y[-1]]+=y[1]*codeclass[y[0]].Inf['cls']*x[codeclass[y[0]].Vari]*codeclass[y[0]].Units
        fucp.append(dic)
    team=[]
    for x in stlist:
        team.append(x[-2])
    for i in range(len(pct_ch[0])):
        dic=crdic(set(team))
        for x in stlist:
            if x[4]==0:
                dic[x[-2]]+=0.035/246*x[1]*x[2]
                continue
            k=sercode.index(x[0])
            if math.isnan(pct_ch[k][i]):
                k=sercode.index(x[-1])
            elif status[k][i]<>u'交易' or status[k][i]<>u'停牌半天':
                k=sercode.index(x[-1])
            ch=pct_ch[k][i]/100
            dic[x[-2]]+=x[1]*x[2]*ch
        stcp.append(dic)
    team=[]
    for x in oplist:
        team.append(x[-1])
    etf50=2.136#20160630
    k=sercode.index('510050.SH')
    for i in range(len(pct_ch[k])):
        dic=crdic(set(team))
        for x in oplist:
            ds=etf50*pct_ch[k][i]/100
            dic[x[-1]]+=10000*x[4]*(x[6]*ds+0.5*x[7]*ds**2+x[9]*1/246)
        opcp.append(dic)
    allt=set(fucp[0].keys()+stcp[0].keys()+opcp[0].keys())
    teamcp=crdic(allt,'list')
    for i in range(len(fucp)):
        cp=crdic(allt,'list')
        for x in fucp[i]:
            cp[x].append(fucp[i][x])
        for x in stcp[i]:
            cp[x].append(stcp[i][x])
        for x in opcp[i]:
            cp[x].append(opcp[i][x])
        for x in cp:
            teamcp[x].append([adt[i]]+cp[x]+[sum(cp[x])])
    for x in teamcp:
        colname= [u'日期']
        if fucp[0].has_key(x):
            colname.append(u'期货盈亏')
        if stcp[0].has_key(x):
            colname.append(u'证券盈亏')
        if opcp[0].has_key(x):
            colname.append(u'期权盈亏')
        colname.append(u'总盈亏')
        df=pd.DataFrame(teamcp[x],columns=colname)
        df.to_excel(x+'20160630.xls')
def test():
    #data,allpch,adt=getfu()
    with open('data.pickle', 'rb') as f:data = pickle.load(f)
    fulist,stlist,oplist=getselfdata()
    #stlist,sercode,pct_ch,status=getst(stlist)
    codeclass=data['codeclass']
    DataPath='\\\\10.100.6.20\\fkfile\VariData'
    filename=DataPath+'\VariUnits.csv'
    Variunits=pd.read_csv(filename)
    for x in codeclass:
        for i in range(len(Variunits)):
            vari=Variunits.vari.at[i]
            if vari==codeclass[x].Vari:
                codeclass[x].Inf['Mrate_all']=Variunits.iloc[i,7:10]
                codeclass[x].Inf['Limt_all']=Variunits.iloc[i,4:7]
                break
    adt=[u'涨1个板',u'涨2个板',u'涨3个板',u'跌1个板',u'跌2个板',u'跌3个板']
    strn=[0.1,0.21,1.1**3-1,-0.1,0.9**2-1,0.9**3-1]
    allpch=[]
    for ud in [1,-1]:
        for k in range(3):
            dic={}
            for x in codeclass:
                if dic.has_key(codeclass[x].Vari):
                    continue
                if codeclass[x].Vari in ['T','TF']:
                    dic[codeclass[x].Vari]=0                
                    continue
                t=1
                for j in range(k+1):
                    t=t*(1+ud*codeclass[x].Inf['Limt_all'][j])
                dic[codeclass[x].Vari]=t-1
            allpch.append(dic)
    fucp=[]
    stcp=[]
    opcp=[]
    team=[]
    for x in fulist:
        team.append(x[-1])
    for x in allpch:
        dic=crdic(set(team))
        for y in fulist:
            dic[y[-1]]+=y[1]*codeclass[y[0]].Inf['cls']*x[codeclass[y[0]].Vari]*codeclass[y[0]].Units
        fucp.append(dic)
    team=[]
    for x in stlist:
        team.append(x[5])
    for i in range(len(strn)):
        dic=crdic(set(team))
        for x in stlist:
            if x[4]==0:
                dic[x[5]]+=0.035/246*x[1]*x[2]
                continue
            ch=strn[i]
            dic[x[5]]+=x[1]*x[2]*ch
        stcp.append(dic)
    team=[]
    for x in oplist:
        team.append(x[-1])
    etf50=2.136#20160630
    for i in range(len(strn)):
        dic=crdic(set(team))
        for x in oplist:
            ds=etf50*strn[i]
            dic[x[-1]]+=10000*x[4]*(x[6]*ds+0.5*x[7]*ds**2+x[9]*1/246)
        opcp.append(dic)
    allt=set(fucp[0].keys()+stcp[0].keys()+opcp[0].keys())
    teamcp=crdic(allt,'list')
    for i in range(len(fucp)):
        cp=crdic(allt,'list')
        for x in fucp[i]:
            cp[x].append(fucp[i][x])
        for x in stcp[i]:
            cp[x].append(stcp[i][x])
        for x in opcp[i]:
            cp[x].append(opcp[i][x])
        for x in cp:
            teamcp[x].append([adt[i]]+cp[x]+[sum(cp[x])])
    for x in teamcp:
        colname= [u'日期']
        if fucp[0].has_key(x):
            colname.append(u'期货盈亏')
        if stcp[0].has_key(x):
            colname.append(u'证券盈亏')
        if opcp[0].has_key(x):
            colname.append(u'期权盈亏')
        colname.append(u'总盈亏')
        df=pd.DataFrame(teamcp[x],columns=colname)
        df.to_excel(x+u'情景测试20160630.xls')
def calwang():
    w.start()
    sdate='2015-05-03'
    edate='2016-07-29'
    windresult=w.wsd('510050.SH',['pct_chg','close'],sdate,edate,'Fill=Previous')        
    wdata=windresult.Data
    date=[datetime.datetime.strftime(x,'%Y%m%d') for x in windresult.Times]
    fname='position.xlsx'
    book = xlrd.open_workbook(fname)   
    posdata={}          
    for sh in book.sheets():
        n=sh.nrows
        shname=sh.name.split(' ')[1]
        v=[]
        for i in range(1,n):
            v.append(sh.row_values(i))
        posdata[shname]=v        
    colname=sh.row_values(0)   
    cdic={'nums':colname.index(u'持仓'),'margin':colname.index(u'保证金'),'delta':colname.index(u'Delta'),'gamma':colname.index(u'Gamma'),'theta':colname.index(u'Theta')}     
    res=[]
    dt=1/252
    for x in posdata:
        i=date.index(x)
        cls=wdata[1][i]
        ptch=wdata[0][:i]
        if i>252:
            ptch=wdata[0][i-252:i]
        pnl=[]
        for ch in ptch:
            ds=cls*ch/100
            pl=0        
            for y in posdata[x]:
                if int(x)>=20160720:
                    pl+=(ds*y[cdic['delta']]+0.5*y[cdic['gamma']]*ds**2+dt*y[cdic['theta']])*10000
                else:
                    pl+=(ds*y[cdic['delta']]+0.5*y[cdic['gamma']]*ds**2+dt*y[cdic['theta']])*y[cdic['nums']]*10000
            pnl.append(pl)
        pnl.sort(reverse=True)
        mar=0
        for y in posdata[x]:
            mar+=y[cdic['margin']]
        k1=int(math.ceil(0.99*len(pnl)))
        k2=int(math.ceil(0.95*len(pnl)))
        k3=int(math.ceil(0.9*len(pnl)))
        v=[pnl[k1],pnl[k2],pnl[k3],mar]
        res.append((x,v))
    res.sort()
    import xlwt
    i=0
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet(u'Var值')
    wcol=[u'日期',u'99%VaR',u'95%VaR',u'90%VaR',u'保证金']
    for j in range(len(wcol)):
        worksheet.write(i, j, label=wcol[j])
    for i in range(len(res)):
        worksheet.write(i+1, 0, label=res[i][0])
        for j in range(len(res[i][1])):
            worksheet.write(i+1, j+1, label=res[i][1][j])
    workbook.save(u'王维扬VaR值.xls')
#data,allpch,adt=getfu()       