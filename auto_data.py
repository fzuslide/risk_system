# -*- coding: utf-8 -*-
"""
Created on Sat Jul 30 10:32:52 2016

@author: IORI
"""
import functions as fn
import pickle
import datetime
import time
import os
time_inf={'tag':0,'starttime':'','endtime':'','count':0}
soursepath=u'\\\\10.100.6.20\\fkfile\\风控终端三表\\读这个'
outpath=u'\\\\10.100.6.20\\fkfile\\python_risksystem\\data'
warmf=soursepath+u'\\放心复制.txt'
warmf2=soursepath+u'\\放心复制,python复制zhong.txt'
sh3_info1=outpath+'\\readable.txt'
sh3_info2=outpath+'\\unreadable.txt'
sh3_datafile=outpath+'\\sh3_data.pickle'
invfname=soursepath+'\\'+u'投资者资金信息.xls'
posfname=soursepath+'\\'+u'持仓查询.xls'
riskfname=soursepath+'\\'+u'实时风控行情.xls' 
def time_count(text=''):
    ''' 函数说明：运行时间计算'''
    if time_inf['tag']:
        time_inf['endtime']= datetime.datetime.now()
        time_inf['tag']=0
        print text+u' 运行完成,耗时：'+str((time_inf['endtime'] - time_inf['starttime']).seconds)+u's'  
        time_inf['count']=time_inf['count']+1
        print u'第'+str(time_inf['count'])+u'次计算时间'
    else:
        time_inf['starttime']= datetime.datetime.now()
        time_inf['tag']=1        
def wait_newsh(oldtime):
    isn=False
    while not isn:
        if datetime.datetime.now().hour>=16:
            return False
        if os.path.exists(warmf):
            tt=[]
            date=datetime.datetime.fromtimestamp(os.path.getmtime(invfname))
            tt.append(u'资金表 '+date.strftime('%Y-%m-%d %H:%M:%S')) 
            date=datetime.datetime.fromtimestamp(os.path.getmtime(posfname))
            tt.append(u'持仓表 '+date.strftime('%Y-%m-%d %H:%M:%S'))
            date=datetime.datetime.fromtimestamp(os.path.getmtime(riskfname))
            tt.append(u'行情表 '+date.strftime('%Y-%m-%d %H:%M:%S')) 
            tag=True
            for i in range(len(oldtime)):
                if oldtime[i]==tt[i]:
                    tag=False
            isn=tag
            if isn:
                break
        time.sleep(5)    
    return True
def ini_data():
    data={}
    time_count()
    ini_codeclass,ini_sporder,ini_belong,ini_seat,ini_sprate,ini_phone,code_house,ini_invtype,tt=fn.read_baicsh()
    data['sheet_info']=tt
    data['ini_codeclass']=ini_codeclass
    data['ini_sporder']=ini_sporder
    data['ini_belong']=ini_belong
    data['ini_seat']=ini_seat
    data['ini_sprate']=ini_sprate
    data['ini_phone']=ini_phone      
    data['ini_invtype']=ini_invtype
    data['code_house']=code_house
    data['delta_price_format']=0#代表结算价有调整时统一向下方式
    data['is_delta_price']=False#代表是否有结算价调整
    data['shfe_unsp'],data['cfe_unsp']=fn.cal_shfe_unsp(ini_codeclass)
    with open(outpath+'\\data.pickle', 'wb') as f:pickle.dump(data, f)
    f.close()
    time_count(u'初始化')
    return data
def auto_sh3(data=''):
    if data=='':
        with open(outpath+'\\data.pickle', 'rb') as f:data = pickle.load(f)   
    while True:
        time_count()
        sh3_data={}
        while not os.path.exists(warmf):
            time.sleep(3)
        os.rename(warmf,warmf2)
        try:
            allinv,allpos,allcode,tt=fn.read_sh3(invfname,posfname,riskfname)
        except:
            os.rename(warmf2,warmf)
            print 'read sh3 error'
            time_count()
            time.sleep(30)
            continue
        os.rename(warmf2,warmf)
        data['allinv']=allinv
        data['allpos']=allpos
        data['allcode']=allcode
        if len(data['sheet_info'])<=6:
            data['sheet_info']=tt+data['sheet_info']
        else:
            data['sheet_info'][:3]=tt
        fn.creat_class(data)
        fn.match_corp_marr(data['invclass'],data['variclass'],data['ini_sprate'])   
        invclass=data['invclass']
        ini_seat=data['ini_seat']
        ini_belong=data['ini_belong']
        ini_phone=data['ini_phone']
        ini_invtype=data['ini_invtype']
        for x in invclass:          
            if ini_seat.has_key(x):
                invclass[x].InvInf['seat']=ini_seat[x]
            else:
                invclass[x].InvInf['seat']=''
                #print u'席位表不存，找不到客户：'+x            
            if ini_belong.has_key(x):
                invclass[x].InvInf['invdepartment']=ini_belong[x]                          
            if ini_phone.has_key(x):
                invclass[x].InvInf['phone']=ini_phone[x]
                invclass[x].InvInf['type']=ini_invtype[x]
            else:
                invclass[x].InvInf['phone']=''
                invclass[x].InvInf['type']=''
                #print  u'信息表不存，找不到客户：'+x    
        sh3_data['invclass']=data['invclass']
        sh3_data['ini_codeclass']=data['ini_codeclass']
        sh3_data['variclass']=data['variclass']
        sh3_data['sheet_info']=data['sheet_info']
        os.rename(sh3_info1,sh3_info2)
        time.sleep(12)
        try:
            with open(sh3_datafile, 'wb') as f:pickle.dump(sh3_data, f)
            f.close()
        except:
            os.rename(sh3_info2,sh3_info1)
            print 'inupt data error'
            time_count()
            time.sleep(30)
            continue            
        os.rename(sh3_info2,sh3_info1)
        time_count(u'读表')
        time_count()
        if not wait_newsh(sh3_data['sheet_info'][:3]):
            print 'end'        
            break
        time_count(u'等待新三表')
#data=ini_data()
#auto_sh3(data)
auto_sh3()   
    
    