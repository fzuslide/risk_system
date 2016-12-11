# -*- coding: utf-8 -*-
"""
Created on Tue Dec 06 09:20:15 2016

@author: IORI
"""

from WindPy import w
from table_gui2 import *
from operator import itemgetter, attrgetter
import time
import datetime
import copy
class market_monitor:
    def __init__(self,codeclass):       
        w.start()
        import sys      
        import pyttsx       
        reload(sys)       
        sys.setdefaultencoding('utf8')
        self.engine = pyttsx.init()
        rate = self.engine.getProperty('rate')
        self.engine.setProperty('rate', rate-30)
        ser_code=[]
        self.unvari=['WR','FU','B','FB','BB','WH','RI','LR','PM','RS','JR']   
        for x in codeclass:
            ser_code.append(x+'.'+codeclass[x].House)
        self.w=w
        self.codeclass=codeclass
        self.ser_code=ser_code
        self.wmk=0.01
        self.c_gui()
    def c_gui(self):
        nulllist=[]
        for i in range(10):
            tp=[]
            for j in range(6):
                tp.append('null')
            nulllist.append(tp)
        self.colname=[u'合约代码',u'交易所',u'涨跌幅',u'停板幅度',u'最新价',u'停板价']
        self.up_gui=tb_gui(self.colname,nulllist,u'上涨合约',w=100)        
        self.down_gui=tb_gui(self.colname,nulllist,u'下跌合约',w=100)        
        self.up_limit_gui=tb_gui(self.colname,nulllist,u'接近涨停合约',w=100)        
        self.down_limit_gui=tb_gui(self.colname,nulllist,u'接近跌停合约',w=100)          
    def gui_config(self):
        for x in [self.up_gui,self.down_gui,self.up_limit_gui,self.down_limit_gui]:
            x.colortree()
            x.tree.tag_configure('red',foreground='red')
            x.tree.tag_configure('green',foreground='darkGreen')
            x.tree.tag_configure('redlimit',foreground='white',background='red')
            x.tree.tag_configure('greenlimit',foreground='white',background='darkGreen')
            x.root.update()
    def get_latest(self):
        windreulst=self.w.wsq(self.ser_code,'rt_latest')
        wind_volume=self.w.wsq(self.ser_code,'rt_vol')
        codeclass=self.codeclass
        up_code={}
        down_code={}
        for i,x in enumerate(self.ser_code):
            cc=x.split('.')[0]
            if codeclass[cc].Vari in self.unvari or windreulst.Data[0][i]==0 or wind_volume.Data[0][i]<=500:
                continue
            ls=codeclass[cc].Inf['lastsettlement']
            if ls==0:
                continue
            r=windreulst.Data[0][i]/ls-1
            if r>0:
                up_code[cc]=[r,windreulst.Data[0][i]]
            else:
                down_code[cc]=[r,windreulst.Data[0][i]]
        return up_code,down_code
    def early_warning(self):
        up_code,down_code=self.get_latest()
        uplist,downlist=[],[]
        codeclass=self.codeclass
        codel1,codel2,coden1,coden2=[],[],[],[]
        wl=0.8#百分比预警
        for x in up_code:
            ul=codeclass[x].Inf['limtup']
            dl=codeclass[x].Inf['limtdown']
            ls=codeclass[x].Inf['lastsettlement']                                  
            #deltar=(ul-up_code[x][1])/ls
            uplr=round(ul/ls-1,3)*100
            deltar=up_code[x][0]/(ul/ls-1)
            tls=[x,self.codeclass[x].House,round(up_code[x][0],4)*100,uplr,up_code[x][1],ul,deltar]
            if abs(up_code[x][1]-ul)<0.0000001:
                tls.append('redlimit')
                codel1.append(x)
            elif deltar>=wl:
                coden1.append(x)
                tls.append('red')
            else:
                tls.append('red')
            uplist.append(tls)
        for x in down_code:
            ul=codeclass[x].Inf['limtup']
            dl=codeclass[x].Inf['limtdown']
            ls=codeclass[x].Inf['lastsettlement']                                   
            #deltar=(down_code[x][1]-dl)/ls
            uplr=round(dl/ls-1,3)*100
            deltar=down_code[x][0]/(dl/ls-1)
            tls=[x,self.codeclass[x].House,round(down_code[x][0],4)*100,uplr,down_code[x][1],dl,deltar]
            if abs(down_code[x][1]-dl)<0.0000001:
                tls.append('greenlimit')
                codel2.append(x)
            elif deltar>wl:
                coden2.append(x)
                tls.append('green')
            else:
                tls.append('green')
            downlist.append(tls)
        uplist.sort(key=itemgetter(2),reverse=True)        
        outdata,tags=[],[]
        for x in uplist:
            outdata.append(x[:-2])
            tags.append(x[-1])
        self.up_gui.resettree(outdata,tags,iscolor=False)
        
        uplist.sort(key=itemgetter(-2),reverse=True)
        outdata,tags=[],[]
        for x in uplist:
            outdata.append(x[:-2])   
            tags.append(x[-1])
        self.up_limit_gui.resettree(outdata,tags,iscolor=False)
        
        downlist.sort(key=itemgetter(2),reverse=False)        
        outdata,tags=[],[]
        for x in downlist:
            outdata.append(x[:-2])
            tags.append(x[-1])
        self.down_gui.resettree(outdata,tags,iscolor=False)
        
        downlist.sort(key=itemgetter(-2),reverse=True)
        outdata,tags=[],[]
        for x in downlist:
            outdata.append(x[:-2])  
            tags.append(x[-1])
        self.down_limit_gui.resettree(outdata,tags,iscolor=False)        
        self.gui_config()
        self.tempcode=[codel1,codel2,coden1,coden2]
        return codel1,codel2,coden1,coden2
    def auto_warn(self):
        edtime=datetime.datetime.now()
        edtime=datetime.datetime(edtime.year,edtime.month,edtime.day,15,15)
        self.warmcode=copy.deepcopy(self.tempcode)
        i=0
        while i<240:#datetime.datetime.now()<edtime
            codel1,codel2,coden1,coden2=self.early_warning()
            scode=self.code_move(i)
            self.sound(scode)
            print u'睇%s次刷新行情'%i
            time.sleep(30)
            i+=1
    def sound(self,scode):
        codel1,codel2,coden1,coden2=scode[0],scode[1],scode[2],scode[3]
        import winsound
        t1=''
        for x in codel1:
            t1+=x+','
        if t1:
            t1+=u'已经涨停'
        t2=''
        for x in codel2:
            t2+=x+','
        if t2:
            t2+=u'已经跌停'
        t3=''
        for x in coden1:
            t3+=x+','
        if t3:
            t3+=u'接近涨停'
        t4=''
        for x in coden2:
            t4+=x+','      
        if t4:
            t4+=u'接近跌停'
        print t1,t2,t3,t4
        if t1 or t2:
            for i in range(2):       
                winsound.PlaySound('alter', winsound.SND_ASYNC)        
                time.sleep(0.5)                                 
            self.engine.say(u'attention，注意，')
            self.engine.say(t1)
            self.engine.say(t2)
        self.engine.runAndWait()
        if t3 or t4:      
            winsound.PlaySound('alter', winsound.SND_ASYNC)                                        
            self.engine.say(u'warming，预警，')
            self.engine.say(t3)
            self.engine.say(t4)            
        self.engine.runAndWait()
    def code_move(self,i):
        scode=[[],[],[],[]]
        if i%40==0:
            self.warmcode=copy.deepcopy(self.tempcode)
            for i in range(4):
                scode[i]=self.tempcode[i]
        else:
            for i in range(4):
                for x in self.tempcode[i]:
                    if x not in self.warmcode[i]:
                        scode[i].append(x)
                        self.warmcode[i].append(x)
        return scode
class tb_gui(table_gui):
    def colortree(self):
        tree=self.tree
        items = tree.get_children()
        for i in range(len(items)):
            tag=tree.item(items[i])['tags']  
            if 'oddrow' in tag:
                tag.remove('oddrow')
                tree.item(items[i], tags=tag)
            if i%2==1:               
                if tag=='':
                    tree.item(items[i], tags=['oddrow'])
                elif 'greenlimit' not in tag and 'redlimit' not in tag:
                    tree.item(items[i], tags=tag+['oddrow'])
                else:                   
                    tree.item(items[i], tags=tag)
        tree.tag_configure('oddrow', background='lavender')   