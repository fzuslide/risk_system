# -*- coding: utf-8 -*-
"""
Created on Mon Aug 08 11:12:21 2016

@author: IORI
"""
from __future__ import division
import Tkinter as tk
import tkSimpleDialog
import tkFileDialog
import ttk
import functions as fn
import simulation as sim
import datetime
import time
import tkMessageBox
import pickle
import xlrd
import os
from operator import itemgetter, attrgetter
from margin_gui import *
from settlement_gui import *
from table_gui2 import *
from single_gui import *
class simulation_gui():
    def __init__(self,parent):
        self.parent=parent
        self.codeclass=parent.data['ini_codeclass']
        self.variclass=parent.data['variclass']
        frame=tk.Frame(parent.note,bg='Azure')
        parent.note.add(frame, text = u'模拟试算参数')
        f1=tk.Frame(frame,bg='Azure')
        f2=tk.Frame(frame,bg='Azure')
        f1.grid(row=0,column=0,sticky='w')
        f2.grid(row=1,column=0,sticky='w')
        self.upgui=f1
        self.downgui=f2
        self.create_upgui()
        self.fr=True
        colname,data=self.get_range()
        self.outdata(colname,data)
        parent.note.select(parent.note.tabs()[-1])
    def create_upgui(self):
        f=self.upgui       
        self.cmbox={}
        cl=[u'置信度',u'区间数量',u'最近交易日']
        for i in range(len(cl)):
            self.cmbox[cl[i]]=ttk.Combobox(f,width=10)
            tk.Label(f,text=cl[i]).grid(row=0,column=i,sticky='w')
            self.cmbox[cl[i]].grid(row=1,column=i,sticky='w')
            self.cmbox[cl[i]].bind('<<ComboboxSelected>>',lambda event:self.cmbox_bu())
        self.cmbox[u'置信度']['values']=['95%',u'90%',u'80%',u'70%',u'50%']
        self.cmbox[u'置信度'].current(1)
        self.cmbox[u'区间数量']['values']=[u'20',u'40',u'60',u'100']
        self.cmbox[u'区间数量'].current(1)
        self.cmbox[u'最近交易日']['values']=[u'130',u'250',u'500',u'750'] 
        self.cmbox[u'最近交易日'].current(1)
        tk.Label(f,text=u'默认参数为置信度：90%，区间数量：40，最近交易日：250').grid(row=2,column=0,columnspan=len(cl)+2,sticky='w')
        tk.Button(f,text=u'确认幅度',command=self.confirm_bu).grid(row=0,column=len(cl)+1,sticky='w')
        tk.Button(f,text=u'显示全部',command=self.visall_bu).grid(row=1,column=len(cl)+1,sticky='w')
        tk.Button(f,text=u'节日试算幅度',command=self.fes_cal_bu).grid(row=0,column=len(cl)+2,sticky='w')
        tk.Button(f,text=u'节日试算',command=self.calfes_bu).grid(row=1,column=len(cl)+2,sticky='w')
    def get_range(self,mk=0.9,n=40,T=250):
        rdic={}
        hdic={}
        sdic={}
        codeclass=self.codeclass
        nonvari=['WR','FU','B','FB','BB','WH','RI','LR','PM','RS','JR']
        for x in self.codeclass:
            vari=codeclass[x].Vari
            if vari in nonvari:
                continue
            if rdic.has_key(vari):
                if codeclass[x].Inf['open_interest']>hdic[vari]:
                    hdic[vari]=codeclass[x].Inf['open_interest']
                    rdic[vari]=codeclass[x].Inf['price_now']/codeclass[x].Inf['lastcls']-1
                    sdic[vari]=codeclass[x].Inf['price_now']/codeclass[x].Inf['lastsettlement']
            else:
                if codeclass[x].Inf['lastcls']=='' or codeclass[x].Inf['lastcls']==0:
                    continue
                hdic[vari]=codeclass[x].Inf['open_interest']
                rdic[vari]=codeclass[x].Inf['price_now']/codeclass[x].Inf['lastcls']-1
                sdic[vari]=codeclass[x].Inf['price_now']/codeclass[x].Inf['meanprice']
        res_range=sim.return_ndays(rdic,n,T)
        for x in res_range:
            reverse=False#从小到大
            if rdic[x]<0:
                reverse=True#从大到小
            for y in res_range[x]:
                temp=float(y[5])
                y[5]=temp
            res_range[x].sort(reverse=reverse,key=itemgetter(5))
        data=[]
        colname=[u'品种',u'收盘价涨跌幅',u'预测涨跌幅',u'真正试算幅度']
        ngvari='p;j;a;b;m;y;jm;i;rm;sr;ta;ma;oi;cf;fg;tc;zc;cu;al;zn;pb;ag;au;ru;rb;hc;bu;ni;sn'
        ngvari=ngvari.upper().split(';')
        for x in rdic:
            v=[x,round(rdic[x],4)*100]
            lk=math.ceil(mk*len(res_range[x]))
            val=res_range[x][int(lk)][5]
            v.append(round(val,4)*100)
            val2=(val+1)*sdic[x]-1
            if abs(val2)>self.variclass[x].Inf['Limt']:
                val2=val2/abs(val2)*self.variclass[x].Inf['Limt']
            v.append(round(val2,4)*100)
            if x in ngvari and abs(val)>0.02:
                data.append(v)
        self.rdic=rdic
        self.res_range=res_range
        self.sdic=sdic
        return colname,data
    def outdata(self,colname,data):
        if self.fr:
            self.sim_tg=table_gui(colname=colname,data=data,isframe=self.downgui,w=120)
            self.sim_tg.sorttree()
            self.sim_tg.canvas.xview_moveto(0.0)
            self.fr=False 
            self.sim_tg.entry_double_click()
            self.sim_tg.tree.bind('<KeyPress-o>',self.open_detail)
            self.sim_tg.tree.bind('<KeyPress-O>',self.open_detail)
        else:          
            self.sim_tg.resettree(data)
            self.sim_tg.canvas.xview_moveto(0.0)
    def cmbox_bu(self):
        cl=[u'置信度',u'区间数量',u'最近交易日']
        mk=float(self.cmbox[u'置信度'].get().split('%')[0])/100
        n=int(self.cmbox[u'区间数量'].get())
        T=int(self.cmbox[u'最近交易日'].get())
        colname,data=self.get_range(mk=mk,n=n,T=T)
        self.outdata(colname,data)
    def confirm_bu(self):
        codeclass=self.codeclass
        tree=self.sim_tg.tree
        items=tree.get_children()
        dv={}
        for x in items:
            dv[tree.set(x,0)]=float(tree.set(x,3))/100
        for x in codeclass:
            vari=codeclass[x].Vari
            if vari in dv.keys():
                codeclass[x].Inf['delta_price']=dv[vari]
        if dv:
            res=tkMessageBox.askquestion(u"检测到合约结算价有调整", u"合约是否停板？")
            if res=='yes':
                self.parent.data['delta_price_format']=1#代表停板
            else:
                self.parent.data['delta_price_format']=0#代表统一向下
            self.parent.data['is_delta_price']=True           
    def visall_bu(self):
        mk=float(self.cmbox[u'置信度'].get().split('%')[0])/100
        data=[]
        colname=[u'品种',u'收盘价涨跌幅',u'预测涨跌幅',u'真正试算幅度']
        rdic=self.rdic
        sdic=self.sdic
        res_range=self.res_range
        for x in rdic:
            v=[x,round(rdic[x],4)*100]
            lk=math.ceil(mk*len(res_range[x]))
            val=res_range[x][int(lk)][5]
            v.append(round(val,4)*100)
            val2=(val+1)*sdic[x]-1
            if abs(val2)>self.variclass[x].Inf['Limt']:
                val2=val2/abs(val2)*self.variclass[x].Inf['Limt']
            v.append(round(val2,4)*100)
            data.append(v)
        self.outdata(colname,data)
    def open_detail(self,event):
        tree=self.sim_tg.tree
        rowid = tree.identify_row(event.y)
        vari=tree.item(rowid)['values'][0]
        outdata=self.res_range[vari]
        for i in range(len(outdata)):
            for j in range(2,7):
                outdata[i][j]=round(float(outdata[i][j])*100,2)
        colname=[u'ext_code',u'交易日期',u'当日涨跌幅',u'次日最大波动',u'次日结算涨跌幅',u'次日收盘涨跌幅',u'次日开盘幅度']
        f=self.parent.add_note(title=vari+u'历史涨跌幅')
        tg=table_gui_sp(colname=colname,data=outdata,isframe=f,w=120)
        tg.sorttree()      
        items = tg.tree.get_children()
        tree=tg.tree
        n=len(items)
        lk1=int(math.floor(0.01*n))
        lk2=int(math.floor(0.05*n))
        lk3=int(math.floor(0.1*n))
        lk4=int(math.floor(0.5*n))
        lk=[lk1,lk2,lk3,lk4]
        for x in lk:
            tag=tree.item(items[x])['tags']
            if tag=='':
                tag=['red2']
            else:
                tag=tag+['red2']
            tree.item(items[x],tags=tag)
        tree.tag_configure('red2',foreground='red')
    def fes_cal_bu(self):
        res=tkMessageBox.askquestion(u"是否进行节日板辐试算？",u"是否进行国庆和春节板辐试算，若是请将调保文件放置个人路径，并输入文件名")
        if res:
            fname=tkSimpleDialog.askstring(u'华泰期货',u'请输入文件名',initialvalue ='2016年国庆调保.xls')
        #self.parent.time_count()
        td,colname=sim.pricevar(self.variclass)
        if res:
            td,colname=sim.pricelimt(td,colname,self.variclass,self.parent.path+'\\'+fname)
        outdata=[]
        for x in td:
            outdata.append([x]+td[x])
        f=self.parent.add_note(title=u'假期试算幅度')
        tg=table_gui(colname=colname,data=outdata,isframe=f,w=90)
        tg.sorttree()
        self.cal_colname=colname
        self.cal_data=td
        #self.parent.time_count()
    def calfes_bu(self):
        colname=self.cal_colname
        cal_r=self.cal_data
        ms=u'请输入试算标准：\n'
        for i in range(1,len(colname)):
            ms+=str(i)+u' FOR '+colname[i]+'\n'
        i=tkSimpleDialog.askinteger(u'华泰期货',ms,initialvalue =1)      
        codeclass=self.parent.data['ini_codeclass']
        for x in codeclass:
            vari=codeclass[x].Vari
            if cal_r.has_key(vari):
                codeclass[x].Inf['delta_price']=cal_r[vari][i-1]
        self.parent.data['is_delta_price']=True
        self.parent.data['delta_price_format']=0
        self.parent.bsh_bu(event=True)
class table_gui_sp(table_gui):
    def treeview_sort_column(self,tv, col, reverse,hl):
        table_gui.treeview_sort_column(self,tv, col, reverse,hl)
        tree=self.tree
        items=tree.get_children()
        for x in items:
            tag=tree.item(x)['tags']
            if 'red2' in tag:
                tag.remove('red2')
                tree.item(x,tags=tag)
        n=len(items)
        lk1=int(math.floor(0.01*n))
        lk2=int(math.floor(0.05*n))
        lk3=int(math.floor(0.1*n))
        lk4=int(math.floor(0.5*n))
        lk=[lk1,lk2,lk3,lk4]
        for x in lk:
            tag=tree.item(items[x])['tags']
            if tag=='':
                tag=['red2']
            else:
                tag=tag+['red2']
            tree.item(items[x],tags=tag)
        tree.tag_configure('red2',foreground='red')            
        