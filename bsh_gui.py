# -*- coding: utf-8 -*-
"""
Created on Mon Aug 01 20:41:31 2016

@author: IORI
"""
import Tkinter as tk
import tkSimpleDialog
import tkFileDialog
import ttk
import functions as fn
import datetime
import time
import tkMessageBox
import pickle
import xlrd
import os
from margin_gui import *
from settlement_gui import *
from table_gui2 import *
from single_gui import *
class bsh_gui():
    def __init__(self,parent):
        self.parent=parent
        frame=tk.Frame(parent.note,bg='Azure')
        parent.note.add(frame, text = u'客户资金信息')
        f1=tk.Frame(frame,bg='Azure')
        f2=tk.Frame(frame,bg='Azure')
        f1.grid(row=0,column=0,sticky='w')
        f2.grid(row=1,column=0,sticky='w')
        self.upgui=f1
        self.downgui=f2
        self.bsh_fr=True
        self.create_bush_upgui()
    def create_bush_upgui(self):
        f=self.upgui
        self.cmbox={}
        cl=[u'持仓方向',u'持仓品种',u'营业部',u'昨风险状态',u'客户类型']
        for i in range(len(cl)):
            self.cmbox[cl[i]]=self.create_cmbox(f,i,cl[i])
        self.cmbox[u'持仓方向']['values']=['ALL',u'多仓',u'空仓']
        self.cmbox[u'昨风险状态']['values']=['ALL',u'强平+追保',u'强平',u'追保']
        self.cmbox[u'客户类型']['values']=['ALL',u'自然人',u'非自然人']
        tk.Label(f,text=u'选中客户数').grid(row=0,column=len(cl),sticky='w')
        self.lbnum=tk.Label(f,text='')
        self.lbnum.grid(row=1,column=len(cl),sticky='w')
        tk.Button(f,text='Mindy营业部',fg='PaleVioletRed',command=self.mindy_bu).grid(row=0,column=len(cl)+1,sticky='w')
        tk.Button(f,text='手输营业部',fg='PaleVioletRed',command=self.handept_bu).grid(row=1,column=len(cl)+1,sticky='w')
        tk.Button(f,text='特殊关注客户',fg='PaleVioletRed',command=self.spclient_bu).grid(row=0,column=len(cl)+2,sticky='w')
        tk.Button(f,text='修改特殊关注',fg='PaleVioletRed',command=self.ch_spclient_bu).grid(row=1,column=len(cl)+2,sticky='w')
        self.lb=tk.Label(f,text='')
        self.lb.grid(row=2,column=0,columnspan=len(cl)+2,sticky='w')
    def create_cmbox(self,f,col,name,value=''):
        tk.Label(f,text=name).grid(row=0,column=col,sticky='w')
        cmbox=ttk.Combobox(f,width=10)
        cmbox['values']=value
        cmbox.grid(row=1,column=col,sticky='w')
        #cmbox.current(0)      
        cmbox.bind('<<ComboboxSelected>>',lambda event,x=name:self.cmbox_bu(x))
        return cmbox
    def cmbox_bu(self,name):
        posdir=self.cmbox[u'持仓方向'].get()
        vari=self.cmbox[u'持仓品种'].get()
        department=self.cmbox[u'营业部'].get()
        status=self.cmbox[u'昨风险状态'].get()
        invtype=self.cmbox[u'客户类型'].get()
        tags=self.tags
        outdata=[]
        pdir={u'多仓':'longnums',u'空仓':'shortnums'}
        isvari=False
        if vari=='ALL':
            isvari=True
        for i in range(len(self.tags)):
            isdata=True
            if not (department=='ALL' or department in tags[i]):
                continue
            if not (vari=='ALL' or vari in tags[i]):
                continue
            if status<>'ALL':
                if status==u'强平+追保':
                    if u'强平' not in tags[i] and u'追保' not in tags[i]:
                        continue
                elif status not in tags[i]:
                    continue
            if invtype<>'ALL':
                if invtype==u'自然人':
                    if self.invclass[self.data[i][0]].InvInf['type']<>u'自然人':
                        continue
                elif self.invclass[self.data[i][0]].InvInf['type']==u'自然人':
                    continue
            if len(self.invclass[self.data[i][0]].Position):
                if posdir=='ALL':
                    outdata.append(self.data[i])
                else:
                    for posinf in self.invclass[self.data[i][0]].Position:
                        if (isvari or posinf['vari']==vari) and posinf[pdir[posdir]]>0:
                                outdata.append(self.data[i])
                                break
        self.bsh_tg.resettree(outdata)
        self.bsh_tg.highlight()
        self.bsh_tg.canvas.xview_moveto(0.0)                                 
        self.lbnum.config(text=str(len(outdata)))                    
    def inputdata(self,title,colname,data,tags='',alldeparment='',allvari='',invclass=''):
        if self.bsh_fr:   
            self.bsh_tg=table_gui(colname=colname,data=data,isframe=self.downgui,parentclass=self.parent,vrows=22)   
            self.bsh_tg.highlight()
            self.bsh_tg.sorttree(hl=True)
            self.bsh_tg.bus_double_click()
            self.bsh_tg.canvas.xview_moveto(0.0)
            self.bsh_fr=False
        else:
            self.bsh_tg.resettree(data)
            self.bsh_tg.highlight()
            self.bsh_tg.canvas.xview_moveto(0.0)
        self.lb.config(text=title,fg='red')
        self.invclass=invclass
        self.data=data
        self.tags=tags
        self.alldeparment=alldeparment
        self.allvari=allvari
        self.setcmbox()
    def setcmbox(self):
        self.alldeparment.sort()
        self.allvari.sort()
        self.cmbox[u'持仓方向'].current(0)  
        self.cmbox[u'昨风险状态'].current(0)
        self.cmbox[u'持仓品种']['values']=['ALL']+self.allvari
        self.cmbox[u'持仓品种'].current(0)
        self.cmbox[u'营业部']['values']=['ALL']+self.alldeparment
        self.cmbox[u'营业部'].current(0)
        self.cmbox[u'营业部'].config(width=15)
        self.cmbox[u'客户类型'].current(0)
    def mindy_bu(self,md=''):
        posdir=self.cmbox[u'持仓方向'].get()
        vari=self.cmbox[u'持仓品种'].get()
        invtype=self.cmbox[u'客户类型'].get()
        alldep=self.cmbox[u'营业部']['values']
        if md=='':
            md=['1108','1131','1133','1178','1165']
        department=[]
        for x in alldep:
            if x[:4] in md:
                department.append(x)
        status=self.cmbox[u'昨风险状态'].get()
        tags=self.tags
        outdata=[]
        pdir={u'多仓':'longnums',u'空仓':'shortnums'}
        isvari=False
        if vari=='ALL':
            isvari=True
        for i in range(len(self.tags)):
            isdata=True
            isde=False
            for x in department:
                if x in tags[i]:
                    isde=True
                    break
            if not isde:
                continue
            if not (vari=='ALL' or vari in tags[i]):
                continue
            if status<>'ALL':
                if status==u'强平+追保':
                    if u'强平' not in tags[i] and u'追保' not in tags[i]:
                        continue
                elif status not in tags[i]:
                    continue
            if invtype<>'ALL':
                if invtype==u'自然人':
                    if self.invclass[self.data[i][0]].InvInf['type']<>u'自然人':
                        continue
                elif self.invclass[self.data[i][0]].InvInf['type']==u'自然人':
                    continue               
            if len(self.invclass[self.data[i][0]].Position):
                if posdir=='ALL':
                    outdata.append(self.data[i])
                else:
                    for posinf in self.invclass[self.data[i][0]].Position:
                        if (isvari or posinf['vari']==vari) and posinf[pdir[posdir]]>0:
                                outdata.append(self.data[i])
                                break
        self.bsh_tg.resettree(outdata)
        self.bsh_tg.highlight()
        self.bsh_tg.canvas.xview_moveto(0.0)                                 
        self.lbnum.config(text=str(len(outdata)))         
    def handept_bu(self):
        dept=tkSimpleDialog.askstring('HUATAI FUTURE',u'Warm：先选定持仓和风险状态，再请输入营业部', initialvalue = '132,158,139,150')
        if dept=='' or dept==None:
            return
        dept=dept.split(',')
        for i,value in enumerate(dept):
            dept[i]='1'+value
        self.mindy_bu(dept)
    def spclient_bu(self):
        path=self.parent.datapath
        spclient=path+'spclient.pickle'
        if os.path.exists(spclient):
            with open(spclient, 'rb') as f:invlist = pickle.load(f)
        else:
            invlist2=tkSimpleDialog.askstring(u'华泰期货',u'首次打开请设置要关注的客户号',initialvalue = '16100007,84700028')
            invlist=invlist2.split(',')
            with open(spclient, 'wb') as f:pickle.dump(invlist, f)
        outdata=[]
        k=0
        for i in range(len(self.data)):           
            if self.data[i][0] in invlist:
                outdata.append(self.data[i])
                k+=1
            if k>len(invlist):
                break
        self.bsh_tg.resettree(outdata)
        self.bsh_tg.highlight()   
        self.bsh_tg.canvas.xview_moveto(0.0)                                 
        self.lbnum.config(text=str(len(outdata)))          
    def ch_spclient_bu(self):
        path=self.parent.datapath
        spclient=path+'spclient.pickle'
        inivalue=''
        if os.path.exists(spclient):
            with open(spclient, 'rb') as f:invlist = pickle.load(f)      
            for x in invlist:
                inivalue=inivalue+','+x
            inivalue=inivalue[1:]
        invlist=tkSimpleDialog.askstring(u'华泰期货',u'请设置要关注的客户号                               ',initialvalue = inivalue)
        invlist=invlist.split(',')
        with open(spclient, 'wb') as f:pickle.dump(invlist, f)