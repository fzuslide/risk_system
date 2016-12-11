# -*- coding: utf-8 -*-
"""
Created on Sun Jun 05 12:04:20 2016

@author: IORI
"""
from __future__ import division
import Tkinter as tk
import tkSimpleDialog
import tkFileDialog
import ttk
import functions as fn
import functions_2 as fn2
import daily_report
import datetime
import time
import tkMessageBox
import pickle
import xlrd
import os
import sys
from margin_gui import *
from settlement_gui import *
from table_gui2 import *
from single_gui import *
from bsh_gui import*
from simulation_gui import*
import realtime_market_monitor
class iorigui:
    def __init__(self):
        self.cloudpath=u'\\\\10.100.6.20\\fkfile\\python_risksystem\\data'
        self.path=u'\\\\10.100.6.20\\fkfile\\IORI-陈志荣\\三表'
        self.datapath=os.getcwd()+'\\data\\'
        self.time_inf={'tag':0,'starttime':'','endtime':''}
        self.message=True#控制计时候是否弹出对话框
        self.isaddrate=True#控制计算大表时是否进行保证金调整
        self.labels={}
        self.buttons={}
        self.data={}
        self.data['delta_price_format']=0#代表统一向下
        self.data['is_delta_price']=False#合约结算价有调整
        self.data['price_format']='settlement'#价格选取     
        self.root=tk.Tk()  
        self.root.title("华泰期货有限公司-风险管理部_by IORI")
        self.root.iconbitmap(self.datapath+'ht_48X48.ico')
        self.root.config(bg='Azure')
        self.root.protocol("WM_DELETE_WINDOW", self.shutdown_ttk_repeat)
        sizex = 800
        sizey = 300
        posx  = 600
        posy  = 100
        #self.root.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, posx, posy))    
        self.create_menu()
        self.creatleftgui()
        self.cleftdown_gui()
        self.createrightgui()    
        self.checkdir()
        #self.outdata_bu('')
    def shutdown_ttk_repeat(self):
        self.root.eval('::ttk::CancelRepeat')
        self.root.destroy()
    def mainloop(self):
        self.root.mainloop()  
    def create_menu(self):
        menubar = tk.Menu(self.root)
        sysmenu =tk.Menu(menubar, tearoff=0)
        rdvar=tk.StringVar()
        self.system_mode='cloud'
        sysmenu.add_radiobutton(label=u"云模式",variable=rdvar,value='cloud',command=lambda :self.system_mode_bu(rdvar))
        sysmenu.add_radiobutton(label=u"本地模式",variable=rdvar,value='local',command=lambda :self.system_mode_bu(rdvar))
        rdvar.set(self.system_mode)
        menubar.add_cascade(label="System Mode", menu=sysmenu)
        
        calcmenu =tk.Menu(menubar, tearoff=0)
        calcmenu.add_command(label=u"导三表+中金所结算价+大表试算（均价）",command=self.autocalc_menubu)
        calcmenu.add_command(label=u"小表",command=lambda :self.ssh_bu(None))
        calcmenu.add_command(label=u"大表",command=lambda :self.bsh_bu(True))
        calcmenu.add_command(label=u"大表（忽略调保）",command=self.bsh_unr_menubu)
        menubar.add_cascade(label="快捷试算功能", menu=calcmenu)
        
        self.root.config(menu=menubar)
        #menubar.grid(row=0,column=0,columnspan=5,sticky='w')
    def autocalc_menubu(self):
        ''' 函数说明：快捷功能菜单,三键齐按'''
        t1= datetime.datetime.now()
        self.message=False
        tg=self.sh3_bu(False)
        if not tg:
            self.message=True
            return
        self.cfesetl_bu(False)
        self.bsh_bu(False)
        t2= datetime.datetime.now()
        self.message=True 
        tkMessageBox.showinfo(title=u'温馨提醒',message=str((t2-t1).seconds)+u's  运行完成')
    def bsh_unr_menubu(self):
        ''' 函数说明：快捷功能菜单,大表（忽略调保）'''
        self.isaddrate=False
        self.bsh_bu(True)
        
    def system_mode_bu(self,rdvar):
        ''' 函数说明：System Mode菜单'''
        self.system_mode=rdvar.get()
    def creatleftgui(self):
        ''' 函数说明：创建左侧gui'''
        leftframe=tk.LabelFrame(self.root,bg='Azure')
        leftframe.grid(row=0,column=0,sticky='NW')
        gridrow=0
        tk.Label(leftframe,text=u'运行状态:',fg='blue').grid(row=0,column=0,sticky='W')        
        self.labels['statu']=tk.Label(leftframe,text='null')
        self.labels['statu'].grid(row=gridrow,column=1,columnspan = 3,sticky='W')
        gridrow+=1
        rdvar=tk.StringVar()
        self.radio_str=rdvar
        rb=tk.Radiobutton(leftframe, text=u"数据管理",fg='blue', variable=rdvar, value=['data_manage'],command =self.radio_bu)
        rb.grid(row=gridrow+0,sticky='W')
        tk.Radiobutton(leftframe, text=u"参数管理",fg='blue', variable=rdvar, value=['cof_manage'],command =self.radio_bu).grid(row=gridrow+1,sticky='W')
        tk.Radiobutton(leftframe, text=u"试算功能",fg='blue', variable=rdvar, value=['cal_fun'],command =self.radio_bu).grid(row=gridrow+2,sticky='W')
        tk.Radiobutton(leftframe, text=u"监控功能",fg='blue', variable=rdvar, value=['monitor_fun'],command =self.radio_bu).grid(row=gridrow+3,sticky='W')
        tk.Radiobutton(leftframe, text=u"其他功能",fg='blue', variable=rdvar, value=['other_fun'],command =self.radio_bu).grid(row=gridrow+4,sticky='W')
        tk.Radiobutton(leftframe, text=u"监控投放",fg='blue', variable=rdvar, value=['btv_monitor_fun'],command =self.radio_bu).grid(row=gridrow+5,sticky='W')
        rb.select()        
        bulist=[u'初始化',u'导三表',u'调保',u'调结算价',u'小表',u'大表',u'席位资金',u'中金所结算价',u'存储数据'
        ,u'读取数据',u'行情系数',u'刷新行情',u'test',u'导出大表',u'导出大表',u'投资者保障基金',u'待处理客户'
        ,u'拆分导出',u'盘后试算',u'打开本地途径',u'行情监控',u'近月不活跃持仓监控',u'月底公司上浮复位',u'修复缺失行情'
        ,u'运行状态复位',u'质押客户监控',u'可出大于可用',u'每日报告',u'客户超仓监控',u'实时行情预警',u'自动行情预警']
        self.bulist=bulist
        for x in bulist:
            bu=tk.Button(leftframe,text=x)
            self.setwidget(bu)
            self.buttons[x]=bu
        self.bindfun()
        self.radio_bu()
    def bindfun(self):
        ''' 函数说明：按钮函数绑定'''
        self.buttons[u'初始化'].bind('<Button-1>',self.ini_bu) 
        self.buttons[u'导三表'].bind('<Button-1>',self.sh3_bu) 
        self.buttons[u'调保'].bind('<Button-1>',self.adjmar_bu)
        self.buttons[u'调结算价'].bind('<Button-1>',self.adjsetl_bu) 
        self.buttons[u'小表'].bind('<Button-1>',self.ssh_bu) 
        self.buttons[u'大表'].bind('<Button-1>',self.bsh_bu) 
        self.buttons[u'席位资金'].bind('<Button-1>',self.seatmoney_bu) 
        self.buttons[u'中金所结算价'].bind('<Button-1>',self.cfesetl_bu) 
        self.buttons[u'存储数据'].bind('<Button-1>',self.indata_bu) 
        self.buttons[u'读取数据'].bind('<Button-1>',self.outdata_bu) 
        self.buttons[u'行情系数'].bind('<Button-1>',self.coefficient_bu) 
        self.buttons[u'刷新行情'].bind('<Button-1>',self.reflash_marketdata_bu) 
        self.buttons[u'导出大表'].bind('<Button-1>',self.outbsh_bu)
        self.buttons[u'投资者保障基金'].bind('<Button-1>',self.invfound_bu)
        self.buttons[u'待处理客户'].bind('<Button-1>',self.riskclient_bu)
        self.buttons[u'拆分导出'].bind('<Button-1>',self.splitout_bu)
        self.buttons[u'test'].bind('<Button-1>',self.test_bu) 
        self.buttons[u'盘后试算'].bind('<Button-1>',self.simulation_bu)
        self.buttons[u'打开本地途径'].bind('<Button-1>',self.open_dir_bu)
        self.buttons[u'行情监控'].bind('<Button-1>',self.codemonitor_bu)
        self.buttons[u'近月不活跃持仓监控'].bind('<Button-1>',self.unactive_monitor_bu)
        self.buttons[u'月底公司上浮复位'].bind('<Button-1>',self.recor_rate_bu)
        self.buttons[u'修复缺失行情'].bind('<Button-1>',self.recover_data_bu)
        self.buttons[u'运行状态复位'].bind('<Button-1>',self.ini_statu_bu)
        self.buttons[u'质押客户监控'].bind('<Button-1>',self.mon_ple_bu)
        self.buttons[u'可出大于可用'].bind('<Button-1>',self.out_left_bu)
        self.buttons[u'每日报告'].bind('<Button-1>',self.daily_report_bu)
        self.buttons[u'客户超仓监控'].bind('<Button-1>',self.cl_pos_bu)
        self.buttons[u'实时行情预警'].bind('<Button-1>',self.realtime_warning_bu)
        self.buttons[u'自动行情预警'].bind('<Button-1>',self.auto_warning_bu)
    def setwidget(self,wid,w=15,h=1,fg='red'):
        ''' 函数说明：按钮参数设置'''
        wid['width']=w
        wid['height']=h
        wid['fg']=fg
        #b1['background']='blue'        
    def radio_bu(self):
        ''' 函数说明：左侧gui radion函数功能，显示按钮'''
        budic={}
        budic['data_manage']=[u'初始化',u'导三表',u'修复缺失行情',u'打开本地途径']
        budic['cof_manage']=[u'调保',u'调结算价',u'行情系数',u'中金所结算价',u'月底公司上浮复位']#,u'刷新行情']
        budic['cal_fun']=[u'小表',u'大表',u'席位资金',u'投资者保障基金',u'盘后试算']
        budic['monitor_fun']=[u'待处理客户',u'行情监控',u'近月不活跃持仓监控',u'质押客户监控',u'可出大于可用',u'客户超仓监控',u'每日报告']
        budic['other_fun']=[u'导出大表',u'拆分导出',u'运行状态复位',u'存储数据',u'读取数据',u'test']
        budic['btv_monitor_fun']=[u'实时行情预警',u'自动行情预警']
        for x in self.bulist:
            self.buttons[x].grid_forget()
        gr=1
        for x in budic[self.radio_str.get()]:
            self.buttons[x].config(relief='raised')
            self.buttons[x].grid(row=gr,column=1)
            gr+=1       
    def cleftdown_gui(self):
        ''' 函数说明：创建左下侧gui'''
        ld_frame=tk.LabelFrame(self.root,bg='Azure')
        ld_frame.grid(row=1,column=0,sticky='NW')
        for x in range(9):
            self.labels[x]=tk.Label(ld_frame,text='')
            #self.labels[x].grid(row=gridrow,column=0,columnspan = 4,sticky='E')
            self.labels[x].grid(row=x+1,column=0,sticky='E')                
    def createrightgui(self):   
        ''' 函数说明：创建右侧gui'''
        rightframe=tk.LabelFrame(self.root,bg='Azure')
        rightframe.grid(row=0,column=1,rowspan=2,sticky='NW')
        note=ttk.Notebook(rightframe)
        note.bind('<Escape>',self.note_delte_bu)
        note.bind('<Delete>',self.note_delte_bu)
        note.bind('<Double-Button-1>',self.note_delte_bu)
        note.grid(row=0,column=0,sticky='NW')
        self.note=note
        self.mk_frame = self.add_note(title=u'合约要素')
        self.mk_fr=True      
        self.bsh_class=bsh_gui(self)
    def note_delte_bu(self,event):
        ''' 函数说明：右侧note删除功能的函数'''
        a=self.note.tabs()
        sel=self.note.select()
        if sel==a[0] or sel==a[1]:
            return
        else:
            self.note.forget(sel)
    def add_note(self,title='newnote',f=1):
        ''' 函数说明：增加note的函数'''
        frame=tk.Frame(self.note,bg='Azure')
        self.note.add(frame,text = title)
        self.note.select(self.note.tabs()[-1])
        if f==1:
            return frame
        else:
            f1=tk.Frame(frame,bg='Azure')
            f2=tk.Frame(frame,bg='Azure')
            f1.grid(row=0,column=0,sticky='w')
            f2.grid(row=1,column=0,sticky='w')
            return f1,f2
    def checkdir(self):
        ''' 函数说明：首次打开设置本地路径'''
        dirinfo=self.datapath+'dir.pickle'
        if os.path.exists(dirinfo):
            with open(dirinfo, 'rb') as f:self.path = pickle.load(f)
        else:
            path=tkSimpleDialog.askstring(u'华泰期货',u'请设置三表路径，用于本地模式以及大表导出（路径用处较多，尽量不要默认确定）。请输入个人三表文件夹存放路径：      ',initialvalue = r'\\10.100.6.20\fkfile\IORI-陈志荣\三表')
            self.path=path
            with open(dirinfo, 'wb') as f:pickle.dump(self.path, f)            
    def out_table(self,tabdata,alldeparment,allvari):
        ''' 函数说明：大表显示'''
        rowname=tabdata.keys()
        rowname.sort()
        #colname=data[rowname[0]].keys()
        rcolname=['invid','inv_name','invdepartment','spmark','riskdegree','cor_riskdegree','leftcapital','cor_leftcapital','seat','margin','capital','lastriskstate','maxcode','posstrut','forcedbound','phone']        
        data=[]
        tags=[]
        for i in range(len(rowname)):
            v=[]
            for x in rcolname:
                temp=tabdata[rowname[i]][x]
                if type(temp)==float:
                    temp=fn.float_to_str(round(temp,2))
                v.append(temp)
            data.append(v)
            tags.append(tabdata[rowname[i]]['tags'])
        imf=[0,0,0,0,0]
        for i in range(len(rowname)):
            if float(tabdata[rowname[i]]['leftcapital'])<0:
                imf[0]+=float(tabdata[rowname[i]]['leftcapital'])
                imf[1]+=1
                if tabdata[rowname[i]]['spmark']=='1':
                    imf[2]+=1
                if float(tabdata[rowname[i]]['capital'])<0:
                    imf[3]+=float(tabdata[rowname[i]]['capital'])
                    imf[4]+=1
            else:
                break
        tt=u'强平金额：'+fn.float_to_str(imf[0])+u'，强平人数：'+fn.float_to_str(imf[1])+u'，其中特殊客户数：'+fn.float_to_str(imf[2])+u'。穿仓金额：'+fn.float_to_str(imf[3])+u'，穿仓人数：'+fn.float_to_str(imf[4])
        #self.labels['bsh_statu'].config(text=tt,fg='red')
        title=u'客户资金风险信息     '+tt+u'    计算时间：'+datetime.datetime.now().strftime('%H:%M')
        
        self.bsh_class.inputdata(title,rcolname,data,tags=tags,alldeparment=alldeparment,allvari=allvari,invclass=self.data['invclass'])
        self.note.select(self.note.tabs()[1])  
        self.time_count()  
    '''          
    def mul_ans(self,k1=0,k2=0):
        #''' '''函数说明：四个计算价格提示框''''''
        t=tk.Toplevel()
        xy=t.winfo_pointerxy()
        sizex = 260
        sizey = 35
        t.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, xy[0], xy[1]))
        t.title(u'请选择价格类型')
        tk.Button(t,text=u'结算价',width=8,height=1,fg='blue',command=lambda :self.getans(t,u'结算价',k1,k2)).grid(row=0,column=0)
        tk.Button(t,text=u'均价',width=8,height=1,fg='blue',command=lambda :self.getans(t,u'均价',k1,k2)).grid(row=0,column=1)
        tk.Button(t,text=u'最新价',width=8,height=1,fg='blue',command=lambda :self.getans(t,u'最新价',k1,k2)).grid(row=0,column=2)
        tk.Button(t,text=u'昨结价',width=8,height=1,fg='blue',command=lambda :self.getans(t,u'昨结价',k1,k2)).grid(row=0,column=3)
        #t.mainloop()
    '''
    def getans(self,t,s,k1,k2):     
        ''' 函数说明：计算大表'''
        self.time_count()
        dic={u'结算价':'settlement',u'均价':'meanprice',u'最新价':'price_now',u'昨结价':'lastsettlement'}
        if t<>'':        
            tk.Toplevel.destroy(t)
        self.data['price_format']=dic[s]
        invclass,tabdata,alldeparment,allvari=fn.allclient_cal(self.data,k1,k2,self.data['price_format'],isaddrate=self.isaddrate)
        self.isaddrate=True
        self.data['invclass2']=invclass       
        self.out_table(tabdata,alldeparment,allvari)               
    def ini_bu(self,event):
        ''' 函数说明：初始化按钮'''
        self.time_count()
        if self.system_mode=='local':
            fn.read_baicsh(self,pak=3)
            self.data['delta_price_format']=0#代表结算价有调整时统一向下方式
            self.data['is_delta_price']=False#代表是否有结算价调整
        elif self.system_mode=='cloud':
            fname=self.cloudpath+'\\data.pickle'
            with open(fname, 'rb') as f:self.data = pickle.load(f)
            f.close()
            for i in range(len(self.data['sheet_info'])):
                self.labels[i+3].config(text=self.data['sheet_info'][i])
        self.time_count()
    def sh3_bu(self,event):  
        ''' 函数说明：导三表按钮'''
        self.time_count()
        if self.system_mode=='local':
            invfname=self.path+'\\'+u'投资者资金信息.xls'
            posfname=self.path+'\\'+u'持仓查询.xls'
            riskfname=self.path+'\\'+u'实时风控行情.xls'
            fn.read_sh3(invfname,posfname,riskfname,self)
            fn.creat_class(self.data)
            fn.match_corp_marr(self.data['invclass'],self.data['variclass'],self.data['ini_sprate'])
            invclass=self.data['invclass']
            ini_seat=self.data['ini_seat']
            ini_belong=self.data['ini_belong']
            ini_phone=self.data['ini_phone']
            ini_invtype=self.data['ini_invtype']
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
        elif self.system_mode=='cloud':
            fname=self.cloudpath+'\\sh3_data.pickle'
            finfo=self.cloudpath+'\\readable.txt'
            isg=False
            for i in range(40):
                if os.path.exists(finfo):
                    with open(fname, 'rb') as f:sh3_data = pickle.load(f)
                    f.close()
                    codeclass=copy.deepcopy(self.data['ini_codeclass'])
                    for x in sh3_data:
                        self.data[x]=sh3_data[x]
                    for x in self.data['ini_codeclass']:
                        self.data['ini_codeclass'][x].Inf['delta_rate']=codeclass[x].Inf['delta_rate']
                        self.data['ini_codeclass'][x].Inf['delta_price']=codeclass[x].Inf['delta_price']
                    isg=True
                    for i in range(len(self.data['sheet_info'])):
                        self.labels[i].config(text=self.data['sheet_info'][i])
                        if i <3:
                            self.labels[i].config(fg='red')
                    break
                time.sleep(1)
            if not isg:                
                tkMessageBox.showinfo(title=u'温馨提醒',message=u'找不到数据，请选择本地模式')
                self.time_count(False)
                return False
        self.time_count()
        return True
    def bsh_bu(self,event):
        ''' 函数说明：大表按钮'''
                
        k1=0
        k2=0       
        if self.data['is_delta_price']:
            res=tkMessageBox.askquestion(u"检测到合约结算价有调整", u"是否进行调整？")
            if res=='yes':
                k1=self.data['delta_price_format']
                k2=self.data['is_delta_price'] 
                #k1最小变动价位选取方向，k2是否引入价格变动
        #self.mul_ans(k1,k2)
        if event:
            t=tk.Toplevel()
            xy=t.winfo_pointerxy()
            sizex = 260
            sizey = 35
            t.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, xy[0], xy[1]))
            t.title(u'请选择价格类型')
            tk.Button(t,text=u'结算价',width=8,height=1,fg='blue',command=lambda :self.getans(t,u'结算价',k1,k2)).grid(row=0,column=0)
            tk.Button(t,text=u'均价',width=8,height=1,fg='blue',command=lambda :self.getans(t,u'均价',k1,k2)).grid(row=0,column=1)
            tk.Button(t,text=u'最新价',width=8,height=1,fg='blue',command=lambda :self.getans(t,u'最新价',k1,k2)).grid(row=0,column=2)
            tk.Button(t,text=u'昨结价',width=8,height=1,fg='blue',command=lambda :self.getans(t,u'昨结价',k1,k2)).grid(row=0,column=3)
        else:
            self.getans('',u'均价',k1,k2)
        
    def ssh_bu(self,event,ivid=''):
        ''' 函数说明：小表按钮'''
        if ivid=='':        
            ivid=tkSimpleDialog.askstring('HUATAI FUTURE',u'请输入客户号',initialvalue = u'10802698')
        if ivid==None:
            return
        if not self.data['invclass'].has_key(ivid):
            tkMessageBox.showinfo(title=u'温馨提醒',message=u'找不到投资者'+ivid)
            return        
        ssh_frame=tk.Frame(self.note,bg='Azure')
        self.note.add(ssh_frame,text=ivid)
        a=single_gui(data=self.data,isframe=ssh_frame,ivid=ivid)
        self.note.select(self.note.tabs()[-1])
        #a.mainloop()                    
    def adjmar_bu(self,event):
        ''' 函数说明：调保按钮'''
        margin_gui(self.data,self.path)
    def adjsetl_bu(self,event):
        ''' 函数说明：调结算价按钮'''
        settlement_gui(self.data['ini_codeclass'],self.data)
    def coefficient_bu(self,event):
        ''' 函数说明：行情系数按钮'''
        self.time_count()
        codeclass=self.data['ini_codeclass']
        colname=['code','r_mrate','Mrate','delta_rate','r_price','settlement','meanprice','price_now','lastsettlement','delta_price','limtup','limtdown']
        outdata=[]
        kk=codeclass.keys()
        kk.sort()
        for x in kk:
            for y in colname:
                if y=='code':
                    td=[x]
                elif y=='r_mrate':
                    td.append(codeclass[x].Inf['Mrate']+codeclass[x].Inf['delta_rate'])    
                elif y=='r_price':
                    td.append(codeclass[x].Price)
                elif y=='vari':
                    td.append(codeclass[x].Vari)
                elif y=='house':
                    td.append(codeclass[x].House)                    
                else:
                    td.append(codeclass[x].Inf[y])
            outdata.append(td)
        if self.mk_fr:
            self.mk_tg=table_gui(s_title=u'合约行情',colname=colname,data=outdata,isframe=self.mk_frame)
            self.mk_tg.sorttree()
            self.mk_fr=False
        else:
            self.mk_tg.resettree(outdata)
        self.note.select(self.note.tabs()[0])
        self.time_count()
    def indata_bu(self,event):
        ''' 函数说明：存储数据按钮'''
        self.time_count()
        with open(self.datapath+'data.pickle', 'wb') as f:pickle.dump(self.data, f)
        #with open('codeclass.pickle', 'wb') as f:pickle.dump(self.data['ini_codeclass'], f)
        self.time_count()
    def outdata_bu(self,event):
        ''' 函数说明：读取数据按钮'''
        self.time_count()
        with open(self.datapath+'data.pickle', 'rb') as f:self.data = pickle.load(f)
        for i in range(len(self.data['sheet_info'])):        
            self.labels[i].config(text=self.data['sheet_info'][i],fg='orangered')
        #with open('codeclass.pickle', 'rb') as f:self.data['ini_codeclass'] = pickle.load(f)
        self.time_count()
    def seatmoney_bu(self,event):
        ''' 函数说明：席位资金按钮'''
        self.time_count()
        invclass,tabdata,alldeparment,allvari=fn.allclient_cal(self.data,0,1,'settlement')
        colname=[u'交易所',u'结算价与持仓价盈亏',u'保证金调整(负为能释放）',u'合计(正为能释放)']
        hac={'DCE':0,'CFE':0,'SHF':0,'CZC':0}
        ham={'DCE':0,'CFE':0,'SHF':0,'CZC':0}
        hamr={'DCE':0,'CFE':0,'SHF':0,'CZC':0}
        for x in invclass:
            for y in hac:
                s='holdprofit_'+y
                hac[y]+=invclass[x].InvInf[s]
                s='margin_'+y
                ham[y]+=invclass[x].InvInf[s]
        hch={'DCE':'DCE','CFFEX':'CFE','SHFE':'SHF','CZCE':'CZC'}
        for x in self.data['invclass']:
            for pos in self.data['invclass'][x].Position:
                hamr[hch[pos['house']]]+=pos['realmargin']
        od=[]
        for x in hac:
            v=[x]
            v.append(round(hac[x],2))
            v.append(round(ham[x]-hamr[x],2))
            v.append(round(v[1]-v[2],2))
            for i in range(len(v)):            
                v[i]=fn.float_to_str(v[i])
            od.append(v)
        self.time_count()
        tg=table_gui(s_title=u'席位资金',colname=colname,data=od,w=150)      
        tg.mainloop()       
    def cfesetl_bu(self,event):
        ''' 函数说明：中金所结算价按钮'''
        self.time_count()        
        fn.get_CFE_setl(self.data['ini_codeclass'],self.datapath)
        self.time_count()        
    def reflash_marketdata_bu(self,event):
        ''' 函数说明：刷新行情系数按钮'''
        self.time_count()
        fn.get_rtdata(self.data['ini_codeclass'])
        self.time_count()
    def outbsh_bu(self,event):
        ''' 函数说明：导出大表按钮'''
        self.time_count()
        fn.out_inv(self.data['invclass2'],self.bsh_class.bsh_tg.tree,self.path)
        self.time_count()
    def invfound_bu(self,event):
        ''' 函数说明：投资者保障基金按钮'''
        self.time_count()
        filename=self.path+'\\'+u'成交查询.xls'
        outcol,outdata=fn.cal_invfound(filename,self.data['ini_codeclass'])
        tg=table_gui(s_title=u'投资者保障基金',colname=outcol,data=outdata,w=100)
        tg.sorttree()
        self.time_count(False)
    def riskclient_bu(self,event):
        ''' 函数说明：待处理客户按钮'''
        if not self.data.has_key('invclass2'):
            tkMessageBox.showinfo(title=u'温馨提醒',message='请先进行大表试算') 
            return
        self.time_count()        
        invclass=self.data['invclass2']       
        dept_client={}
        for x in invclass:
            inf=invclass[x].InvInf
            dept=inf['invdepartment']
            if inf['riskdegree']>=100 or (inf['lastriskstate'] in [u'追保',u'强平'] and inf['cor_riskdegree']>=100):
                if dept_client.has_key(dept):
                    dept_client[dept].append(x)
                else:
                    dept_client[dept]=[x]
        self.data['risk_client']=dept_client
        frame=tk.Frame(self.note,bg='Azure')
        self.note.add(frame, text = u"待处理风险客户")
        colname=[u'营业部',u'数量']
        n=0
        outdata=[]
        for x in dept_client:
            outdata.append([x,len(dept_client[x])])
            n+=len(dept_client[x])
        outdata.append([u'所有客户',n])
        tg=table_gui(colname=colname,data=outdata,isframe=frame,w=200,parentclass=self)
        tg.sorttree()
        tg.risk_client_double_click()
        self.note.select(self.note.tabs()[-1])
        self.time_count()
    def riskclient_visiable(self,dept):
        ''' 函数说明：按营业部显示风险客户'''
        self.time_count()
        if self.data['risk_client'].has_key(dept):
            ids=self.data['risk_client'][dept]
        else:
            ids=[]
            for x in self.data['risk_client']:
                ids=ids+self.data['risk_client'][x]
        frame=tk.Frame(self.note,bg='Azure')
        self.note.add(frame, text =dept)
        self.note.select(self.note.tabs()[-1])  
        invclass=self.data['invclass2']
        l1=[]
        l2=[]
        for x in ids:
            if invclass[x].InvInf['riskdegree']>=100:
                l1.append((invclass[x].InvInf['riskdegree'],x))
            else:
                l2.append((invclass[x].InvInf['cor_riskdegree'],x))
        l1.sort(reverse=True)
        l2.sort(reverse=True)
        colname=self.bsh_class.bsh_tg.tree['columns']
        outdata=[]
        for x in l1+l2:
            ivid=x[1]
            td=[]
            for y in colname:
                if y=='invid':
                    td.append(ivid)
                    continue
                try:
                    td.append(fn.float_to_str(round(invclass[ivid].InvInf[y],2)))
                except:
                    td.append(invclass[ivid].InvInf[y])
            outdata.append(td)
        
        tg=table_gui(colname=colname,data=outdata,isframe=frame,parentclass=self)
        tg.highlight()
        tg.sorttree()
        tg.bus_double_click()
        tg.canvas.xview_moveto(0.0)
        self.time_count(False)        
            
    def time_count(self,outmessage=True):
        ''' 函数说明：运行时间计算'''
        #self.time_tg=0第一次计算
        if self.time_inf['tag']:
            self.time_inf['endtime']= datetime.datetime.now()
            self.time_inf['tag']=0
            self.labels['statu'].config(text='System Ready',fg='black')          
            self.root.update()
            if outmessage and self.message:
                tkMessageBox.showinfo(title=u'温馨提醒',message=str((self.time_inf['endtime'] - self.time_inf['starttime']).seconds)+u's  运行完成') 
        else:
            self.time_inf['starttime']= datetime.datetime.now()
            self.time_inf['tag']=1
            self.labels['statu'].config(text=u'拼命运行中......',fg='saddlebrown')
            #if event<>'':
            #    event.widget['relief']='sunken'
            self.root.update()
    def splitout_bu(self,event):
        ''' 函数说明：拆分导出按钮'''
        self.time_count()
        a=fn.splitout_inv(self.data['invclass2'],self.bsh_class.bsh_tg.tree,self.path,self.datapath)
        self.time_count()   
    def simulation_bu(self,event):
        ''' 函数说明：盘后试算按钮'''
        self.time_count()
        simulation_gui(self)
        self.time_count()     
    def open_dir_bu(self,event):
        ''' 函数说明：打开本地路径按钮'''
        #path=self.path.replace('\\','/')
        reload(sys)
        sys.setdefaultencoding('utf8')
        #os.system('explorer.exe %s' % self.path.encode('cp936'))
        a=os.popen('explorer.exe %s' % self.path.encode('cp936'))
    def codemonitor_bu(self,event):
        '''函数说明：行情监控按钮'''
        self.time_count()
        res,tags=fn.get_leastdata(self.data['ini_codeclass'])
        colname=[u'合约',u'最新价',u'停板价',u'涨跌幅（%）',u'停板幅度']
        f=self.add_note(title=u'最新行情')
        a=table_gui(colname,res,isframe=f,w=100)
        a.resettree(res,tags)
        a.tree.tag_configure('red',foreground='red')
        a.tree.tag_configure('green',foreground='darkGreen')
        a.sorttree()
        self.time_count()
    def unactive_monitor_bu(self,event):
        '''函数说明：近月不活跃持仓监控按钮'''
        self.time_count()
        codeclass=self.data['ini_codeclass']
        #s1,s2,s3=fn.get_lastmonth()
        mcode=tkSimpleDialog.askstring(u'华泰期货',u'对哪个月份进行监控',initialvalue ='1608,1609')
        vol=tkSimpleDialog.askfloat(u'华泰期货',u'成交量少于多少进行监控',initialvalue =100)
        opn=tkSimpleDialog.askfloat(u'华泰期货',u'持仓量少于多少进行监控',initialvalue =100)
        mmcode=mcode.split(',')
        tarcode=[]
        for x in codeclass:
            if codeclass[x].Inf['volume']<=vol or codeclass[x].Inf['open_interest']<=opn:
                codenum='1'+x[-3:]
                if codenum in mmcode or '*' in mmcode:#and codeclass[x].House<>'CFE':
                    tarcode.append(x)
        invclass=self.data['invclass']
        invclass2=self.data['invclass2']
        outdata=[]
        colname=[u'投资者代码',u'投资者名称',u'交易所风险度',u'合约代码',u'买持',u'卖持',u'持仓结构',u'服务营业部',u'客户类型']
        if tarcode:        
            for x in invclass:
                for posinf in invclass[x].Position:
                    if posinf['code'] in tarcode:
                        pos=[x,invclass[x].Name,round(invclass2[x].InvInf['riskdegree'],2),posinf['code'],posinf['longnums'],posinf['shortnums'],invclass2[x].InvInf['posstrut'],invclass[x].InvInf['invdepartment'],invclass[x].InvInf['type']]
                        outdata.append(pos)
        f=self.add_note(title='不活跃持仓监控')
        a=table_gui(colname,outdata,isframe=f,w=100)
        a.sorttree() 
        self.time_count()       
    def recor_rate_bu(self,event):
        '''函数说明：月底公司上浮复位按钮'''
        self.time_count()
        fn.match_corp_marr(self.data['invclass'],self.data['variclass'],self.data['ini_sprate'],isforce=True)
        self.time_count()
    def recover_data_bu(self,event):
        '''函数说明：修复缺失行情按钮'''
        self.time_count()
        fn.recover_data(self.data['ini_codeclass'],self.data['invclass'])
        self.time_count(False)
    def ini_statu_bu(self,event):
        '''函数说明：运行状态复位按钮'''
        if self.time_inf['tag']==1:
            self.time_count(False)
    def mon_ple_bu(self,event):
        '''函数说明：质押客户监控按钮'''
        if not self.data.has_key('invclass2'):
            tkMessageBox.showinfo(title=u'温馨提醒',message='请先进行大表试算') 
            return        
        self.time_count()
        colname,outdata,colname1=fn.mon_ple_client(self.data['invclass2'])
        f=self.add_note(u'质押客户监控')
        tg=table_gui(colname,outdata,isframe=f,w=100)
        tg.sorttree()
        self.time_count()
    def out_left_bu(self,event):
        '''函数说明：可出大于可用按钮'''
        if not self.data.has_key('invclass2'):
            tkMessageBox.showinfo(title=u'温馨提醒',message='请先进行大表试算') 
            return        
        self.time_count()
        colname,outdata=fn.out_left_client(self.data['invclass'],self.data['invclass2'])
        f=self.add_note(u'可出大于可用客户')
        tg=table_gui(colname,outdata,isframe=f,w=100)
        tg.sorttree()
        self.time_count()        
    def daily_report_bu(self,event):
        ''' 函数说明：每日报告生成按钮'''
        self.time_count()
        daily_report.create_report(self.data,self.datapath,self.path)
        self.time_count() 
    def cl_pos_bu(self,event):
        '''函数说明：客户超仓监控'''
        self.time_count()
        colname,outdata,colname1=daily_report.client_pos_mon(self.data['invclass'],self.data['ini_codeclass'],self.datapath)
        f=self.add_note(u'客户超仓监控')
        tg=table_gui(colname,outdata,isframe=f,w=100)
        tg.sorttree()        
        self.time_count()
    def realtime_warning_bu(self,event):
        '''实时行情预警按钮'''
        a=realtime_market_monitor.market_monitor(self.data['ini_codeclass'])
        a.early_warning()
        self.market_warning=a
    def auto_warning_bu(self,event):
        '''自动预警按钮'''
        self.market_warning.auto_warn()
        
    def test_bu(self,event):
        '''test'''
        self.time_count()
        res=fn.get_leastdata(self.data['ini_codeclass'])
        colname=[u'合约',u'最新价',u'停板价',u'涨跌幅（%）',u'停板幅度']
        f=self.add_note(title=u'最新行情')
        a=table_gui(colname,res,isframe=f)
        a.sorttree()
        self.time_count()