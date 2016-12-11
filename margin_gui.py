# -*- coding: utf-8 -*-
"""
Created on Wed Jun 22 14:27:05 2016

@author: IORI
"""
from __future__ import division
import Tkinter as tk
import tkSimpleDialog
import ttk
import tkMessageBox
import datetime
import functions as fn
import table_gui2
import os
import xlwt
import copy
import xlrd
import codecs
import csv
from mysql_conn import *
class margin_gui:
    def __init__(self,data,path):
        self.codeclass=data['ini_codeclass']
        self.variclass=data['variclass']
        self.path=path
        self.ini_sprate=data['ini_sprate']
        self.data=data
        self.weight={}
        self.framesize=0
        self.framew={}
        self.root=tk.Toplevel()
        self.creat_gui()
        self.root.mainloop()
    def creat_gui(self):
        self.root.title(u'保证金率调整')
        self.root.iconbitmap(os.getcwd()+'\\data\\'+'ht_48X48.ico')
        sizex = 600
        sizey = 440
        posx  = 1600
        posy  = 100
        xy=self.root.winfo_pointerxy()
        self.root.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, xy[0], xy[1]))
        f1=tk.LabelFrame(self.root,relief='groove',heigh=150,width=450,text='f1')#
        f2=tk.LabelFrame(self.root,relief='groove',heigh=290,width=400,text='f2')
        f1.grid(row=0,column=0,sticky='W')
        f2.grid(row=1,column=0,sticky='W')
        rdvar=tk.StringVar()
        tk.Radiobutton(f1, text="CFFEX",fg='blue', variable=rdvar, value=['CFE']).grid(row=0,sticky='W')
        tk.Radiobutton(f1, text="SHFE",fg='blue', variable=rdvar, value=['SHF']).grid(row=1,sticky='W')
        tk.Radiobutton(f1, text="DCE",fg='blue', variable=rdvar, value=['DCE']).grid(row=2,sticky='W')
        tk.Radiobutton(f1, text="CZCE",fg='blue', variable=rdvar, value=['CZC']).grid(row=3,sticky='W')
        rb=tk.Radiobutton(f1, text="ALL",fg='blue', variable=rdvar, value=['CZC','CFE','SHF','DCE'])
        rb.grid(row=4,sticky='W')
        rb.select()
        tk.Label(f1,text=u'品种或合约').grid(row=0,column=1,columnspan=2,sticky='W')
        tk.Label(f1,text=u'调整幅度/调整目标(近月调保)').grid(row=2,column=1,columnspan=2,sticky='W')
        svar1=tk.StringVar()
        svar2=tk.StringVar()        
        self.weight['rdvar']=rdvar
        self.weight['svar1']=svar1
        self.weight['svar2']=svar2
        e1=tk.Entry(f1,textvariable=svar1,width=15)
        e2=tk.Entry(f1,textvariable=svar2,width=15)
        e1.grid(row=1,column=1,columnspan=2,sticky='W')
        e2.grid(row=3,column=1,columnspan=2,sticky='W')
        e1.insert(0, '*609,rm*')
        e2.insert(0, '0')
        b=tk.Button(f1,text=u'查询',fg='red',width=8,command=self.query_bu)
        b.grid(row=4,column=1)
        b=tk.Button(f1,text=u'近月调保',fg='red',width=8,command=self.lastmoth_bu)
        b.grid(row=4,column=2)
        gridcl=4
        b=tk.Button(f1,text=u'显示调保',fg='red',width=10,command=self.visiable_bu)
        b.grid(row=0,column=gridcl)     
        b=tk.Button(f1,text=u'保证金对比',fg='red',width=10,command=self.compare_bu)
        b.grid(row=1,column=gridcl)          
        b=tk.Button(f1,text=u'夜盘调保',fg='red',width=10,command=self.ch_night_bu)
        b.grid(row=2,column=gridcl)   
        b=tk.Button(f1,text=u'确认调整',fg='saddlebrown',width=10,command=self.confirm_bu)
        b.grid(row=3,column=gridcl)  
        b=tk.Button(f1,text=u'退出',fg='saddlebrown',width=10,command=self.exit_bu)
        b.grid(row=4,column=gridcl) 
        
        gridcl=5        
        b=tk.Button(f1,text=u'导出保证金率',fg='saddlebrown',width=10,command=self.outrate_bu)
        b.grid(row=0,column=gridcl)   
        b=tk.Button(f1,text=u'公司标准修改',fg='saddlebrown',width=10,command=self.ch_cor_bu)
        b.grid(row=1,column=gridcl)        
        b=tk.Button(f1,text=u'节日优惠匹配',fg='saddlebrown',width=10,command=self.fes_match_bu)
        b.grid(row=2,column=gridcl)
        b=tk.Button(f1,text=u'节日调保指引',fg='saddlebrown',width=10,command=self.guide_bu)
        b.grid(row=3,column=gridcl)         
        
        gridcl=6
        b=tk.Button(f1,text=u'保证金率核对',fg='blue',width=10,command=self.rate_check_bu)
        b.grid(row=0,column=gridcl)   
        b=tk.Button(f1,text=u'上传保证金率',fg='blue',width=10,command=self.upload_bu)
        b.grid(row=1,column=gridcl)        
        b=tk.Button(f1,text=u'下载保证金率',fg='blue',width=10,command=self.download_bu)
        b.grid(row=2,column=gridcl)
       
        #c=tk.Canvas(f2,heigh=250,width=300)
        tree=ttk.Treeview(f2)
        tree.grid(row=0,column=0)
        ysb=tk.Scrollbar(f2,orient='vertical',command=tree.yview)
        tree.config(yscrollcommand=ysb.set)
        ysb.grid(row=0,column=1,sticky='NS')
        colname=[u'合约代码',u'原保证金率',u'保证金调整',u'调整后保证金率']
        tree.config(columns=colname)
        for x in colname:
            tree.column(x,width=75,anchor='center')
            tree.heading(x,text=x)
        tree.column(u'调整后保证金率',width=90)
        tree.config(displaycolumns='#all',show="headings")
        tree.bind("<Double-1>", self.on_detail_bom_line_db_click)
        self.weight['tree']=tree
    def query_bu(self):
        ''' 函数说明：查询按钮'''
        house=self.weight['rdvar'].get()
        sers=self.weight['svar1'].get().upper().split(',')
        delta=float(self.weight['svar2'].get())/100
        effcode=[x for x in self.codeclass.keys() if self.codeclass[x].House in house]
        rescode=[]
        if type(sers)<>list:
            sers=[sers]
        for x in sers:
            if x[0]=='*':
                dt=x[1:]
                for y in effcode:
                    if dt in y:
                        rescode.append(y)    
            elif x[-1]=='*':
                vari=x[:-1]
                for y in effcode:
                    if self.codeclass[y].Vari==vari:
                        rescode.append(y)
            else:
                if x in effcode:
                    rescode.append(x)
        codelist=list(set(rescode))
        codelist.sort()
        valuelist=[delta for i in range(len(codelist))]
        tree=self.weight['tree']
        for x in tree['columns']:
            tree.heading(x,text=x)
        items=tree.get_children()
        [tree.delete(item) for item in items]
        for i in range(len(codelist)):
            r=self.codeclass[codelist[i]].Inf['Mrate']*100
            v=[codelist[i],r,valuelist[i]*100,r+valuelist[i]*100]
            tree.insert('',i+1,values=v)
    def exit_bu(self):
        ''' 函数说明：退出按钮'''
        self.root.destroy()
    def confirm_bu(self):
        ''' 函数说明：确认调保按钮'''
        tree=self.weight['tree']
        codeclass=self.codeclass
        code=[]
        delta=[]
        for k in tree.get_children(''):
            item=tree.item(k)
            code.append(item['values'][0])
            delta.append(item['values'][2])
        for i in range(len(code)):
            codeclass[code[i]].setvalue('delta_rate',float(delta[i])/100)
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'完成调整')
        self.root.wm_attributes('-topmost',1)
    def visiable_bu(self):
        ''' 函数说明：显示调保按钮'''
        tree=self.weight['tree']
        items=tree.get_children()
        [tree.delete(item) for item in items]
        for x in tree['columns']:
            tree.heading(x,text=x)
        codeclass=self.codeclass
        s=0
        n=0
        clist=codeclass.keys()
        clist.sort()
        for x in clist:
            if codeclass[x].Inf['delta_rate']<>0:
                value=[x,codeclass[x].Inf['Mrate']*100,codeclass[x].Inf['delta_rate']*100
                ,(codeclass[x].Inf['Mrate']+codeclass[x].Inf['delta_rate'])*100]
                tree.insert('','end',values=value)
                s+=codeclass[x].Inf['delta_rate']
                n+=1
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'共调了 '+str(n)+u' 个合约，合计调整幅度：'+str(s))
        self.root.wm_attributes('-topmost',1)
    def ch_night_bu(self):
        ''' 函数说明：夜盘调保按钮'''        
        ncode=tkSimpleDialog.askstring('HUATAI FUTURE',u'请输入调整幅度'
        ,initialvalue = 'p;j;a;b;m;y;jm;i;rm;sr;ta;ma;oi;cf;fg;tc;zc;cu;al;zn;pb;ag;au;ru;rb;hc;bu;ni;sn')
        sr=tkSimpleDialog.askfloat('HUATAI FUTURE',u'请输入调整幅度', initialvalue = '2')
        ncode=ncode.upper().split(';')
        rescode=[]
        tree=self.weight['tree']
        codeclass=self.codeclass
        for x in codeclass:
            if codeclass[x].Vari in ncode:
                rescode.append(x)        
        rescode.sort()
        valuelist=[sr/100 for i in range(len(rescode))]
        tree=self.weight['tree']
        items=tree.get_children()
        [tree.delete(item) for item in items]
        for i in range(len(rescode)):
            r=self.codeclass[rescode[i]].Inf['Mrate']*100
            v=[rescode[i],r,valuelist[i]*100,r+valuelist[i]*100]
            tree.insert('',i+1,values=v)
        self.root.wm_attributes('-topmost',1)
    def lastmoth_bu(self):
        ''' 函数说明：近月调保按钮'''
        tree=self.weight['tree']
        house=self.weight['rdvar'].get()
        sers=self.weight['svar1'].get().upper().split(',')
        tarval=float(self.weight['svar2'].get())
        effcode=[x for x in self.codeclass.keys() if self.codeclass[x].House in house]
        rescode=[]
        if type(sers)<>list:
            sers=[sers]
        for x in sers:
            if x[0]=='*':
                dt=x[1:]
                for y in effcode:
                    if dt in y:
                        rescode.append(y)    
            elif x[-1]=='*':
                vari=x[:-1]
                for y in effcode:
                    if self.codeclass[y].Vari==vari:
                        rescode.append(y)
            else:
                if x in effcode:
                    rescode.append(x)
        codelist=list(set(rescode))
        codelist.sort()
        valuelist=[tarval for i in range(len(codelist))]
        items=tree.get_children()
        [tree.delete(item) for item in items]
        for i in range(len(codelist)):
            r=self.codeclass[codelist[i]].Inf['Mrate']*100
            if r>valuelist[i]:
                valuelist[i]=r
            v=[codelist[i],r,valuelist[i]-r,valuelist[i]]
            tree.insert('',i+1,values=v)
        self.root.wm_attributes('-topmost',1)
    def on_detail_bom_line_db_click(self, event):
        ''' Executed, when a row is double-clicked. Opens
        read-only EntryPopup above the item's column, so it is possible
        to select text '''
        tree=self.weight['tree']
        # close previous popups
        #if tree.entryPopup:
            #tree.entryPopup.destroy()    
        # what row and column was clicked on   
        rowid = tree.identify_row(event.y)
        column = tree.identify_column(event.x)   
        # clicked row parent id
        parent = tree.parent(rowid)
        #print 'parent:'+parent
        # do nothing if item is top-level
        if parent == '':
            pass    
        # get column position info
        x,y,width,height = tree.bbox(rowid, column)   
        # y-axis offset
        pady = height // 2    
        # place Entry popup properly
        url = tree.item(rowid, 'text')   
        tree.entryPopup = StickyEntry(tree, url, width=10)
        tree.entryPopup.place( x=x, y=y+pady, anchor='w')
        tree.entryPopup.bind("<Return>",lambda event,x=tree.entryPopup:self.entry_enter(x,rowid))
    def entry_enter(self,entryPopup,rowid):
        tree=self.weight['tree']
        value=entryPopup.get()
        item=tree.item(rowid)
        vs=[x for x in item['values']]
        if value=='':           
            value=item['values'][2]
        vs[2]=value
        vs[3]=float(value)+float(item['values'][1])
        ttk.Treeview.item(tree,rowid,value=vs)
        entryPopup.destroy()
    def compare_bu(self):
        ''' 函数说明：保证金率对比按钮'''
        import calendar
        res=[]
        for x in self.codeclass:
            if self.codeclass[x].Inf['Mrate']<>self.variclass[self.codeclass[x].Vari].Inf['Mrate']:
                res.append(x)
        res.sort()
        m=datetime.datetime.now()
        m1=m.strftime('%y%m')[1:]
        m2=datetime.datetime(m.year,m.month,1)+datetime.timedelta(calendar.monthrange(m.year,m.month)[1])
        m2=m2.strftime('%y%m')[1:]
        czc=False
        dce=False
        if m.day>=16:
            czc=True
        n=fn.cal_tradeday(datetime.datetime(m.year,m.month,1),m)
        if n>=15:
            dce=True
        tree=self.weight['tree']
        items=tree.get_children()
        [tree.delete(item) for item in items]
        k=1
        for i in range(len(res)):
            if m1 in res[i]:
                continue
            if m2 in res[i]:
                if self.codeclass[res[i]].House=='SHF':
                    continue
                if self.codeclass[res[i]].House=='CZC' and czc:
                    continue
                if self.codeclass[res[i]].House=='DCE' and dce:
                    continue              
            r=self.codeclass[res[i]].Inf['Mrate']*100
            r1=self.variclass[self.codeclass[res[i]].Vari].Inf['Mrate']*100
            d=self.codeclass[res[i]].Inf['delta_rate']*100
            v=[res[i],r1,d,r]
            tree.insert('',k,values=v)       
            k+=1
        tree.heading(u'原保证金率',text=u'品种保证金率')
        tree.heading(u'调整后保证金率',text=u'交易所保证金率')
        self.root.wm_attributes('-topmost',1)
    def outrate_bu(self):
        ''' 函数说明：导出保证金率按钮'''
        codeclass=copy.deepcopy(self.codeclass)
        colname=[u'合约',u'交易所保证金率',u'涨跌停板幅度%',u'合约数量乘数',u'最小变动价位',u'交易所',u'保证金率调整']
        ecol=['code','Mrate','Limt','Units','Minprice','house','delta_rate']
        workbook = xlwt.Workbook(encoding = 'ascii')
        worksheet = workbook.add_sheet(u'new保证金率')
        for i in range(len(colname)):
            worksheet.write(1, i, label =colname[i])
        for i,x in enumerate(codeclass.keys()):
            #codeclass[x].Inf['Mrate']=codeclass[x].Inf['Mrate']+codeclass[x].Inf['delta_rate']
            for j,y in enumerate(ecol):
                if y=='Units':
                    worksheet.write(i+2, j, label =codeclass[x].Units)
                elif y=='house':
                    worksheet.write(i+2, j, label =codeclass[x].House)
                elif y=='code':
                    worksheet.write(i+2, j, label =codeclass[x].Code)
                else:
                    worksheet.write(i+2, j, label =codeclass[x].Inf[y])
        outfilename=self.path+u'\\new保证金率.xls'
        workbook.save(outfilename)
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'完成')
        self.root.wm_attributes('-topmost',1)
    def fes_match_bu(self):
        ''' 函数说明：节日优惠匹配按钮'''
        ini_sprate=self.ini_sprate 
        ms=u'注意格式，保留持仓和保留优惠的数据均从第三列开始\n'
        ms=ms+u'保留持仓审批标准请只有：同意、不同意，不要有额外的文字\n'
        ms=ms+u'保留优惠审批标准请只有这些：维持原标准、交易所+1、交易所+2...，不要有额外的文字\n'
        ms=ms+u'请把国庆（春节）保留保证金、持仓申请表放在设置好的本地路径，并选择。'
        tkMessageBox.showinfo(title=u'温馨提醒',message=ms)
        #fname=tkSimpleDialog.askstring(u'华泰期货',u'请输入文件名',initialvalue ='(总表0929早)2016年国庆保留保证金、持仓申请表.xlsx')
        allfile = os.listdir(self.path)
        xlsfile,ms=[],''
        for x in allfile:
            if '.xls' in x:
                xlsfile.append(x)
                ms=ms+str(len(xlsfile))+u': '+x+'\n'
        fnamei=tkSimpleDialog.askinteger(u'请输入对应序号',ms,initialvalue =0)
        if fnamei==None:
            return
        fname=xlsfile[fnamei-1]
        colname=ini_sprate[0].replace('"','').replace('\r\n','').split(',')
        colk={'vari':colname.index(u'合约代码'),'ivid':colname.index(u'投资者代码')
        ,'mname':colname.index(u'保证金分段名称'),'rate':colname.index(u'投机多头保证金率')
        ,'house':colname.index(u'\ufeff'+u'交易所代码'),'ivr':colname.index(u'投资者范围')}      
        dp=self.path+'\\'+fname
        book = xlrd.open_workbook(dp)
        sh=book.sheets()[0]
        n=sh.nrows   
        res={}
        for i in range(n):
            if sh.cell_value(i,4)==u'同意':
                ivid=str(int(sh.cell_value(i,0)))
                res[ivid]=0
        book = xlrd.open_workbook(dp)
        sh=book.sheets()[1]
        n=sh.nrows
        for i in range(2,n):
            ivid=str(int(sh.cell_value(i,0)))
            if sh.cell_value(i,6)==u'不同意':
                continue
            elif sh.cell_value(i,6)==u'维持原标准':
                res[ivid]=u'维持原标准'
            else:
                res[ivid]=float(sh.cell_value(i,6)[-1])/100
        outdata=[colname]
        temp=copy.deepcopy(res)
        inihouse={}
        for i in range(1,len(ini_sprate)):
            line=ini_sprate[i].replace('"','').replace('\r\n','').split(',')
            ivid=line[colk['ivid']]
            if line[colk['ivr']]==u'公司标准':
                outdata.append(line)   
                inihouse[line[colk['vari']].upper()]=line[colk['house']]
            elif line[colk['house']]==u'CFFEX' and line[colk['ivr']]==u'单一投资者':
                outdata.append(line)
            elif ivid[:3]=='800':
                outdata.append(line)               
            else:        
                if res.has_key(ivid):
                    if res[ivid]<>u'维持原标准':
                        line[colk['rate']]=res[ivid]
                    outdata.append(line)
            if temp.has_key(ivid):
                temp.pop(ivid)       
        allvari=self.variclass.keys()
        for x in temp:
            if temp[x]==u'维持原标准':
                continue
            for vari in allvari:
                if vari in inihouse:
                    ol=copy.deepcopy(line)
                    ol[colk['ivid']]=x
                    ol[colk['vari']]=vari
                    ol[colk['rate']]=0
                    ol[colk['house']]=inihouse[vari]
                    outdata.append(ol)
        '''
        workbook = xlwt.Workbook(encoding = 'ascii')
        worksheet = workbook.add_sheet(u'sheet1')
        for i in range(len(outdata)):
            for j in range(len(outdata[i])):
                worksheet.write(i, j, label =outdata[i][j])
        workbook.save(self.path+u'\\投资者保证金率属性.csv')  
        '''
        import sys
        reload(sys) 
        sys.setdefaultencoding( "utf-8" )
        filename=self.path+u'\\投资者保证金率属性.csv'
        csvFile=file(filename,'wb')
        csvWriter = csv.writer(csvFile,delimiter=',') 
        #csvWriter.writerows(outdata)
        for data in outdata:
            csvWriter.writerow(data)
        csvFile.close()
        f=codecs.open(filename,'rb','utf-8')
        #ini_sprate=f.readlines()
        ini_sprate=[]
        for line in f:
            ini_sprate.append(line)
        f.close()
        self.data['ini_sprate']=ini_sprate
        self.ini_sprate=ini_sprate
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'完成')
        self.root.wm_attributes('-topmost',1)
    def guide_bu(self):
        ''' 函数说明：调保指引按钮'''
        ms=u'国庆和春节调保指引：\n'
        ms=ms+u'  分三步调，第一步将交易所标准调至节时交易所标准。第二步将公司上浮调至节时公司上浮。第三步将重新生成优惠客户标准\n'
        ms=ms+u'  第一步：利用 调保 模块，进行调保，调至节时交易所标准，并按 导出保证金率 按钮导出。所得保证金率表为节时交易所标准，以后可直接读取\n'
        ms=ms+u'  第二步：按 公司标准修改 按钮，生成节时公司标准，新的 投资者保证金率属性 表会存放至所设置的本地路径\n'
        ms=ms+u'  第三步：按 节日优惠匹配 按钮，生成节时优惠客户标准,新的 投资者保证金率属性 表会存放至所设置的本地路径\n'
        ms=ms+u'  读取新的 new保证金表.xls 以及 投资者保证金率属性.csv 进行试算(暂时需手动操作)，所得出的数据交易所线和公司线，均为过节时数据，即是正确对应的标准\n' 
        tkMessageBox.showinfo(title=u'温馨提醒',message=ms)
    def ch_cor_bu(self):
        ''' 函数说明：公司标准修改按钮'''
        ini_sprate=self.ini_sprate        
        ms=u'注意格式，节日调保文件格式如“2016年国庆调保.xls”，不要合并单元格，不要有文字“保持不变”\n'
        ms=ms+u'请把节日调保表放在设置好的本地路径，并选择'
        tkMessageBox.showinfo(title=u'温馨提醒',message=ms)
        #fname=tkSimpleDialog.askstring(u'华泰期货',u'请输入文件名',initialvalue ='2016年国庆调保.xls')
        allfile = os.listdir(self.path)
        xlsfile,ms=[],''
        for x in allfile:
            if '.xls' in x:
                xlsfile.append(x)
                ms=ms+str(len(xlsfile))+u': '+x+'\n'
        fnamei=tkSimpleDialog.askinteger(u'请输入对应序号',ms,initialvalue =0)
        if fnamei==None:
            return
        fname=xlsfile[fnamei-1]        
        codeclass=copy.deepcopy(self.codeclass)
        sdate=datetime.datetime.now()
        edate=sdate-datetime.timedelta(days=sdate.day-1)+datetime.timedelta(days=31)
        edate2=datetime.datetime(edate.year,edate.month,1)+datetime.timedelta(days=31)
        lcd=datetime.datetime.strftime(edate,"%Y%m%d")[2:6]
        llcd=lcd[1:]
        blcd=datetime.datetime.strftime(edate2,"%Y%m%d")[2:6]
        lblcd=blcd[1:]        
        
        colname=ini_sprate[0].replace('"','').replace('\r\n','').split(',')
        colk={'vari':colname.index(u'合约代码'),'ivid':colname.index(u'投资者代码')
        ,'mname':colname.index(u'保证金分段名称'),'rate':colname.index(u'投机多头保证金率')
        ,'house':colname.index(u'\ufeff'+u'交易所代码'),'ivr':colname.index(u'投资者范围')}          
        dp=self.path+'\\'+fname
        book = xlrd.open_workbook(dp)
        sh=book.sheets()[0]
        n=sh.nrows   
        fesrate={}#节时交易所标准
        fescor={}#节时公司标准
        delcor={}#节前公司上浮
        for i in range(4,n):
            if sh.cell_value(i,1).upper()=='IF':
                break
            vari=sh.cell_value(i,1).upper()
            fesrate[vari]=sh.cell_value(i,5)
            fescor[vari]=sh.cell_value(i,6)
            delcor[vari]=sh.cell_value(i,4)-sh.cell_value(i,3)
        for x in codeclass:
            if codeclass[x].House<>'CFE' and codeclass[x].Inf['Mrate']<fesrate[codeclass[x].Vari]:
                codeclass[x].Inf['Mrate']=fesrate[codeclass[x].Vari]
            if x[-3:]==llcd:
                a=0.2
                if codeclass[x].House=='SHF':
                    a=0.15
                if codeclass[x].Inf['Mrate']<a:
                    codeclass[x].Inf['Mrate']=a
            if x[-3:]==lblcd and codeclass[x].House=='SHF':
                a=0.10
                if codeclass[x].Inf['Mrate']<a:
                    codeclass[x].Inf['Mrate']=a                
        outdata=[colname] 
        for i in range(1,len(ini_sprate)):
            line=ini_sprate[i].replace('"','').replace('\r\n','').split(',')
            ivid=line[colk['ivid']]
            if line[colk['ivr']]==u'公司标准' and line[colk['house']]<>'CFFEX':
                vari=line[colk['vari']].upper()
                if len(vari)>2:
                    line[colk['rate']]=fescor[codeclass[vari].Vari]-codeclass[vari].Inf['Mrate']                   
                    if vari[-3:]==llcd or (vari[-3:]==lblcd and codeclass[vari].House=='SHF'):
                        if delcor[codeclass[vari].Vari]>line[colk['rate']]:
                            line[colk['rate']]=delcor[codeclass[vari].Vari]
                        '''
                        if vari=='CU1611':
                            print 'fescor',fescor[codeclass[vari].Vari]
                            print 'fes',codeclass[vari].Inf['Mrate']
                            print 'old',delcor[codeclass[vari].Vari]
                        '''
                elif vari in fescor.keys():                    
                    if line[colk['mname']]==u'上市月后含1个交易日':
                        line[colk['rate']]=fescor[vari]-fesrate[vari]
                    if line[colk['mname']]==u'交割月后含1个交易日前2个交易日' or line[colk['mname']]==u'交割月后含1个公历日前2个交易日':
                        line[colk['rate']]=delcor[vari]
            outdata.append(line) 
        import sys
        reload(sys) 
        sys.setdefaultencoding( "utf-8" )
        filename=self.path+u'\\投资者保证金率属性.csv'
        csvFile=file(filename,'wb')
        csvWriter = csv.writer(csvFile,delimiter=',') 
        #csvWriter.writerows(outdata)
        for data in outdata:
            csvWriter.writerow(data)
        csvFile.close()
        f=codecs.open(filename,'rb','utf-8')
        #ini_sprate=f.readlines()
        ini_sprate=[]
        for line in f:
            ini_sprate.append(line)
        f.close()
        self.data['ini_sprate']=ini_sprate
        self.ini_sprate=ini_sprate
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'完成')        
        self.root.wm_attributes('-topmost',1)
    def rate_check_bu(self):
        '''保证金核对按钮'''
        now=datetime.datetime.now().strftime('%Y-%m-%d')
        mysql=mysql_conn()
        sql='select * from code_rate where tr_date=\'%s\'' %now
        mysql.execute(sql)
        rs=mysql.fetchall()
        mysql.close()
        if not rs:
            tkMessageBox.showinfo(title=u'温馨提醒',message=u'数据库没找到当日保证金率数据,请先上传至至数据库')
            self.root.wm_attributes('-topmost',1)
            return
        colname=[u'合约代码',u'交易所',u'数据库-保证金率调整',u'本系统-保证金率调整',u'原保证金率',u'数据库-调整后保证金率',u'本系统-调整后保证金率']
        outdata=[]
        for x in rs:
            if x['code'] in self.codeclass:
                tpinf=self.codeclass[x['code']].Inf
                if tpinf['delta_rate']<>x['delta_rate']:
                    data=[x['code'],x['house'],x['delta_rate'],tpinf['delta_rate'],x['Mrate'],x['margin_rate'],tpinf['delta_rate']+tpinf['Mrate']]
                    for i in range(2,len(data)):
                        data[i]=data[i]*100
                    outdata.append(data)
        a=table_gui2.table_gui(colname,outdata,'保证金核对',w=120)
        a.sorttree()
        a.mainloop()                
    def upload_bu(self):
        '''上传保证金率按钮'''
        now=datetime.datetime.now().strftime('%Y-%m-%d')
        mysql=mysql_conn()
        sql='select * from code_rate where tr_date=\'%s\'' %now
        mysql.execute(sql)
        rs=mysql.fetchall()
        if rs:
            ques=tkMessageBox.askquestion(u"警告", u"检测到数据库已有保证金率，是否仍要上传并且覆盖原有保证金率？") 
            if ques=='no':
                mysql.close()
                return
            else:
                sql='DELETE FROM  code_rate where tr_date=\'%s\'' %now
                mysql.execute(sql)
        colname=['tr_date','code','house','delta_rate','delta_price','Mrate','margin_rate']
        outdata=[]        
        for x in self.codeclass:
            tpinf=self.codeclass[x].Inf
            data=[now,x,self.codeclass[x].House,tpinf['delta_rate'],tpinf['delta_price'],tpinf['Mrate'],tpinf['Mrate']+tpinf['delta_rate']]
            outdata.append(data)
        mysql.insert_table('code_rate',colname,outdata)
        mysql.commit()
        mysql.close()
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'完成')
        self.root.wm_attributes('-topmost',1)
    def download_bu(self):
        '''下载保证金率按钮'''
        now=datetime.datetime.now().strftime('%Y-%m-%d')
        mysql=mysql_conn()
        sql='select * from code_rate where tr_date=\'%s\'' %now
        mysql.execute(sql)
        rs=mysql.fetchall()
        if not rs:
            mysql.close()
            tkMessageBox.showinfo(title=u'温馨提醒',message=u'数据库没找到当日保证金率数据,请先上传至至数据库')
            self.root.wm_attributes('-topmost',1)
            return
        for x in rs:
            tpinf=self.codeclass[x['code']].Inf
            tpinf['delta_rate']=x['delta_rate']
            tpinf['delta_price']=x['delta_price']
            tpinf['Mrate']=x['Mrate']            
        mysql.close()
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'完成')
        self.root.wm_attributes('-topmost',1)        
class StickyEntry(tk.Entry):
 
    def __init__(self, parent, text, **kw):
        ''' If relwidth is set, then width is ignored '''
        #fa = super(self,StickyEntry)
        #fa.__init__(parent, **kw)
        apply(tk.Entry.__init__, (self, parent), kw)
 
        self.insert(0, text)
        #self['state'] = 'readonly'
        self['readonlybackground'] = 'white'
        self['selectbackground'] = '#1BA1E2'
        self['exportselection'] = False
 
        self.focus_force()
        self.bind("<Control-a>", self.selectAll)
        self.bind("<Escape>", lambda *ignore: self.destroy())
 
    def selectAll(self, *ignore):
        ''' Set selection on the whole text '''
        self.selection_range(0, 'end')
 
        # returns 'break' to interrupt default key-bindings
        return 'break'
