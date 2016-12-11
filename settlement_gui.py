# -*- coding: utf-8 -*-
"""
Created on Wed Jun 29 10:06:20 2016

@author: Administrator
"""

from __future__ import division
import Tkinter as tk
import tkSimpleDialog
import ttk
import tkMessageBox
import functions as fn
import os
class settlement_gui:
    def __init__(self,codeclass,data):
        self.codeclass=codeclass
        self.data=data
        self.weight={}
        self.framesize=0
        self.framew={}
        self.root=tk.Toplevel()
        self.creat_gui()
        self.root.mainloop()        
    def creat_gui(self):
        self.root.title(u'结算价调整')
        self.root.iconbitmap(os.getcwd()+'\\data\\'+'ht_48X48.ico')
        sizex = 300
        sizey = 440
        xy=self.root.winfo_pointerxy()
        #posx  = 1600
        #posy  = 100
        self.root.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, xy[0], xy[1]))
        f1=tk.LabelFrame(self.root,relief='groove',heigh=150,width=300,text='f1')
        f2=tk.LabelFrame(self.root,relief='groove',heigh=290,width=300,text='f2')
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
        tk.Label(f1,text=u'调整幅度').grid(row=2,column=1,columnspan=2,sticky='W')
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
        b=tk.Button(f1,text=u'停板价',fg='red',width=8,command=self.limt_bu)
        b.grid(row=4,column=2) 
        
        b=tk.Button(f1,text=u'显示调结',fg='red',width=10,command=self.visiable_bu)
        b.grid(row=0,column=3)     
        b=tk.Button(f1,text=u'清空调整',fg='red',width=10,command=self.clear_bu)
        b.grid(row=1,column=3)          
        b=tk.Button(f1,text=u'夜盘调结',fg='red',width=10,command=self.ch_night_bu)
        b.grid(row=2,column=3)   
        b=tk.Button(f1,text=u'确认调整',fg='saddlebrown',width=10,command=self.confirm_bu)
        b.grid(row=3,column=3)  
        b=tk.Button(f1,text=u'退出',fg='saddlebrown',width=10,command=self.exit_bu)
        b.grid(row=4,column=3) 
        #c=tk.Canvas(f2,heigh=250,width=300)
        tree=ttk.Treeview(f2)
        tree.grid(row=0,column=0)
        ysb=tk.Scrollbar(f2,orient='vertical',command=tree.yview)
        tree.config(yscrollcommand=ysb.set)
        ysb.grid(row=0,column=1,sticky='NS')
        colname=[u'合约代码',u'原结算价',u'调整幅度(%)',u'调整后结算价']
        tree.config(columns=colname)
        for x in colname:
            tree.column(x,width=70,anchor='center')
            tree.heading(x,text=x)
        tree.config(displaycolumns='#all',show="headings")
        tree.bind("<Double-1>", self.on_detail_bom_line_db_click)
        self.weight['tree']=tree
    def query_bu(self):
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
        items=tree.get_children()
        [tree.delete(item) for item in items]
        for i in range(len(codelist)):
            setl=self.codeclass[codelist[i]].Inf['settlement']
            if setl<>'':
                v=[codelist[i],setl,valuelist[i]*100,setl*(1+valuelist[i])]
            else:
                v=[codelist[i],setl,valuelist[i]*100,'']
            tree.insert('',i+1,values=v)
        self.root.wm_attributes('-topmost',1)
    def exit_bu(self):
        self.data['is_delta_price']=False
        for x in self.codeclass:
            if self.codeclass[x].Inf['delta_price']<>0:
                res=tkMessageBox.askquestion(u"检测到合约结算价有调整", u"合约是否停板？")
                if res=='yes':
                    self.data['delta_price_format']=1#代表停板
                else:
                    self.data['delta_price_format']=0#代表统一向下
                self.data['is_delta_price']=True
                break
        self.root.destroy()
    def confirm_bu(self):
        tree=self.weight['tree']
        codeclass=self.codeclass
        code=[]
        delta=[]
        for k in tree.get_children(''):
            item=tree.item(k)
            code.append(item['values'][0])
            delta.append(item['values'][2])
        for i in range(len(code)):
            codeclass[code[i]].setvalue('delta_price',float(delta[i])/100)
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'完成调整')
        self.root.wm_attributes('-topmost',1)
    def visiable_bu(self):
        tree=self.weight['tree']
        items=tree.get_children()
        [tree.delete(item) for item in items]
        codeclass=self.codeclass
        s=0
        n=0
        for x in codeclass:
            if codeclass[x].Inf['delta_price']<>0:
                try:
                    value=[x,codeclass[x].Inf['settlement'],codeclass[x].Inf['delta_price']*100
                    ,codeclass[x].Inf['settlement']*(1+codeclass[x].Inf['delta_price'])]
                except:
                    value=[x,codeclass[x].Inf['settlement'],codeclass[x].Inf['delta_price']*100,'']                    
                tree.insert('','end',values=value)
                s+=codeclass[x].Inf['delta_price']
                n+=1
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'共调了 '+str(n)+u' 个合约，合计调整幅度：'+str(s))
        self.root.wm_attributes('-topmost',1)
    def ch_night_bu(self):        
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
            r=self.codeclass[rescode[i]].Inf['settlement']
            if r<>'':
                v=[rescode[i],r,valuelist[i]*100,r*(1+valuelist[i])]
            else:
                v=[rescode[i],r,valuelist[i]*100,'']
            tree.insert('',i+1,values=v)
        self.root.wm_attributes('-topmost',1)
    def clear_bu(self):
        for x in self.codeclass:
            self.codeclass[x].Inf['delta_price']=0
        tkMessageBox.showinfo(title=u'温馨提醒',message=u'清空完毕')
        self.root.wm_attributes('-topmost',1)
    def limt_bu(self):
        res=tkMessageBox.askquestion(u"请选择", u"选择涨停（yes)还是跌停(no)？")
        k=0
        if res=='yes':
            k=1
        house=self.weight['rdvar'].get()
        sers=self.weight['svar1'].get().upper().split(',')        
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
        valuelist=[]
        for x in codelist:
            if k:
                ep=self.codeclass[x].Inf['limtup']
            else:
                ep=self.codeclass[x].Inf['limtdown']
            valuelist.append((ep-self.codeclass[x].Inf['meanprice'])/self.codeclass[x].Inf['meanprice'])
        tree=self.weight['tree']
        items=tree.get_children()
        [tree.delete(item) for item in items]
        for i in range(len(codelist)):
            setl=self.codeclass[codelist[i]].Inf['settlement']
            if setl<>'':
                v=[codelist[i],setl,valuelist[i]*100,setl*(1+valuelist[i])]
            else:
                v=[codelist[i],setl,valuelist[i]*100,'']
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
        try:
            vs[3]=(1+float(value)/100)*float(item['values'][1])
        except:
            vs[3]=''
        ttk.Treeview.item(tree,rowid,value=vs)
        entryPopup.destroy()
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
