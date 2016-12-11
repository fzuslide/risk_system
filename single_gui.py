# -*- coding: utf-8 -*-
"""
Created on Wed Jul 13 22:48:02 2016

@author: Administrator
"""

import Tkinter as tk
import ttk
import copy
import functions as fn
import tkSimpleDialog
import tkMessageBox
import cal_margin
from table_gui2 import *
class single_gui:
    def __init__(self,data,ivid,isframe=''):
        invclass=data['invclass']
        '''
        ivid=tkSimpleDialog.askstring('HUATAI FUTURE',u'请输入客户号',initialvalue = u'10802698')
        if ivid==None:
            return
        if not invclass.has_key(ivid):
            tkMessageBox.showinfo(title=u'温馨提醒',message=u'找不到投资者'+ivid)
            return
        '''   
        inv=copy.deepcopy(invclass[ivid])
        if isframe=='':
            root=tk.Toplevel()
            root.iconbitmap('ht_48X48.ico')
            root.title(u'单个客户风险试算')                       
        else:
            root=isframe
        self.root=root
        self.data=data
        self.inv=inv
        self.inv2=copy.deepcopy(inv)
        self.ivid=ivid
        self.codeclass=copy.deepcopy(data['ini_codeclass'])
        self.ps='meanprice'
        self.recal()
        self.calforced()
        inv=self.inv2
        labels={}

        gr=0
        gl=0
        tk.Label(root,text=ivid,fg='red').grid(row=gr,column=0,sticky='W')
        tk.Label(root,text=inv.InvInf['inv_name'],fg='red').grid(row=gr,column=3,columnspan=3,sticky='W')
        tk.Label(root,text=inv.InvInf['invdepartment']).grid(row=gr,column=1,sticky='W')
        tk.Label(root,text=inv.InvInf['seat']).grid(row=gr,column=2,sticky='W')
        lbname=[[u'交易所可用',u'交易所风险度',u'公司可用',u'公司风险度'],
                [u'出入金',u'平仓盈亏',u'持仓盈亏',u'手续费'],
                [u'昨权益',u'权益',u'交易所保证金',u'公司保证金'],
                [u'质押金额',u'资金冻结',u'银期可取',u'OA可取']]
        for i in range(len(lbname)):
            gr=1            
            for x in lbname[i]:
                lb= tk.Label(root,text=x)
                lb.grid(row=gr,column=i*2,sticky='W')
                labels[x]=lb
                gr+=1
        
        cl=[['leftcapital','riskdegree','cor_leftcapital','cor_riskdegree'],
            ['cashmove','closeprofit','holdprofit','fee'],
            ['lastcapital','capital','margin','cor_margin'],
            ['mortgagemoney','frozenmoney','bankmoney','oa_bankmoney']]
        self.up_info=cl
        self.edit_info=['cashmove','mortgagemoney','closeprofit','fee','frozenmoney']
        for i in range(len(cl)):
            gr=1
            for x in cl[i]:
                lb=tk.Label(root,bg='white',text=fn.float_to_str(round(inv.InvInf[x],2)))
                lb.grid(row=gr,column=i*2+1,sticky='W')
                labels[x]=lb
                gr+=1
        labels[u'交易所风险度'].config(fg='red')
        labels['riskdegree'].config(fg='red')       
        self.labels=labels        
        self.mainframe=tk.LabelFrame(root)
        self.mainframe.grid(row=gr,column=0,columnspan=100,sticky='W')
        self.outtree(1)
        gr+=1
        posframe=tk.LabelFrame(root)
        posframe.grid(row=gr,column=0,columnspan=100,sticky='W')
        self.posframe=posframe
        self.outpos(1)   
        self.edit_label()
        
        cmbox=ttk.Combobox(root,width=5)
        cmbox['values']=[u'均价',u'今结算',u'昨结算',u'最新价']
        cmbox.grid(row=0,column=6)
        cmbox.current(0)
        self.cmbox=cmbox        
        cmbox.bind('<<ComboboxSelected>>',self.cmbox_bu)
        pos_cmbox=ttk.Combobox(root,width=8)
        pos_cmbox['values']=[u'正常持仓价',u'不利持仓价']
        pos_cmbox.grid(row=0,column=7)
        pos_cmbox.current(0)
        self.pos_cmbox=pos_cmbox        
        pos_cmbox.bind('<<ComboboxSelected>>',self.pos_cmbox_bu)
        
                
        
    def cmbox_bu(self,event):
        sdic={u'均价':'meanprice',u'今结算':'settlement',u'昨结算':'lastsettlement',u'最新价':'price_now'}
        ps=sdic[self.cmbox.get()]
        self.ps=ps
        tree=self.main_gui.tree
        codeclass=self.codeclass
        items=tree.get_children()
        k=tree['columns'].index(u'结算价')
        for x in items:
            val=codeclass[tree.item(x)['values'][0]].Inf[ps]
            tree.set(x,column=k,value=val) 
        self.get_recal()
    def pos_cmbox_bu(self,event):
        tree=self.main_gui.tree
        items=tree.get_children('')
        pos=self.inv.Position
        i=0
        colname=tree['columns']       
        for k in items:
            while True:
                if len(pos[i]['code'])<=6:
                    break
                i+=1
            if self.pos_cmbox.get()==u'正常持仓价':
                tree.set(k,colname.index(u'买持均价'),pos[i]['longholdprice'])
                tree.set(k,colname.index(u'卖持均价'),pos[i]['shortholdprice'])                             
            else:
                if pos[i]['longnums']>0:
                    tree.set(k,colname.index(u'买持均价'),pos[i]['longholdprice']+self.codeclass[pos[i]['code']].Inf['Minprice'])
                if pos[i]['shortnums']>0:
                    tree.set(k,colname.index(u'卖持均价'),pos[i]['shortholdprice']-self.codeclass[pos[i]['code']].Inf['Minprice']) 
            i+=1 
        self.get_recal()
    def edit_label(self):
        cl=self.edit_info
        for x in cl:
            self.labels[x].config(bg='lightgreen')
            self.labels[x].bind("<Double-1>", lambda event,lb=x:self.label_double_click(lb))
    def label_double_click(self,lname):
        label=self.labels[lname]
        lbinfo=label.grid_info()
        entry=StickyEntry(self.root,'', width=8)
        entry.grid(row=lbinfo['row'],column=lbinfo['column'],columnspan=lbinfo['columnspan'],rowspan=lbinfo['rowspan'],sticky=lbinfo['sticky'])
        entry.bind("<Return>",lambda event:self.label_double_click_enter(entry,lname))
    def label_double_click_enter(self,entry,lname):
        val=entry.get()
        entry.destroy()
        self.inv.InvInf[lname]=float(val)
        self.inv2.InvInf[lname]=float(val)
        self.recal()
        self.outinf()
        self.outtree()
        self.outpos()        
    def recal(self):
        data=self.data
        [self.codeclass[x].adj_price(ptag=0,ctag=1,ps=self.ps) for x in self.codeclass]
        invcalss={}
        invcalss[self.ivid]=self.inv2
        invclass2=cal_margin.allcal(invcalss,self.codeclass,data['ini_sporder'],data['code_house'],data['variclass'],data['shfe_unsp'],data['cfe_unsp'])
        fn.cal_otherinf(invclass2,invcalss,self.codeclass)
        self.inv2.InvInf=invclass2[self.ivid].InvInf
        self.inv3=invclass2[self.ivid]
    def calforced(self,):
        inv=self.inv2
        codeclass=self.codeclass
        for x in inv.Position:
            if len(x['code'])>6:
                x['forcedbound']=0
                x['outbound']=0
                x['forcednums']=0
                continue
            cclass=codeclass[x['code']]
            codenums=x['longnums']-x['shortnums']
            x['codenums']=codenums
            pl=x['longnums']
            ps=x['shortnums']
            x['codedir']=1
            if codenums<0:
                x['codedir']=-1
            if codenums<>0:                
                x['outbound']=cclass.Price-inv.InvInf['capital']/(cclass.Units*codenums)
                if codenums-cclass.Rate*max(pl,ps)<>0:
                    x['forcedbound']=cclass.Price-inv.InvInf['leftcapital']/(cclass.Units*(codenums-cclass.Rate*max(pl,ps)))
                else:
                    x['forcedbound']=0              
                k=abs(codenums)/codenums
                x['forcednums']=inv.InvInf['leftcapital']/(-k*(cclass.Inf['price_now']-cclass.Price)*cclass.Units-cclass.Rate*cclass.Price*cclass.Units)      
            else:
                x['outbound']=0
                x['forcedbound']=0
                x['forcednums']=0
                
    def myfunction(self,event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"),width=self.w,height=self.h)
    def myfunction2(self,event):
        self.treecanvas.configure(scrollregion=self.treecanvas.bbox("all"),width=self.wt,height=self.ht)        
    def mainloop(self):
        self.root.mainloop()
    def outtree(self,fs=0):
        #fs是否首次显示
        cl=[u'合约',u'投保',u'强平边界',u'穿仓边界',u'应平多',u'应平空',u'调整后结算价',u'买持',u'买持均价',u'卖持',u'卖持均价',
            u'调保后保证金率',u'公司上调',u'调保',u'结算价',u'结算价调整',u'强平价格',u'平多仓',u'平空仓',u'Code']
        cl2=['code','posident','forcedbound','outbound','forcednums','codeclass_Price','longnums','longholdprice','shortnums','shortholdprice',
            'codeclass_Rate','cor_rate','codeclass_delta_rate','codeclass_ps','codeclass_delta_price','codeclass_price_now']
        self.calforced()
        inv=self.inv2
        codeclass=self.codeclass
        
        if not fs:
            tree=self.main_gui.tree
            colname=tree['columns']
            items=tree.get_children()
            lsinf=[]
            for item in items:
                value=tree.item(item)['values']
                lsinf.append([value[colname.index(u'平多仓')],value[colname.index(u'平空仓')]])        
        
        outdata=[]
        i=1        
        for x in inv.Position:
            code=x['code']           
            if len(code)>6:
                continue
            else:
                codeinf=codeclass[code].Inf
                v=[]
                for y in cl2:
                    if 'codeclass_' in y:
                        cname=y.replace('codeclass_','')
                        if cname=='ps':                        
                            v.append(codeinf[self.ps])
                        elif cname in ['Rate','delta_rate','delta_price']:
                            v.append(codeinf[cname]*100)
                        else:
                            v.append(codeinf[cname])   
                    else:
                        if y=='forcednums':
                            v.append(round(max(0,x['codedir']*x['forcednums']),2))
                            v.append(round(max(0,-x['codedir']*x['forcednums']),2))
                        elif y in ['forcedbound','outbound']:
                            v.append(round(x[y],2))
                        elif y=='cor_rate':
                            v.append(x[y]*100)
                        else:
                            v.append(x[y])
                v.append(0)
                v.append(0)
                v.append(code)
            if not fs:
                v[-3:-1]=lsinf[i-1]
            outdata.append(v) 
            i+=1
        if fs:
            tg=table_gui(colname=cl,data=outdata,isframe=self.mainframe,w=70,vrows=8)
            self.main_gui=tg
            tg.tree.bind("<Double-1>", self.on_detail_bom_line_db_click)
            #tg.col_config([u'调保后保证金率',u'调保',u'结算价',u'结算价调整',u'强平价格',u'平多仓',u'平空仓'])
        else:
            self.main_gui.resettree(outdata)
            
    def outpos(self,fs=0):
        frame=self.posframe
        cl=['code','longnums','shortnums','longholdprice','shortholdprice','longmar','shortmar','cor_longmar','cor_shortmar']       
        inv=self.inv3
        outdata=[]
        for x in inv.Position:
            v=[]
            for y in cl:
                v.append(fn.float_to_str(x[y]))
            outdata.append(v)
        if fs:
            tg=table_gui(colname=cl,data=outdata,isframe=frame,w=90,vrows=8) 
            self.postree_gui=tg
        else:
            self.postree_gui.resettree(outdata)

        
    def outinf(self):
        inv=self.inv2        
        for x in self.up_info:   
            for y in x:
                self.labels[y].config(text=fn.float_to_str(round(inv.InvInf[y],2)))
    def on_detail_bom_line_db_click(self, event):
        ''' Executed, when a row is double-clicked. Opens
        read-only EntryPopup above the item's column, so it is possible
        to select text '''
        tree=self.main_gui.tree
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
        tree.entryPopup.bind("<Return>",lambda event,x=tree.entryPopup:self.entry_enter(x,rowid,column))
    def entry_enter(self,entryPopup,rowid,col):
        tree=self.main_gui.tree
        val=entryPopup.get()
        tree.set(rowid,column=col,value=val) 
        entryPopup.destroy()
        self.get_recal(tick=tree.column(col)['id'])
    def get_recal(self,tick=''):
        tree=self.main_gui.tree
        self.inv2=copy.deepcopy(self.inv)
        items=tree.get_children()
        pos=self.inv2.Position
        if len(pos)==1:
            items=[items]
        codeclass=copy.deepcopy(self.codeclass)
        colname=tree['columns']
        for item in items:
            value=tree.item(item)['values']
            code=value[colname.index(u'合约')]
            if len(code)>6:
                continue
            isfind=False
            for i,x in enumerate(pos):
                if x['code']==code and x['posident']==value[colname.index(u'投保')]:
                    isfind=True
                    break
            if not isfind:
                tkMessageBox.showinfo(title=u'温馨提醒',message=code+u'持仓对不上')
                return
            posinf=pos[i]
            codeclass[code].Inf['delta_rate']=float(value[colname.index(u'调保')])/100
            codeclass[code].Inf['delta_price']=float(value[colname.index(u'结算价调整')])/100
            codeclass[code].Inf[self.ps]=float(value[colname.index(u'结算价')])
            codeclass[code].Inf['price_now']=float(value[colname.index(u'强平价格')])
            posinf['longholdprice']=float(value[colname.index(u'买持均价')])
            posinf['shortholdprice']=float(value[colname.index(u'卖持均价')])
            
            ln=float(value[colname.index(u'平多仓')])
            sn=float(value[colname.index(u'平空仓')])
            self.inv2.InvInf['closeprofit']+=ln*codeclass[code].Units*(codeclass[code].Inf['price_now']-posinf['longholdprice'])
            self.inv2.InvInf['closeprofit']+=-sn*codeclass[code].Units*(codeclass[code].Inf['price_now']-posinf['shortholdprice'])
            posinf['longnums']-=ln
            posinf['shortnums']-=sn
            if tick in [u'买持',u'卖持']:
                posinf['longnums']=float(value[colname.index(u'买持')])
                posinf['shortnums']=float(value[colname.index(u'卖持')])    
                tree.set(item,column=colname.index(u'平多仓'),value=0)
                tree.set(item,column=colname.index(u'平空仓'),value=0)
            
            i+=1
        self.codeclass=codeclass
        self.recal()
        self.outinf()
        self.outtree()
        self.outpos()
       
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
            
