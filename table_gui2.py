# -*- coding: utf-8 -*-
"""
Created on Thu Jul 14 23:47:54 2016

@author: Administrator
"""

# -*- coding: utf-8 -*-
"""
Created on Mon Jul 11 22:44:07 2016

@author: Administrator
"""
import Tkinter as tk
import ttk
import functions as fn
import tkSimpleDialog
import math
import re
import os
class table_gui:
    def __init__(self,colname,data,s_title='',nums=1000,w=70,isframe='',width=1050,height=420,vrows=25,parentclass=''):
        self.w=w
        self.width=width
        self.data=data
        self.height=round(20.84*vrows)
        self.parentclass=parentclass
        if isframe=='':
            root=tk.Toplevel()
            root.iconbitmap(os.getcwd()+'\\data\\'+'ht_48X48.ico')
            root.title(s_title)
            xy=root.winfo_pointerxy()
            #sizex = 260
            #sizey = 35
            #root.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, xy[0], xy[1]))
            f=tk.LabelFrame(root)
            f.pack()
            self.root=root
        else:
            f=isframe
        canvas=tk.Canvas(f)
        xsb=tk.Scrollbar(f,orient='horizontal',command=canvas.xview)
        canvas.config(xscrollcommand=xsb.set)      
        canvas.grid(row=0, column=0)       
        xsb.grid(row=1,column=0, sticky='ew')                        
        f2=tk.Frame(canvas)
        tree=ttk.Treeview(f2,height=vrows)
        ysb=tk.Scrollbar(f2,orient='vertical',command=tree.yview)
        self.tree=tree
        
        self.colname=colname
        self.canvas=canvas        
        tree.config(yscrollcommand=ysb.set)
        tree.grid(row=0,column=1)
        ysb.grid(row=0,column=0, sticky='ns')
        canvas.create_window(0,0,window=f2)
        self.settree(tree,colname,data,nums)
        self.colortree()       
        f.bind("<Configure>",self.myfunction)
        tree.bind_all('<Control-KeyPress-F>',self.search)
        tree.bind_all('<Control-KeyPress-f>',self.search)
        items=tree.get_children()
        if len(items)<vrows:
            tree.config(height=len(items))
        tree.update()
        if f2.winfo_width()<800:
            self.width=f2.winfo_width()
            canvas.config(width=self.width)            
        if tree.winfo_height()<625:
            self.height=tree.winfo_height()
            canvas.config(height=self.height)
        canvas.xview_moveto(0.0)    
    def myfunction(self,event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"),width=self.width,height=self.height)    
    def resettree(self,data,tags=[],iscolor=True):
        tree=self.tree
        items=tree.get_children()
        [tree.delete(item) for item in items]
        if tags:
            for i in range(len(data)):
                tree.insert('',i+1,values=data[i],tags=tags[i]) 
        else:
            for i in range(len(data)):
                tree.insert('',i+1,values=data[i]) 
        if iscolor:
            self.colortree()
    def settree(self,tree,colname,data,nums):        
        tree.config(columns=colname)
        for x in colname:
            tree.column(x,width=self.w,anchor='center')
            tree.heading(x,text=fn.english_to_ch(x))       
        for i in range(len(data)):
            tree.insert('',i+1,values=data[i])
            if i>nums:
                break
        tree.config(displaycolumns='#all',show="headings")
    def sorttree(self,hl=False):
        tree=self.tree
        colname=tree['columns']
        for x in colname:
            tree.heading(x,command=lambda col=x:self.treeview_sort_column(tree,col, False,hl))
    def highlight(self,name='leftcapital'):
        tree=self.tree
        nk=self.colname.index(name)
        
        for k in tree.get_children(''):
            item=tree.item(k)
            if item['values'][nk][0]=='-':               
                if tree.tag_has('oddrow',k):
                    tree.item(k, tags=['red','oddrow'])
                else:
                    tree.item(k, tags=['red'])
        tree.tag_configure('red',foreground='red')
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
                else:
                    tree.item(items[i], tags=tag+['oddrow'])
        tree.tag_configure('oddrow', background='lavender')        
    def search(self,event):
        text=tkSimpleDialog.askstring('HUATAI FUTURE',u'请输入第一列搜寻文本',initialvalue = '')
        if text==None:
            return
        text=text.upper()
        tree=self.tree
        for k in tree.get_children(''):
            item=tree.item(k)
            if text in str(item['values'][0]):
                tree.selection_set(k)
                tree.see(k)
                break
    def treeview_sort_column(self,tv, col, reverse,hl):

        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        if re.match('[0-9-]',l[0][0]):            
            ll=[]
            for x in l:
                try:
                    ll.append((float(x[0].replace(',','')),x[1]))
                except:
                    ll.append(x)
            l=ll
        l.sort(reverse=reverse)
        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            tv.insert('','end',values=tv.item(k)['values'],tags=tv.item(k)['tags'])
            tv.delete(k)
        # reverse sort next time
        tv.heading(col, command=lambda:self.treeview_sort_column(tv, col, not reverse,hl))
        self.colortree()
        if hl:
            self.highlight()
    def bus_double_click(self):
        self.tree.bind('<Double-Button-1>',self.bus_double_click_bu)
    def risk_client_double_click(self):
        self.tree.bind('<Double-Button-1>',self.risk_client_double_click_bu)
    def entry_double_click(self):
        self.tree.bind('<Double-Button-1>',self.on_detail_bom_line_db_click)        
    def bus_double_click_bu(self,event):
        tree=self.tree
        rowid = tree.identify_row(event.y)
        ivid=str(tree.item(rowid)['values'][0])   
        self.parentclass.ssh_bu(event,ivid)
    def risk_client_double_click_bu(self,event):
        tree=self.tree
        rowid = tree.identify_row(event.y)
        dept=tree.item(rowid)['values'][0]
        self.parentclass.riskclient_visiable(dept)
    def col_config(self,collist,w=90):
        colname=self.tree['columns']
        for x in collist:
            if x in colname:
                self.tree.column(x)
                self.tree.heading(x,text=x+u'(*)')
    def mainloop(self):
        self.root.mainloop()
        
        
        
    def on_detail_bom_line_db_click(self, event):
        ''' 双击进入编辑'''
        tree=self.tree
        rowid = tree.identify_row(event.y)
        column = tree.identify_column(event.x)   
        parent = tree.parent(rowid)
        if parent == '':
            pass    
        x,y,width,height = tree.bbox(rowid, column)   
        pady = height // 2    
        url = tree.item(rowid, 'text')   
        tree.entryPopup = StickyEntry(tree, '', width=10)
        tree.entryPopup.place( x=x, y=y+pady, anchor='w')
        tree.entryPopup.bind("<Return>",lambda event,x=tree.entryPopup:self.entry_enter(x,rowid,column))
    def entry_enter(self,entryPopup,rowid,column):
        tree=self.tree
        value=entryPopup.get()
        tree.set(rowid,column,value)
        entryPopup.destroy()
class StickyEntry(tk.Entry):
    def __init__(self, parent, text, **kw):
        ''' If relwidth is set, then width is ignored '''
        apply(tk.Entry.__init__, (self, parent), kw) 
        self.insert(0, text)
        self['readonlybackground'] = 'white'
        self['selectbackground'] = '#1BA1E2'
        self['exportselection'] = False
        self.focus_force()
        self.bind("<Control-a>", self.selectAll)
        self.bind("<Escape>", lambda *ignore: self.destroy()) 
    def selectAll(self, *ignore):
        ''' Set selection on the whole text '''
        self.selection_range(0, 'end')
        return 'break'
       