#!/bin/python
"""
Hello World, but with more meat.
"""

import wx
import wx.grid as gridlib
import pandas as pd
import numpy as np

class Tab(wx.Panel):
    def __init__(self, parent,filename):
        wx.Panel.__init__(self, parent)

        # hyperpara
        k_items = 3

        self.filename = filename
        self.df = self.read_db(filename)
        self.transform_date()
        self.vbox = wx.BoxSizer(wx.VERTICAL)
        self.current_lists = []

        self.columns = [col for col in self.df.columns if not col.startswith('Unnamed:')]
        self.textAreas = dict()
        for idx, col in enumerate(self.columns):
            col = col.strip()
            if idx % k_items == 0:
                self.hbox = wx.BoxSizer(wx.HORIZONTAL) 
            l1 = wx.StaticText(self,label = col)  
            t1 = wx.TextCtrl(self,-1) 
            self.textAreas[col] = t1
            self.hbox.Add(l1, proportion=0, flag= wx.EXPAND | wx.ALL, border=5)
            self.hbox.Add(t1,proportion=1, flag=wx.EXPAND | wx.ALL, border=5) 
            if idx%k_items == 1 or idx == len(self.columns) - 1:
                self.vbox.Add(self.hbox,proportion = 0,flag=wx.EXPAND)

        self.hbox = wx.BoxSizer(wx.HORIZONTAL) 
        self.btnSearch = wx.Button(self,-1,"查詢")  
        self.btnSearch.Bind(wx.EVT_BUTTON,self.OnSearch) 
        self.btnEdit = wx.Button(self,-1,"更改")  
        self.btnEdit.Bind(wx.EVT_BUTTON,self.OnEdit) 
        self.btnInsert = wx.Button(self,-1,"新增")  
        self.btnInsert.Bind(wx.EVT_BUTTON,self.OnInsert) 
        self.btnDelete = wx.Button(self,-1,"刪除")  
        self.btnDelete.Bind(wx.EVT_BUTTON,self.OnDelete) 
        self.btnSave = wx.Button(self,-1,"儲存")  
        self.btnSave.Bind(wx.EVT_BUTTON,self.OnSave) 
        self.hbox.Add(self.btnSearch, proportion=0, flag= wx.EXPAND | wx.ALL, border=5)
        self.hbox.Add(self.btnEdit, proportion=0, flag= wx.EXPAND | wx.ALL, border=5)
        self.hbox.Add(self.btnInsert, proportion=0, flag= wx.EXPAND | wx.ALL, border=5)
        self.hbox.Add(self.btnSave, proportion=0, flag= wx.EXPAND | wx.ALL, border=5)
        self.hbox.Add(self.btnDelete, proportion=0, flag= wx.EXPAND | wx.ALL, border=5)
        self.vbox.Add(self.hbox,proportion = 0,flag=wx.ALL | wx.ALIGN_RIGHT,border=8)

        self.hbox = wx.BoxSizer(wx.HORIZONTAL) 
        self.addColTextArea = wx.TextCtrl(self,-1)
        self.hbox.Add(self.addColTextArea,proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        self.btnAddCol = wx.Button(self,-1,"增加欄位")
        self.btnAddCol.Bind(wx.EVT_BUTTON,self.OnAddCol)
        self.hbox.Add(self.btnAddCol, proportion=0, flag= wx.EXPAND | wx.ALL, border=5)
        self.delColTextArea = wx.TextCtrl(self,-1)
        self.hbox.Add(self.delColTextArea,proportion=0, flag=wx.EXPAND | wx.ALL, border=5) 
        self.btnDeleteCol = wx.Button(self,-1,"刪除欄位")
        self.btnDeleteCol.Bind(wx.EVT_BUTTON,self.OnDelCol)
        self.hbox.Add(self.btnDeleteCol, proportion=0, flag= wx.EXPAND | wx.ALL, border=5)
        self.btnShowAll = wx.Button(self,-1,"顯示所有資料")  
        self.btnShowAll.Bind(wx.EVT_BUTTON,self.OnShowAll) 
        self.hbox.Add(self.btnShowAll, proportion=0, flag= wx.EXPAND | wx.ALL, border=5)
        self.vbox.Add(self.hbox,proportion = 0,flag=wx.ALL | wx.ALIGN_RIGHT,border=8)

        self.myGrid = gridlib.Grid(self)
        self.myGrid.CreateGrid(self.df.shape[0]+20 ,len(self.columns)+5)
        self.vbox.Add(self.myGrid, proportion = 1, flag=wx.TOP | wx.EXPAND,border=10)
        self.Bind(gridlib.EVT_GRID_CELL_LEFT_CLICK, self.OnCellLeftClick)

        # self.datatype = self.set_datatype()
        self.init_table()
        self.SetSizer(self.vbox)

    def read_db(self,name):
        if name[-3:] == 'csv':
            return pd.read_csv(name,encoding = 'big5')  
        else:
            return pd.read_excel(name)

    def transform_date(self):
        mandarin = [i for i in '一二三四五六七八九十零O吉正元']
        integer = '1|2|3|4|5|6|7|8|9|10|0|0|-1|1|1'.split('|')
        mandarin2int = {}
        for man, i in zip(mandarin,integer):
            mandarin2int[man] = int(i)

        for item in ['出生年','月','日']:
            if item in self.df.columns:
                for index, mandInt in enumerate(self.df[item]):
                    if type(mandInt) == str or str(mandInt).lower() == 'nan':
                        if str(mandInt).lower() == 'nan':
                            num = -1
                        else:
                            mandInt = ''.join([y for y in mandInt.replace('(潤)','').replace('潤','').replace('初','').split() if y.strip()])
                            num = 0
                            for idx, y in enumerate(mandInt):
                                if num == 0:
                                        num = mandarin2int[y]
                                else:  
                                    if mandarin2int[y] in [10,0]:
                                        num *= 10
                                    else:
                                        if mandarin2int[mandInt[idx-1]] == 10: 
                                            num += mandarin2int[y]
                                        else:
                                            num = num*10 + mandarin2int[y]
                        self.df.at[index, item] = num
                        
    
    def init_table(self):
        # set col title
        for idx, val in enumerate(self.columns):
            self.myGrid.SetColLabelValue(idx,val)

        self.current_lists = []

        for row_id, row in self.df.iterrows():
            self.myGrid.SetReadOnly(0,row_id,False)
            self.current_lists.append(row_id)
            for col_id, val in enumerate(self.columns):
                if self.df[val].dtype == 'object':
                    self.myGrid.SetCellEditor(row_id,col_id, gridlib.GridCellTextEditor())
                else:
                    self.myGrid.SetCellEditor(row_id,col_id, gridlib.GridCellNumberEditor())

                if str(row[val]) != 'nan':
                    self.myGrid.SetCellValue(row_id,col_id,str(row[val]))
                else:
                    self.myGrid.SetCellValue(row_id,col_id,"")

    def myFilter(self):
        res = self.df
        for label, text in self.textAreas.items():
            text = text.GetValue().strip()
            if text:
                if self.df[label].dtype == 'object':
                    res = res[res[label].str.contains(text)==True]
                else:
                    if '<<' in text or '>>' in text:
                        if '<<' in text:
                            ranges = text.split('<<')
                            lower, upper = int(ranges[0]),int(ranges[1])
                        else:
                            ranges = text.split('>>')
                            upper, lower = int(ranges[0]),int(ranges[1])
                        res = res.loc[(res[label] >= lower) & (res[label] <= upper)]
                    elif '>' in text:
                        lower = int(text[1:])
                        res = res.loc[res[label] >= lower]
                    elif '<' in text:
                        upper = int(text[1:])
                        res = res.loc[res[label] <= upper]
                    else:
                        equal = int(text)
                        res = res.loc[pd.to_numeric(res[label]) == equal]
        return res

    def OnSearch(self, event): 
        res = self.myFilter()
        if not res.empty:
            self.current_lists = []
            self.myGrid.ClearGrid()
            for row_id, (idx,row) in enumerate(res.iterrows()):
                self.current_lists.append(idx)
                for col_id, val in enumerate(self.columns):
                    self.myGrid.SetCellValue(row_id,col_id,str(row[val]))
    
    def OnEdit(self, event): 
        edit_id = [idx for idx in self.df['編號'] if idx == int(self.textAreas['編號'].strip())]
        if not self.textAreas['編號'].strip():
            error_msg(1)
        elif any(edit_id):
            for val, text in self.textAreas.items():
                new_value = text.GetValue().strip()
                if self.df[val].dtype == np.int64 or val in ['出生年','月','日']:
                    try:
                        self.df.at[edit_id,val] = int(new_value)
                    except:
                        if new_value:
                            error_msg(1)
                elif self.df[val].dtype == np.float64:
                    try:
                        self.df.at[edit_id,val] = float(new_value)
                    except:
                        if new_value:
                            error_msg(1)
                else:
                    self.df.at[edit_id,val] = new_value
        else:
            error_msg(1)

    def error_msg(self,code,label=None,value=None):
        if code == 0:
            wx.MessageBox('{}的{}輸入型態錯誤'.format(val,new_value), '輸入錯誤', wx.OK | wx.ICON_INFORMATION)
        elif code == 1:
            wx.MessageBox('請在編號內打入正確數字,才可更改資料', '查詢錯誤', wx.OK | wx.ICON_INFORMATION)
        elif code == 'add':
            wx.MessageBox('若要新增{}的值，請重啟視窗'.format(label), '提醒', wx.OK | wx.ICON_INFORMATION)

    def OnInsert(self, event): 
        record = []
        for val, text in self.textAreas.items():
            new_value = text.GetValue().strip()
            if val in self.df:
                if self.df[val].dtype == np.int64 or val in ['出生年','月','日']:
                    try:
                        record.append(int(new_value))
                    except:
                        if new_value:
                            self.error_msg(0,val,new_value)
                        else:
                            record.append('')
                elif self.df[val].dtype == np.float64:
                    try:
                        record.append(float(new_value))
                    except:
                        if new_value:
                            error_msg(0,val,new_value)
                        else:
                            record.append('')
                else:
                    record.append(new_value)
        df2 = pd.DataFrame([tuple(record)], columns=self.columns)
        self.df = self.df.append(df2,ignore_index=True)

    def OnDelete(self, event): 
        res = self.myFilter()
        if not res.empty:
            for idx,_ in res.iterrows():
                self.df = self.df.drop(self.df.index[idx])

    def OnCellLeftClick(self, evt):
        self.myGrid.SetGridCursor(evt.GetRow(),evt.GetCol())

    def OnAddCol(self,event):
        name = self.addColTextArea.GetValue()
        if not name in self.df:
            self.df[name] = [''] * self.df.shape[0]
            self.save2csv()

            self.columns += [name]
            self.myGrid.SetColLabelValue(len(self.columns)-1,name)
            self.error_msg("add",name)

    def OnDelCol(self,event):
        name = self.delColTextArea.GetValue()
        if name in self.df:
            self.df = self.df.drop(columns=[name])
            self.save2csv()
            self.columns.remove(name)

            for idx, val in enumerate(self.columns):
                self.myGrid.SetColLabelValue(idx,val)
            self.myGrid.SetColLabelValue(idx+1,"")
            self.myGrid.ClearGrid()
            for row_id in self.current_lists:
                for col_id, val in enumerate(self.columns):
                    self.myGrid.SetCellValue(row_id,col_id,str(self.df[val][row_id]))
            self.textAreas[name].Hide()

        
    def OnShowAll(self,event):
        self.current_lists = []
        self.myGrid.ClearGrid()
        for row_id, (_,row) in enumerate(self.df.iterrows()):
            self.current_lists.append(row_id)
            for col_id, val in enumerate(self.columns):
                self.myGrid.SetCellValue(row_id,col_id,str(row[val]))
    
    def OnSave(self,event):
        self.save2csv()

    def save2csv(self):
        for label in self.df:
            if self.df[label].dtype in [np.int64,np.float64] or any([l for l in label if l[-1] in ['年','月','日','號']]):
                self.df[label] = pd.to_numeric(self.df[label], errors='coerce')

        self.df.to_csv(self.filename,index=False, encoding='big5')

class MainFrame(wx.Frame):
    """
    A Frame that says Hello World
    """

    def __init__(self, *args, **kw):
        # ensure the panel's __init__ is called
        super(MainFrame, self).__init__(*args, **kw)
        # create a self.panel in the frame
        self.panel = wx.Panel(self,-1)
        self.nb = wx.Notebook(self.panel)
        # Create the tab windows
        filenames = ['test.csv','普渡樣本.xlsx']
        # tab1 = Tab(self.nb,"補財庫樣本.xlsx")
        tab1 = Tab(self.nb,filenames[0])
        tab2 = Tab(self.nb,filenames[1])

        self.nb.AddPage(tab1, filenames[0].split('.')[0])
        self.nb.AddPage(tab2, filenames[1].split('.')[0])

        # Set noteboook in a sizer to create the layout
        sizer = wx.BoxSizer()
        sizer.Add(self.nb, 1, wx.EXPAND)
        self.panel.SetSizer(sizer)


if __name__ == '__main__':
    # When this module is run (not imported) then create the app, the
    # frame, show it, and start the event loop.
    app = wx.App()
    frm = MainFrame(None, title=u'寶元宮',size=(778, 494),style=wx.DEFAULT_FRAME_STYLE | wx.RESIZE_BORDER | wx.MAXIMIZE_BOX)
    frm.Show(True)
    app.MainLoop()
