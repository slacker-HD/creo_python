# -*- coding: utf8 -*-
import win32com
from win32com import client
import VBAPI
import os
from tkinter import scrolledtext, messagebox, filedialog, Tk, Button, Entry, Label
CREO_APP = 'C:/PTC/Creo 2.0/Parametric/bin/parametric.exe'
INPUT_DIR = 'D:/test/'

win = Tk()
win.title("批量关系操作")
win.resizable(0, 0)

Label(win, text="Creo程序路径", padx=5, pady=5).grid(row=0, column=0, sticky='W')
Label(win, text="包含prt文件的目录", padx=5, pady=5).grid(row=1, column=0, sticky='W')
Label(win, text="在此编辑关系：", padx=5, pady=5).grid(row=2, column=0, sticky='W', columnspan=3)


e1 = Entry(win, width="55")
e2 = Entry(win, width="55")
e1.grid(row=0, column=1, padx=5, pady=5)
e2.grid(row=1, column=1, padx=5, pady=5)
e1.insert(0, CREO_APP)
e2.insert(0, INPUT_DIR)

st3 = scrolledtext.ScrolledText(win, width=85, height=13)
st3.grid(row=3, column=0, padx=5, pady=5, columnspan=3, sticky='W')


def addrel():
    rel_contents = (st3.get("0.0", "end").replace(" ", "")).split("\n")
    rel_contents.pop()
    relations = client.Dispatch(VBAPI.Cstringseq)
    cAC = client.Dispatch(VBAPI.CCpfcAsyncConnection)
    AsyncConnection = cAC.Start(CREO_APP + ' -g:no_graphics -i:rpc_input', '')
    files = AsyncConnection.Session.ListFiles("*.prt", getattr(VBAPI.constants, "EpfcFILE_LIST_LATEST"), INPUT_DIR)
    for i in range(0, files.Count):
        ModelDescriptor = client.Dispatch(VBAPI.CCpfcModelDescriptor)
        mdlDescr = ModelDescriptor.Create(getattr(VBAPI.constants, "EpfcMDL_PART"), "", None)
        mdlDescr.Path = files.Item(i)
        RetrieveModelOptions = client.Dispatch(VBAPI.CCpfcRetrieveModelOptions)
        options = RetrieveModelOptions.Create()
        options.AskUserAboutReps = False
        model = AsyncConnection.Session.RetrieveModelWithOpts(mdlDescr, options)
        originrels = model.Relations
        for j in range(0, originrels.Count):
            relations.Append(originrels.Item(j))
        for line in rel_contents:
            relations.Append(line)
        model.Relations = relations
        model.Save()
    AsyncConnection.End()
    messagebox.showinfo('提示', '关系已全部清空')


def delrel():
    cAC = client.Dispatch(VBAPI.CCpfcAsyncConnection)
    AsyncConnection = cAC.Start(CREO_APP + ' -g:no_graphics -i:rpc_input', '')
    files = AsyncConnection.Session.ListFiles("*.prt", getattr(VBAPI.constants, "EpfcFILE_LIST_LATEST"), INPUT_DIR)
    for i in range(0, files.Count):
        ModelDescriptor = client.Dispatch(VBAPI.CCpfcModelDescriptor)
        mdlDescr = ModelDescriptor.Create(getattr(VBAPI.constants, "EpfcMDL_PART"), "", None)
        mdlDescr.Path = files.Item(i)
        RetrieveModelOptions = client.Dispatch(VBAPI.CCpfcRetrieveModelOptions)
        options = RetrieveModelOptions.Create()
        options.AskUserAboutReps = False
        model = AsyncConnection.Session.RetrieveModelWithOpts(mdlDescr, options)
        model.DeleteRelations()
        model.Save()
    AsyncConnection.End()
    messagebox.showinfo('提示', '关系已全部清空')


def chooseapp():
    filename = filedialog.askopenfilename()
    if filename != '':
        CREO_APP = filename
        e1.delete('0', 'end')
        e1.insert(0, CREO_APP)


def choosedir():
    dirname = filedialog.askdirectory()
    if dirname != '':
        INPUT_DIR = dirname
        e2.delete('0', 'end')
        e2.insert(0, INPUT_DIR)


Button(win, text="选择文件", command=chooseapp).grid(row=0, column=2, padx=5, pady=5, sticky='E')
Button(win, text="选择路径", command=choosedir).grid(row=1, column=2, padx=5, pady=5, sticky='E')
Button(win, text="批量添加关系", command=addrel).grid(row=4, column=0, sticky='W', padx=5, pady=5)
Button(win, text="批量清空关系", command=delrel).grid(row=4, column=2, sticky='E', padx=5, pady=5)

win.mainloop()
