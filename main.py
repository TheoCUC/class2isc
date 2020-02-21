# -*- coding: utf-8 -*-
import functions as f
import sys
from tkinter import filedialog
from tkinter import *


def loadcsv():
    csvinputE.delete(0,END)
    csvinputE.insert(0,filedialog.askopenfilename())

def saveisc():
    iscoutputE.delete(0,END)
    iscoutputE.insert(0,filedialog.askdirectory())

def doit():
    
    date = yearE.get() + '-' + monthE.get() + '-' + dayE.get()
    if (len(date)!=10 or int(date[5:7])>12 or int(date[8:10])>31 or len(csvinputE.get())==0 or len(iscoutputE.get())==0):
        stateL.config(text='请检查输入!')
    elif ( csvinputE.get().find('.csv',len(csvinputE.get())-4) == -1 and  csvinputE.get().find('.xls',len(csvinputE.get())-4) == -1 ):
        stateL.config(text='格式不受支持!')
    elif( csvinputE.get().find('.csv',len(csvinputE.get())-4) >0):
        f.writeisc(date,f.read_csv(csvinputE.get()),iscoutputE.get())
        stateL.config(text='转换完毕')
        if (sys.platform == 'darwin'):
            f.os.system('open '+iscoutputE.get())
    elif( csvinputE.get().find('.xls',len(csvinputE.get())-4) >0):
        try:
            import xlrd
        except ImportError:
            stateL.config(text='未安装xlrd库')
        else:
            f.writeisc(date,f.read_xls(csvinputE.get()),iscoutputE.get())
            stateL.config(text='转换完毕')
            if (sys.platform == 'darwin'):
                f.os.system('open '+iscoutputE.get())
    # else:
    #     f.writeisc(date,f.read_csv(csvinputE.get()),iscoutputE.get())
    #     stateL.config(text='转换完毕')
    #     if (sys.platform == 'darwin'):
    #         f.os.system('open '+iscoutputE.get())

def clear():
    csvinputE.delete(0,END)
    iscoutputE.delete(0,END)
    stateL.config(text='...')

if __name__=='__main__':
    root = Tk()
    root.title('非官方中国传媒大学课表转日历工具')

    csvinputL = Label(root,text='教务下载的课表文件：')
    csvinputL.grid(row=1)
    iscoutputL = Label(root,text='日历文件保存地址：')
    iscoutputL.grid(row=2)

    csvinputE = Entry(root)
    iscoutputE = Entry(root)
    csvinputE.grid(row=1,column=1,columnspan=4)
    iscoutputE.grid(row=2,column=1,columnspan=4)

    csvbtn = Button(root,text='...',command=loadcsv)
    csvbtn.grid(row=1,column=5)

    iscbtn = Button(root,text='...',command=saveisc)
    iscbtn.grid(row=2,column=5)

    dobtn = Button(root,text='转换',command=doit)
    dobtn.grid(row=3,column=2)
    stateL = Label(root,text='...')
    stateL.grid(row=3,column=1)
    clrbtn = Button(root,text='清空',command=clear)
    clrbtn.grid(row=3,column=0)

    DateL = Label(root,text='第一周第一天日期(yyyyMMdd)：').grid(row=0,column=0)
    yearE = Entry(root,width=6,justify='center')
    yearE.grid(row=0,column=1)
    Label(root,text='-').grid(row=0,column=2)
    monthE = Entry(root,width=4,justify='center')
    monthE.grid(row=0,column=3)
    Label(root,text='-').grid(row=0,column=4)
    dayE = Entry(root,width=4,justify='center')
    dayE.grid(row=0,column=5)

    root.mainloop()