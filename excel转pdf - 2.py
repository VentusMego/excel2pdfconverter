#!Python 3
#网上找的代码，适当改动
#源代码依然保留如下，按excel确认xls与xlsm文件所属位置，pdf确认文件输出路径，单击第三个按钮完成转换。
#在excel下（包括子文件夹）都会被转化
import sys
import os
import win32com.client
import tkinter.filedialog
from os.path import splitext
from tkinter import *
import tkinter.messagebox
root = Tk()
root.title('excel转pdf')
def callback1():
    global fpath
    fpath = tkinter.filedialog.askdirectory()
    print('已确认文件所属位置')
def callback2():
    global fpath2
    fpath2 = tkinter.filedialog.askdirectory()
    print('已确认df输出位置')
def callback3():
    for root, dirs, files in os.walk(fpath):
        for file in files:
            if file.endswith('.xls') or file.endswith('.xlsx'): #or file.endswith('.xlsm'):
                file1 = splitext(file)
                out_file_path = fpath2 + '/' + file1[0]
                filename = os.path.join(root, file)
                in_file = os.path.abspath(filename)
                out_file = os.path.abspath(out_file_path)
                o = win32com.client.Dispatch("Excel.Application")
                o.Visible=False
                wb = o.Workbooks.Open(in_file)
               # o.Run("tc") # 不知道是作者写的什么宏
                wb.ActiveSheet.ExportAsFixedFormat(0, out_file+'.pdf')
                print(out_file)
                wb.Close(SaveChanges=0)
def callback4():
    for root, dirs, files in os.walk(fpath):
        for file in files:
            if file.endswith('.xls') or file.endswith('.xlsx'): #or file.endswith('.xlsm'):
                file1 = splitext(file)
                out_file_path = root + '/' + file1[0]
                print(out_file_path)
                filename = os.path.join(root, file)
                in_file = os.path.abspath(filename) #绝对路径
                print(in_file)
                out_file = os.path.abspath(out_file_path)
                o = win32com.client.Dispatch("Excel.Application")
                o.Visible=False
                wb = o.Workbooks.Open(in_file)
               # o.Run("tc") # 不知道是作者写的什么宏
                wb.ActiveSheet.ExportAsFixedFormat(0, out_file+'.pdf')
                print(out_file)
                wb.Close(SaveChanges=0)
Button(root, text="EXCEL路径", fg="blue",bd=2,width=28,command=callback1).pack()
Button(root, text="PDF输出路径", fg="blue",bd=2,width=28,command=callback2).pack()
Button(root, text="转换", fg="blue",bd=2,width=28,command=callback3).pack()
Button(root, text="xls原路径一键输出", fg="red", bd=2,width=28,command=callback4).pack()
root.mainloop()
