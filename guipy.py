# conding=utf-8
import tkinter as tk
import tkinter.filedialog
# -*- coding:=utf8 -*-
import io
import re
import sys
import os
import glob
import platform

from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.platypus import SimpleDocTemplate, Image
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab import rl_settings

from PIL import Image
# 改变标准输出的默认编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')


window = tk.Tk()
window.title("批量脚注")
window.geometry('400x300')
var1 = tk.StringVar()
fileName = tk.Entry(window, textvariable=var1, width=40).place(x=100, y=20)


def chooseFile():
    fn1 = tk.filedialog.askdirectory()
    var1.set(fn1)


tk.Button(window, text='选择文件夹', command=chooseFile).place(x=10, y=20)

tk.Label(window, text='页脚左边的文字').place(x=10, y=80)
var2 = tk.StringVar()
var2.set('左边的文字')
tk.Entry(window, textvariable=var2, width=40).place(x=100, y=80)

tk.Label(window, text='页脚中间的文字').place(x=10, y=140)
var3 = tk.StringVar()
var3.set('页脚中间的文字')
fileName = tk.Entry(window, textvariable=var3, width=40).place(x=100, y=140)

tk.Label(window, text='页脚右边的文字').place(x=10, y=200)
var4 = tk.StringVar()
var4.set('页脚右边的文字')
fileName = tk.Entry(window, textvariable=var4, width=40).place(x=100, y=200)


def imgtopdf():
    filePath = var1.get()
    # myPathRoot(filePath)
    bianLi(filePath)
    topdf(filePath, pictureType=['png', 'jpg'], save=filePath)


def bianLi(rootDir):
    for root, dirs, files in os.walk(rootDir):
        for dir in dirs:
            fn1 = os.path.join(root, dir)
            topdf(fn1, pictureType=['png', 'jpg'], save=fn1)


def topdf(path, recursion=None, pictureType=None, sizeMode=None, width=None, height=None, fit=None, save=None):
    if platform.system() == 'Windows':
        path = path.replace('\\', '/')
    if path[-1] != '/':
        path = (path + '/')
    # print(path)
    if recursion == True:
        for i in os.listdir(path):
            if os.path.isdir(os.path.abspath(os.path.join(path, i))):
                topdf(path+i, recursion, pictureType,
                      sizeMode, width, height, fit, save)
    filelist = []
    if pictureType == None:
        filelist = glob.glob(os.path.join(path, '*.jpg'))
    else:
        picName = var2.get()
        for i in pictureType:
            hellopic = glob.glob(os.path.join(path, '*.'+i))
            hellopic.sort(key=lambda x: len(x))

            if picName == '':
                filelist.extend(hellopic)
            else:
                for j in hellopic:
                    file_path = os.path.split(j)  # 分割出目录与文件
                    file_path1 = file_path[0].split('.')
                    lists = file_path[1].split('.')  # 分割出文件与文件扩展名
                    file_ext = lists[0]  # 取出后缀名(列表切片操作)
                    if picName in file_ext:
                        filelist.append(j)

    maxw = 0
    maxh = 0
    if sizeMode == None or sizeMode == 0:
        for i in filelist:
            im = Image.open(i)
            if maxw < im.size[0]:
                maxw = im.size[0]
            if maxh < im.size[1]:
                maxh = im.size[1]
    elif sizeMode == 1:
        maxw = 999999
        maxh = 999999
        for i in filelist:
            im = Image.open(i)
            print(i)
            if maxw > im.size[0]:
                maxw = im.size[0]
            if maxh > im.size[1]:
                maxh = im.size[1]
    else:
        if width == None or height == None:
            raise Exception("没有提供宽度或者高度")
        maxw = width
        maxh = height

    maxsize = (maxw, maxh)
    save1 = var3.get()
    fileLength = filelist.__len__()
    if(fileLength > 0):
        filename_pdf = path+save1 + '.pdf'
        c = canvas.Canvas(filename_pdf, pagesize=maxsize)

        l = len(filelist)
        for i in range(l):
            (w, h) = maxsize
            width, height = letter
            if fit == True:
                c.drawImage(filelist[i], 0, 0)
            else:
                c.drawImage(filelist[i], 0, 0, maxw, maxh)
            c.showPage()
        c.save()
        tk.Label(window, text='成功转换！').place(x=10, y=290)


tk.Button(window, text='ImgToPdf', width=50,
          command=imgtopdf).place(x=10, y=260)
window.mainloop()
