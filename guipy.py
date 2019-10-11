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


from PIL import Image
# 改变标准输出的默认编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')


window = tk.Tk()
window.title("批量脚注")
window.geometry('450x360')
var1 = tk.StringVar()
fileName = tk.Entry(window, textvariable=var1, width=40).place(x=100, y=20)

statusStr = tk.StringVar()    # 将label标签的内容设置为字符类型，用var来接收函数的传出内容用以显示在标签上
l = tk.Label(window, textvariable=statusStr, bg='green',
             fg='white', font=('Arial', 12), width=30, height=2)
l.pack(side='bottom')


def chooseFile():
    fn1 = tk.filedialog.askdirectory()
    var1.set(fn1)
    statusStr.set('you hit me')


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
    if filePath == "":
        statusStr.set('you hit me1111111111111111')
        return

    print(filePath, "================")
    # topdf(filePath, pictureType=['png', 'jpg'], save=filePath)


# def bianLi(rootDir):
#     for root, dirs, files in os.walk(rootDir):
#         for dir in dirs:
#             fn1 = os.path.join(root, dir)
#             topdf(fn1, pictureType=['png', 'jpg'], save=fn1)


tk.Button(window, text='ImgToPdf', width=50,
          command=imgtopdf).place(x=10, y=260)
window.mainloop()
