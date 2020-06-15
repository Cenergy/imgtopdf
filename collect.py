# conding=utf-8
import tkinter as tk
import tkinter.filedialog
import os
import glob

import pandas as pd
import openpyxl

import docx
import re
from docx import Document

from openpyxl.worksheet.header_footer import _HeaderFooterPart


window = tk.Tk()
window.title("auto excel")
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
    statusStr.set('')





tk.Button(window, text='选择文件夹', command=chooseFile).place(x=10, y=20)

# tk.Label(window, text='固定的行数').place(x=10, y=80)
# var2 = tk.StringVar()
# var2.set('4')
# tk.Entry(window, textvariable=var2, width=40).place(x=100, y=80)


def read_docx(file_name):
    doc = docx.Document(file_name)
    content = '\n'.join([para.text for para in doc.paragraphs])
    return content

def convertCore(filePath, name):

    text=read_docx(filePath)

    regText="原有竣工图工程(.*)实际工程量"
    targetText=re.findall(regText, text,re.S)
    
    return {"name":name,"value":targetText[0]}
    

res=[]

def bianLi(rootDir):
    for root, dirs, files in os.walk(rootDir):
        for file in files:
            filePath = os.path.join(root, file)
            if filePath.endswith('.docx'):
                try:
                    tempName = os.path.splitext(file)[0]
                    result=convertCore(filePath, tempName)
                    res.append(result)
                except:
                    pass

        for dir in dirs:
            bianLi(dir)


def batchConvert():
    filePath = var1.get()
    # myPathRoot(filePath)
    if filePath == "":
        statusStr.set('请选择文件夹！！')
        return

    xlsx_file_number = glob.glob(pathname=filePath+'/' + r'*.docx')


    if len(xlsx_file_number) == 0:
        statusStr.set('不存在.docx的文件')
        return
    bianLi(filePath)
    dataFrame=pd.DataFrame(res)
    dataFrame.to_excel(filePath+'/'+'result.xlsx')
    statusStr.set('恭喜，转换完毕！')


tk.Button(window, text='开始', width=50,
          command=batchConvert).place(x=10, y=260)
window.mainloop()
