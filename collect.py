# pip install python_docx
# conding=utf-8
import tkinter as tk
import tkinter.filedialog
import os
import glob


from pandas import DataFrame
import re
from docx import Document




window = tk.Tk()
window.title("批量获取docx中的文字")
window.geometry('450x360')
var1 = tk.StringVar()
res=[]

fileName = tk.Entry(window, textvariable=var1, width=40).place(x=100, y=20)

statusStr = tk.StringVar()    # 将label标签的内容设置为字符类型，用var来接收函数的传出内容用以显示在标签上
l = tk.Label(window, textvariable=statusStr, bg='green',
             fg='white', font=('Arial', 12), width=30, height=2)
l.pack(side='bottom')

tk.Label(window,text='开始截取的文字').place(x=10,y=80)
var2=tk.StringVar()
var2.set('原有竣工图工程')
tk.Entry(window,textvariable=var2,width=40).place(x=100,y=80)

tk.Label(window,text='结束截取的文字').place(x=10,y=140)
var3=tk.StringVar()
var3.set('实际工程量')
fileName=tk.Entry(window,textvariable=var3,width=40).place(x=100,y=140)


def chooseFile():
    fn1 = tk.filedialog.askdirectory()
    var1.set(fn1)
    statusStr.set('')





tk.Button(window, text='选择文件夹', command=chooseFile).place(x=10, y=20)



def read_docx(file_name):
    doc = Document(file_name)
    content = '\n'.join([para.text for para in doc.paragraphs])
    return content

def returnValue(targetList):
    if targetList==None:
        return float(0)
    if len(targetList):
        return float(targetList[0])
    return float(0)
    

def convertCore(filePath, name,startText,endText):

    text=read_docx(filePath)

    regText="{}(.*){}".format(startText,endText)
    targetText=re.findall(regText, text,re.S)
    if len(targetText):
        pipeLengthReg=r"{}(-?\d+\.?\d*e?-?\d*?)".format("排水管道长度共计")
        rainReg=r"{}(-?\d+\.?\d*e?-?\d*?)".format("雨水管道")
        pollutionReg=r"{}(-?\d+\.?\d*e?-?\d*?)".format("污水管道")
        stationReg=r"{}(-?\d+\.?\d*e?-?\d*?)".format("雨水篦等连接管")
        wcReg=r"{}(-?\d+\.?\d*e?-?\d*?)".format("化粪池等连接管")
        allLength=re.findall(pipeLengthReg, targetText[0],re.S)
        rainLength=re.findall(rainReg, targetText[0],re.S)
        pollutionLength=re.findall(pollutionReg, targetText[0],re.S)
        stationLength=re.findall(stationReg, targetText[0],re.S)
        wcLength=re.findall(wcReg, targetText[0],re.S)
        nameReg=r'[(](.*?)[)]'
        realName=re.findall(nameReg, name,re.S)
        return {"名称":realName[1],"name":name,"排水管道长度共计":returnValue(allLength),"雨水管道":returnValue(rainLength),"污水管道":returnValue(pollutionLength),"雨水篦等连接管":returnValue(stationLength),"化粪池等连接管":returnValue(wcLength),"原始值":targetText[0]}
    



def bianLi(rootDir):
    startText=var2.get()
    endText=var3.get()
    for root, dirs, files in os.walk(rootDir):
        for file in files:
            filePath = os.path.join(root, file)
            if filePath.endswith('.docx'):
                try:
                    tempName = os.path.splitext(file)[0]
                    result=convertCore(filePath, tempName,startText,endText)
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
    dataFrame=DataFrame(res)
    dataFrame.to_excel(filePath+'/'+'result.xlsx')
    statusStr.set('恭喜，转换完毕！')


tk.Button(window, text='开始', width=50,
          command=batchConvert).place(x=10, y=260)
window.mainloop()
