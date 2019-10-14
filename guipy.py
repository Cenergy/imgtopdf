# conding=utf-8
import tkinter as tk
import tkinter.filedialog
import os

import openpyxl
from openpyxl.worksheet.header_footer import _HeaderFooterPart
import glob

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
    statusStr.set('')


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


def convertCore(xlsxFiles, path):
    for xlsxFile in xlsxFiles:
        fileDir = path+'/'+xlsxFile

        wb = openpyxl.load_workbook(fileDir)
        for ws in wb.worksheets:

            ws.HeaderFooter.differentOddEven = True
            ws.oddFooter.left = _HeaderFooterPart('探测单位：中国电建集团华东勘测设计研究院有限公司')

            ws.oddFooter.center = _HeaderFooterPart('制表者：段磊仔    校核者：陈皎')

            ws.oddFooter.right.size = 8
            ws.oddFooter.left.size = 8
            ws.oddFooter.center.size = 8
            ws.oddFooter.right.font = "宋体"
            ws.oddFooter.right.text = "日期:2019-5      第 &[Page]-2 页    共 &N-2 页"

            ws.evenFooter.left = _HeaderFooterPart('探测单位：中国电建集团华东勘测设计研究院有限公司')
            ws.evenFooter.center = _HeaderFooterPart('制表者：段磊仔    校核者：陈皎')
            ws.evenFooter.right.size = 8
            ws.evenFooter.left.size = 8
            ws.evenFooter.center.size = 8
            ws.evenFooter.right.font = "宋体"
            ws.evenFooter.right.text = "日期:2019-5      第 &[Page]-2 页    共 &N-2 页"

        ws1 = wb.create_sheet("Mysheet", 0)
        ws1['A3'] = "封面"

        newFileDir = path+'/' + 'new_'+xlsxFile

        wb.save(newFileDir)
        statusStr.set('转换成功')


def batchConvert():
    filePath = var1.get()
    # myPathRoot(filePath)
    if filePath == "":
        statusStr.set('请选择文件夹！！')
        return

    xlsx_file_number = glob.glob(pathname=filePath+'/' + r'*.xlsx')
    print(len(xlsx_file_number))

    if len(xlsx_file_number) == 0:
        statusStr.set('不存在.xlsx的文件')
        return

    xlsxFiles = (fn for fn in os.listdir(filePath) if fn.endswith('.xlsx'))
    convertCore(xlsxFiles, filePath)


tk.Button(window, text='开始脚注', width=50,
          command=batchConvert).place(x=10, y=260)
window.mainloop()
