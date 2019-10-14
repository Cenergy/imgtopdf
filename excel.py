
import os
import openpyxl
from openpyxl.worksheet.header_footer import _HeaderFooterPart
import pandas as pd

from openpyxl.drawing.image import Image

from openpyxl.styles import Color, Font, Alignment
xlsxFiles = (fn for fn in os.listdir('.') if fn.endswith('.xlsx'))


sSourceFile = "./templates/templatexls.xlsx"
wbCopy = openpyxl.load_workbook(sSourceFile)


for xlsxFile in xlsxFiles:
    wb = openpyxl.load_workbook(xlsxFile)

    for ws in wb.worksheets:
        # ws.HeaderFooter.differentFirst = True

        ws.HeaderFooter.differentOddEven = True

        # ws.oddHeader.right = _HeaderFooterPart('奇数页右页眉')
        ws.oddFooter.left = _HeaderFooterPart('探测单位：中国电建集团华东勘测设计研究院有限公司')

        ws.oddFooter.center = _HeaderFooterPart('制表者：段磊仔    校核者：陈皎')

        ws.oddFooter.right.size = 8
        ws.oddFooter.left.size = 8
        ws.oddFooter.center.size = 8
        ws.oddFooter.right.font = "宋体"
        ws.oddFooter.right.text = "日期:2019-5      第 &[Page]-2 页    共 &N-2 页"

        # ws.evenHeader.left = _HeaderFooterPart('偶数页左页眉')
        ws.evenFooter.left = _HeaderFooterPart('探测单位：中国电建集团华东勘测设计研究院有限公司')
        ws.evenFooter.center = _HeaderFooterPart('制表者：段磊仔    校核者：陈皎')
        # ws.evenFooter.right = _HeaderFooterPart(
        #     '日期:2019-5      第&[页码]-2页    共&[总页数]-2页')
        ws.evenFooter.right.size = 8
        ws.evenFooter.left.size = 8
        ws.evenFooter.center.size = 8
        ws.evenFooter.right.font = "宋体"
        ws.evenFooter.right.text = "日期:2019-5      第 &[Page]-2 页    共 &N-2 页"

    ws0 = wb.get_sheet_by_name('目录')
    dataFrame = pd.read_excel(xlsxFile, sheet_name='目录')

    print(ws0.max_row)
    value12 = ws0['A12'].value.split('…………………………………………………')
    value13 = ws0['A13'].value.split('…………………………………………………')
    value14 = ws0['A14'].value.split('…………………………………………………')
    value15 = ws0['A15'].value.split('…………………………………………………')
    value16 = ws0['A16'].value.split('…………………………………………………')
    value17 = ws0['A17'].value.split('…………………………………………………')
    value18 = ws0['A18'].value.split('…………………………………………………')
    value19 = ws0['A19'].value.split('…………………………………………………')
    ws00 = wb.create_sheet("封面0", 0)

    varStr = '………………………………………………………………………………………………………………'

    ws00['B4'].value = varStr
    ws00['A4'].value = value12[0]
    ws00['C4'].value = value12[1]
    ws00['B5'].value = varStr
    ws00['A5'].value = value13[0]
    ws00['C5'].value = value13[1]
    ws00['B6'].value = varStr
    ws00['A6'].value = value14[0]
    ws00['C6'].value = value14[1]
    ws00['B7'].value = varStr
    ws00['A7'].value = value15[0]
    ws00['C7'].value = value15[1]
    ws00['B8'].value = varStr
    ws00['A8'].value = value16[0]
    ws00['C8'].value = value16[1]
    ws00['B9'].value = varStr
    ws00['A9'].value = value17[0]
    ws00['C9'].value = value17[1]
    ws00['B10'].value = varStr
    ws00['A10'].value = value18[0]
    ws00['C10'].value = value18[1]
    ws00['B11'].value = varStr
    ws00['A11'].value = value19[0]
    ws00['C11'].value = value19[1]

    print(value12, "=============")

    ws1 = wb.create_sheet("封面", 0)

    # 调整列宽
    ws1.column_dimensions['A'].width = 117.88

    # 调整行高
    for i in range(1, 9):
        ws1.row_dimensions[i].height = 60

    fontObj1 = Font(name=u'楷体_GB2312', bold=True, italic=False, size=24)
    fontObj2 = Font(name=u'楷体_GB2312', bold=True, italic=False, size=30)
    fontObj3 = Font(name=u'楷体', bold=True, italic=False, size=18)
    alignObj1 = Alignment(horizontal='center',
                          vertical='center', wrapText=True)
    ws1['A2'] = "市政文锦渠及东湖公园暗涵综合整治和清污剥离工程—文锦渡口岸泵站"
    ws1['A2'].font = fontObj1
    ws1['A2'].alignment = alignObj1
    ws1['A3'] = "管线点成果表"
    ws1['A3'].font = fontObj2
    ws1['A3'].alignment = alignObj1
    ws1['A8'] = "二〇一九年九月"
    ws1['A8'].font = fontObj3
    ws1['A8'].alignment = alignObj1

    img = Image('images/img.png')
    newsize = (567, 71)
    img.width, img.height = newsize  # 这两个属性分别是对应添加图片的宽高

    img.anchor = 'A7'

    ws1.add_image(img)

    wb.save('new_'+xlsxFile)
