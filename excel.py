
import os
import openpyxl
from openpyxl.worksheet.header_footer import _HeaderFooterPart

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
