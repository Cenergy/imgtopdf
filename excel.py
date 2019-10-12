
import os
import openpyxl
from openpyxl.worksheet.header_footer import _HeaderFooterPart
xlsxFiles = (fn for fn in os.listdir('.') if fn.endswith('.xlsx'))


sSourceFile = "./templates/templatexls.xlsx"
wb = openpyxl.load_workbook(sSourceFile)
copy_sheet1 = wb.copy_worksheet(wb.worksheets[0])
print(copy_sheet1)


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

    ws1 = wb.create_sheet("Mysheet", 0)
    ws1['A3'] = "封面"

    wb.save('new_'+xlsxFile)
