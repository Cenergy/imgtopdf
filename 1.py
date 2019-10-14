import xlwings as xw

path1 = 'ExcelMod.xlsx'
path2 = 'Target.xlsx'

wb1 = xw.Book(path1)
wb2 = xw.Book(path2)

ws1 = wb1.sheets(1)
ws1.api.Copy(Before=wb2.sheets(1).api)
wb2.save()
wb2.app.quit()
