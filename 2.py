import pandas as pd


path1 = './templates/templatexls.xls'
# change xxx with the sheet name that includes the data
data = pd.read_excel(path1, sheet_name='封面')


path2 = 'Target.xlsx'

data.to_excel(path2, sheet_name='new_tab')

# save it to the 'new_tab' in destfile
# data.to_excel(destfile, sheet_name='new_tab')
