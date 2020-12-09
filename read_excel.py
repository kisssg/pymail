import xlrd
import pandas as pd
book = xlrd.open_workbook('template.xlsx')
sheet = book.sheet_by_name('Sheet1')
data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
data=pd.DataFrame(data,None,data[0])
print(xlrd.xldate.xldate_as_datetime((data['外访日期'][1]), 0).time())