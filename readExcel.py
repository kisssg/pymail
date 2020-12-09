import xlwings as xw
import pandas as pd
# app = xw.App(visible= True, add_book= False)
workbook = xw.Book('template.xlsx')
worksheet = workbook.sheets
column = ["name","msg1","msg2","email"]
data=[]
for i in worksheet:
    values = i.range('A1').expand().options(pd.DataFrame,index=False).value
    print(values)
    filtered=values[column]
    data.append(filtered)
workbook.close()