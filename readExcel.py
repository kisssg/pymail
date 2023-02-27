import xlwings as xw
import pandas as pd
# app = xw.App(visible= True, add_book= False)
worksheet = xw.Book('template.xlsx').sheets[0]
column = ["employee_id","mail_to"]
values = worksheet.range('A1').expand().options(pd.DataFrame,index=False).value
filtered=values[column]
print(filtered)
# workbook.close()
