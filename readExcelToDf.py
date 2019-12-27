import win32com.client
import csv
import sys
import pandas as pd
# from xlrd import *


xlApp = win32com.client.Dispatch("Excel.Application")
print ("Excel library version:", xlApp.Version)

filename,password = r"path_to_excel_file", 'my_password'
# xlwb = xlApp.Workbooks.Open(filename, Password=password)
xlwb = xlApp.Workbooks.Open(filename, False, True, None, password)

xlApp.Visible = False
//read multiple sheets
month = ['January', "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
final_df = pd.DataFrame() 

for i in month:
    xl_sh = xlwb.Worksheets(i)

    # Get last_row
    row_num = 0
    cell_val = ''
    while cell_val != None:
        row_num += 1
        cell_val = xl_sh.Cells(row_num, 1).Value
        # print(row_num, '|', cell_val, type(cell_val))
    last_row = row_num - 1
    # print(last_row)

    # Get last_column
    col_num = 0
    cell_val = ''
    while cell_val != None:
        col_num += 1
        cell_val = xl_sh.Cells(1, col_num).Value
        # print(col_num, '|', cell_val, type(cell_val))
    last_col = col_num - 1

    content = xl_sh.Range(xl_sh.Cells(1, 1), xl_sh.Cells(last_row, last_col)).Value

    df = pd.DataFrame(list(content[1:]), columns=content[0])
    
    final_df=final_df.append(df)

print(final_df.head())
xlwb.Close(False)
