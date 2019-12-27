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

# read multiple sheets, rename sheets according to sheet name found in file
month = ['January', "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
final_df = pd.DataFrame() 

final_df = pd.DataFrame() 

for sh in xlwb.Sheets:
    
    if sh.Name in month:
        xlwb.Worksheets(sh.Name).Activate()
        readData = xlwb.Worksheets(sh.Name)
        ws = xlwb.ActiveSheet
            
        print ('Active sheet name : ', ws.Name)
        allData = readData.UsedRange

        EndRow = allData.Rows.Count
        EndCol = allData.Columns.Count
        content = ws.Range(ws.Cells(1, 1), ws.Cells(EndRow, EndCol)).Value

        df = pd.DataFrame(list(content))
        
        # set column as header
        df.columns = df.iloc[0]
    
        df.drop(df.index[1])
        
        df.fillna(value=pd.np.nan, inplace=True)
        #  drop null by subset 
        df=df.dropna(subset=['ID'])
        df=df.drop_duplicates(keep=False)

        df=df[["column_1", "column_2"]]

        final_df=final_df.append(df)
    
finaldf = final_df.reset_index(drop=True)

print(final_df.head())
xlwb.Close(False)
