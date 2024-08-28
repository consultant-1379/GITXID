import xlwt#Please install xlwt module
from tempfile import TemporaryFile#Please install TemporaryFile module
import os#Please install os module
import pandas as pd#Please install pandas module

#---------------------------------------------------------------------------------------------#
#Excel1.xlsx-->reference list
#Excel2.xlsx-->our list
#---------------------------------------------------------------------------------------------#
excel_1 = 'C:\Python27\excelcomparission\Excel1.xlsx'
excel_1_wb = pd.ExcelFile(excel_1)
xl1 = excel_1_wb.parse("Sheet1")
Mail_excel_1 = xl1['SIGNUM'].tolist()
#---------------------------------------------------------------------------------------------#
excel_2 = 'C:\Python27\excelcomparission\Excel2.xlsx'
excel_2_wb = pd.ExcelFile(excel_2)
xl2 = excel_2_wb.parse("Sheet1")
Mail_excel_2 = xl2['SIGNUM'].tolist()
#---------------------------------------------------------------------------------------------#
MissingId = xlwt.Workbook()
Finalsheet = MissingId.add_sheet('sheet1')
req_list = (list(set(Mail_excel_1) - set(Mail_excel_2)))
Final = ["SIGNUM"]+req_list

for i,e in enumerate(Final):
    Finalsheet.write(i,0,e)
if os.path.exists("MissingID.xls"):
    os.remove("MissingID.xls")
MissingIdfile = "MissingID.xls"
MissingId.save(MissingIdfile)
MissingId.save(TemporaryFile())

		

