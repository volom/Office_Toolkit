# change extention of Excel file from xlsb to xlsx

import win32com.client as win32com

excel = win32com.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True # delete in case of .exe transformation
excel.Application.EnableEvents = False
excel.Application.ScreenUpdating = True
excel.Application.DisplayAlerts = False

# file formats
# https://learn.microsoft.com/uk-ua/office/vba/api/Excel.XlFileFormat

def xlsb2xlsx(file):
    wb = excel.Workbooks.Open(file, Local=False)
    wb.SaveAs(file[:-5], FileFormat=win32com.constants.xlOpenXMLWorkbook)
