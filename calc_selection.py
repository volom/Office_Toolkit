# -*- coding: utf-8 -*-
 
import win32com.client as win32com
try:
    excel = win32com.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True # delete in case of .exe transformation
    excel.Selection.Calculate()
    print("Calculation of selected range was done successfully")
    while True:
        1 + 1
except Exception as e:
    print(e)
    while True:
        1 + 1




