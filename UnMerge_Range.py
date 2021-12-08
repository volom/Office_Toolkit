# -*- coding: utf-8 -*-
"""
Created on Wed Jun 23 08:42:36 2021
unmerging cells and putting value from merged in each separated ones
@author: volom
"""
import os
import win32com.client as win32com
from tqdm import tqdm

try:
    dirr = input("Enter dir with file 'D:\': ")
    book = input("Enter file name 'file.xlsx': ")
    sheet = input("Enter sheet name 'Sheet1': ")
    xlrange = input("Enter range 'A1:B2': ").upper()
    
    os.chdir(dirr)
    
    excel = win32com.gencache.EnsureDispatch('Excel.Application')
    book = os.getcwd() + '\\' + book
    wb = excel.Workbooks.Open(book)
    sheet = wb.Worksheets(sheet)
    sheet.Select()
    for cell in tqdm(sheet.Range(xlrange), position=0, leave=False):
        value = cell.Value
        cell.Select()
        cell.MergeCells = False
        excel.Selection.Value = value
    wb.Close(SaveChanges=1)
    excel.Quit()
    print("The procedure was done successfully!")
except:
    print("Cannot run...")
    print("Possible errors: ")
    print("--- Wrong dirs")
    print("--- Wrong range")
    print("--- Book and/or sheet are restricted from editing")
    print("--- The script can not run if the excel file is open while running")
while True:
    1 + 1
