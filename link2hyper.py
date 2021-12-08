# -*- coding: utf-8 -*-
"""
Created on Tue Jun 15 14:54:02 2021

@author: volom
"""
import os
import win32com.client as win32com

path = input(r"Enter path 'D:\file' ")
file = input("Enter file name 'file.xlsx' ")
sheet = input("Enter sheet name 'Sheet1' ")
xlrange = input("Enter range 'A1'")
os.chdir(path)
def get_range(cell_r, get='A'):
    decs = ['1', '2', '3', '4', '5', '6', '7', '8', '9']
    res = ''
    if get == 'A':
        for i in cell_r:
            if i not in decs:
                res += i
    else:
        for i in cell_r:
            if i in decs:
                res += i
    return res
    
def link_to_hyperlink(book: str, sheet: str, start_range='A1'):
    excel = win32com.gencache.EnsureDispatch('Excel.Application')
    book = os.getcwd() + '\\' + book
    wb = excel.Workbooks.Open(book)
    sheet = wb.Worksheets(sheet)
    column = get_range(start_range, get='A')
    cell = int(get_range(start_range, get='1'))
    # last_cell = sheet.Range("A" + str(sheet.Rows.Count)).End(win32com.constants.xlUp).Row
    while sheet.Range(f'{column}{cell}').Value != None:
        try:
            sheet.Hyperlinks.Add(Anchor=sheet.Range(f'{column}{cell}'), Address=sheet.Range(f'{column}{cell}').Value)
        except:
            pass
        cell += 1
    wb.Close(SaveChanges=1)
    excel.Quit()
    
# cd  C:\Users\B51\Desktop\link2
link_to_hyperlink(file, sheet, xlrange)
os.system("taskkill /f /im  EXCEL.EXE")
print("The convertation was done successfully!")
while True:
    1 == 1


