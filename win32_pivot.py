# -*- coding: utf-8 -*-
"""
Created on Tue Jun 22 13:52:53 2021

@author: volom
"""
import os
import pandas as pd
import win32com.client as win32com
win32c = win32com.constants


file = 'pivot_tables.xlsx'
excel = win32com.gencache.EnsureDispatch('Excel.Application')
book = os.getcwd() + '\\' + file
wb = excel.Workbooks.Open(book, Local=True)
sheet = wb.Worksheets('Лист20')
sheet.Activate()


ws1 = wb.Worksheets('Sheet1')


ws2_name = 'pivot_table'
wb.Sheets.Add().Name = ws2_name
ws2 = wb.Sheets(ws2_name)

pt_name = 'pivot_table'
pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)

# create the pivot table object
pc.CreatePivotTable(TableDestination=f'{ws2_name}!R1C1', TableName= pt_name)

# Adding columns, rows and filters to pivot table
# (([FILTERS], win32c.xlPageField), ([ROWS], win32c.xlRowField), ([COLUMNS], win32c.xlColumnField))
for field_list, field_r in ((['AAA'], win32c.xlPageField), (['BBB'], win32c.xlRowField), (['CCC'], win32c.xlColumnField)):
    for i, value in enumerate(field_list):
        ws2.PivotTables(pt_name).PivotFields(value).Orientation = field_r
        ws2.PivotTables(pt_name).PivotFields(value).Position = i + 1

# Adding values to pivot tables
ws2.PivotTables(pt_name).AddDataField(ws2.PivotTables(pt_name).PivotFields("A"), "Sum by A", win32c.xlSum)
ws2.PivotTables(pt_name).AddDataField(ws2.PivotTables(pt_name).PivotFields("B"), "Sum by B", win32c.xlSum)


 
        
