# -*- coding: utf-8 -*-
"""
Created on Wed Jun 23 08:42:36 2021
unmerging cells and putting value from merged in each separated ones.
Just select excel range you want to procede and run the script
@author: volom
"""
import os
import win32com.client as win32com
from tqdm import tqdm

excel = win32com.gencache.EnsureDispatch('Excel.Application')

for cell in tqdm(excel.Selection, position=0, leave=False):
    value = cell.Value
    cell.Select()
    cell.MergeCells = False
    excel.Selection.Value = value

print("The procedure was done successfully!")
