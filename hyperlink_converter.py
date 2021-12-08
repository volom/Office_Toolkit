# -*- coding: utf-8 -*-
"""
Created on Fri Jun 11 13:39:46 2021

@author: volom
"""
import os
import xlwings as xw
import re
from tqdm import tqdm

def link_to_hyperlink(book: str, sheet: str, start_range='A1'):
    app = xw.App(visible=False, add_book=False)
    book = xw.Book(book, local=False, update_links=False)
    sheet = book.sheets(sheet)
    column = re.match(r'.*[^\d*]', start_range).group(0)
    cell = int(re.findall('(\d*)', start_range)[1])
    last_cell = sheet.range(column + str(sheet.cells.last_cell.row)).end('up').row
    for cell_n in tqdm(range(cell, last_cell+1), position=0, leave=False):
        try:
            xw.Range(f'{column}{cell_n}').add_hyperlink(xw.Range(f'{column}{cell_n}').value)   
        except:
            pass
    book.save()
