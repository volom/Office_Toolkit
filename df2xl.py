#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Dec  8 20:28:50 2021
function for export DataFrame to Excel app using win32com
@author: volom
"""

def df2xl(df, xlworksheet, start_range_row=1, start_range_col="A", columns=True):
    xlworksheet.Activate()
    last_row = len(df) + (start_range_row-1)
    if len(df) != 0:
        xlworksheet.Range(f"{start_range_col}{start_range_row}:{start_range_col}{last_row}").Select()
        lst_columns = list(df.columns)
        if columns:
            for column in lst_columns:
                xlworksheet.Cells(excel.Selection.Row, excel.Selection.Column).Value = column
                tuplecol = tuple([(str(x),) for x in list(df.loc[:, column])])
                tuplecol = [('',) if x == ('None',) or x == ('NaT',) or x == ('nan',) else x for x in tuplecol]
                try:
                    xlworksheet.Range(xlworksheet.Cells(excel.Selection.Row+1, excel.Selection.Column), xlworksheet.Cells(last_row+1, excel.Selection.Column)).Value = tuplecol
                except:
                    pass
                excel.Selection.Cells(1, 2).Select()
        else:
            for column in lst_columns:
                tuplecol = tuple([(str(x),) for x in list(df.loc[:, column])])
                tuplecol = [('',) if x == ('None',) or x == ('NaT',) or x == ('nan',) else x for x in tuplecol]
                try:
                    xlworksheet.Range(xlworksheet.Cells(excel.Selection.Row, excel.Selection.Column), xlworksheet.Cells(last_row, excel.Selection.Column)).Value = tuplecol
                except:
                    pass
                excel.Selection.Cells(1, 2).Select()  