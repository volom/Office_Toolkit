# -*- coding: utf-8 -*-
"""
the macros can help you to merge cells with value in each ones
and add merged value to merged cells

@author: volom
"""

import os
import re
import win32com.client as win32com
win32c = win32com.constants



excel = win32com.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True # delete in case of .exe transformation
excel.Application.EnableEvents = False
excel.Application.ScreenUpdating = True
excel.Application.DisplayAlerts = False

# create variable with value in selected cells
value_selection = excel.Selection.Value


# merge cells
excel.Selection.Merge()


# Add merged value to merged cells

excel.Selection.Value = ' '.join([str(j) for i in value_selection for j in i if j != None])
excel.Application.DisplayAlerts = True

