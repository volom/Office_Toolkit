# -*- coding: utf-8 -*-
"""
Created on Wed Jun 23 09:24:19 2021

@author: volom
"""
import os
import win32com.client as win32com
win32c = win32com.constants



pp = win32com.gencache.EnsureDispatch('PowerPoint.Application')
pp.Visible = True
presentation = pp.Presentations.Open(r'', ReadOnly=False)

Slide = presentation.Slides.Range(1)
Slide.Select()
Slide.Shapes.Range(2).TextFrame.TextRange.Text = "Hello, world\nHi\Gracia"
presentation.Save()
presentation.Close()
pp.Quit()
os.system("taskkill /f /im  POWERPNT.EXE")




