# -*- coding: utf-8 -*-
"""
Created on Tue Jun 29 17:06:29 2021
script which helps to delete metadata from ms office files located in 
current directory
@author: volom
"""
import win32com.client as win32com
import os
import re

list_files = []
for (dirpath, dirnames, filenames) in os.walk(os.getcwd()):
    list_files += [os.path.join(dirpath, file) for file in filenames]

list_files = [x for x in list_files if re.findall(r'\.(.*)', x)[0] in ['xlsx', 'xls', 'docx', 'doc', 'pptx', 'ppt']]

for book in list_files:
	try:
		try:
			excel = win32com.gencache.EnsureDispatch('Excel.Application')
			try:
				wb = excel.Workbooks.Open(book, Local=True)
			except:
				wb = excel.Workbooks.Open(book, Local=False)
			wb.RemovePersonalInformation = True
			wb.Close(SaveChanges=1)
			excel.Quit()
		except:
			word = win32com.gencache.EnsureDispatch('Word.Application')
			wb = word.Documents.Open(book)
			wb.RemovePersonalInformation = True
			wb.Save()
			word.Quit()
		finally:
			pp = win32com.gencache.EnsureDispatch('PowerPoint.Application')
			wb = pp.Presentations.Open(book, WithWindow=False)
			wb.RemovePersonalInformation = True
			wb.Save()
			wb.Close()
	except:
		pass

print("Deleting metadata was done successfully!")
while True:
    1 + 1
