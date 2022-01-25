# -*- coding: utf-8 -*-
"""
Script for pressing hotkeys to save (Ctrl-S) with time interval in given seconds
(!) It works only on Windows OS

@author: volom
"""
import time
import win32com.client
shell = win32com.client.Dispatch("WScript.Shell")
sleep_time = int(input("Put time interval to save (seconds)"))
while True:
   shell.SendKeys('^s')
   time.sleep(sleep_time)
   print("SAVED")
   
   
   