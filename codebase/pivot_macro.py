from apinfvi import eSight
from DateTime import DateTime
from shutil import copyfile
import datetime
#from assetpivot import *
import vars
import pandas as pd
import json
import os, sys
import glob
import numpy as np
import openpyxl
import win32com.client as win32
import win32api
import win32con
import time
import shutil
import timeit

print("Start Running Script for PIVOT ASSET INVENTORY\n-----------------------------------------------------\n")
start = timeit.default_timer()
current_time = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
print('Start Time: {}'.format (current_time))

print('Converting xlsx to macro ===============================\n')
wb = win32.Dispatch('Excel.Application')
wb.Visible = True
ws = wb.Workbooks.Open(r'//10.150.20.30/SIM Working Directory/Scripts/Asset Inventory/Project Pivot/Script/pivot_report/PIVOT ASSET INVENTORY.xlsx')
ws.SaveAs(Filename='PIVOT ASSET INVENTORY - {}.xlsm'.format(current_time[:8]), FileFormat = 52)
wb.Workbooks(1).Close(SaveChanges=1)
wb.Application.Quit()

print('Copying to log file ===============================\n')
src = 'C:/Users/Administrator/Documents/'
dtn = '//10.150.20.30/SIM Working Directory/Scripts/Asset Inventory/Project Pivot/Log/'
current_path = os.chdir(src)
shutil.move('{}PIVOT ASSET INVENTORY - {}.xlsm'.format(src,current_time[:8]), dtn)

print('Running macro. PLEASE DO NOT TERMINATE EXCEL. ===============================\n')
with open('//10.150.20.30/SIM Working Directory/Scripts/Asset Inventory/Project Pivot/Script/pivot_macro/macro.txt', "r") as myfile:
    print('reading macro into string from: ' + str(myfile))
    macro=myfile.read()
    xl = win32.Dispatch('Excel.Application')
    xl.Visible = True
    wk = xl.Workbooks.Open(r'{}PIVOT ASSET INVENTORY - {}.xlsm'.format(dtn,current_time[:8]))
    #wkt = wk.Worksheets('Sheet1')
    key = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER,
                                "Software\\Microsoft\\Office\\16.0\\Excel"
                                + "\\Security", 0, win32con.KEY_ALL_ACCESS)
    win32api.RegSetValueEx(key, "AccessVBOM", 0, win32con.REG_DWORD, 1)
    xlmodule = wk.VBProject.VBComponents.Add(1)
    xlmodule.CodeModule.AddFromString(macro)
    xl.Application.Run('pivot')
    xl.Workbooks(1).Close(SaveChanges=1)
    xl.Application.Quit()

print('Copying to routine log ===============================\n')
s_log = '//10.150.20.30/SIM Working Directory/Scripts/Asset Inventory/Project Pivot/Log/'
d_log = '//10.150.20.30/SIM Working Directory/Routines/Asset Inventory/Project Pivot/'
current_path = os.chdir(s_log)
file = os.listdir(current_path)[-1]
c_file = '{}PIVOT ASSET INVENTORY - {}.xlsm'.format(d_log, current_time[:8])
shutil.copyfile(file, c_file)
shutil.copystat(file, c_file)

current_time2 = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
print('End Time: {}'.format(current_time2))
stop = timeit.default_timer()
print('Total time to run the program: ', stop - start)
print("_________________________END_________________________")



