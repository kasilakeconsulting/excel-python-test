'''
Open an Excel file and read content from certain cells.
'''

import xlrd
import platform
import os
from os import path

print("-----\n\n")

nodename = platform.node()

filename = "excel1.xls"

sheetname = "Sheet 1"

print("os: " + os.name)
print(os.uname())
print('system: ' + platform.system())
print("current directory: " + path.curdir)
print("default path: " + path.defpath)
print("file path: " + path.realpath(filename))

print("\nfile exists? " + str(path.exists(filename)))

if path.exists(filename) is False:
    print("*** Check if the file needs to be downloaded from the cloud")
else:
    print("is a file? " + str(path.isfile(filename)))

    workbook = xlrd.open_workbook(filename)
    # sheet = workbook.sheet_by_name(sheetname)
    sheet = workbook.sheet_by_index(0)

    print (sheet.cell_value(0, 0))
    print (sheet.cell_value(1, 0))
    print (sheet.cell_value(2, 0))
