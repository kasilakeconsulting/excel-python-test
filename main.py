'''
Open an Excel file, read and output the existing values, update some values and save, then output the revised values.
'''

import xlrd
import xlwt
from xlutils.copy import copy
from os import path
from random import randint


def output_bal_and_sum(filenameRd, sheetname):
    workbookRd = xlrd.open_workbook(filenameRd)

    sheetRd = workbookRd.sheet_by_name(sheetname)
    # You can also reference a sheet by index:
    # sheet = workbook.sheet_by_index(0)

    # Loop through the sheet.
    print("   Number of rows and columns: ", sheetRd.nrows, ",", sheetRd.ncols)
    i = 2
    # Skip the first 2 header rows.
    maxRows = sheetRd.nrows - 1

    while i <= maxRows:
        balance = sheetRd.cell_value(i, 2)
        # Check if the custID is not blank; if it is, skip the row
        # (Don't check the balance value for blank because 0.0 is a valid balance but it's considered blank.)
        if sheetRd.cell_value(i, 0):
            if balance and balance < 0:
                comment = " *** below zero"
            else:
                comment = ""
            print("  ", sheetRd.cell_value(i, 1), ":", balance, comment)
        i += 1

print("\n-----\n")

filename = "excel1.xls"
# If you cannot or do not want to use the OS package, delete the following 2 lines and uncomment the 3rd line.
filePath = path.realpath(filename)
print("File path: " + filePath)
#filePath = filename

if path.exists(filename) is False:
    print("*** File does not exist: " + filePath)
else:
    if path.isfile(filename) is False:
        print("*** This is not a file: " + filePath)
    else:
        print("\nExisting sheet values:\n")
        output_bal_and_sum(filename, "Sheet1")

        # BEGIN: If you cannot or do not want to use the xlwt package, delete the lines between here and END.
        # Update some values and save.

        print("\nUpdating...")
        wkRd = xlrd.open_workbook(filename)
        # If you want to test the write, the xlutils package must be included to be able to copy.
        workbookWt = copy(wkRd)
        sheetWt = workbookWt.get_sheet(0)

        # If you cannot or do not want to use the random package, change the following parameters to something literal.
        sheetWt.write(2, 2, randint(-10, 10))
        sheetWt.write(4, 2, randint(-10, 10))

        # Warning - saving may reformat the spreadsheet.
        workbookWt.save(filename)

        print("\nRevised sheet values:\n")
        output_bal_and_sum(filename, "Sheet1")

        # END: Delete up to here if you cannot or do not want to se ethe xlwt package.
