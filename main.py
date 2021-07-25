'''
Open an Excel file, read and output the existing values, update some values and save, then output the revised values.
'''

import xlrd
from os import path


def output_bal_and_sum(sheetname):
    workbook = xlrd.open_workbook(filename)

    sheet = workbook.sheet_by_name(sheetname)
    # You can also reference a sheet by index:
    # sheet = workbook.sheet_by_index(0)

    # Loop through the sheet.
    print("   Number of rows and columns: ", sheet.nrows, ",", sheet.ncols)
    i = 2
    # Skip the first 2 header rows.
    maxRows = sheet.nrows - 2

    while i <= maxRows:
        balance = sheet.cell_value(i, 2)
        # Check if the custID is not blank; if it is, skip the row
        # (Don't check the balance value for blank because 0.0 is a valid balance but it's considered blank.)
        if sheet.cell_value(i, 0):
            if balance and balance < 0:
                comment = " *** below zero"
            else:
                comment = ""
            print("  ", sheet.cell_value(i, 1), ":", balance, comment)
        i += 1

    print("   Sum: ", sheet.cell_value(1, 5))


print("\n-----\n")

filename = "excel1.xls"
filePath = path.realpath(filename)
print("File path: " + filePath)

if path.exists(filename) is False:
    print("*** File does not exist: " + filePath)
else:
    if path.isfile(filename) is False:
        print("*** This is not a file: " + filePath)
    else:
        print("\nExisting sheet values:\n")
        output_bal_and_sum("Sheet1")

        # Update some values and save.



        print("\nRevised sheet values:\n")
        output_bal_and_sum("Sheet1")
