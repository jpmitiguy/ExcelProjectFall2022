# This file contains a function that will update and save StockAndMutualFundInfo.xlsx

# Sources: https://www.geeksforgeeks.org/python-script-to-automate-refreshing-an-excel-spreadsheet/
# https://stackoverflow.com/questions/45183713/open-excel-file-to-run-macro-from-relative-file-path-in-python

# Importing the pywin32 module
import win32com.client
# Import pre-installed libraries
from time import sleep
from datetime import date
import os

# function that updates StockAndMutualFundInfo.xlsx
def update_file():
  
    # Opening Excel software using the win32com
    File = win32com.client.Dispatch("Excel.Application")
    
    # Optional line to show the Excel software
    print("Opening Excel...")
    File.Visible = 1
    
    # Opening your workbook
    # solve to relative filepath errors
    # https://stackoverflow.com/questions/45183713/open-excel-file-to-run-macro-from-relative-file-path-in-python 
    Filedir = os.path.dirname(os.path.realpath('__file__'))
    filename = os.path.join(Filedir, "StockAndMutualFundInfo.xlsx")
    print("Opening " + filename)
    Workbook = File.Workbooks.Open(filename)

    # Waits 2 sec (uses time library)
    sleep(2)

    # Refeshing all the sheets
    print("Refreshing data...")
    Workbook.RefreshAll()

    # Holds program until the refresh is completed
    File.CalculateUntilAsyncQueriesDone()

    # Waits 5 sec (uses time library)
    sleep(5)

    # Finds active sheet in workbook
    Worksheet = Workbook.ActiveSheet

    # date is a class from the datetime library, today is a method in that class
    today_date = date.today()
            
            # # Changes the layout of the date to mm/dd/yy
            # d1 = today_date.strftime("%m/%d/%y")

    # creates new date that fits the excel date data type
    new_today_date = str(today_date) + str(" 00:00:00+00:00")

    # compares the excel data in A2, (2, 1) in its syntax, with today's date
    if str(Worksheet.Cells(2, 1).Value) == new_today_date:
        print("Appears to have worked!")
        # Saves the workbook
        Workbook.Save()
        print("Saving file...")
        # Closes the Excel file
        File.Quit()
        print("Quitting Excel...")
    # if the excel data in A2 doesn't match today's data, an error may have occured (e.g., not connected to internet)
    else:
        print("Check for possible errors")