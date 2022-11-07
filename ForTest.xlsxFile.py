# This file will update and save Test.xlsx

# Sources: https://www.geeksforgeeks.org/python-script-to-automate-refreshing-an-excel-spreadsheet/

# pip install pywin32

# Importing the pywin32 module
import win32com.client
from time import sleep
from datetime import date
  
# Opening Excel software using the win32com
File = win32com.client.Dispatch("Excel.Application")
  
# Optional line to show the Excel software
print("Opening Excel...")
File.Visible = 1
  
# Opening your workbook
print("Opening your excel file...")
Workbook = File.Workbooks.Open("C:/Users/JP.Mitiguy23/Mitiguy/Test.xlsx")

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

# compares the excel data in A2, (2, 1) in its language, with today's date
if str(Worksheet.Cells(2, 1).Value) == new_today_date:
    print("Appears to have worked!")
    # Saves the workbook
    Workbook.Save()
    print("Saving file...")
    # Closes the Excel file
    File.Quit()
    print("Quitting Excel...")
else:
    print("Check for possible errors")