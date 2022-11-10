# https://docs.xlwings.org/en/latest/api.html
# https://www.geeksforgeeks.org/working-with-excel-files-in-python-using-xlwings/

# Importing the pywin32 module
import xlwings
from time import sleep
from datetime import date
import ForTestxlsxFile

# From ForTestxlsxFile, run function that downloads most recent stock prices
ForTestxlsxFile.update_file()

# sleep for 3 seconds to ensure smooth transition from openpyxl to xlwings code
sleep(3)

############# Load xlwings ##################
print("xlwings_____________________________________")

# Load workbook
print("Loading workbook...")
Workbook = xlwings.Book("C:/Users/JP.Mitiguy23/Mitiguy/FidelityHoldingsForTest.xlsm")

# Finds active sheet in workbook
print("Pulling up Main Sheet...")
Worksheet = Workbook.sheets['MainSheet']

# Refresh data
print("Refreshing data...")
Workbook.api.RefreshAll()

################ NEW ACTIVITY ###################

# Find blank row
print("Finding blank row...")
e = Worksheet.range("E:E").value
searching = True
while searching == True:
    for i in range(800):
        cell = e[i]
        if cell == None:
            global blank_row
            blank_row = i + 1
            print(blank_row)
            searching = False
            break

# Add recent activities
num_add = int(input("How many activity additions would you like to input? "))
for n in range(num_add):
    trade_date = input("Trade Date (if any) mm/dd/yy: ")
    settlement_date = input("Settlement Date mm/dd/yy: ")
    description = input("Enter activity description (e.g., You Sold Transaction Profit: $3.25): ")
    quantity = input("Enter Quantity (negative for sold): ")
    price = input("Enter price: ")
    cost = input("Enter cost (if any): ")
    transaction_cost = input("Enter transaction cost (if any): ")
    amount = input("Enter amount (negative for buy): ")
    # if ref. num. exists, fill in other values automatically
    ref_num = input("Enter reference number (if any): ")
    if ref_num != "":
        type = "1*"
        reg_rep = "E##"
        order_num = input("Enter order number: ")
    else:
        type = ""
        reg_rep = ""
        order_num = ""

    # while loop takes in fund/stock code data & adds known info accordingly
    take_info = True
    while take_info:
        information_input = input("Fund/Stock Code: ")
        if information_input == "SPAXX":
            information_output = "Fidelity Government Money Market (SPAXX)"
            if description != "Check Received":
                symbol_cusip = "31617H102"
            else:
                symbol_cusip = ""
            take_info = False
        elif information_input == "FNCMX":
            information_output = "Fidelity Nasdaq Composite Index Fund (FNCMX)"
            take_info = False
            symbol_cusip = "315912709"
        elif information_input == "FBGRX":
            information_output = "Fidelity Blue Chip Growth Fund (FBGRX)"
            symbol_cusip = "316389303"
            take_info = False
        elif information_input == "FOCPX":
            information_output = "Fidelity OTC Portfolio (FOCPX)"
            symbol_cusip = "316389105"
            take_info = False
        elif information_input == "FNILX":
            information_output = "Fidelity Zero Large Cap Index Fund (FNILX)"
            symbol_cusip = "315911628"
            take_info = False
        elif information_input == "FLCEX":
            information_output = "Fidelity Large Cap Core Enhanced Index Fund (FLCEX)"
            symbol_cusip = "31606X100"
            take_info = False
        elif information_input == "FFGCX":
            information_output = "Fidelity Global Commodity Stock Fund (FFGCX)"
            symbol_cusip = "31618H606"
            take_info = False

    # add data to blank rows in excel
    print("Filling in data...")
    input_cell = "A" + str(blank_row + n)
    Worksheet.range(input_cell).value = [trade_date, settlement_date, information_output, symbol_cusip, description, quantity, price, cost, transaction_cost, amount, ref_num, type, reg_rep, order_num]
    

################ Update Table #########################
# Find the date in spot V2 (last updated date)
print("Old date: ")
old_date = Worksheet.range("V2").value
print(old_date)

# Create new worksheet loaded to the final worksheet
Worksheet_Test = Workbook.sheets["Final (2)"]
# Find value in B2 of Final (2) worksheet (may be blank)
print("Value in B2 of Final (2): ")
test_B2 = Worksheet_Test.range("B2").value
print(test_B2)

# find the latest price to update the main sheet with
print("Latest price: ")
if test_B2 == None:
    latest_price = Worksheet_Test.range("B3").value
    latest_row = "3"
else:
    latest_price = test_B2
    latest_row = "2"
print(latest_price)

# Determine which the old date needs to go once the new info is added
print("Old date will now go in this row: ")
a = Worksheet_Test.range("A:A").value
searching = True
while searching == True:
    # searches through A column to find the date
    for i in range(100):
        cell = a[i]
        if cell == old_date:
            global old_date_new_row
            # Determines if a one cell adjustment needs to be made to account for blank 1st row
            if test_B2 == None:
                old_date_new_row = i
            else:
                old_date_new_row = i + 1
            print(old_date_new_row)
            searching = False
            break

# first cell of old date
print("Old table will start in this cell: ")
old_date_new_cell = "V" + str(old_date_new_row)
print(old_date_new_cell)
# select the entire table
print("Pasting table below...")
table = Worksheet.range("V2").expand().formula
# paste the table into its new spot
Worksheet.range(old_date_new_cell).expand().formula = table

# Macro Code
'''

Sub DeleteExtraAtSymbol()
'
' DeleteExtraAtSymbol Macro
' When I copy and paste using xlwings and python, it inputs an "@" symbol into some formulas. This macro is designed to remove these..
'

'
    Columns("V:AQ").Select
    Selection.Replace What:="@$", Replacement:="$", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

'''


# use macro to delete extra at symbol that has popped up
print("Running macro to delete extra '@' symbol...")
DeleteExtraAtSymbolMacro = Workbook.macro("DeleteExtraAtSymbol")
DeleteExtraAtSymbolMacro()

# copy and paste the column of dates
new_dates_test_column = "A" + latest_row + ":A" + str(old_date_new_row)
# ndim=2 ensures copied column is pasted as column and not row:
# https://github.com/xlwings/xlwings/issues/398#:~:text=Note%20that%20currently%2C%201d%20arrays%20still%20require%20ndim%3D2%20to%20preserve%20the%20column%20orientation
new_dates_test = Worksheet_Test.range(new_dates_test_column).options(ndim=2).value
new_dates_column = "V2:V" + str(old_date_new_row)
print("Pasting new dates from 'Final (2)' cells " + new_dates_test_column + " in 'MainSheet' cells " + new_dates_column)
Worksheet.range(new_dates_column).value = new_dates_test

# copy and paste the column of FNCMX prices
new_fncmx_test_column = "B" + latest_row + ":B" + str(old_date_new_row)
new_fncmx_test = Worksheet_Test.range(new_fncmx_test_column).options(ndim=2).value
new_fncmx_column = "W2:W" + str(old_date_new_row)
print("Pasting new FNCMX prices from 'Final (2)' cells " + new_fncmx_test_column + " in 'MainSheet' cells " + new_fncmx_column)
Worksheet.range(new_fncmx_column).value = new_fncmx_test

# copy and paste the column of FBGRX prices
new_fbgrx_test_column = "C" + latest_row + ":C" + str(old_date_new_row)
new_fbgrx_test = Worksheet_Test.range(new_fbgrx_test_column).options(ndim=2).value
new_fbgrx_column = "Z2:Z" + str(old_date_new_row)
print("Pasting new FBGRX prices from 'Final (2)' cells " + new_fbgrx_test_column + " in 'MainSheet' cells " + new_fbgrx_column)
Worksheet.range(new_fbgrx_column).value = new_fbgrx_test

# copy and paste the column of FOCPX prices
new_focpx_test_column = "D" + latest_row + ":D" + str(old_date_new_row)
new_focpx_test = Worksheet_Test.range(new_focpx_test_column).options(ndim=2).value
new_focpx_column = "AC2:AC" + str(old_date_new_row)
print("Pasting new FOCPX prices from 'Final (2)' cells " + new_focpx_test_column + " in 'MainSheet' cells " + new_focpx_column)
Worksheet.range(new_focpx_column).value = new_focpx_test

# copy and paste the column of FNILX prices
new_fnilx_test_column = "E" + latest_row + ":E" + str(old_date_new_row)
new_fnilx_test = Worksheet_Test.range(new_fnilx_test_column).options(ndim=2).value
new_fnilx_column = "AF2:AF" + str(old_date_new_row)
print("Pasting new FNILX prices from 'Final (2)' cells " + new_fnilx_test_column + " in 'MainSheet' cells " + new_fnilx_column)
Worksheet.range(new_fnilx_column).value = new_fnilx_test

# copy and paste the column of FNILX prices
new_fnilx_test_column = "E" + latest_row + ":E" + str(old_date_new_row)
new_fnilx_test = Worksheet_Test.range(new_fnilx_test_column).options(ndim=2).value
new_fnilx_column = "AF2:AF" + str(old_date_new_row)
print("Pasting new FNILX prices from 'Final (2)' cells " + new_fnilx_test_column + " in 'MainSheet' cells " + new_fnilx_column)
Worksheet.range(new_fnilx_column).value = new_fnilx_test




# Update lines Y, AB, AE, AH, AK, AN