'''
My final project is to create a robust, concise, simple way to manage updated stock and mutual fund activity
'''

'''
Sources:
I used https://docs.xlwings.org/en/latest/api.html to examine the full documentaion of xlwings library
I sued https://www.geeksforgeeks.org/working-with-excel-files-in-python-using-xlwings/ to understand the basics of using Excel with xlwings
I used https://www.w3schools.com/python/python_try_except.asp to learn use try & except 
I used https://www.geeksforgeeks.org/python-reversing-list/#:~:text=Using%20reversed()%20we%20can,to%20reverse%20list%20in%2Dplace. to reverse lists
'''


# Import download libraries
import xlwings
# Import built-in libraries
from time import sleep
from datetime import date
# Import created module
import StockInfo

# global variables

# utility functions

# classes

# From StockInfo, run function that downloads most recent stock prices
# try:
#     StockInfo.update_file()
# except IndentationError:
#     print("Encountered an Indentation Error.")
# except:
#     print("Encountered unknown error.")


# sleep for 3 seconds to ensure smooth transition from openpyxl to xlwings code
sleep(3)

############# LOAD XLWINGS ##################
print("xlwings_____________________________________")

# Load workbook
print("Loading workbook...")
Workbook = xlwings.Book("../FidelityHoldingsProject.xlsm")

# Finds active sheet in workbook
print("Pulling up Main Sheet...")
Worksheet = Workbook.sheets['MainSheet']

# Refresh data
print("Refreshing data...")
Workbook.api.RefreshAll()

################ NEW ACTIVITY ###################

# Find blank row
print("Finding blank row: ")
# variable holds list of values in the E column of the main worksheet
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
    print("__________Activity #" + str(n + 1) + "__________")
    trade_date = input("Trade Date (if any) mm/dd/yy: ")
    settlement_date = input("* Settlement Date mm/dd/yy: ")
    description = input("* Enter activity description (e.g., You Sold Transaction Profit: $3.25): ")
    quantity = input("* Enter Quantity (negative for sold): ")
    price = input("Enter price: ")
    cost = input("Enter cost (if any): ")
    transaction_cost = input("Enter transaction cost (if any): ")
    amount = input("* Enter amount (negative for buy): ")
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
        information_input = input("* Fund/Stock Code: ")
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
    

################ UPDATE TABLE #########################
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
    latest_row_test = "3"
else:
    latest_price = test_B2
    latest_row_test = "2"
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

# Update lines Y, AB, AE, AH, AK, AN


# BRING RECENT STOCK DATA INTO WORKSHEET

# copy and paste the column of dates
new_dates_test_column = "A" + latest_row_test + ":A" + str(old_date_new_row)
# ndim=2 ensures copied column is pasted as column and not row:
# https://github.com/xlwings/xlwings/issues/398#:~:text=Note%20that%20currently%2C%201d%20arrays%20still%20require%20ndim%3D2%20to%20preserve%20the%20column%20orientation
new_dates_test = Worksheet_Test.range(new_dates_test_column).options(ndim=2).value
new_dates_column = "V2:V" + str(old_date_new_row)
print("Pasting new dates from 'Final (2)' cells " + new_dates_test_column + " in 'MainSheet' cells " + new_dates_column)
Worksheet.range(new_dates_column).value = new_dates_test

# copy and paste the column of FNCMX prices
new_fncmx_test_column = "B" + latest_row_test + ":B" + str(old_date_new_row)
new_fncmx_test = Worksheet_Test.range(new_fncmx_test_column).options(ndim=2).value
new_fncmx_column = "W2:W" + str(old_date_new_row)
print("Pasting new FNCMX prices from 'Final (2)' cells " + new_fncmx_test_column + " in 'MainSheet' cells " + new_fncmx_column)
Worksheet.range(new_fncmx_column).value = new_fncmx_test

# copy and paste the column of FBGRX prices
new_fbgrx_test_column = "C" + latest_row_test + ":C" + str(old_date_new_row)
new_fbgrx_test = Worksheet_Test.range(new_fbgrx_test_column).options(ndim=2).value
new_fbgrx_column = "Z2:Z" + str(old_date_new_row)
print("Pasting new FBGRX prices from 'Final (2)' cells " + new_fbgrx_test_column + " in 'MainSheet' cells " + new_fbgrx_column)
Worksheet.range(new_fbgrx_column).value = new_fbgrx_test

# copy and paste the column of FOCPX prices
new_focpx_test_column = "D" + latest_row_test + ":D" + str(old_date_new_row)
new_focpx_test = Worksheet_Test.range(new_focpx_test_column).options(ndim=2).value
new_focpx_column = "AC2:AC" + str(old_date_new_row)
print("Pasting new FOCPX prices from 'Final (2)' cells " + new_focpx_test_column + " in 'MainSheet' cells " + new_focpx_column)
Worksheet.range(new_focpx_column).value = new_focpx_test

# copy and paste the column of FNILX prices
new_fnilx_test_column = "E" + latest_row_test + ":E" + str(old_date_new_row)
new_fnilx_test = Worksheet_Test.range(new_fnilx_test_column).options(ndim=2).value
new_fnilx_column = "AF2:AF" + str(old_date_new_row)
print("Pasting new FNILX prices from 'Final (2)' cells " + new_fnilx_test_column + " in 'MainSheet' cells " + new_fnilx_column)
Worksheet.range(new_fnilx_column).value = new_fnilx_test


# UPDATE TABLE ROWS BASED ON ACTIVIIES

# Update Shares of Mutual Funds/Stocks, SPAXX total, and Investment Increase


# # Update FNCMX shares column to fit latest activities
# latest_shares_fncmx = "=SUM($F$2:$F$" + str(latest_activity_row) + "*($C$2:$C$" + str(latest_activity_row) + "=$Q$2))"
# print("Updating new FNCMX shares from Activity List in cells X2:X" + str(old_date_new_row - 1))
# for xc in range(old_date_new_row - 2):
#     latest_shares_fncmx_cell = "X" + str(xc + 2)
#     Worksheet.range(latest_shares_fncmx_cell).formula = latest_shares_fncmx

# # Update FBGRX shares column to fit latest activities
# latest_shares_fbgrx = "=SUM($F$2:$F$" + str(latest_activity_row) + "*($C$2:$C$" + str(latest_activity_row) + "=$Q$3))"
# print("Updating new FBGRX shares from Activity List in cells AA2:AA" + str(old_date_new_row - 1))
# for aac in range(old_date_new_row - 2):
#     latest_shares_fbgrx_cell = "AA" + str(aac + 2)
#     Worksheet.range(latest_shares_fbgrx_cell).formula = latest_shares_fbgrx

# # Update FOCPX shares column to fit latest activities
# latest_shares_focpx = "=SUM($F$2:$F$" + str(latest_activity_row) + "*($C$2:$C$" + str(latest_activity_row) + "=$Q$4))"
# print("Updating new FOCPX shares from Activity List in cells AD2:AD" + str(old_date_new_row - 1))
# for adc in range(old_date_new_row - 2):
#     latest_shares_focpx_cell = "AD" + str(adc + 2)
#     Worksheet.range(latest_shares_focpx_cell).formula = latest_shares_focpx

# # Update FNILX shares column to fit latest activities
# latest_shares_fnilx = "=SUM($F$2:$F$" + str(latest_activity_row) + "*($C$2:$C$" + str(latest_activity_row) + "=$Q$5))"
# print("Updating new FNILX shares from Activity List in cells AG2:AG" + str(old_date_new_row - 1))
# for agc in range(old_date_new_row - 2):
#     latest_shares_fnilx_cell = "AG" + str(agc + 2)
#     Worksheet.range(latest_shares_fnilx_cell).formula = latest_shares_fnilx

# # Update FLCEX shares column to fit latest activities
# latest_shares_flcex = "=SUM($F$2:$F$" + str(latest_activity_row) + "*($C$2:$C$" + str(latest_activity_row) + "=$Q$6))"
# print("Updating new FLCEX shares from Activity List in cells AJ2:AJ" + str(old_date_new_row - 1))
# for ajc in range(old_date_new_row - 2):
#     latest_shares_flcex_cell = "AJ" + str(ajc + 2)
#     Worksheet.range(latest_shares_flcex_cell).formula = latest_shares_flcex

# # Update FFGCX shares column to fit latest activities
# latest_shares_ffgcx = "=SUM($F$2:$F$" + str(latest_activity_row) + "*($C$2:$C$" + str(latest_activity_row) + "=$Q$8))"
# print("Updating new FFGCX shares from Activity List in cells AM2:AM" + str(old_date_new_row - 1))
# for amc in range(old_date_new_row - 2):
#     latest_shares_ffgcx_cell = "AM" + str(amc + 2)
#     Worksheet.range(latest_shares_ffgcx_cell).formula = latest_shares_ffgcx

# # Update SPAXX value column to fit latest activities
# latest_spaxx_total = "=SUM(J2:J" + str(latest_activity_row) + ")"
# print("Updating new SPAXX value from Activity List in cells AO2:AO" + str(old_date_new_row - 1))
# for aoc in range(old_date_new_row - 2):
#     latest_spaxx_total_cell = "AO" + str(aoc + 2)
#     Worksheet.range(latest_spaxx_total_cell).formula = latest_spaxx_total

# # Update Investment Increase column to fit latest activities
# latest_investment_increase = "=OFFSET([@[Investment Increase]],0,-1)-SUM($J$2:$J$" + str(latest_activity_row) + "*($E$2:$E$" + str(latest_activity_row) + "=$Q$31))"
# print("Updating investment increase from Activity List in cells AQ2:AQ" + str(old_date_new_row - 1))
# for aqc in range(old_date_new_row - 2):
#     latest_investment_increase_cell = "AQ" + str(aqc + 2)
#     Worksheet.range(latest_investment_increase_cell).formula = latest_investment_increase


    
# function updates table share, SPAXX, and Investment Increase formulas to be equal to formula below it
def table_activity_update_by_row(row_to_edit):
    # Update FNCMX shares column to fit latest activities
    table_row_below_formula = Worksheet.range("X" + str(row_to_edit + 1)).formula
    Worksheet.range("X" + str(row_to_edit)).formula = table_row_below_formula
    # Update FBGRX shares column to fit latest activities
    table_row_below_formula = Worksheet.range("AA" + str(row_to_edit + 1)).formula
    Worksheet.range("AA" + str(row_to_edit)).formula = table_row_below_formula
    # Update FOCPX shares column to fit latest activities
    table_row_below_formula = Worksheet.range("AD" + str(row_to_edit + 1)).formula
    Worksheet.range("AD" + str(row_to_edit)).formula = table_row_below_formula
    # Update FNILX shares column to fit latest activities
    table_row_below_formula = Worksheet.range("AG" + str(row_to_edit + 1)).formula
    Worksheet.range("AG" + str(row_to_edit)).formula = table_row_below_formula
    # Update FLCEX shares column to fit latest activities
    table_row_below_formula = Worksheet.range("AJ" + str(row_to_edit + 1)).formula
    Worksheet.range("AJ" + str(row_to_edit)).formula = table_row_below_formula
    # Update FFGCX shares column to fit latest activities
    table_row_below_formula = Worksheet.range("AM" + str(row_to_edit + 1)).formula
    Worksheet.range("AM" + str(row_to_edit)).formula = table_row_below_formula
    # Update SPAXX values column to fit latest activities
    table_row_below_formula = Worksheet.range("AO" + str(row_to_edit + 1)).formula
    Worksheet.range("AO" + str(row_to_edit)).formula = table_row_below_formula
    # Update Investment Increase column to fit latest activities
    table_row_below_formula = Worksheet.range("AQ" + str(row_to_edit + 1)).formula
    Worksheet.range("AQ" + str(row_to_edit)).formula = table_row_below_formula

# Determine which rows to update the table based on new activities
latest_activity_row = blank_row + num_add - 1
for a in range(num_add):
    print(a)
    row = blank_row + a
    print("Settlement date of " + str(row) + ":")
    new_settlement_date = Worksheet.range("B" + str(row)).value
    print(new_settlement_date)
    v = Worksheet.range("V2:V" + str(old_date_new_row + 40)).value
    # reverse() reverses the order of a list
    v.reverse()
    for d in range(len(v)):
        table_date = v[d]
        print(table_date)
        print("Table row to edit:")
        table_row_to_edit = old_date_new_row + 40 - d
        print(table_row_to_edit)
        if table_date == new_settlement_date:
            global iterate_set_point
            iterate_set_point = d
            in_table_activity_row = str(latest_activity_row - (num_add - (a + 1)))

            # Update FNCMX shares column to fit latest activities
            latest_shares_fncmx = "=SUM($F$2:$F$" + in_table_activity_row + "*($C$2:$C$" + in_table_activity_row + "=$Q$2))"
            print("Updating new FNCMX shares from Activity List in cell X" + str(table_row_to_edit))
            Worksheet.range("X" + str(table_row_to_edit)).formula = latest_shares_fncmx

            # Update FBGRX shares column to fit latest activities
            latest_shares_fbgrx = "=SUM($F$2:$F$" + in_table_activity_row + "*($C$2:$C$" + in_table_activity_row + "=$Q$3))"
            print("Updating new FBGRX shares from Activity List in cell AA" + str(table_row_to_edit))
            Worksheet.range("AA" + str(table_row_to_edit)).formula = latest_shares_fbgrx

            # Update FOCPX shares column to fit latest activities
            latest_shares_focpx = "=SUM($F$2:$F$" +in_table_activity_row + "*($C$2:$C$" +in_table_activity_row + "=$Q$4))"
            print("Updating new FOCPX shares from Activity List in cell AD" + str(table_row_to_edit))
            Worksheet.range("AD" + str(table_row_to_edit)).formula = latest_shares_focpx

            # Update FNILX shares column to fit latest activities
            latest_shares_fnilx = "=SUM($F$2:$F$" +in_table_activity_row + "*($C$2:$C$" +in_table_activity_row + "=$Q$5))"
            print("Updating new FNILX shares from Activity List in cells AG" + str(table_row_to_edit))
            Worksheet.range("AG" + str(table_row_to_edit)).formula = latest_shares_fnilx

            # Update FLCEX shares column to fit latest activities
            latest_shares_flcex = "=SUM($F$2:$F$" +in_table_activity_row + "*($C$2:$C$" +in_table_activity_row + "=$Q$6))"
            print("Updating new FLCEX shares from Activity List in cells AJ" + str(table_row_to_edit))
            Worksheet.range("AJ" + str(table_row_to_edit)).formula = latest_shares_flcex

            # Update FFGCX shares column to fit latest activities
            latest_shares_ffgcx = "=SUM($F$2:$F$" +in_table_activity_row + "*($C$2:$C$" +in_table_activity_row + "=$Q$8))"
            print("Updating new FFGCX shares from Activity List in cells AM" + str(table_row_to_edit))
            Worksheet.range("AM" + str(table_row_to_edit)).formula = latest_shares_ffgcx

            # Update SPAXX value column to fit latest activities
            latest_spaxx_total = "=SUM(J2:J" + in_table_activity_row + ")"
            print("Updating new SPAXX value from Activity List in cells AO" + str(table_row_to_edit))
            Worksheet.range("AO" + str(table_row_to_edit)).formula = latest_spaxx_total

            # Update Investment Increase column to fit latest activities
            latest_investment_increase = "=OFFSET([@[Investment Increase]],0,-1)-SUM($J$2:$J$" + in_table_activity_row + "*($E$2:$E$" + in_table_activity_row + "=$Q$31))"
            print("Updating investment increase from Activity List in cells AQ" + str(table_row_to_edit))
            Worksheet.range("AQ" + str(table_row_to_edit)).formula = latest_investment_increase

            print("Yay!!")
        elif table_date != new_settlement_date:
            # uses try because iterate_set_point may not be assigned a value yet
            try:
                if d <= iterate_set_point:
                    pass
                elif d > iterate_set_point:
                    table_activity_update_by_row(table_row_to_edit)
            except:
                if d <= 40:
                    pass
                elif d > 40:
                    table_activity_update_by_row(table_row_to_edit)
            


# use macro to delete extra at symbol that has been created from copying & pasting formulas
print("Running macro to delete extra '@' symbol...")
DeleteExtraAtSymbolMacro = Workbook.macro("DeleteExtraAtSymbol")
DeleteExtraAtSymbolMacro()