# JP Mitiguy

# Goal
'''
My final project is to create a robust, concise, simple way to manage updated stock and mutual fund activity
'''

# Sources
'''
I used https://docs.xlwings.org/en/latest/api.html to examine the full documentaion of xlwings library
I sued https://www.geeksforgeeks.org/working-with-excel-files-in-python-using-xlwings/ to understand the basics of using Excel with xlwings
I used https://www.w3schools.com/python/python_try_except.asp to learn use try & except 
I used https://www.geeksforgeeks.org/python-reversing-list/#:~:text=Using%20reversed()%20we%20can,to%20reverse%20list%20in%2Dplace. to reverse lists
I used https://stackoverflow.com/questions/41977016/xlwings-using-api-autofill-how-to-pass-a-range-as-argument-for-the-range-autofil to determine how to autofill columns
I used https://www.dataquest.io/blog/python-excel-xlwings-tutorial/#:~:text=It%20will%20be%20useful%20to%20be%20able%20to%20tell%20where%20our%20table%20ends. to find a quick way to find the last data in a column
I used https://github.com/xlwings/xlwings/issues/1284 to determine how to insert rows
'''

## Import

# Import download libraries
import xlwings
from xlwings.constants import AutoFillType
# Import built-in libraries
from time import sleep
# Import created modules
import StockInfo
from settings import *

## Global Variables/Lists
# see settings.py

# dummy variable to be used in later for loops
iterate_set_point = "lorem ipsum"

## Utility Functions

# function updates table share, SPAXX, and Investment Increase formulas to be equal to formula below it
def table_activity_update_by_row(row_to_edit):

    # Update Stock & Mutual fund shares columns to fit latest activities
    for i in range(len(STOCKS_AND_MUTUAL_FUND_CODES)):
        # 'Worksheet.range("A1").formula' syntax either finds or sets the formula in cell A1 of the Excel Worksheet assigned the variable 'Worksheet'
        table_row_below_formula = Worksheet.range(STOCK_AND_FUND_SHARES_COLUMNS[i] + str(row_to_edit + 1)).formula
        Worksheet.range(STOCK_AND_FUND_SHARES_COLUMNS[i] + str(row_to_edit)).formula = table_row_below_formula

    # Update SPAXX values column to fit latest activities
    table_row_below_formula = Worksheet.range(MONEY_MARKET_SPAXX_COLUMN + str(row_to_edit + 1)).formula
    Worksheet.range(MONEY_MARKET_SPAXX_COLUMN + str(row_to_edit)).formula = table_row_below_formula
    # Update Investment Increase column to fit latest activities
    table_row_below_formula = Worksheet.range(INVESTMENT_INCREASE_COLUMN + str(row_to_edit + 1)).formula
    Worksheet.range(INVESTMENT_INCREASE_COLUMN + str(row_to_edit)).formula = table_row_below_formula

# use macro to delete extra at symbol that has been created from copying & pasting formulas
def delete_at_macro():
    print("Running macro to delete extra '@' symbol...")
    # closes StockAndMutualFundInfo.xlsx to ensure macro runs in correct Workbook
    Workbook_Stock_Info.close()
    DeleteExtraAtSymbolMacro = Workbook.macro("DeleteExtraAtSymbol")
    DeleteExtraAtSymbolMacro()

# create class of Activity Table
class Actv_Table:
    # table has rows and columns
    def __init__(self, columns, rows):
        self.columns = columns
        self.rows = rows

    # create method that adds new activities to the table
    def new_activity(self):
        # Add recent activities
        global num_add
        num_add = int(input("How many activity additions would you like to input? "))
        # add new activities to number of rows in the table
        self.rows += num_add
        print("Rows in activity table" + str(self.rows))
        for n in range(num_add):
            print("__________Activity #" + str(n + 1) + "__________")
            trade_date = input("Trade Date (if any) mm/dd/yy: ")
            settlement_date = input("* Settlement Date mm/dd/yy: ")
            description = input("* Enter activity description (e.g., You Sold Transaction Profit: $3.25): ")
            quantity = input("Enter Quantity (negative for sold): ")
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
                # determines if input matches known stock/mutual fund
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


# Macro Code to delete extra "@" symbol that disrupts excel formulas
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

## Run StockInfo

# From StockInfo, run function that downloads most recent stock prices
print("win32/pywin32_________________________________")
StockInfo.update_file()

# sleep for 3 seconds to ensure smooth transition from openpyxl to xlwings code
sleep(3)

############# LOAD XLWINGS ##################
print("xlwings_____________________________________")

# Load workbooks
print("Loading workbooks...")
Workbook = xlwings.Book("FidelityHoldingsProject.xlsm")
Workbook_Stock_Info = xlwings.Book("StockAndMutualFundInfo.xlsx")

# Refresh data
print("Refreshing data...")
Workbook.api.RefreshAll()

# Pause to ensure time to Refresh sheets
sleep(4)

# Finds Sheets in workbooks
print("Pulling up Sheets...")
Worksheet = Workbook.sheets['MainSheet']
Worksheet_Stock_Info = Workbook_Stock_Info.sheets["Final (2)"]
Worksheet_Test = Workbook.sheets["Final (2)"]

################ NEW ACTIVITY ###################

## Find blank row
print("Finding blank row: ")
# finds first blank row in the Main Worksheet without data in column E
# https://www.dataquest.io/blog/python-excel-xlwings-tutorial/#:~:text=It%20will%20be%20useful%20to%20be%20able%20to%20tell%20where%20our%20table%20ends.
blank_row = Worksheet.range("E2").end('down').row + 1
print(blank_row)

# instantiate Actv_Table class w/14 columns and a certain number of rows
activity_table = Actv_Table(14, blank_row - 1)
# run new activity method
activity_table.new_activity()

################ UPDATE TABLE #########################
# Find the date in spot V2 (last updated date)
print("Old date: ")
old_date = Worksheet.range("V2").value
print(old_date)

# Find value in B2 (latest price) of Final (2) worksheet (may be blank)
print("Value in B2 of Final (2) from StockAndMutualFundInfo.xlsx:")
test_B2 = Worksheet_Stock_Info.range("B2").value
print(test_B2)

# find the latest price to update the main sheet with
print("Latest price: ")
# if/else acommodates if first data row of excel file is blank
if test_B2 == None:
    latest_price = Worksheet_Stock_Info.range("B3").value
    latest_row_test = "3"
else:
    latest_price = test_B2
    latest_row_test = "2"
print(latest_price)

# Determine which row the old date needs to go once the new info is added
print("Old date will now go in this row: ")
# 'a' is a list of all the data values in column A
a = Worksheet_Stock_Info.range("A:A").value
# searches through A column until date found
for i in range(MAX_LENGTH_SINCE_UPDATE):
    cell = a[i]
    if cell == old_date:
        global old_date_new_row
        # Determines if a one cell adjustment needs to be made to account for blank 1st row
        if test_B2 == None:
            old_date_new_row = i
        else:
            old_date_new_row = i + 1
        print(old_date_new_row)
        break

# get first cell of old date
print("Old table will start in this cell: ")
old_date_new_cell = "V" + str(old_date_new_row)
print(old_date_new_cell)

# Copy table into designated rows below
table = Worksheet.range("V2").expand().formula
# insert rows into 2nd to last row of table (to maintain table format)
# https://github.com/xlwings/xlwings/issues/1284
print("Inserting blank rows in rows:")
row_inserts = str(len(table) - 1) + ":" + str(len(table) + old_date_new_row - 4)
print(row_inserts)
Worksheet.range(row_inserts).insert()
# paste the table into its new spot
print("Pasting original table into:")
table_paste_location = old_date_new_cell + ":" + INVESTMENT_INCREASE_COLUMN + str(Worksheet.range(old_date_new_cell).end('down').row + old_date_new_row)
print(table_paste_location)
Worksheet.range(table_paste_location).formula = table

# use macro to delete extra at symbol that has been created
delete_at_macro()
# re-open StockAndMutualFundInfo.xlsx after function above closed
Workbook_Stock_Info = xlwings.Book("StockAndMutualFundInfo.xlsx")
Worksheet_Stock_Info = Workbook_Stock_Info.sheets["Final (2)"]

## UPDATE COLUMNS Y, AB, AE, AH, AK, AN after cut & paste creates unconsistent formulas in these columns

# find last row in table
print("Finding last table row...")
last_table_row = Worksheet.range("Y2").end('down').row
print("Last table row: " + str(last_table_row))

# Update Stock/Mutual Fund Value Columns so formula autofills
for i in range(len(STOCKS_AND_MUTUAL_FUND_CODES)):
    print("Autofilling column " + STOCK_AND_FUND_VALUE_COLUMNS[i])
    cells_to_autofill = STOCK_AND_FUND_VALUE_COLUMNS[i] + "2:" + STOCK_AND_FUND_VALUE_COLUMNS[i] + str(last_table_row)
    # Autofill formulas
    # https://stackoverflow.com/questions/41977016/xlwings-using-api-autofill-how-to-pass-a-range-as-argument-for-the-range-autofil
    Worksheet.range(STOCK_AND_FUND_VALUE_COLUMNS[i] + "2").api.AutoFill(Worksheet.range(cells_to_autofill).api,AutoFillType.xlFillDefault)

# Copy information from 'Final (2)' Worksheet in StockAndMutualFundInfo.xlsx to 'Final (2)' Worksheet in FidelityHoldingsProject.xlsm
stock_info_table = Worksheet_Stock_Info.range("A1").expand().value
print("Number of blank rows inserted:")
print(row_inserts)
Worksheet_Test.range(row_inserts).insert()
# paste the table into its new spot
print("Pasting original table into:")
stock_info_table_range = "A1:" + STOCK_AND_FUND_PRICE_COLUMNS_STOCK_INFO[len(STOCKS_AND_MUTUAL_FUND_CODES) - 1] + str(Worksheet_Test.range("A1").end('down').row)
print(stock_info_table_range)
Worksheet_Test.range(stock_info_table_range).expand().value = stock_info_table
# TO DO___________ Close StockAndMutualFundInfo.xlsx

## BRING RECENT STOCK DATA INTO WORKSHEET

# copy and paste the column of dates from other sheet
new_dates_test_column = "A" + latest_row_test + ":A" + str(old_date_new_row)
# ndim=2 ensures copied column is pasted as column and not row:
# https://github.com/xlwings/xlwings/issues/398#:~:text=Note%20that%20currently%2C%201d%20arrays%20still%20require%20ndim%3D2%20to%20preserve%20the%20column%20orientation
new_dates_test = Worksheet_Test.range(new_dates_test_column).options(ndim=2).value
new_dates_column = "V2:V" + str(old_date_new_row)
print("Pasting new dates from 'Final (2)' cells " + new_dates_test_column + " in 'MainSheet' cells " + new_dates_column)
Worksheet.range(new_dates_column).value = new_dates_test

# copy and paste columns of stock & mutual fund prices
for i in range(len(STOCKS_AND_MUTUAL_FUND_CODES)):
    new_stock_mutual_fund_test_column = STOCK_AND_FUND_PRICE_COLUMNS_STOCK_INFO[i] + latest_row_test + ":" + STOCK_AND_FUND_PRICE_COLUMNS_STOCK_INFO[i] + str(old_date_new_row)
    new_stock_mutual_fund_test = Worksheet_Test.range(new_stock_mutual_fund_test_column).options(ndim=2).value
    new_fncmx_column = STOCK_AND_FUND_PRICE_COLUMNS[i] + "2:" + STOCK_AND_FUND_PRICE_COLUMNS[i] + str(old_date_new_row)
    print("Pasting new " + STOCKS_AND_MUTUAL_FUND_CODES[i] + " prices from 'Final (2)' cells " + new_stock_mutual_fund_test_column + " in 'MainSheet' cells " + new_fncmx_column)
    Worksheet.range(new_fncmx_column).value = new_stock_mutual_fund_test

# Update Table Shares of Mutual Funds/Stocks, SPAXX total, and Investment Increase based on new activity additions
latest_activity_row = blank_row + num_add - 1
# runs for the number of times there's a new activiy
for a in range(num_add):
    print("Activity #" + str(a + 1))
    row = blank_row + a
    print("Settlement date of " + str(row) + ":")
    new_settlement_date = Worksheet.range("B" + str(row)).value
    print(new_settlement_date)
    v = Worksheet.range("V2:V" + str(old_date_new_row + ACTIVITY_AGE_LIBERTY)).value
    # reverse() reverses the order of a list so that list's dates go from old to current (bottom to top)
    # https://www.geeksforgeeks.org/python-reversing-list/#:~:text=Using%20reversed()%20we%20can,to%20reverse%20list%20in%2Dplace.
    v.reverse()
    # iterates through all cells in dates column and updates to include latest activites
    for d in range(len(v)):
        table_date = v[d]
        table_row_to_edit = old_date_new_row + ACTIVITY_AGE_LIBERTY - d
        print("Updating row " + str(table_row_to_edit) + " with date " + str(table_date))
        # if activity date matches current row's date, update that row
        if table_date == new_settlement_date:
            # assign 'iterate_set_point' so that future runs in 'for loop' with 'a' doesn't start from bottom of list of v
            iterate_set_point = d
            in_table_activity_row = str(latest_activity_row - (num_add - (a + 1)))

            # Update stock & mutual fund shares columns to fit latest activities
            for i in range(len(STOCKS_AND_MUTUAL_FUND_CODES)):
                latest_shares_stock_mutual_fund = "=SUM($F$2:$F$" + in_table_activity_row + "*($C$2:$C$" + in_table_activity_row + "=$Q$" + STOCK_AND_MUTUAL_FUND_FULL_NAME_ROW[i] + "))"
                print("Updating new " + STOCKS_AND_MUTUAL_FUND_CODES[i] + " shares from Activity List in cell " + STOCK_AND_FUND_SHARES_COLUMNS[i] + str(table_row_to_edit))
                Worksheet.range(STOCK_AND_FUND_SHARES_COLUMNS[i] + str(table_row_to_edit)).formula = latest_shares_stock_mutual_fund

            # Update SPAXX value column to fit latest activities
            latest_spaxx_total = "=SUM(J2:J" + in_table_activity_row + ")"
            print("Updating new SPAXX value from Activity List in cells " + MONEY_MARKET_SPAXX_COLUMN + str(table_row_to_edit))
            Worksheet.range(MONEY_MARKET_SPAXX_COLUMN + str(table_row_to_edit)).formula = latest_spaxx_total

            # Update Investment Increase column to fit latest activities
            latest_investment_increase = "=OFFSET([@[Investment Increase]],0,-1)-SUM($J$2:$J$" + in_table_activity_row + "*($E$2:$E$" + in_table_activity_row + "=$Q$31))"
            print("Updating investment increase from Activity List in cells " + INVESTMENT_INCREASE_COLUMN + str(table_row_to_edit))
            Worksheet.range(INVESTMENT_INCREASE_COLUMN + str(table_row_to_edit)).formula = latest_investment_increase

            print("Yay!!")
        # if there have been no updates since the current 'table date'
        elif table_date != new_settlement_date and iterate_set_point == "lorem ipsum":
            # if current list item is less than/equal to activity age wiggle room, don't edit
            if d <= ACTIVITY_AGE_LIBERTY:
                pass
            # otherwise, copy formula from cell below
            elif d > ACTIVITY_AGE_LIBERTY:
                table_activity_update_by_row(table_row_to_edit)
        # else - there have been edits to the table but the current 'table date' does not equal the latest 'new_settlement_date'
        else:
            # if current list item is less than/equal to last set point, don't edit
            if d <= iterate_set_point:
                pass
            # otherwise, copy formula from cell below
            elif d > iterate_set_point:
                table_activity_update_by_row(table_row_to_edit)

# use macro to delete extra at symbol that has been created from copying & pasting formulas
delete_at_macro()
# re-open StockAndMutualFundInfo.xlsx after function above closed
Workbook_Stock_Info = xlwings.Book("StockAndMutualFundInfo.xlsx")
Worksheet_Stock_Info = Workbook_Stock_Info.sheets["Final (2)"]