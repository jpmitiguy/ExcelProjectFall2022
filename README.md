# Intro To Computer Programming Final Project Fall 2022

# Excel Stock Project

Creating a robust, concise, simple way to deal with stock and mutual fund activity

## Description

A python file that asks for recent stock/mutual fund activity, and updates an Excel file to display current information about the account 

## Getting Started

### Dependencies

* Requires Python installation. Version 3.10 recommended
* Requires xlwings library
* Requires time & datetime libraries (generally pre-installed with Python installation)
* Requires Excel files
    1. Excel file with up-to-date stock/mutual fund information
    2. Excel file with bare bones of account information

### Installing

1. Install xlwings
2. Download ForFidelityHoldingsxlsm.py and ForTestxlsxFile.py files to the same folder
3. Change line 23 of ForTestxlsxFile.py to call the Excel File with stock/mutual fund prices (see example below)
    ```
    Workbook = File.Workbooks.Open("C:/Users/Joe/Investments/AccountInfo.xlsm")
    ```
4. Change line 21 of ForFidelityHoldingsxlsm.py to call the Excel file with account information (see example below)
    ```
    Workbook = xlwings.Book("C:/Users/Joe/Investments/StockAndMutualFundPrices.xlsm")
    ```
5. Create Excel macro in file with account information (copy and paste code below)
    ```
        
    Sub DeleteExtraAtSymbol()
    '
    ' DeleteExtraAtSymbol Macro
    ' When copying and pasting using xlwings and python, it inputs an "@" symbol into some formulas. This macro is designed to remove these..
    '

    '
        Columns("V:AQ").Select
        Selection.Replace What:="@$", Replacement:="$", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    End Sub

    ```
    

### Executing program

* Ensure internet connection for up-to-date information (not required)
* Run ForFidelityHoldingsxlsm.py
* Enter number of activities when prompted (see example below)
    ```
    How many activity additions would you like to input? 3
    ```
    * If input more than 0 new activities, follow prompts (see examples below)
        ```
        Trade Date (if any) mm/dd/yy: 05/02/22
        * Settlement Date mm/dd/yy: : 05/03/22
        Enter activity description (e.g., You Sold Transaction Profit: $3.25): You Sold Transaction Profit: $33.11 Transaction Loss: $3.94
        * Enter Quantity (negative for sold): -42.211
        Enter price: 18.72
        Enter cost (if any): 575.38
        Enter transaction cost (if any): 
        * Enter amount (negative for buy): 790.19
        Enter reference number (if any): 10482-GHSW9T
        Enter order number: 20481-J2IAWB
        Fund/Stock Code: FNILX
        ```

## Help

* Program will delay after printing "Refreshing data..." if there's no internet connection. It will continue to run after a short time.
* If program prints "Check for possible errors", there may not be up-to-date stock/mutual fund information

## Authors

Contributors names and contact info

JP Mitiguy

jpm.mitiguy01@gmail.com

## Version History

Coming soon!
<!-- * 0.2
    * Various bug fixes and optimizations
    * See [commit change]() or See [release history]()
* 0.1
    * Initial Release -->

## License

This project is soon planned to be licensed under the MIT License - see the LICENSE file for current details

## Acknowledgments

Inspiration, code snippets, etc.
* [w3Schools](https://www.w3schools.com/python/default.asp)
* [Automate The Boring Stuff](https://automatetheboringstuff.com/)
* [xlwings Documentation](https://docs.xlwings.org/en/latest/api.html)
* [Geeks for Geeks Guide to xlwings](https://www.geeksforgeeks.org/working-with-excel-files-in-python-using-xlwings/)
* [xlwings 1-D array solution](https://github.com/xlwings/xlwings/issues/398#:~:text=Note%20that%20currently%2C%201d%20arrays%20still%20require%20ndim%3D2%20to%20preserve%20the%20column%20orientation)

