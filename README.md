# Intro To Computer Programming Final Project Fall 2022

# Excel Stock Project

Creating a robust, concise, simple way to manage updated stock and mutual fund activity

## Description

A python file that asks for recent stock/mutual fund activity, and updates an Excel file to display current information about the account 

## Getting Started

### Dependencies

* Requires Python installation
* Requires xlwings library (and pywin32 library)
* Requires time & datetime libraries (generally pre-installed with Python installation)
* Requires Excel

### Installing

1. Install xlwings library (& pywin32 library; installing xlwings will also install pywin32)
    ```
    pip install xlwings
    ```
2. Download main.py, StockInfo.py, FidelityHoldingsProject.xlsm, and StockAndMutualFundInfo.xlsx files to the same folder
3. Create macro titled "DeleteExtraAtSymbol" with this code:
    ```
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
    ```
4. Use FidelityHoldingsProject.xlsm and StockAndMutualFundInfo.xlsx as templates to adjust to your own criteria

### Before First Run of Program
1. Open StockAndMutualFundInfo.xlsx
2. Navigate to the tab titled "Data" on the ribbon
3. Click "Refresh All" under the group "Queries & Connections"
4. Follow prompts to allow connections to yfinance
    * Click "Connect"
    * Click "Private" or desired privacy setting under the drop-down menu
5. Execute the program (see below)

### Executing program

* Ensure internet connection for up-to-date information (not required)
* Run main.py
* Enter number of activities when prompted (see example below)
    ```
    How many activity additions would you like to input? 3
    ```
    * If input more than 0 new activities, follow prompts (see examples below)
        ```
        Trade Date (if any) mm/dd/yy: 05/02/22
        * Settlement Date mm/dd/yy: : 05/03/22
        * Enter activity description (e.g., You Sold Transaction Profit: $3.25): You Sold Transaction Profit: $33.11 Transaction Loss: $3.94
        Enter Quantity (negative for sold): -42.211
        Enter price: 18.72
        Enter cost (if any): 575.38
        Enter transaction cost (if any): 
        * Enter amount (negative for buy): 790.19
        Enter reference number (if any): 10482-GHSW9T
        Enter order number: 20481-J2IAWB
        * Fund/Stock Code: FNILX
        ```

## Help

* If program delays after printing "Refreshing data..." there may be no internet connection. It will continue to run after a short time.
* If program still delays after printing "Refreshing data..." Excel may be unresponsive. If you've already followed the steps under "Before First Run of Program Above," try closing the excel file, killing the active python terminal, and running main.py again.
* If program prints "Check for possible errors", there may not be up-to-date stock/mutual fund information

## Authors

Contributors names and contact info

JP Mitiguy

jpm.mitiguy01@gmail.com

<!--## Version History

Coming soon!
* 0.2
    * Various bug fixes and optimizations
    * See [commit change]() or See [release history]()
* 0.1
    * Initial Release -->

## License

This project is licensed under the MIT License - see the LICENSE file for current details

## Acknowledgments

Inspiration, code snippets, etc.
* [w3Schools](https://www.w3schools.com/python/default.asp)
* [Automate The Boring Stuff](https://automatetheboringstuff.com/)
* [xlwings Documentation](https://docs.xlwings.org/en/latest/api.html)
* [Geeks for Geeks Guide to xlwings](https://www.geeksforgeeks.org/working-with-excel-files-in-python-using-xlwings/)
* [xlwings 1-D array solution](https://github.com/xlwings/xlwings/issues/398#:~:text=Note%20that%20currently%2C%201d%20arrays%20still%20require%20ndim%3D2%20to%20preserve%20the%20column%20orientation)
* [Reverse a list](https://www.geeksforgeeks.org/python-reversing-list/#:~:text=Using%20reversed()%20we%20can,to%20reverse%20list%20in%2Dplace.)
* [Autofill columns](https://stackoverflow.com/questions/41977016/xlwings-using-api-autofill-how-to-pass-a-range-as-argument-for-the-range-autofill)
* [Using a relative file path with pywin32](https://stackoverflow.com/questions/45183713/open-excel-file-to-run-macro-from-relative-file-path-in-python)
* [Find last row with data in a column](https://www.dataquest.io/blog/python-excel-xlwings-tutorial/#:~:text=It%20will%20be%20useful%20to%20be%20able%20to%20tell%20where%20our%20table%20ends.)
* [Insert rows](https://github.com/xlwings/xlwings/issues/1284)