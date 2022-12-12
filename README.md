# Intro To Computer Programming Final Project Fall 2022

# Excel Stock Project

Creating a robust, concise, simple way to manage updated stock and mutual fund activity

## Description

A python file that asks for recent stock/mutual fund activity, and updates an Excel file to display current information about the account 

## Getting Started

### Dependencies

* Requires Python installation
* Requires xlwings library and pywin32 library
* Requires time & datetime libraries (generally pre-installed with Python installation)
* Requires Excel files
    1. Excel file with up-to-date stock/mutual fund information
    2. Excel file with bare bones of account information

### Installing

1. Install xlwings & pywin32 libraries
    ```
    pip install xlwings
    ```
    ```
    pip install pywin32
    ```
2. Download main.py, StockInfo.py, FidelityHoldingsProject.xlsm, and StockAndMutualFundInfo.xlsx files to the same folder
3. Use FidelityHoldingsProject.xlsm and StockAndMutualFundInfo.xlsx as templates to adjust to your own criteria

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
* [Reverse a list](https://www.geeksforgeeks.org/python-reversing-list/#:~:text=Using%20reversed()%20we%20can,to%20reverse%20list%20in%2Dplace.)
* [Autofill columns](https://stackoverflow.com/questions/41977016/xlwings-using-api-autofill-how-to-pass-a-range-as-argument-for-the-range-autofill)
* [Using a relative file path with win32](https://stackoverflow.com/questions/45183713/open-excel-file-to-run-macro-from-relative-file-path-in-python)
* [Find last row with data in a column](https://www.dataquest.io/blog/python-excel-xlwings-tutorial/#:~:text=It%20will%20be%20useful%20to%20be%20able%20to%20tell%20where%20our%20table%20ends.)
* [Insert rows](https://github.com/xlwings/xlwings/issues/1284)