# Intro To Computer Programming Final Project Fall 2022

# Excel Stock Project

Creating a robust, concise, simple way to deal with stock and mutual fund activity

## Description

A python file that asks for recent stock/mutual fund activity, and updates an Excel file to display current information about the account 

## Getting Started

### Dependencies

* Requires Python installation. Version 3.10 recommended
* Requires xlwings library
* Requires Excel files. (1) Up-to-date stock/mutual fund information (2) Bare bones of account information

### Installing

* Download ForFidelityHoldingsxlsm.py and ForTestxlsxFile.py files to the same folder
* Change line 23 of ForTestxlsxFile.py to call the Excel File with stock/mutual fund prices
```
Workbook = File.Workbooks.Open("[File Path]")
```
* Change line 21 of ForFidelityHoldingsxlsm.py to call the Excel file with account information
```
Workbook = xlwings.Book("[File Path]")
```

### Executing program

* Run ForFidelityHoldingsxlsm.py
* Enter number of activities
```
code blocks for commands
```

## Help

Any advise for common problems or issues.
```
command to run if program contains helper info
```

## Authors

Contributors names and contact info

ex. JP M 
ex. email@email.org

## Version History

* 0.2
    * Various bug fixes and optimizations
    * See [commit change]() or See [release history]()
* 0.1
    * Initial Release

## License

This project is licensed under the [NAME HERE] License - see the LICENSE file for details

## Acknowledgments

Inspiration, code snippets, etc.
* [w3Schools](https://www.w3schools.com/python/default.asp)
* [Automate The Boring Stuff](https://automatetheboringstuff.com/)
* [xlwings Documentation](https://docs.xlwings.org/en/latest/api.html)
* [Geeks for Geeks Guide to xlwings](https://www.geeksforgeeks.org/working-with-excel-files-in-python-using-xlwings/)
* [xlwings 1-D array solution](https://github.com/xlwings/xlwings/issues/398#:~:text=Note%20that%20currently%2C%201d%20arrays%20still%20require%20ndim%3D2%20to%20preserve%20the%20column%20orientation)

