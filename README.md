# EasyExcel

[![Build Status](https://travis-ci.org/JoshHumpherey/EasyExcel.svg?branch=master)](https://travis-ci.org/JoshHumpherey/EasyExcel)  [![Coverage Status](https://coveralls.io/repos/github/JoshHumpherey/EasyExcel/badge.svg?branch=master)](https://coveralls.io/github/JoshHumpherey/EasyExcel?branch=master)

## Why Another Python-Excel Library?
For a project I found myself needing to interact with some old excel spreadsheets and the common alternatives such as pandas and openpyxl weren't enough for me. The biggest need this library fills for me is being able to easily append my work to existing spreadsheets. Other libraries would overwrite the whole spreadsheet which doesn't do much good when you are trying to preserve the original data. This library also excels (pun intended) at interacting with non-standard formatting in spreadsheets. Pandas is a great library if you are in control of the data formatting but often that's not possible. All-in-all this project aims to help you get up and running in as little time as possible while maintaining an extremely easy to use interface. Enjoy!

## Getting started
Starting development using EasyExcel takes almost no time! Simply install via the pip command:
```Python
pip install EasyExcel
```
From here you can import it into your source file by calling:
```Python
import EasyExcel #you will call this by using EasyExcel.PythonExcelConverter(args*)
 ```

 ## Documentation
 1. #### PythonExcelConverter(filename, extension, sheet_name, output_filename)
* filename (string): The name of the file you want to read-in without the extension
* extension (string): The file extension (.xlsx)
* sheet_name (string): The name of the sheet you want to pull data from (Sheet1, Sheet2, etc...)
* output_filename: This is the filename that  EasyExcel will write to upon a save. To append data to your spreadsheet simply use the same name as your filename

2. #### query_column(column_number)
 * column_number (int): The number of the column you want to grab data from (A=1, B=2, C=3, etc...)

3. #### write_list_to_column(the_list, col_letter, title, start):
 * the_list (list): This is the list/array that you want to write as a column in the spreadsheet
 * col_letter (string): This is the letter of the column you want to write to
 * title (string): This will be your column header
 * start (int): This is the row number that your header will appear on and your data will appear below

4. #### query_row(row_number):
 * row_number (int): The number of the row you want to get data from

5. #### query_cell(col_letter, row_number):
 * col_letter (string): The letter of the column you want to grab data from
 * row_number (int): The number of the row you want to grab data from

6. #### write_to_cell(col_letter, row_number, val):
 * col_letter (string): The letter of the column you want to write data to
 * row_number (int): The number of the row you want to write data to
 * val (string/int): The data you want written to the cell

7. #### save_file()
 * Requires no arguments. Simply saves the file as ouput_name in your working directory

 ## Contributions
 Want to help contribute? This is just the start for a small project I needed but I would love to expand it. There is a need for better excel libraries in the Python ecosystem and I'd love to help! To contribute, open a pull request with a well-documented change-log and I'll merge it. I would request that if you do, please use Pylint to make sure that this project follows the PEP8 standards for readability and future collaboration.

 ## Possible Additions
 * Query a section of data instead of a whole column
 * Make spreadsheet cell addressing consistent instead of currently using both letters and numbers for column heads
 * add more flexibility for opening and saving spreadsheets
