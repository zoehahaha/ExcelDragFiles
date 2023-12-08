# Excel Drag Files from Different Folders

## Introduction

This program will develop an user-friendly excel interface for people who are not familiar with Python. People can make dual choices in excel or even update new changes in the excel without updating the Python code. Its dragging file running time is also less than using macro.

The main function is to copy and paste files with different file names and folders in to related excel sheets (with formulas drag down). Given the background of monthly data consolidation, the users want to drag files with specific date and paste to the excle file to do a monthly check. In addition, the dates for each file should be adjusted due to the lagged trading days (either $T-1$ or $T-2$)

## Functions

In details, the program will drag files with specific date determined from the folders with adjusted trading date. For example, the specific date is August 16th and the adjusted trading dates is $n$, the program will copy the files with date that is $n$ business date before the specific date. However, if users determine that the correct date for this file would be Aug 3rd, they can fill in that in the column of "replacement date". 

<!--- ![IMG_0617](https://github.com/zoehahaha/ExcelDragFiles/assets/133292874/3597349c-1b4d-4474-b902-39567b9dc647) -->

In this specific case, people can modify
* load this file or not: choose 1 or 0
* The adjusted trading date $T-n$
* File name prefix
* Is dragging first of month of the file?: choose 1 or 0

## Installation

Before running the program, make sure to install the following packages:
1. win32com.client
   ```
   pip install pywin32
   ```
3. datetime
   ```
   pip install datetime
   ```

When installing packages with workplace PCs may encounter "Coundn't find...", can use the following line to solve
```
pip install --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org #Package_Name
```

## Before Run

Then, change the file path for the excel that you are copying to, remember to double all the "\" in the path. Otherwise, Python will fail to recogize the the sign. For example, 
   ```
   wrong_one = "\\c:\python_automation\test.xlsx"
   correct_one = "\\\\c:\\python_automation\\test.xlsx"
   ```
## Maintainence

### Holidays Excluding
Users need to maintain the a list of string of the holidays. Users can consider this list as a collections of gaps that they would like to exclude. For example, even if Oct $10^{th}$ is not a holiday, since the team don't have data for that day. Team can include "2023-10-10" in the string to exclude it.
```
   holiday_2023 = ["2023-01-01", ..]
```
### More files need reading

