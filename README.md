# Excel Drag Files from Different Folders

This program will develop an user-friendly excel interface for people who are not familiar with Python. People can make dual choices in excel or even update new changes in the excel without updating the Python code. Its dragging file running time is also less than using macro.

The main function is to copy and paste files with different file names and folders in to related excel sheets (with formulas drag down). Given the background of monthly data consolidation, the users want to drag files with specific date


In details, the program will drag files with specific date determined from the folders with adjusted trading date. For example, the specific date is August 16th and the adjusted trading dates is $n$, the program will copy the files with date that is $n$ business date before the specific date. However, if users determine that the 

<!--- ![IMG_0617](https://github.com/zoehahaha/ExcelDragFiles/assets/133292874/3597349c-1b4d-4474-b902-39567b9dc647) -->

In this specific case, people can modify
* load this file or not
* The adjusted trading date $T-n$
* File name of
* Is dragging first of month of the file?
  

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

