################################################################################################################################
# the big excel file name
bigfile_name ="Excel File Path.xlsx"
holiday_2023 = ["2023-01-01","2023-02-20","2023-04-07","2023-05-22","2023-07-01","2023-09-04","2023-10-09","2023-12-25","2023-12-26","2023-11-13"] # manually update

################################################################################################################################


# importing openpyxl module
import win32com.client as win32
from time import strftime, localtime
from datetime import datetime
from datetime import timedelta

# no sunday / no holidays
def add_days(date,days):
    date = str(date)
    d = datetime.strptime(date[:10], "%Y-%m-%d")
    d += timedelta(days)
    weekday_name = d.strftime('%A')
    string = datetime.strftime(d, "%Y-%m-%d")
    while weekday_name == "Sunday" or string in holiday_2023:
        d += timedelta(1)
        string = datetime.strftime(d, "%Y-%m-%d")
        weekday_name = d.strftime('%A')
    return datetime.strftime(d, "%Y-%m-%d")

#get the first day of month, skipping for sunday and holidays
def BD1(date):
    d = datetime.strptime(date[:10], "%Y-%m-%d")
    d =  d.replace(day=1)
    weekday_name = d.strftime('%A')
    string = datetime.strftime(d, "%Y-%m-%d")
    while weekday_name == "Sunday" or string in holiday_2023:
        d += timedelta(1)
        string = datetime.strftime(d, "%Y-%m-%d")
        weekday_name = d.strftime('%A')
    return datetime.strftime(d, "%Y-%m-%d")

# convert 2023-08-16 to 20230816
def no_dash(date):
    date = date[:10]
    text = date[:4]+date[5:7]+date[-2:]
    return text

# numbers to column letter
def n2a(n):
    d, m = divmod(n,26) # 26 is the number of ASCII letters
    return '' if n < 0 else n2a(d-1)+chr(m+65) # chr(65) = 'A'

# get folder name like "2023\\10_Aug_2023"
def folder_date(date):
    d = datetime.strptime(date[:10], "%Y-%m-%d")
    return str(d.year) + "\\\\" + str(d.month + 2) +"_" + d.strftime("%b") + "_" +  str(d.year)


# can copy all filled worksheet
# wb0 is our big file
def paste_autofill(wb0,cop_filename,des_wsname,cop_wsname):
    
    # destination file
    excel.DisplayAlerts = False
    ws0 = wb0.Worksheets(des_wsname)

    # copied file
    wb2 = excel.Workbooks.Open(cop_filename)
    ws2 = wb2.Worksheets(cop_wsname)
    LastRow2 = ws2.UsedRange.Rows.Count # last row
    LastCol = ws2.UsedRange.Columns.Count # lasr column in number format

    copy_range = "A1:" + n2a(LastCol-1) + str(LastRow2)
    paste_range = "B1:" + n2a(LastCol) + str(LastRow2+1)
    
    # static info pasted
    ws2.Range(copy_range).Copy(ws0.Range(paste_range))

    # formula drag down
    if LastRow2 > 2:
        ws0.Range("A2").AutoFill(ws0.Range("A2:A" + str(LastRow2)), win32.constants.xlFillDefault)

    wb2.Close()
    wb0.Save()

# only copy
def paste_autofill1(wb0,cop_filename,des_wsname,cop_wsname):
    
    # destination file
    excel.DisplayAlerts = False
    ws0 = wb0.Worksheets(des_wsname)

    # copied file
    wb2 = excel.Workbooks.Open(cop_filename)
    ws2 = wb2.Worksheets(cop_wsname)
    LastRow2 = ws2.UsedRange.Rows.Count # last row
    LastCol = ws2.UsedRange.Columns.Count # lasr column in number format

    copy_range = "A2:" + n2a(LastCol-1) + str(LastRow2)
    paste_range = "A2:" + n2a(LastCol) + str(LastRow2)
    
    # static info pasted
    ws2.Range(copy_range).Copy(ws0.Range(paste_range))

    wb2.Close()
    wb0.Save()
    
# copy + different formula drag down
def paste_autofill2(wb0,cop_filename,des_wsname,cop_wsname):
    
    # destination file
    excel.DisplayAlerts = False
    ws0 = wb0.Worksheets(des_wsname)

    # copied file
    wb2 = excel.Workbooks.Open(cop_filename)
    ws2 = wb2.Worksheets(cop_wsname)
    LastRow2 = ws2.UsedRange.Rows.Count # last row
    LastCol = ws2.UsedRange.Columns.Count # lasr column in number format

    copy_range = "A2:" + n2a(LastCol-1) + str(LastRow2)
    paste_range = "B2:" + n2a(LastCol) + str(LastRow2+1)
    
    # static info pasted
    ws2.Range(copy_range).Copy(ws0.Range(paste_range))
    
    if LastRow2 > 2:
        ws0.Range("L2:M2").AutoFill(ws0.Range("L2:M" + str(LastRow2)), win32.constants.xlFillDefault)
        ws0.Range("A2").AutoFill(ws0.Range("A2:A" + str(LastRow2)), win32.constants.xlFillDefault)

    wb2.Close()
    wb0.Save()


excel = win32.gencache.EnsureDispatch('Excel.Application')

folder_prefix = "\\\\SIWBPVSC03MRG\\Processed\\"
bigfile_wb = excel.Workbooks.Open(bigfile_name)

# read which file want to be copied
Load_ws = bigfile_wb.Worksheets("Load Function")

print("The securitization date is "+ str(Load_ws.Cells(1,2).Value)) # print the date for reference
sec_date = str(Load_ws.Cells(1,2).Value)[:10]

for i in range(5, Load_ws.UsedRange.Rows.Count+1):
    if Load_ws.Cells(i,2).Value == 1:
        print("Loading " + Load_ws.Cells(i,1).Value)
        
        # for the red tabs
        if Load_ws.Cells(i,1).Value == "Pre Securitization PV Report":
            filename = no_dash(sec_date) + " PLSummary Pre.csv"
            folder_name = "\\\\cibg-srv-tor08\\cibg_dpss-grps\\GROUPS\\CMB Swaptions\\Daily PnL\\Fiscal " + folder_date(sec_date) + "\\Working\\\PV Report\\"
            ws_name =  no_dash(sec_date) + " PLSummary Pre"
            file_path = folder_name + filename
            paste_autofill2(bigfile_wb, file_path,Load_ws.Cells(i,1).Value,ws_name)
            
        elif Load_ws.Cells(i,1).Value == "Post Securitization PV Report":
            filename = no_dash(sec_date) + " PLSummary.csv"
            folder_name = "\\\\cibg-srv-tor08\\cibg_dpss-grps\\GROUPS\\CMB Swaptions\\Daily PnL\\Fiscal " + folder_date(sec_date) + "\\Working\\\PV Report\\"
            ws_name =  no_dash(sec_date) + " PLSummary"
            file_path = folder_name + filename
            paste_autofill2(bigfile_wb, file_path,Load_ws.Cells(i,1).Value,ws_name)
            
        elif Load_ws.Cells(i,1).Value == "Inception PnL":
            filename = no_dash(sec_date) + " Inception P&L Detailed Report Pre.csv"
            folder_name = "\\\\cibg-srv-tor08\\cibg_dpss-grps\\GROUPS\\CMB Swaptions\\Daily PnL\\Fiscal " + folder_date(sec_date) + "\\Working\\Inception PnL\\"
            ws_name =  no_dash(sec_date) + " Inception P&L Detailed"
            file_path = folder_name + filename
            paste_autofill1(bigfile_wb, file_path,Load_ws.Cells(i,1).Value,ws_name)
            
        # then for the blue tab and purple tab
        else:
            # is first day of month?
            if Load_ws.Cells(i,5).Value == 1: # yes, use BD1
                sec_date = BD1(sec_date)
            
            # else, use securitization date
            
            # replace date!
            if Load_ws.Cells(i,6).Value != 0:
                folder_name =  folder_prefix + add_days(Load_ws.Cells(i,6).Value,0)
            else:
                added_day = int(Load_ws.Cells(i,3).Value)
                folder_name = folder_prefix + add_days(sec_date,added_day)
            
            #print(Load_ws.Cells(i,1).Value)
            if Load_ws.Cells(i,1).Value == "Nesto Own File - 1st of Month" or Load_ws.Cells(i,1).Value == "Nesto Own File - Sec Date":
                filename = Load_ws.Cells(i,4).Value + sec_date
            else:
                filename = Load_ws.Cells(i,4).Value + no_dash(sec_date)
            print(folder_name + "  " + filename)  # for debug purpose
            
            file_path = folder_name + "\\" + filename + ".csv"
            
            sec_date = str(Load_ws.Cells(1,2).Value)[:10]
            
            # copy & paste
            paste_autofill(bigfile_wb, file_path, Load_ws.Cells(i,1).Value, filename)
            

bigfile_wb.Close()
#excel.Application.Quit()
