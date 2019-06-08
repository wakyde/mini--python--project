# encoding:utf-8

import xlwings as xw
import xlrd
import datetime
import pytz





# 搜索----------------------------------------------------------
def search(sheet,sht,empname,sheet_name):
    nrows = sht.nrows
    rng_name = "B1:E" +  str(nrows)
    rng = sheet.range(rng_name)
    try:
        row_num = rng.api.Find(empname,SearchOrder=1).row
        row_num1 = rng.api.FindPrevious().row
    except AttributeError:
            print("%s表中没有这个人的名字" % sheet_name)
            exit()

    if row_num != row_num1:
        print("请输员工工号:")
        person_id = input().strip()
        try:
            row_num2 = rng.api.Find(person_id,SearchOrder=2).row
        except AttributeError:
            print("%s表中没有这个人的工号" % sheet_name)
            exit()
        return row_num2
    else:
        return row_num
       
    
# 添加离职信息---------------------------------------------------------        
def departure_info(sheet,sht,row_num,sheet_name):
    # 判断该人有没有离职
    
    today = datetime.datetime.now()
    formatted_today = today.strftime('%s/%s/%s' % (today.year,today.month,today.day))
    print("请输入离值日期:")    
    leave_date = input()
              
    print("请输入离值原因:") 
    leave_reson = input().strip()
    
    if sheet_name == "Perm":
        col_num = sheet.range("A1").api.EntireRow.Find("离职日期").column
        sheet.api.Cells(row_num,col_num).value = leave_date
        sheet.api.Cells(row_num,col_num).NumberFormat = "yyyy/m/d"
        sheet.api.Cells(row_num,col_num+1).value = leave_reson      
        return col_num


            
    elif sheet_name == "Temp DVC" or sheet_name == "Hourly Worker":
        col_num = sheet.range("A1").api.EntireRow.Find("实际离职日期").column
        sheet.api.Cells(row_num,col_num).value = leave_date
        sheet.api.Cells(row_num,col_num).NumberFormat = "yyyy/m/d"
        sheet.api.Cells(row_num,col_num).Copy()
        sheet.api.Cells(row_num,col_num+1).PasteSpecial()       
        sheet.api.Cells(row_num,col_num+2).value = formatted_today
        sheet.api.Cells(row_num,col_num+2).NumberFormat = "yyyy/m/d"
        sheet.api.Cells(row_num,col_num+3).value = leave_reson             
        return col_num


#转劳务工--------------------------------------------------------
def change_TempDVC(sheet,sht,row_num,sheet_name):
    print("请输入转换日期:")    
    leave_date = input()
    col_num = sheet.range("A1").api.EntireRow.Find("实际离职日期").column
    sheet.api.Cells(row_num,col_num).value = leave_date
    sheet.api.Cells(row_num,col_num).NumberFormat = "yyyy/m/d"
    sheet.api.Cells(row_num,col_num+3).value = "转劳务工"         
    return col_num
 
    
#劳务工转正---------------------------------------------------------
def change_Perm(sheet,sht,row_num,sheet_name):
    print("请输入转正日期:")    
    leave_date = input()
    col_num = sheet.range("A1").api.EntireRow.Find("实际离职日期").column
    sheet.api.Cells(row_num,col_num).value = leave_date
    sheet.api.Cells(row_num,col_num).NumberFormat = "yyyy/m/d"
    sheet.api.Cells(row_num,col_num+3).value = "劳务工转正"         
    return col_num 


# 剪切和插入适当位置------------------------------------------------        
def shear_insert(sheet,sht,row_num,col_num):
    sheet.range("A" + str(row_num)).api.EntireRow.Cut()
    nrows = sht.nrows 
    target_time = sheet.api.Cells(row_num,col_num).value
    if sheet_name == "Perm":
        nrows = nrows - 2
        for nrows_i in range(nrows,0,-1):
            
            if sheet.api.Cells(nrows_i,col_num).value < target_time:
                sheet.api.Rows(nrows_i+1).Insert()
                break
           
    
    elif sheet_name == "Temp DVC" or sheet_name == "Hourly Worker":  
        for nrows_i in range(nrows,0,-1):
            if sheet.api.Cells(nrows_i,col_num).value < target_time:
                sheet.api.Rows(nrows_i+1).Insert()
                break 


# 按照日期将在职改为离职-----------------------------------------------------------
def on_job_change(sheet,sht,sheet_name,row_num,col_num):
    sheet.api.Cells(825, 30).value = "离职"
    nrows = sht.nrows
    today1 = datetime.datetime.now()
    today = today1.replace(tzinfo=pytz.timezone("UTC"))
    formatted_today = today1.strftime('%s/%s/%s' % (today.year,today.month,today.day))
    if sheet_name == "Perm":
        nrows = nrows - 2
        for row_i in range(nrows,1,-1):
            time_str= sheet.api.Cells(row_i, col_num).value
            if type(time_str).__name__ == 'datetime':
                if time_str <= today and sheet.api.Cells(row_i, col_num - 1).value != "离职":
                    sheet.api.Cells(row_i, col_num - 1).value = "离职"
            elif type(time_str).__name__ == 'str':
                if time_str <= formatted_today and sheet.api.Cells(row_i, col_num - 1).value != "离职":
                    sheet.api.Cells(row_i, col_num - 1).value = "离职"
            elif type(time_str).__name__ == 'NoneType':
                break
                    
    elif sheet_name == "Temp DVC" or sheet_name == "Hourly Worker":
        for row_i in range(nrows,1,-1):
            time_str = sheet.api.Cells(row_i, col_num).value

            if type(time_str).__name__ == 'datetime':
                if time_str <= today and sheet.api.Cells(row_num, col_num-1).value != "离职" and sheet.api.Cells(row_i, col_num+3).value != "劳务工转正" and sheet.api.Cells(row_i, col_num+3).value != "转劳务工":
                    sheet.api.Cells(row_i, col_num - 1).value = "离职"
            elif type(time_str).__name__ == 'str':    
                if time_str <= formatted_today and sheet.api.Cells(row_num, col_num-1).value != "离职" and sheet.api.Cells(row_i, col_num+3).value != "劳务工转正" and sheet.api.Cells(row_i, col_num+3).value != "转劳务工":
                    sheet.api.Cells(row_i, col_num - 1).value = "离职"  
            elif type(time_str).__name__ == 'NoneType':
                break          
            
            
            
filename = 'F:/Python/Homework/auto_work_excel/WuHan Plant Staff NameList- 20180720 .xlsx'
excel_file1 = xlrd.open_workbook(filename)
excel_file = xw.Book(filename)       
while(1):
    
    print("请输入表格名:")
    
    sheet_name = input().strip()
    
    sheet = excel_file.sheets[sheet_name]
    sht = excel_file1.sheet_by_name(sheet_name)
    
    print("请输入员工名字:")
    empname = input().strip()
    row_num = search(sheet,sht,empname,sheet_name)
    col_num = departure_info(sheet,sht,row_num,sheet_name)
    shear_insert(sheet,sht,row_num,col_num)
    
    on_job_change(sheet,sht,sheet_name,row_num,col_num)
    print("是否确定完成?(0:完成  1:继续)")
    finish = int(input())
    if finish == 0:
        break