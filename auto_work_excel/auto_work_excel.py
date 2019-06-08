# encoding:utf-8

import xlrd
import xlwings

# 打开excel文件
excel_file = xlrd.open_workbook('F:/Python/Homework/auto_work_excel/WuHan Plant Staff NameList- 20180720 .xlsx')
print('请输入表格名称:')

#  确定是哪个表格
sheet_name = input()
sheet = excel_file.sheet_by_name()




# 搜索名字
def search_name(empname,sheet):
    person_count = 0
    # 得到excel中首行的数据
    first_row_values = sheet.row_values(0)
    # 遍历列表的元素
    for cols_i in range(1,len(first_row_values)):
        # 要求匹配姓名那一列
        if first_row_values[cols_i].split()[0] == "姓名" :
            # 进入姓名那一列之后,寻找你要找的人
               # 所有人的名字
            
            col_values = sheet.col_values(cols_i)
            for rows_i in range(1,len(col_values)):
                if empname == col_values[rows_i]:
                    person_count += 1
                    empname_data = sheet.row_values(rows_i)
                
                    
            if person_count != 1 and person_count != 0:
                print("跟此人同姓名的有%d个" % person_count)
                print("请输入工号查询:")
                if sheet_name == "Perm":
                    person_id = int(input())
                elif sheet_name == "Temp DVC" or sheet_name == "Hourly Worker":
                    person_id = input()
                ncols = sheet.ncols
                for cols_j in range(2,ncols):
                    col_values_every = sheet.col_values(cols_j)
                    for rows_j in range(1,len(col_values_every)):
                        if person_id == col_values_every[rows_j]:
                            empname_data = sheet.row_values(rows_j)
                            return empname_data
                                   
            else:
                return empname_data
        



# empname = input()
# print(search_name(empname,sheet))

    