import openpyxl


workbook = openpyxl.load_workbook("Assignment_Timecard.xlsx")

sheets = workbook.sheetnames
Sheet1= workbook['Sheet1']


sheet_obj = workbook.active
row = sheet_obj.max_row     # find the no of rows
column = sheet_obj.max_column      #find the column



print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>  Position id and name of employees who has worked for 7 consecutive days")

print("************************************************************************************")
counter=1
for i in range(2, row + 1):
    
    time_in = sheet_obj.cell(row=i, column=3)
    time_in_next = sheet_obj.cell(row=i+1, column=3)
    value1=time_in.value
    value2=time_in_next.value

    name = sheet_obj.cell(row=i, column=8)        # for the name of the employee
        
    name_next = sheet_obj.cell(row=i+1, column=8)     # for the next Employee which we will compare 
    emp_name = name.value                                 #cleared name of employee
    emp_name_next = name_next.value                       #next Employee name

    postion_id=sheet_obj.cell(row=i, column=1)
    
    
    
    
    if value1=="" or value2=="":               #for Eleminate the empty values
        counter=1
        continue
    
    else:
        date = value1.date()
        date_next=value2.date()              # here is decleae only the date from the date and time formate
        day=date.day                         # here is find the day from the date
        day_next=date_next.day                

        if day==day_next and emp_name==emp_name_next:
            continue
        elif day+1==day_next and emp_name==emp_name_next:
            counter=counter+1
            if counter==7:
                pos_id = postion_id.value
                print("ID is ->          ",pos_id)
                print("Employee name ->  ",emp_name)
        else:
            counter=1
            continue
