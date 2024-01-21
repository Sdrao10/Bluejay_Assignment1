import openpyxl


workbook = openpyxl.load_workbook("Assignment_Timecard.xlsx")

sheets = workbook.sheetnames
Sheet1= workbook['Sheet1']


sheet_obj = workbook.active