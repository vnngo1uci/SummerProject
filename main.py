import openpyxl

data_file = '/Users/vincentngo/Desktop/Summer23/SummerProject/ABCDCatering.xls'

workbook_obj = openpyxl.load_workbook(data_file)

sheet_obj = workbook_obj.active
 

cell_obj = sheet_obj.cell(row = 1, column = 1)
 

print(cell_obj.value)