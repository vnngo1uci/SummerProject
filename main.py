import openpyxl

data_file = '/Users/vincentngo/Desktop/Summer23/SummerProject/MultiplicationTable.xlsx'

workbook_obj = openpyxl.load_workbook(data_file, data_only = True)

sheet_obj = workbook_obj.active
 

for cell in sheet_obj['C']:
    if cell == None:
        pass
    elif cell.coordinate == "C2":
        print("3")
    else:
        print(cell.value)