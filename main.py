import openpyxl
from collections import defaultdict
import xlsxwriter

DATA_FILE = '/Users/vincentngo/Desktop/Summer23/SummerProject/ABCDCatering.xlsx'

class OrderingList(object):

    def __init__ (self, data):
        self.workbook_obj = openpyxl.load_workbook(data, data_only = True)
        self.sheet_obj = self.workbook_obj.active
        self.structure = defaultdict(defaultdict)

    def create_reverse_index(self):
       print(self.sheet_obj["A"+"100"])
       for cell in self.sheet_obj['G']:
            if cell.value != None and cell.value != "Order Number":
                item = str(cell.row)
                self.structure[cell.value]["Fiscal Year"] = self.sheet_obj["A"+item].value
                self.structure[cell.value]["Fiscal Quarter"] = self.sheet_obj["B"+item].value
                self.structure[cell.value]["Food"] = self.sheet_obj["C"+item].value
                self.structure[cell.value]["Company"] = self.sheet_obj["E"+item].value
                self.structure[cell.value]["Order Date"] = self.sheet_obj["H"+item].value
                self.structure[cell.value]["Sales Rep"] = self.sheet_obj["I"+item].value
                self.structure[cell.value]["Net"] = self.sheet_obj["K"+item].value
                self.structure[cell.value]["Gross"] = self.sheet_obj["L"+item].value
                self.structure[cell.value]["Quantity"] = self.sheet_obj["M"+item].value
        
    def helper_print(self):
        print(self.structure)
        
    def get_num_items_per_quarter(self) -> dict:
        structure = defaultdict(int)
        for cell in self.sheet_obj['B']:
            if cell.value == None:
                pass
            else:
                if "Q" in cell.value and "-" in cell.value:
                    structure[cell.value] += 1
        return structure
    
    def get_num_foods(self) -> dict:
        structure = defaultdict(int)
        for cell in self.sheet_obj['D']:
            if cell.value == None:
                pass
            else:
                structure[cell.value] += 1
        return structure
    
    
    def print_results(self) -> None:
        print("-----Welcome to the Excel Database-----")
        print("-----   Input Choices Are Below   -----")
        print("1. Total Per Quarter")
        print("2. Total Food")
        print("3. Exit")
        while True:
            user_input = str(input("Please Enter What Information You Need: "))
            print()
            if user_input == "3" or user_input == "Exit":
                print("Thank You For Using The Excel Database\n")
                break
            elif user_input == "2" or user_input == "Total Food":
                temp = self.get_num_foods()
                for i in temp.keys():
                    print(f"We Have Sold {temp[i]} {i}s This Year")
            elif user_input == "1" or user_input == "Total Per Quarter":
                temp = self.get_num_items_per_quarter()
                for i in temp.keys():
                    print(f"For Quarter {i[-1]}, We have Sold {temp[i]} Items")
            else:
                print("Please Give A Valid Command")
            print()

def main():
    Excel_Sheet = OrderingList(DATA_FILE)
    #Excel_Sheet.print_results()
    Excel_Sheet.create_reverse_index()
    Excel_Sheet.helper_print()

if __name__ == "__main__":
    main()


