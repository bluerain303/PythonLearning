from openpyxl import Workbook

#==========
def print_rows():
    print ("========================")
    for row in mySheet.iter_rows(values_only=True):
        print(row)
    print ("========================")
#==========

fileFullPath = "/Users/yishi/Documents/GitHub/PythonLearning/Excel/Cars.xlsx"

myWorkbook = Workbook()
mySheet = myWorkbook.active

mySheet["A1"] = "Brand"
mySheet["B1"] = "Model"
mySheet["C1"] = "Price"
print_rows()

# Change value of a cell
mySheet["C1"] = "EUR"
print_rows()
# Or
myCell = mySheet["C1"] = "SEK"
print_rows()

# Insert rows
mySheet.insert_rows(idx=1,amount=2)
print_rows()
# Delete rows
mySheet.delete_rows(idx=1,amount=2)
print_rows()
# Insert columns
mySheet.insert_cols(idx=2, amount=1)
print_rows()
# Delete columns
mySheet.delete_cols(idx=2,amount=1)
print_rows()


myWorkbook.save(fileFullPath)

