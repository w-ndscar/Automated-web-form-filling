from selenium import webdriver
import xlrd
 
driver = webdriver.Chrome( executable_path="")
driver.get("")

name = []
short_form = []
e_Address = []
passwd = []

file_location = ""
workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)
for col in range(sheet.ncol):
    name.append(sheet.cell_value(col,0))
    short_form.append(sheet.cell_value(col,1))
    e_Address.append(sheet.cell_value (col,2))
    passwd.append(sheet.cell_value(col,3))

print (name)
print (short_form)
print (e_Address)
print (passwd)

  
