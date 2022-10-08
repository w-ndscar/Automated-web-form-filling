from time import sleep
from selenium import webdriver
import openpyxl
from openpyxl.utils import get_column_letter
 
driver = webdriver.Chrome(executable_path="C:/chromedriver.exe")
driver.get("https://projectmanagement-8d5c4.web.app/userManager")

name = []
short_form = []
e_Address = []
passwd = []


file_location = "C:/Users/Arun/Documents/Timesheet/Timesheet_Team_Copy.xlsx"
workbook = openpyxl.load_workbook(file_location)
sheet = workbook.active
col = 2
for row in range(3, 17):
    char = get_column_letter(col)
    cell_name = char + str(row)
    name.append(sheet[cell_name].value)
    
col=col+1
for row in range(3, 17):
    char = get_column_letter(col)
    cell_name = char + str(row)
    short_form.append(sheet[cell_name].value)

col=col+1
for row in range(3, 17):
    char = get_column_letter(col)
    cell_name = char + str(row)
    e_Address.append(sheet[cell_name].value)

col=col+1
for row in range(3, 17):
    char = get_column_letter(col)
    cell_name = char + str(row)
    passwd.append(sheet[cell_name].value)

print (name)
print (short_form)
print (e_Address)
print (passwd)

elem = driver.find_element("xpath", '/html/body/app-root/app-landing/div/div/app-login/div/div[1]/input')
elem.send_keys("arun.venkatesh@csiglobal.net")

elem = driver.find_element("xpath", '/html/body/app-root/app-landing/div/div/app-login/div/div[2]/input')
elem.send_keys("projectinc0.")

sleep(0.5)

elem = driver.find_element("xpath", '/html/body/app-root/app-landing/div/div/app-login/div/div[3]/input')
elem.click()

driver.get("https://projectmanagement-8d5c4.web.app/userManager")