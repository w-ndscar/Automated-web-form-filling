from time import sleep
from turtle import clear
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import keyboard
import openpyxl
from openpyxl.utils import get_column_letter

#Chrome driver
chrome_driver = "C:/chromedriver.exe"

#Chrome - Debugger Mode
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(chrome_driver, chrome_options=chrome_options)

#Just for a reference - see console
print(driver.title)

#Lists to store the Excel values
proj = []
dwg_name = []
dwg_desc = []


#Load workbook and sheet
file_location = "C:/Users/Arun/Documents/Timesheet/Timesheet_Team_Copy.xlsx"
workbook = openpyxl.load_workbook(file_location)
sheet = workbook["Team7"]

#Loops to store values in the lists
col = 2
for row in range(3, 11):
    char = get_column_letter(col)
    cell_name = char + str(row)
    proj.append(sheet[cell_name].value)
    
col=col+1
for row in range(3, 11):
    char = get_column_letter(col)
    cell_name = char + str(row)
    dwg_name.append(sheet[cell_name].value)

col=col+1
for row in range(3, 11):
    char = get_column_letter(col)
    cell_name = char + str(row)
    dwg_desc.append(sheet[cell_name].value)


#Printing just for a reference
print (proj)
print (dwg_name)
print (dwg_desc)


sleep(1)


for i in range(len(proj)):
    #Click - New button
    elem2 = driver.find_element("xpath", '/html/body/app-root/app-project-detail/div/div[4]/div/button')
    elem2.click()
    sleep(1)
    #Fields
    elem3 = driver.find_element("xpath", '/html/body/app-root/app-project-detail/div/div[4]/div/div/div/div/div[1]/ng-autocomplete/div[1]/div[1]/input')
    elem3.send_keys(proj[i])
    elem4 = driver.find_element("xpath", '/html/body/app-root/app-project-detail/div/div[4]/div/div/div/div/input[1]')
    elem4.send_keys(dwg_name[i])
    elem5 = driver.find_element("xpath", '/html/body/app-root/app-project-detail/div/div[4]/div/div/div/div/input[2]')
    elem5.send_keys(dwg_desc[i])
    
    #Hotkey to wait before saving
    keyboard.wait("esc")

    #Dropdown
    elem6 = driver.find_element("xpath", '/html/body/app-root/app-project-detail/div/div[4]/div/div/div/div/div[2]/button[1]')
    elem6.click
    sleep(0.5)
    
    #Hotkey to wait before the next iteration
    keyboard.wait("esc")

