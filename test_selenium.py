from selenium import webdriver
#object of ChromeOptions class
o = webdriver.ChromeOptions()
#adding specific Chrome Profile Path
o.add_arguments = {'user-data-dir': "C:/UsersArun/AppData/Local/Google/Chrome/User Data/Default"}
#set chromedriver.exe path
driver = webdriver.Chrome(executable_path="C:\chromedriver.exe", options=o)
#maximize browser
driver.maximize_window()
#launch URL
driver.get("https://www.tutorialspoint.com/index.htm")
#get browser title
print(driver.title)
#close browser
driver.close()