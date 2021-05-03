import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import selenium.webdriver.chrome.options
import time
import os
from selenium.webdriver.chrome.options import Options
import pyautogui
from selenium.webdriver.support.ui import Select
import xlwt
from xlwt import Workbook
name = input("Enter the intials of the state you want to download data of? (Check from website for intials!!!)")
driverrange1 = int(input("Enter Min Number if you want to apply driver filter else type (none)!"))
driverrange2 = int(input("Enter Max Number if you want to apply driver filter else type (none)!"))
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')

PATH = "C:\Program Files (x86)\chromedriver.exe"
options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-user-media-security=true")
options.add_argument("--use-fake-ui-for-media-stream")
options.add_argument("--disable-popup-blocking")
driver = webdriver.Chrome(PATH, options=options)
driver.get("https://directory.dat.com/?t=%22CARRIER%22")
driver.maximize_window()
time.sleep(5)
email=driver.find_element_by_id("mat-input-1")
email.send_keys('')# Add username
pass1=driver.find_element_by_id("mat-input-0")
pass1.send_keys('')#add password
driver.find_element_by_id("submit-button").click()
time.sleep(7)
close=driver.find_elements_by_class_name("mat-icon-button")
driver.find_element_by_class_name("mat-checkbox-inner-container").click()
list1=[]
count=0
time.sleep(3)
for i in close:
    count=count+1
    list1.append(i)
    try:
        i.click()
        print("work")
    except:
        print("not work")

driver.get("https://directory.dat.com/?t=%22CARRIER%22&q=%22%22&ls=%5B%22"+name+"%22%5D")
time.sleep(5)
if driverrange1 and driverrange2 == 0:
    print(driverrange1)
    print(driverrange2)
    print("no range")
else:
    print('range avaiable')
    driver.find_element_by_xpath("//*[contains(text(), 'Drivers')]").click()
    data4=driver.find_element_by_id("mat-input-3")
    data5 = driver.find_element_by_id("mat-input-4")
    data4.send_keys(driverrange1)
    data5.send_keys(driverrange2)

close=driver.find_elements_by_class_name("mat-icon-button")
list1=[]
count=0
for i in close:
    count=count+1
    list1.append(i)
    try:
        i.click()
        print("work")
    except:
        print("not work")
page_number=driver.find_element_by_class_name("mat-paginator-range-label")
page_number1=page_number.text
page_number1=page_number1[-2:]
print(page_number1)
page_number1=page_number1.strip()

sheet1.write(0,0,'First Name')
sheet1.write(0,1,'Last Name')
sheet1.write(0,2,'Docket')
sheet1.write(0,3,'DOT number')
sheet1.write(0,4,'Company Type')
sheet1.write(0,5,'Office Phone')
count=1
count1 = 1
count2 = 1
count3 = 1
count4 = 1
count5 = 1
count6 = 1
count7=0

for i in range(int(page_number1)):
    a=[]
    b=[]
    c=[]
    d=[]
    e=[]
    f = []
    g = []
    a = driver.find_elements_by_class_name("company-list__name")
    b = driver.find_elements_by_class_name("cdk-column-docket")
    c = driver.find_elements_by_class_name("cdk-column-dotNumber")
    d = driver.find_elements_by_class_name("cdk-column-companyType")
    e = driver.find_elements_by_class_name("cdk-column-officePhone")
    f=driver.find_elements_by_class_name("company-list__location")



    driver.find_element_by_xpath("//body").click()
    time.sleep(3)


    for i in a:
        sheet1.write(count, 0, i.text)
        count = count + 1

    for i in f[1:]:
        sheet1.write(count1, 1, i.text)
        count1 = count1 + 1

    for i in b[1:]:
        sheet1.write(count2, 2, i.text)
        count2 = count2 + 1

    for i in c[1:]:
        sheet1.write(count3, 3, i.text)
        count3 = count3 + 1

    for i in d[1:]:
        sheet1.write(count4, 4, i.text)
        count4 = count4 + 1

    for i in e[1:]:
        sheet1.write(count5, 5, i.text)
        count5 = count5 + 1

    driver.find_element_by_xpath("//body").click()
    driver.find_element_by_class_name("mat-paginator-navigation-next").click()

    time.sleep(5)







wb.save('xlwt example.xls')



#select.select_by_index(i)
#print(len(list1))

