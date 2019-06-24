###This program is for saving/downloading BRC details from the DGFT website.

import openpyxl, time, easygui, win32api
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from pynput.keyboard import Key, Controller
import subprocess
#these libraries need to be installed before running the program


wb = openpyxl.load_workbook('C:\\Users\\DELL\\Desktop\\seleniumtesting.xlsx')
sheet = wb['Sheet1']
#location of the file 'seleniumtesting.xlsx' should be maintained correctly with "//" instead of single "/"

row_count = sheet.max_row

n = row_count + 1
print(n)


keyboard = Controller()
driver = webdriver.Firefox(executable_path="C:\\Users\\DELL\\Downloads\\geckodriver-v0.24.0-win64\\geckodriver.exe")
#driver.get('http://dgftebrc.nic.in:8100/BRCQueryTrade/index.jsp')


def fillform(i):
    driver.get('http://dgftebrc.nic.in:8100/BRCQueryTrade/index.jsp')

    driver.find_element_by_name('iec').send_keys(sheet['E'+str(i)].value)

    driver.find_element_by_name('ifsc').send_keys(sheet['D'+str(i)].value)

    driver.find_element_by_name('brcno').send_keys(sheet['C'+str(i)].value)

    captchaelement = driver.find_element_by_name('captext')
	
    #Enter CAPTCHA manually here
    capvar = easygui.enterbox("Enter CAPTCHA code here: ");
    captchaelement.send_keys(capvar)


    driver.find_element_by_xpath("/html/body/form/div/center/table/tbody/tr[7]/td/p/input[1]").click() 
    # detailselement = driver.find_element_by_value('Show Details')
                
    windowBefore = driver.window_handles[0]

    try:
        driver.find_element_by_xpath("/html/body/div[2]/center/table/tbody/tr[2]/td[11]/form/font/input[4]").click()
        #print button on next screen
    except NoSuchElementException:
        win32api.MessageBox(0, "Please fill in the details correctly", "Incorrect CAPTCHA")
        #Doesn't appear if CAPTCHA is filled incorrectly
        fillform(i)

    #windowAfter = driver.window_handles[1]
    #driver.switch_to_window(windowAfter)
	
    #make sure the file location of autoIT's application file are correct.
    subprocess.call("C:\\Users\\DELL\\Desktop\\hello.exe")
    print("going ahead")
    keyboard.type(sheet['C'+str(i)].value)
    print("going ahead= saved?")
    subprocess.call("C:\\Users\\DELL\\Desktop\\hello2.exe")
    print("going ahead=closed")
    eleback = driver.find_element_by_link_text('Modify Query')
    eleback.click()
    print("what?")

    sheet['F'+str(i)] = "done"
    wb.save('C:\\Users\\DELL\\Desktop\\seleniumtesting.xlsx')

#sheet['F'+str(i)] = "done"
       
#windowAfter = driver.window_handles[1]

#driver.switch_to_window(windowAfter)


for x in range(2, n):

	fillform(x)
    
print('done')
wb.save('C:\\Users\\DELL\\Desktop\\seleniumtesting.xlsx')
win32api.MessageBox(0, 'The script was implemented successfully', 'Success')
driver.close()

