#! python3
#Web scraper that pulls info from TN license database and adds to excel spreadsheet
#

import webbrowser, pyperclip, requests, bs4, openpyxl, os, logging, requests, re, getpass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC # available since 2.26.0

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s- %(message)s')

url = 'https://www.k-12.state.tn.us/authorize/'
#open firefox to tn database
browser = webdriver.Firefox()
browser.get(url)

# locate the username and password fields
username = browser.find_element_by_id("ContentPlaceHolder1_txtUserID")
password = browser.find_element_by_id("ContentPlaceHolder1_txtPassword")

# insert the username and password 
uname = input("Please input your username \n")
username.send_keys(uname)
pword = getpass.getpass("Please Enter Password \n")
password.send_keys(pword)
filename = input("What would like to name the file?")

#locate the login button
login = browser.find_element_by_id("ContentPlaceHolder1_cmdLogin")

#click the login button
login.click()
#locate the district user link
district = browser.find_element_by_link_text('District User')
#click on the district user click 
district.click()

# change cwd

os.chdir('M:\Assessment (New)')


#load work book
wb = openpyxl.load_workbook('test.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')

"""
#To do.  Ask for user input regarding filename and path
filename = input("Please enter the new working directory that you would like to use")
os.chdir(filename)"""
#create new workbook for data 
nwb = openpyxl.Workbook()
sheet1 = nwb.get_active_sheet()

#take SS for each student
try:
        for i in range(2, int(sheet.get_highest_row())):
                #input SS into the SS browser box
                ss = sheet.cell(row=i, column=3).value
                box = browser.find_element_by_id('MainContent_txtSSN')
                box.clear()
                box.send_keys(ss)
                search = browser.find_element_by_id('MainContent_btnSearch')
                search.click()
                try:
                        WebDriverWait(browser, 4).until(EC.element_to_be_clickable((By.ID,'MainContent_GridView1_LinkButton1_0')))
                        teacherid = browser.find_element_by_id('MainContent_GridView1_LinkButton1_0')
                        teacherid.click()
                        logging.info('Found teacher id and clicked')
						#gather teacher text
                        WebDriverWait(browser, 4).until(EC.text_to_be_present_in_element((By.CLASS_NAME, 'main'), 'purged'))
                        element = browser.find_element_by_id('ctl00_MainContent_ReportViewer1')
                        logging.info('Has found element')
                        infoneeded = element.text
                        logging.info(infoneeded)
                        wordregex = re.compile(r'\w+')
                        words = wordregex.findall(infoneeded)
                        columnnumber = 1
                        for line in words:
                                sheet1.cell(row=i, column=columnnumber).value = line
                                columnnumber += 1
                                
                except Exception as int:
                        print(int)
except:
        pass

nwb.save(filename + '.xlsx')