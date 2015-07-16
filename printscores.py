#! python3
# program that logs into tn license database and prints the most recent Praxis scores for college 

import bs4, threading,time, getpass
from selenium import webdriver 
from selenium.webdriver.support.ui import WebDriverWait # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.common.by import By
import logging
import pyautogui
import time
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s- %(message)s')

url = 'https://tlcs.ets.org/clientservices/profile/login/login.do'
uname = input("Username: ")
pword = getpass.getpass("Password: ")
printcode = input("Please enter the print code:" )
browser = webdriver.Firefox()


def login():
	browser.get(url)
	unameinput = browser.find_element_by_id('eias-user-name')
	unameinput.send_keys(uname)
	pwordinput = browser.find_element_by_id('eias-user-password')
	pwordinput.send_keys(pword)
	login = browser.find_element_by_xpath('//*[@id="greds-sign-in"]/table/tbody/tr[3]/td[3]/input')
	login.click()

def findscores():
	scorereport = WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.LINK_TEXT, "Test Taker Score Reports")))
	scorereport.click()
	#TODO change based upon the week.  Need to increment the id up 2 for every week. 
	currentweek = WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.ID,"m_m_mid_mid_rptResult_ctl00_ctl04_GridClientSelectColumnSelectCheckBox")))
	currentweek.click()
	next = WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.ID,"m_m_mid_mid_bNext0")))
	next.click()
	
def printscores():
	#note only works in Firefox
	#selects all of the scores
	try:
		selectall =  WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.ID, "m_m_mid_mid_ReportByDate2_rptDataView_ctl00_ctl02_ctl01_GridClientSelectColumnSelectCheckBox")))   
		selectall.click()
	except:
		pass
	#exports to pdf
	export = WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.ID, "m_m_top_wel_imgBtnExport"))) 
	export.click()
	logging.debug("Still working")
	#chooses pdf from the popup
	browser.switch_to.frame('rwPreview')
	choosepdf = WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.ID, "radio4")))
	choosepdf.click()
	donebutton = browser.find_element_by_id('imgBtnDone')
	donebutton.click()
	#todo let it continue when the pdf downloads and break if it does not download
	try:
		
	time.sleep(5)
	pyautogui.press('enter')
	time.sleep(2)
	pyautogui.hotkey('ctrl', 'j')
	time.sleep(2)
	pyautogui.press('enter')
	time.sleep(4)
	pyautogui.hotkey('ctrl', 'p')
	time.sleep(4)
	pyautogui.press('enter')
	time.sleep(4)
	pyautogui.typewrite(printcode)
	pyautogui.press('enter')
	
	
	
	
login()
findscores()
printscores()

