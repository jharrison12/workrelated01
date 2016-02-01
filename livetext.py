#! python3
# program that opens Livetext automatically so I don't have to do it

import threading, time, getpass, logging
# bs4
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait  # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s- %(message)s')

url = 'http://www.livetext.com/#'
course = input('Please enter the course number Jonathan: ')
subject = input('Please enter the subject (e.g. EG, ED, EGSE, etc.): ')
livetextuname = input('Please enter your username: ')
livetextpword = getpass.getpass('Enter your Livetext password: ')
semestercode = input('Enter the semester code: ')

# open course list for Lipscomb University banner page
logging.debug("Has accepted the variables")


def opencourse():
    logging.debug("Opening Firefox")
    cbrowser = webdriver.Firefox()
    cbrowser.get('https://bannerweb.lipscomb.edu/ssb_prod/stw_lu_schedule1.PW_SelSchClass')
    cn = cbrowser.find_element_by_name('sel_crse')
    # insert the course number in the search bar
    cn.send_keys(course)
    # Chooses Summer 2015 using XPATH.
    # Todo change so it DOESN'T choose all subjects
    semestercodedict = {'201620': 2, '201610': 4, '201530': 5, '201520': 6  }
    if semestercode in semestercodedict:
        semester = cbrowser.find_element_by_xpath(
            '/html/body/form/table/tbody/tr[1]/td[2]/select/option[' + str(
                    semestercodedict[semestercode]) + ']')
        semester.click()
    else:
        print('Please choose a valid semester')
    # Chooses EG using Xpath
    subjectdict = {'EG': 33, 'EGSE': 35, 'ED': 31, 'EGEL': 34}
    # unclick default all subjects
    unclickallsubjects = cbrowser.find_element_by_xpath(
        '/html/body/form/table/tbody/tr[2]/td[2]/select/option[1]')
    unclickallsubjects.click()
    if subject in subjectdict:
        education = cbrowser.find_element_by_xpath(
            '/html/body/form/table/tbody/tr[2]/td[2]/select/option[' + str(
                    subjectdict[subject]) + ']')
        education.click()
        #/html/body/form/table/tbody/tr[2]/td[2]/select/option[33]

    else:
        pass
    # Submits the searchclass button
    searchclass = cbrowser.find_element_by_xpath("/html/body/form/pre/center/input[1]")
    searchclass.click()


threadObj = threading.Thread(target=opencourse)


# Open livetext.com and find the appropriate course

def openlivetext():
    browser = webdriver.Firefox()

    browser.get(url)
    loginform = browser.find_element_by_link_text("Login")
    loginform.click()
    # insert username
    username = browser.find_element_by_id('user')
    username.send_keys(livetextuname)
    # insert password
    password = browser.find_element_by_id('pswd1')
    password.send_keys(livetextpword)
    # submit uname and password
    button = browser.find_element_by_tag_name('Button')
    button.click()
    # Journey to course admin page
    courseadmin = browser.find_element_by_id('topLevelNavCourseAdminButton')
    courseadmin.click()
    # Wait for "Course Editor and Exporter" link to load
    coursecatalog = WebDriverWait(browser, 3).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Course Editor and Exporter")))
    coursecatalog.click()
    browser.implicitly_wait(2)  # seconds
    semesterinput = browser.find_element_by_name('search')
    # input the semester needed
    semesterinput.send_keys(semestercode)
    # click search button
    browser.find_element_by_link_text("Search").click()
    links = browser.find_elements_by_class_name('action')
    # obtain list of links by their class name, click on the #2 index since there are multiple links that meet this criteria
    links[2].click()
    egclass = browser.find_element_by_name('search')
    egclass.clear()
    egclass.send_keys(course)
    WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.LINK_TEXT, "Search"))).click()


threadObj1 = threading.Thread(target=openlivetext)

threadObj.start()
time.sleep(4)
threadObj1.start()
