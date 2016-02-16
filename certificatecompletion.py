#! python3
# Program to query Argos to determine if someone is a non-degree completer
#
import threading, time, getpass, logging
# bs4
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait  # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

website = "https://argos.lipscomb.edu/"
argosUname = input('Please enter your username: ')
argosPasswd = getpass.getpass("Please enter your password: ")
cbrowser = webdriver.Firefox()


def openArgos():
    #Open Argos web viewer in Firefox and login
    logging.debug("Opening Firefox")
    cbrowser.implicitly_wait(5)
    cbrowser.get(website)
    loginformUname = WebDriverWait(cbrowser,5).until(EC.element_to_be_clickable((By.ID,'loginUsername')))
    loginformUname.click()
    loginformUname.send_keys(argosUname)
    loginformPasswd = WebDriverWait(cbrowser,5).until(EC.element_to_be_clickable((By.ID,'loginPassword')))
    loginformPasswd.click()
    loginformPasswd.send_keys(argosPasswd)
    cbrowser.find_element_by_xpath("//*[@id=\"modalLogin\"]/div[3]/button").click()
    cbrowser.find_element_by_link_text("Argos Web Viewer").click()
    cbrowser.switch_to.window(cbrowser.window_handles[-1])
    cbrowser.find_element_by_xpath("//*[@id=\"mCSB_1\"]/div[1]/div[2]/div[2]").click()
    cbrowser.find_element_by_xpath("//*[@id=\"mCSB_2\"]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]").click()
    studentID = cbrowser.find_element_by_xpath('//*[@id="datablock"]/div[5]/div[1]/div[25]/div[11]/input')
    studentID.click()
    studentID.send_keys("L21743148")
    #TODO iterate over a list of keys and copy the courses to an excel file
    Semester = cbrowser.find_element_by_xpath("//*[@id=\"datablock\"]/div[5]/div[1]/div[25]/div[6]/select/option[7]")
    Semester.click()
    Search = cbrowser.find_element_by_xpath("//*[@id=\"datablock\"]/div[5]/div[1]/div[25]/div[9]/div/span")
    Search.click()
    print(cbrowser.find_element_by_xpath("//*[@id=\"datablock\"]/div[5]/div[1]/div[25]/div[12]/div/div[1]/div[2]/div/div[2]/div[1]/div[4]").text)
    for i in range(1,10):
        if str(cbrowser.find_element_by_xpath("//*[@id=\"datablock\"]/div[5]/div[1]/div[25]/div[12]/div/div[1]/div[2]/div/div[2]/div[" + i + "]/div[4]").text) == "G":
            cbrowser.find_element_by_xpath("//*[@id=\"datablock\"]/div[5]/div[1]/div[25]/div[12]/div/div[1]/div[2]/div/div[2]/div[" + i + "]/div[4]").click()
        else:
            continue


#//*[@id="datablock"]/div[5]/div[1]/div[25]/div[12]/div/div[1]/div[2]/div/div[2]/div[1]/div[4]/span
#xpath selector for whether the student is grad or ug
#//*[@id="datablock"]/div[5]/div[1]/div[25]/div[12]/div/div[1]/div[2]/div/div[2]/div[2]/div[4]






def main():
    openArgos()

main()




#TODO Open a file with certificate enrolles and pull their course information from the Argos web viewer
#TODO Write mannnnnnny functions to determine if the enrolle is a completer or not


