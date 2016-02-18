#! python3
# Program to query Argos to determine if someone is a non-degree completer
# 2/18/2016 Paused the development of this to see if there was a quicker solution
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
    cbrowser.get(website)
    loginformUname = WebDriverWait(cbrowser,5).until(EC.element_to_be_clickable((By.ID,'loginUsername')))
    loginformUname.click()
    loginformUname.send_keys(argosUname)
    loginformPasswd = WebDriverWait(cbrowser,5).until(EC.element_to_be_clickable((By.ID,'loginPassword')))
    loginformPasswd.click()
    loginformPasswd.send_keys(argosPasswd)
    cbrowser.find_element_by_xpath("//*[@id=\"modalLogin\"]/div[3]/button").click()
    argosWebViewer = WebDriverWait(cbrowser,4).until(EC.element_to_be_clickable((By.LINK_TEXT,'Argos Web Viewer')))
    argosWebViewer.click()
    #Switch to newest window with the code below
    cbrowser.switch_to.window(cbrowser.window_handles[-1])
    argosStudentView = WebDriverWait(cbrowser,4).until(EC.element_to_be_clickable((By.XPATH, "//*[@id=\"mCSB_1\"]/div[1]/div[2]/div[2]")))
    argosStudentView.click()
    ultimateStudentview = WebDriverWait(cbrowser, 4).until(EC.element_to_be_clickable((By.XPATH, "//*[@id=\"mCSB_2\"]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]")))
    ultimateStudentview.click()
    studentID = WebDriverWait(cbrowser,3).until(EC.element_to_be_clickable((By.XPATH, "//*[@id=\"datablock\"]/div[5]/div[1]/div[25]/div[11]/input")))
    studentID.click()
    studentID.send_keys("L21743148")
    #TODO iterate over a list of keys and copy the courses to an excel file
    Semester = cbrowser.find_element_by_xpath("//*[@id=\"datablock\"]/div[5]/div[1]/div[25]/div[6]/select/option[7]")
    Semester.click()
    Search = cbrowser.find_element_by_xpath("//*[@id=\"datablock\"]/div[5]/div[1]/div[25]/div[9]/div/span")
    Search.click()
    #TODO maybe try except clause here
    for i in range(1,10):
        if WebDriverWait(cbrowser,4).until(EC.element_to_be_clickable((By.XPATH, "//*[@id=\"datablock\"]/div[5]/div[1]/div[25]/div[12]/div/div[1]/div[2]/div/div[2]/div[" + str(i) + "]/div[4]"))).text == "G":
            cbrowser.find_element_by_xpath("//*[@id=\"datablock\"]/div[5]/div[1]/div[25]/div[12]/div/div[1]/div[2]/div/div[2]/div[" + str(i) + "]/div[4]").click()
        else:
            continue
        break
    # May not need the lines below because the Classes may already be visibile
    #Search = cbrowser.find_element_by_xpath("//*[@id=\"datablock\"]/div[5]/div[1]/div[18]/div/span")
    #Search = WebDriverWait(cbrowser,2).until(EC.element_to_be_clickable((By.XPATH,"//*[@id=\"datablock\"]/div[5]/div[1]/div[18]/div/span")))
    #Search.click()
    print(WebDriverWait(cbrowser,4).until(EC.element_to_be_clickable((By.XPATH,"//*[@id=\"datablock\"]/div[5]/div[6]/div[25]/div[4]/div/div[1]/div[2]/div/div[2]"))).text)


def main():
    openArgos()

main()




#TODO Open a file with certificate enrolles and pull their course information from the Argos web viewer
#TODO Write mannnnnnny functions to determine if the enrolle is a completer or not


