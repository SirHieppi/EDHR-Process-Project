import glob
import os
import selenium
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from pathlib import Path

path_to_download_folder = str(os.path.join(Path.home(), "Downloads"))
chromedriver = os.getcwd() + "\\chromedriver_win32\\chromedriver.exe"

# Urls
edhrUrl = 'http://ussd-illuminareporting.illumina.com/Reports/Pages/Report.aspx?ItemPath=%2fCamstar+Reports%2fProduction%2feDHR+-+Instrument+-+Detail'

def getLatestDownloadedPDF():
    listOfFiles = glob.glob((path_to_download_folder + "/*")) # * means all if need specific format then *.csv
    latestFile = max(listOfFiles, key=os.path.getctime)
    
    print(latestFile)

    return latestFile

def getEdhrPDF(instrumentSerialNumber):
    browser = visit(edhrUrl)

    delay = 5

    # Enter serial number into search box
    searchBox = browser.find_element_by_id('ctl32_ctl04_ctl03_txtValue')
    searchBox.send_keys(instrumentSerialNumber)

    time.sleep(3)

    # # Click on view report button
    submitButton = browser.find_element_by_id('ctl32_ctl04_ctl00')
    submitButton.click()

    time.sleep(3)

    # # Find drop down export menu
    menu = browser.find_element_by_id('ctl32_ctl05_ctl04_ctl00_Menu')
    browser.execute_script("arguments[0].style.visibility = 'visible'; arguments[0].style.display = 'block'", menu)
    time.sleep(3)

    # # Click save eDHR report as PDF
    pdfButton = menu.find_element_by_xpath("//div//a[@title='PDF']")
    print(pdfButton.get_attribute('title'))
    pdfButton.click()

    time.sleep(3)

    browser.close()

    return  getLatestDownloadedPDF()

def visit(url):
    print("[INFO] Visiting {}".format(url))
    browser = webdriver.Chrome(executable_path=r'{}'.format(chromedriver))         
    browser.get(url)
    
    return browser

def main():
    getEdhrPDF('A01420')

main()
