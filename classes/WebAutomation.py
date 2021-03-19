import glob
import os
import selenium
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from pathlib import Path


class WebAutomation:
    def __init__(self):
        self.path_to_download_folder = str(os.path.join(Path.home(), "Downloads"))
        self.chromedriver = os.getcwd() + "\\chromedriver_win32\\chromedriver.exe"

        # Urls
        self.edhrUrl = 'http://ussd-illuminareporting.illumina.com/Reports/Pages/Report.aspx?ItemPath=%2fCamstar+Reports%2fProduction%2feDHR+-+Instrument+-+Detail'

    def getLatestDownloadedExcel(self):
        listOfFiles = glob.glob((self.path_to_download_folder + "/*.xlsx")) # * means all if need specific format then *.csv
        latestFile = max(listOfFiles, key=os.path.getctime)
        
        print(latestFile)

        return latestFile

    def getEdhrExcel(self, instrumentSerialNumber):
        browser = self.visit(self.edhrUrl)

        delay = 3

        # Enter serial number into search box
        searchBox = browser.find_element_by_id('ctl32_ctl04_ctl03_txtValue')
        searchBox.send_keys(instrumentSerialNumber)

        time.sleep(delay)

        # # Click on view report button
        submitButton = browser.find_element_by_id('ctl32_ctl04_ctl00')
        submitButton.click()

        time.sleep(delay)

        # # Find drop down export menu
        menu = browser.find_element_by_id('ctl32_ctl05_ctl04_ctl00_Menu')
        browser.execute_script("arguments[0].style.visibility = 'visible'; arguments[0].style.display = 'block'", menu)
        
        time.sleep(delay)

        # # Click save eDHR report as PDF
        pdfButton = menu.find_element_by_xpath("//div//a[@title='Excel']")
        print(pdfButton.get_attribute('title'))
        pdfButton.click()

        time.sleep(delay)

        browser.close()

        return  self.getLatestDownloadedExcel()

    def visit(self, url):
        print("[INFO] Visiting {}".format(url))
        browser = webdriver.Chrome(executable_path=r'{}'.format(self.chromedriver))         
        browser.get(url)
        
        return browser
