import os
from classes.WebAutomation import WebAutomation
from classes.ExcelHandler import ExcelHandler
from os import path

def main():
    # wa = WebAutomation()
    # excel = wa.getEdhrExcel('A01400')

    eh = ExcelHandler()
    eh.processExcel(os.getcwd() + "\\eDHR_A1409.xlsx")

main()
