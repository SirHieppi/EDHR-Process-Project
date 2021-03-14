from classes.WebAutomation import WebAutomation
from classes.ExcelHandler import ExcelHandler

def main():
    wa = WebAutomation()
    excel = wa.getEdhrExcel('A01420')

    eh = ExcelHandler()
    eh.processExcel(excel)

main()
