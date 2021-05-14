import os
from classes.WebAutomation import WebAutomation
from classes.ExcelHandler import ExcelHandler
from classes.InstrQualTable import InstrQualTable
from os import path

def main():
    # wa = WebAutomation()
    # excel = wa.getEdhrExcel('A01400')

    eh = ExcelHandler()
    eh.processExcel(os.getcwd() + "\\eDHR_A1409.xlsx")

    iq = InstrQualTable()
    

    # iq.selectQualification(["REAGENT CHILLER ASSEMBLY (RCA)"])
    
    iq.selectQualification(["REAGENT CHILLER ASSEMBLY (RCA)", 
                            "FLUIDICS MODULE (FLM)",
                            "DUAL ACTUATION DECK (DAD)",
                            "CABLE TRACK ASSEMBLY (CTA)"])
    
    # print(iq.getQualificationSteps(["REAGENT CHILLER ASSEMBLY (RCA)"],["FUNC"]))

    # iq.getQualificationSteps(["REAGENT CHILLER ASSEMBLY (RCA)"])

    # print(iq.getQualificationSteps(["REAGENT CHILLER ASSEMBLY (RCA)", 
    #                                 "FLUIDICS MODULE (FLM)",
    #                                 "DUAL ACTUATION DECK (DAD)",
    #                                 "CABLE TRACK ASSEMBLY (CTA)"]))


main()
