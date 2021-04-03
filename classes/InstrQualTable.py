import win32com.client
import win32api
import os
from os import path

class InstrQualTable:
    def __init__(self):
        self.qualificationIndexes = {
            "MIB BOARD": 4,
            "FCH BOARD": 5,
            "FIB BOARD": 6,
            "SBC BOARD": 7,
            "SYS BOARD": 8,
            
            "SYRINGE PUMP": 9,
            "FLOW RATE SENSOR": 10,
            "FLUIDICS MODULE (FLM)": 11,
            "REAGENT CHILLER ASSEMBLY (RCA)": 12,
            "BUFFER INTERFACE MODULE (BIM)": 13,
            "BIM DRAWER": 14,
            "DEGASSER": 15,
            "ASSY, TUBING KIT, COMMON LINES, SLEEVE": 16,
            
            "DUAL ACTUATION DECK (DAD)": 17,
            "CABLE TRACK ASSEMBLY (CTA)": 18,
            
            "ASSY, XY STAGE MODULE (XYA)": 19,
            "TIP TILT ASSEMBLY (TTA)": 20,
            
            "ASSY, CAMERA MODULE (CAM)": 21,
            "ASSY, FOCUS TRACKING MODULE (FTM)": 22,
            "ASSY, EMISSION OPTICS MODULE (EOM)": 23,
            "LGM_HT_V2 ACTUATORS (LGM)": 24,
            "BIRD UBERTARGERT": 25,
            "Z STAGE MOTION CONTROLLER": 26,
            
            "FC ENCLOSURE": 27,
            "ASSY, LIGHTBAND": 28,
            "ASSY, OPA W/ NOZZLE (OPA)": 29,
            "COMPUTE ENGINE (CE)": 30,
            "OPTICAL AIR SHIELD (OAS)": 31,
            "COOLANT RESERVOIR TRAY ASSEMBLY": 32,
            "PSU": 33
        }

        self.instrQualTableExcelPath = os.getcwd() + "\\1000000052652_v01_TD,NovaSeq,MATL-RPLC-QUAL.xlsx"

        self.userChoices = []

        self.qualificationTable = {}

        self.combinedStepsInOrder = {}

    def getQualificationSteps(self, qualifications, operations=[]):
        steps = {}

        for operation in operations:
            if not operation:
                return self.qualificationTable[qualification]

            for q in qualifications:
                for op in operations:
                    if op not in steps:
                        steps[op] = self.qualificationTable[q][op]
                    else:
                        for step in self.qualificationTable[q][op]:
                            if step not in steps[op]:
                                steps[op].append(step)
                    steps[op] = sorted(steps[op], key = lambda x: x[1])

        return steps

    def openInstrQualTableExcel(self):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(r'{}'.format(self.instrQualTableExcelPath))        
        
        ws = wb.Worksheets["Qualification Table"]
        
        return wb, ws, excel

    def selectQualification(self, qualifications):
        wb, ws, excel = self.openInstrQualTableExcel()
        
        for q in qualifications:
            col = self.qualificationIndexes[q]

            if q not in self.qualificationTable:
                self.qualificationTable[q] = {}

            for row in range(3, 99):
                cell = ws.Cells(row, col).Value
                task = ws.Cells(row, 3).Value
                operation = ws.Cells(row, 1).Value

                if operation not in self.qualificationTable[q]:
                    self.qualificationTable[q][operation] = []
                
                if cell:
                    self.qualificationTable[q][operation].append(tuple((task, row)))

        wb.Close(True)
        excel.Quit()

        self.printSelectedQualifications()

    def removeQualification(self, qualifications):
        pass

    def checkQualification(self, performedTasks):
        pass

    def printSelectedQualifications(self):
        print("Qualification table:")

        for q in self.qualificationTable:
            print(q + ": ")
            for operation in self.qualificationTable[q]:
                print("\tOP " + operation + ": " + str(self.qualificationTable[q][operation]))
                # print(self.qualificationTable[q][operation])
                print("")
