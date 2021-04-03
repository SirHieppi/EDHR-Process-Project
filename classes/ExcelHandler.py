import win32com.client
import win32api
import os
from os import path

class ExcelHandler:
    def __init__(self):
        self.edhrCheckExcelPath = os.getcwd() + "\\eDHR_check2.xlsx"
        # key = tool name; val = list of ILM #s
        self.tools = {}
        self.incorrectILMTools = {}
        self.ilmDoesNotExist = {}
        self.reworks = {}

    def processExcel(self, excelPath):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(r'{}'.format(excelPath))        
        
        self.transferToEDHR(wb, excel, excelPath)

        wb.Close(True)
        
        excel.Quit()

    def getStartOfMESRowNum(self, edhrWs):
        mesSectionFound = False
        row = 1

        while not mesSectionFound:
            if edhrWs.Cells(row, 3).Value == "6. MES Transaction Details":
                mesSectionFound = True
                break
            row += 1

        return row + 3

    def transferToEDHR(self, edhrWb, excel, excelPath):
        edhrWs = edhrWb.Worksheets["eDHR - Instrument - Detail"]

        edhrCheckWb = excel.Workbooks.Open(r'{}'.format(self.edhrCheckExcelPath))        
        edhrCheckWs = edhrCheckWb.Worksheets["eDHR - Instrument - Detail"]

        # Transfers eDHR to eDHR_check
        edhrCheckCurrentRow = 2
        edhrCurrentRow = self.getStartOfMESRowNum(edhrWs)
        print("start of mes section is at row {}".format(edhrCurrentRow))
        endOfEdhr = False

        taskName = "XXXXX"
        reworkID = 0

        while not endOfEdhr:
            stepDescriptionCell = edhrWs.Cells(edhrCurrentRow, 5).Value
            taskNameCell = edhrWs.Cells(edhrCurrentRow, 8).Value

            if edhrWs.Cells(edhrCurrentRow, 3).Value == None:
                endOfEdhr = True
                break
            
            if taskNameCell[:len(taskName)] == taskName:
                # Copy row from edhr to edhrcheck
                edhrCheckWsRangeStr = "C" + str(edhrCheckCurrentRow) + ":AT" + str(edhrCheckCurrentRow)
                edhrWsRangeStr = "C" + str(edhrCurrentRow) + ":AT" + str(edhrCurrentRow)
                edhrCheckWs.Range(edhrCheckWsRangeStr).Value = edhrWs.Range(edhrWsRangeStr).Value

            
                # Add tool to dictionary
                toolName = edhrWs.Cells(edhrCurrentRow, 13).Value[:-5] 
                toolNum = edhrWs.Cells(edhrCurrentRow, 33).Value
            
                self.checkToolRow(edhrCheckWs, edhrCheckCurrentRow, toolName, toolNum)
                edhrCheckCurrentRow += 1

                if not toolName in self.tools:
                    self.tools[toolName] = []

                if toolNum != "N/A":
                    self.tools[toolName].append(toolNum)

            if stepDescriptionCell and "Rework" in stepDescriptionCell:
                edhrCheckCurrentRow += 1

                # Copy row from edhr to edhrcheck
                edhrCheckWsRangeStr = "C" + str(edhrCheckCurrentRow) + ":AT" + str(edhrCheckCurrentRow)
                edhrWsRangeStr = "C" + str(edhrCurrentRow) + ":AT" + str(edhrCurrentRow)
                edhrCheckWs.Range(edhrCheckWsRangeStr).Value = edhrWs.Range(edhrWsRangeStr).Value

                if not reworkID in self.reworks:
                    self.reworks[reworkID] = []

                if taskNameCell != "N/A" and not taskNameCell in self.reworks[reworkID]:
                    self.reworks[reworkID].append(taskNameCell)

                if edhrWs.Cells(edhrCurrentRow, 33).Value == "Exit":
                    reworkID += 1

            edhrCurrentRow += 1

        edhrCheckWb.Save()
        edhrCheckWb.Close(True)

        self.printDictionaries()

    def checkToolRow(self, edhrWs, row, toolName, toolNum):
        status = edhrWs.Cells(row, 50).Value
        
        # if status:
        #     print("status: " + status)

        if status == "Incorrect ILM#":
            if not toolName in self.incorrectILMTools:
                    self.incorrectILMTools[toolName] = []

            if toolNum != "N/A":
                self.incorrectILMTools[toolName].append(toolNum)

        elif status == "ILM# not exist. Update Cal table if needed":
            if not toolName in self.ilmDoesNotExist:
                    self.ilmDoesNotExist[toolName] = []

            if toolNum != "N/A":
                self.ilmDoesNotExist[toolName].append(toolNum)

    def getILMs(self):
        return self.tools, self.incorrectILMTools, self.ilmDoesNotExist

    def getToolsWithoutILMs(self):
        toolsWithoutILMs = []

        for tool in self.tools.keys():
            if not self.tools[tool]:
                toolsWithoutILMs.append(tool)
            
        return toolsWithoutILMs

    def printDictionaries(self):

        for key in self.tools.keys():
            string = key + ": {}".format(self.tools[key])
            print(string)

        print("\nIncorrect:")

        for key in self.incorrectILMTools.keys():
            string = key + ": {}".format(self.incorrectILMTools[key])
            print(string)

        print("\nDoes not exist:")

        for key in self.ilmDoesNotExist.keys():
            string = key + ": {}".format(self.ilmDoesNotExist[key])
            print(string)

        print("\nReworks:")

        for key in self.reworks.keys():
            string = str(key) + ": {}".format(self.reworks[key])
            print("\t" + string + "\n")

        print("\n")