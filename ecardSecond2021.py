#!/usr/bin/python  



import sys
import subprocess
import importlib.util

packages = ['openpyxl', 'pandas', 'xlsx2csv']
for package_name in packages:
    spec = importlib.util.find_spec(package_name)
    if spec is None:
        print("Installing python packages. Please wait.")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package_name]) 
 
import openpyxl
import pandas as pd
import os
import shutil
import ecardFirst2021

pathCards = "2nd-grading-cards"
tempPath = "temp"
infosExcel = "ecard-infos.xlsx"
cardTemplate = "ecard-template.xlsx"
infosSheets = ["Infos-Card-Male", "Infos-Card-Female", "1st-Summary-Male", "1st-Summary-Female", "2nd-Summary-Male", "2nd-Summary-Female"]
sheetsCard = ["infosMale", "infosFemale", "firstGradesMale", "firstGradesFemale", "secondGradesMale", "secondGradesFemale"]
firstGradesDFIndex = [2, 3]
secondGradesDFIndex = [4, 5]
dfList = []


def createCard():
    os.makedirs(pathCards, exist_ok=True)
    for i in range(0, 2):
        for k in range(0, len(dfList[i])):
            studentName = dfList[i].iloc[k,1]
            studentCode = dfList[i].iloc[k,0]
            fileName = studentCode + "-" + studentName + ".xlsx"
            print("Creating file:", fileName)
            template = openpyxl.load_workbook(cardTemplate) 
            infoSheet = template["Infos"]
            gradeSheet = template["Grades"]
            studentRow = k + 2
            selectedInfoRow = ecardFirst2021.copyRange(1, studentRow, 11, studentRow, sheetsCard[i]) 
            pasteInfoRow = ecardFirst2021.pasteRange(1, 2, 11, 2, infoSheet, selectedInfoRow)
            selectedFirstGradeRow = ecardFirst2021.copyRange(1, studentRow, 21, studentRow, sheetsCard[firstGradesDFIndex[i]]) 
            pasteFirstGradeRow = ecardFirst2021.pasteRange(1, 2, 21, 2, gradeSheet, selectedFirstGradeRow)
            selectedSecondGradeRow = ecardFirst2021.copyRange(1, studentRow, 21, studentRow, sheetsCard[secondGradesDFIndex[i]]) 
            pasteSecondGradeRow = ecardFirst2021.pasteRange(1, 3, 21, 3, gradeSheet, selectedSecondGradeRow)
            template.save(os.path.join(pathCards,fileName))
    


if __name__ == '__main__':
    ecardFirst2021.readSheets(infosExcel, tempPath, infosSheets, dfList, sheetsCard)
    createCard()
    print("Done!")