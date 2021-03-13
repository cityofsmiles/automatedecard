#!/usr/bin/python  

import sys
import subprocess
import importlib.util

packages = ['openpyxl']
for package_name in packages:
    spec = importlib.util.find_spec(package_name)
    if spec is None:
        print("Installing python packages. Please wait...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package_name]) 

import openpyxl
import os
import ecardFirst2021


pathCards = "2nd-grading-cards"
infosExcel = "ecard-infos.xlsx"
cardTemplate = "ecard-template.xlsx"
infosSheets = ["Infos-Card-Male", "Infos-Card-Female", "1st-Summary-Male", "1st-Summary-Female", "2nd-Summary-Male", "2nd-Summary-Female"]
sheetsCard = ["infosMale", "infosFemale", "firstGradesMale", "firstGradesFemale", "secondGradesMale", "secondGradesFemale"]

# always keep same length
selectedEndCol = [12, 21, 21]
pasteStartRow = [2, 2, 3]
templateSheets = ["Infos", "Grades", "Grades"]
pasteSheet = ["infoSheet", "gradeSheet", "gradeSheet"]
addSheetNum = [0, 2, 4]


if __name__ == '__main__':
    ecardFirst2021.loadSheets(infosExcel, infosSheets, sheetsCard)
    ecardFirst2021.makeCard(pathCards, sheetsCard, cardTemplate, selectedEndCol, pasteStartRow, templateSheets, pasteSheet, addSheetNum)
    print("Done!")