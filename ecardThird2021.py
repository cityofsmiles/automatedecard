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


pathCards = "3rd-grading-cards"
infosExcel = "ecard-infos.xlsx"
cardTemplate = "ecard-template.xlsx"
infosSheets = ["Infos-Card-Male", "Infos-Card-Female", "1st-Summary-Male", "1st-Summary-Female", "2nd-Summary-Male", "2nd-Summary-Female", "3rd-Summary-Male", "3rd-Summary-Female"]


# always keep same length
selectedEndCol = [12, 21, 21, 21]
pasteStartRow = [2, 2, 3, 4]
templateSheets = ["Infos", "Grades", "Grades", "Grades"]
pasteSheet = ["infoSheet", "gradeSheet", "gradeSheet", "gradeSheet"]


if __name__ == '__main__':
    sheetsCard = ecardFirst2021.loadSheets(infosExcel, infosSheets)
    ecardFirst2021.makeCard(pathCards, sheetsCard, cardTemplate, infosSheets, selectedEndCol, pasteStartRow, templateSheets, pasteSheet)
    print("Done!")