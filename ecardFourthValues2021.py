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


pathCards = "4th-grading-cards"
infosExcel = "ecard-infos.xlsx"
cardTemplate = "ecard-template.xlsx"


# always keep same length
selectedEndCol = [11, 14, 14, 14, 14, 31]
pasteStartRow = [2, 2, 3, 4, 5, 2]
templateSheets = ["Infos", "Grades", "Grades", "Grades", "Grades", "Values"]


if __name__ == '__main__':
    sheetsCard = ecardFirst2021.loadSheets(infosExcel)
    ecardFirst2021.makeCard(pathCards, sheetsCard, cardTemplate, selectedEndCol, pasteStartRow, templateSheets)
    print("Done!")