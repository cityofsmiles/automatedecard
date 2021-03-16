# Automated E-Card for High School Teachers
These are simple Python scripts for producing high school report cards.

## Installation
1. In Windows 10, launch the Microsoft Store and install Python. As of this writing, the latest available version is Python 3.9. To install in any Android smartphone, launch the Google Play Store to install the F-Droid app. Then launch the F-Droid app to install the Termux app. In the Termux app, install Python. Click [here](https://itrendbuzz.com/install-python-on-android-using-termux/) for a tutorial on Termux.

2. Click [here](https://github.com/cityofsmiles/Grade8Lessons/raw/assets/miscellaneous/ecard2021.zip) to download this repository as a zip file. Unzip the file to any folder you choose. This folder will be your working directory.

## Usage
1. Edit the data and grades in the *ecard-infos.xlsx* file.

2. Launch the Command Prompt app* or the Termux app and navigate to your working directory. For example, if your working directory is the Downloads folder, type the following in the command prompt or the Termux app and hit "Enter".
```
cd Downloads
```

3. Run** the *ecard2021.py* script by typing the following in the command prompt or the Termux app and hitting "Enter".
```
python ecard2021.py
```
The finished cards will be saved in the *Cards* folder in your working directory.

*Note: To open a windows command prompt, press the “Windows Key+R” to open a “Run” dialog box. Next, type in “cmd”, and then click “OK”. This opens a normal Command Prompt.

**Note: During the first run, the script will install the openpyxl package so it may take several minutes depending on your internet connection. Please be patient.
