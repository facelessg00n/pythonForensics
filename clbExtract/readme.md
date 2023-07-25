# Cellebrite Contact Extractor

---
Extracts contacts from Cellebrite formatted Excel files. The data in these files is nested within the column of an excel file which can cause issues when analysing them with third party tools.

This tool exports contacts on a per app basis into flat .CSV files for use with third party analysis tools. This was built to handle Excel files as this is typically what analysts will receive unless they have received a 'reader' file.

## Usage

This folder contains 2 python scripts. One is an optional GUI if you wish to build this into a portable .exe fi.

These instruction assume you are utilising *VS Code* (<https://code.visualstudio.com/>) and have a Python environment setup.

Download the contents of this folder and open the folder in VS code.

Create and activate a virtual environment.

<https://code.visualstudio.com/docs/python/environments>

install the requirements packages, this will include the tools to turn this into a portable exe.

`pip install -r .\requirements.txt`

The standalone script may then be run from the command line.

options:

- -h show this help message and exit
- -f path to the input file
- -b process all files in the working directory
- -p add data provenance from one of the pre approved items

Place the Excel files in the folder where the script is located to process the files in bulk.

## Building the exe

A portable exe can be build utilising PyInstaller.

The exe must be built on the same OS it is intended to be run on or it will not work. For example if you intend to build this for use on Windows machines it must be built on a windows machine.

The resulting exe will be locayed in the /dist folder of the working directory after it has been built.

### **With GUI**

`pyinstaller --onefile .\clbExtractGUI.py`

### **Without GUI**

`pyinstaller --onefile .\clbExtract.py`

---

## Current known issues

- Native contacts does not currently export email addresses
- Depending on which version of Cellebrite was used, or what type of extraction was perfomed some social media SUer ID's may not be available in the Excel files.

## Newtork Analysis tools

----
**Constellation**

<https://www.constellation-app.com/>

**Maltego**

<https://www.maltego.com/>
