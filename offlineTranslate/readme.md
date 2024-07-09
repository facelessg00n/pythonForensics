# Offline Translation

---

Many forensic tools have inbuilt translation offerings however my experience shows they can be slow or unreliable. As an offline translation option is often required I began to seek other means of translation. Enter LibreTranslate, an self hosted Machine translation API.

<https://github.com/LibreTranslate/LibreTranslate>

## Installation

Installation options will depend on your environment however to test the proof of concept LibreTranslate can be installed with the following command from an internet connected machine.

    `pip install libretranslate`

The server can then be started on localhost with the following command. On first run it will pull down the language packages. The machine can then be taken offline.

    `libretranslate`
---

## Modification

You may need to change the `serverURL = "http://localhost:5000"` value to match where you libretranslate instance is hosted

## Script usage

The Python script loads the specified Excel file and looks for a column named 'Messages' as per the Magnet AXIOM formatted excel sheets. At this time it will only handle Excel documents with a single sheet.

In reality it will take any Excel spreadsheet with a Column named messages.

- Auto detection of language is much faster however not as accurate. If you know the language it is best to select one of the language codes manually. To retrieve the language codes run `python3 bulk_translate_v3 -g` and the available languages from the server will be listed.
- The generated CSV files may not open in Microsoft Excel however will open in LibreOffice Calc. It will however also attempt to output Excel files.
- Defaults to English translation but other languages are possible

Example usage

    Auto Detect
    python3 bulk_translate_v3.py -f excel.xlsx
    
    Manually Select language
    python3 python3 bulk_translate_v3 -f excel.xlsx -l zh

![screenshot](https://github.com/facelessg00n/pythonForensics/blob/main/offlineTranslate/images/offlineTranslate.jpg)

## Other usage

options:

  -h, --help            show this help message and exit

  -f INPUTFILEPATH, --file INPUTFILEPATH
                        Path to Excel File

  -s TRANSLATIONSERVER, --server TRANSLATIONSERVER
                        Address of translation server if not localhost or hardcoded

  -l {}, --language {}  Language code for input text - optional but can greatly improve accuracy

  -e {Chats,Instant Messages}, --excelSheet {Chats,Instant Messages}
                        Sheet name within Excel file to be translated

  -c, --isCellebrite    If file originated from Cellebrite, header starts at 1, and message column is called 'Body'
  
  -g, --getlangs        Get supported language codes and names from server

## Building the exe

A portable exe can be build utilising PyInstaller.

The exe must be built on the same OS it is intended to be run on or it will not work. For example if you intend to build this for use on Windows machines it must be built on a windows machine.

The resulting exe will be located in the /dist folder of the working directory after it has been built.

### **With GUI**

`pyinstaller --onefile .\translateGUI.py`

### **Without GUI**

`pyinstaller --onefile .\bulk_translate_v3.py`
