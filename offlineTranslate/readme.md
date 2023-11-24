# Offline Translation

---

Many forensic tools such as Cellebrite Physical Analyzer have inbuilt translation offerings however my experience shows they can be slow or unreliable. As an offline translation option is often required I began to seek other means of translation. Enter LibreTranslate, an self hosted Machine translation API.

<https://github.com/LibreTranslate/LibreTranslate>

## Installation

Installation options will depend on your enviroment however to test the proof of concept LibreTrtanslate can be installed with the following command from an internet conneted machine.

    `pip install libretranslate`

The server can then be started on localhost with the following command. On first run it will pul down the language packages. The machine can then be taken offline.

    `libretranslate`
---

## Script usage

The Python script loads the specified Excel file and looks for a column named 'Messages' as per the Magnet AXIOM formatted excel sheets. At this time it will only handle Excel documents with a single sheet.

In reality it will take any Excel spreadsheet with a Column named messages.

- Auto detection of language is much faster however not as accurate. If you know the language it is best to select one of the language codes manually. To retrieve the language codes run `bulk_translate.py -g` and the available languages from the server will be listed.
- The generated CSV files may not open in Microsoft Excel however will open in LibreOffice Calc. It will however also attempt to output Excel files.
- Defaults to English translation but other languages are possible

Example usage

    Auto Detect
    python3 bulk_translate.py -f excel.xlsx
    
    Manually Select language
    python3 bulk_translate.py -f excel.xlsx -l zh

![screenshot](https://github.com/facelessg00n/pythonForensics/blob/main/offlineTranslate/images/offlineTranslate.jpg)
