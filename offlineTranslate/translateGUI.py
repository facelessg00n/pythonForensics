### GUI for Offline Translation
# Only tested with Windows
# Known display issues with OSX

# Changelog
# v0.2 - Update function names and handle Cellebrite formatted files.
#      - Language selection menu
# v0.1 - Initial concept

import bulk_translate_v3

import os
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

LIGHT_GREY = "#BEBFC7"
LIGHT_BLUE = "#307FE2"
DARK_BLUE = "#024DA1"
DARK_RED = "#FF5342"
FONT_1 = "Roboto Condensed"

isCellebrite = False
SERVER_CONNECTED = False


inputLanguages = [
    "auto",
    "en",
    "sq",
    "ar",
    "az",
    "bn",
    "bg",
    "ca",
    "zh",
    "zt",
    "cs",
    "da",
    "nl",
    "eo",
    "et",
    "fi",
    "fr",
    "de",
    "el",
    "he",
    "hi",
    "hu",
    "id",
    "ga",
    "it",
    "ja",
    "ko",
    "lv",
    "lt",
    "ms",
    "nb",
    "fa",
    "pl",
    "pt",
    "ro",
    "ru",
    "sr",
    "sk",
    "sl",
    "es",
    "sv",
    "tl",
    "th",
    "tr",
    "uk",
]

## _______________Functions live here___________________________________________________________


# Process a selected file
def get_selection():
    selected_file = lbox.curselection()
    print(lbox.get(selected_file))
    print(inputSheetMenu.get())
    bulk_translate_v3.loadAndTranslate(
        lbox.get(selected_file),
        inputLangMenu.get(),
        inputSheetMenu.get(),
        isCellebrite.get(),
    )


def inputComboSelection(event):
    selectedProvenance = inputSheetMenu.get()


def langComboSelection(event):
    selectedProvenance = inputLangMenu.get()
    # messagebox.showinfo(message=f"The Selected value is {selectedProvenance}",title='Selection')


### _____Create interface______________________________________________________________________

# Show list of Excel files in the current working directory
candidateFiles = os.listdir(os.getcwd())
file_list = []
for candidateFiles in candidateFiles:
    if candidateFiles.endswith(".xlsx"):
        file_list.append(candidateFiles)
fileListingDisplay = "\n".join(file_list)

# Test Connectivity
# bulk_translate_v3.serverCheck())
if bulk_translate_v3.serverCheck(bulk_translate_v3.serverURL) == "SERVER_OK":
    print("Connected to server")
    SERVER_CONNECTED = True
    serverButtonColour = LIGHT_BLUE
    serverStatus = "Online"
else:
    print("Server connection failed")
    SERVER_CONNECTED = False
    serverButtonColour = DARK_RED
    serverStatus = "Offline"

# Create box
root = Tk()
root.geometry("580x650")
root.minsize(458, 580)
root.maxsize(780, 780)
root.configure(bg=LIGHT_GREY)

prog_name = Label(
    text="Offline Translation",
    anchor=W,
    padx=10,
    pady=10,
    background=DARK_BLUE,
    width=480,
    font=(FONT_1, 20),
)
prog_name.pack()

sideFrame = Frame(master=root, width=100, height=100, bg=LIGHT_BLUE)
sideFrame.pack(fill=Y, side=LEFT)
sideFrame.pack()

servAdd = Label(
    text="Server Address: {} Server Status: {}".format(
        str(bulk_translate_v3.serverURL), serverStatus
    ),
    padx=10,
    pady=00,
    bg=serverButtonColour,
)
servAdd.pack()
# User instructions
prog_data = Label(
    text="For procesisng of files place this program in the folder\n containing your Excel files.",
    font=(FONT_1, 10),
    anchor=W,
    padx=5,
    pady=5,
    bg=LIGHT_GREY,
)
prog_data.pack()

app_data_heading = Label(sideFrame, text="   ", bg=LIGHT_BLUE, font=(FONT_1, 10))
app_data_heading.pack()

# app_data.pack()

## Show Auto located files
auto_locate_data = Label(
    text="{} candidate files located at path: \n{}".format(
        str(len(file_list)), str(os.getcwd())
    ),
    anchor=W,
    padx=10,
    pady=10,
    bg=LIGHT_GREY,
)
auto_locate_data.pack(pady=10, padx=10)

# Tick box if file is a Cellebrite file, the header in these files starts at 1
isCellebrite = IntVar()
c1 = Checkbutton(text="Cellebrite file?", variable=isCellebrite, onvalue=1, offvalue=0)
c1.pack()

# Select an input Datasheet
inputSheetName = Label(
    text="Input Sheet name if multiple sheets exist",
    padx=10,
    pady=10,
    bg=LIGHT_GREY,
)
inputSheetName.pack()

# Input sheet selection menu
inputSheetVar = StringVar()
inputSheetMenu = ttk.Combobox(
    values=bulk_translate_v3.inputSheets, textvariable=inputSheetVar, state="readonly"
)

inputSheetMenu.bind("<<ComboboxSelected>>", inputComboSelection)

# inputSheetMenu.set("Chats")
inputSheetMenu.pack(side="top")

# File selection label
filesLabel = Label(
    text="Select File",
    padx=10,
    pady=10,
    bg=LIGHT_GREY,
)

# ____________Language selection menu_______________________________________
inputLangName = Label(
    text="Input Language",
    padx=10,
    pady=10,
    bg=LIGHT_GREY,
)
inputLangName.pack()
langVar = StringVar()
inputLangMenu = ttk.Combobox(
    values=inputLanguages, textvariable=langVar, state="readonly"
)

inputLangMenu.bind("<<ComboboxSelected>>", langComboSelection)
inputLangMenu.set("auto")
inputLangMenu.pack(side="top")

# ____________File selction menu_______________________________________
filesLabel = Label(
    text="Select File",
    padx=10,
    pady=10,
    bg=LIGHT_GREY,
)
filesLabel.pack()

# Select file names
fNames = StringVar(value=fileListingDisplay)
lbox = Listbox(root, listvariable=fNames, height=5, width=200)
scroll_bar = Scrollbar(root)
scroll_bar.pack(side=RIGHT, fill=Y)
lbox.pack()
scroll_bar.config(command=lbox.yview)

### Buttons for processing selected files
processSelectedBtn = Button(
    root,
    text="Process Selected",
    command=get_selection,
    bg=LIGHT_GREY,
    padx=10,
)
processSelectedBtn.pack(side="top")

# Exit Program
exitBtn = Button(root, text="Exit", command=root.destroy, bg=LIGHT_GREY)
exitBtn.pack(side=TOP, pady=20, padx=10)

# Display version info
verLabel = Label(
    text="Version {}".format(str(bulk_translate_v3.__version__)),
    padx=10,
    pady=10,
    bg=LIGHT_GREY,
)
verLabel.pack()

root.mainloop()
