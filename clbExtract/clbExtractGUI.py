### GUI for Cellebrite File Flattener
# Only tested with Windows


# Changelog
# v0.1 - Initial concept

import clbExtract

import os
from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

LIGHT_GREY = "#BEBFC7"
LIGHT_BLUE = "#307FE2"
DARK_BLUE = "#024DA1"
FONT_1 = "Roboto Condensed"

# Auto locate list of files
x = os.listdir(os.getcwd())
file_list = []
for x in x:
    if x.endswith(".xlsx"):
        file_list.append(x)
y1 = "\n".join(file_list)

# list of handled apps
y2 = "\n".join(clbExtract.parsedApps)


## _____Functions live here_____
def process_all():
    print("Process all selected")
    clbExtract.bulkProcessor()


def select_file():
    filetypes = [("Excel Files", "*.xlsx")]

    filename = fd.askopenfile(
        title="Open a file",
        initialdir=os.listdir(os.getcwd()),
        filetypes=filetypes,
        multiple=False,
    )
    if filename:
        print(filename.name)
        showinfo(
            title="Selected File",
            message=filename.name,
        )
        clbExtract.processMetadata(filename.name)


def get_selection():
    selected_file = lbox.curselection()
    print(lbox.get(selected_file))
    clbExtract.processMetadata(lbox.get(selected_file))


### _____Create interface_____
root = Tk()
root.geometry("580x580")
root.minsize(458, 580)
root.maxsize(780, 780)
root.configure(bg=LIGHT_GREY)

prog_name = Label(
    text="Cellebrite Contact Extractor",
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

prog_data = Label(
    text="For bulk processing of files place this program in the folder\n containing your Cellebrite formatted Excel files. ",
    font=(FONT_1, 10),
    anchor=W,
    padx=10,
    pady=10,
    bg=LIGHT_GREY,
)
prog_data.pack()

app_data_heading = Label(
    sideFrame, text="Handled apps:", bg=LIGHT_BLUE, font=(FONT_1, 10)
)
app_data_heading.pack()
app_data = Label(sideFrame, text=y2, bg=LIGHT_BLUE, font=(FONT_1, 10))
app_data.pack()


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

fnames = StringVar(value=y1)
lbox = Listbox(root, listvariable=fnames, height=5, width=200)
scroll_bar = Scrollbar(root)
scroll_bar.pack(side=RIGHT, fill=Y)
lbox.pack()
scroll_bar.config(command=lbox.yview)


btn2 = Button(root, text="Process Selected", command=get_selection, bg=LIGHT_GREY)
btn2.pack(side="top")
btn3 = Button(root, text="Process all files", command=process_all, bg=LIGHT_GREY)
btn3.pack(side="top")


prog_data = Label(
    text="Manually select file a file to extract \n Output files will be located at: \n {}".format(
        str(os.getcwd())
    ),
    anchor=W,
    padx=10,
    pady=10,
    font=(FONT_1, 10),
    bg=LIGHT_GREY,
)
prog_data.pack()

btn = Button(root, text="Locate file", command=select_file, bg=LIGHT_GREY)
btn.pack(side=TOP, pady=10, padx=10)
# Exit Program
exitBtn = Button(root, text="Exit", command=root.destroy, bg=LIGHT_GREY)
exitBtn.pack(side=TOP, pady=20, padx=10)
# Display version info
verLabel = Label(
    text="Version {}\n".format(str(clbExtract.__version__)),
    padx=10,
    pady=10,
    bg=LIGHT_GREY,
)
verLabel.pack()


root.mainloop()
