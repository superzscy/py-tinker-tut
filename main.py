from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from openpyxl import load_workbook

root = Tk()
root.grid_anchor("center")


def open_openfilename_dialog(var):
    ftypes = [("all files", "*")]
    var1 = StringVar()
    openfilename = filedialog.askopenfilename(
        parent=root,
        filetypes=ftypes,
        title="New title for open file name dialog box",
        typevariable=var1,
    )
    var.set(f"{openfilename}")


path_raw_sheet_var = StringVar()
path_raw_sheet_var.set("NA")
path_summary_sheet_var = StringVar()
path_summary_sheet_var.set("NA")

label_raw_sheet = Label(
    root,
    textvariable=path_raw_sheet_var,
    padding=(50, 10),
    font="Arial 14 bold",
    background="yellow",
)
label_raw_sheet.grid(row=1, column=1, pady=5)
btn_choose_raw_sheet = Button(
    root,
    text="选择原始数据表",
    command=lambda: open_openfilename_dialog(path_raw_sheet_var),
    padding=15,
)
btn_choose_raw_sheet.grid(row=1, column=0, pady=5)

label_summary_sheet = Label(
    root,
    textvariable=path_summary_sheet_var,
    padding=(50, 10),
    font="Arial 14 bold",
    background="yellow",
)
label_summary_sheet.grid(row=2, column=1, pady=5)
btn_choose_summary_sheet = Button(
    root,
    text="选择汇总表",
    command=lambda: open_openfilename_dialog(path_summary_sheet_var),
    padding=15,
)
btn_choose_summary_sheet.grid(row=2, column=0, pady=5)

def start_process(path_raw_sheet_var, path_summary_sheet_var):
    print(path_raw_sheet_var.get(), path_summary_sheet_var.get())

btn_process = Button(
    root,
    text="开始汇总",
    command=lambda: start_process(path_raw_sheet_var, path_summary_sheet_var),
    padding=15,
)
btn_process.grid(row=3, column=0, pady=5)

root.mainloop()
