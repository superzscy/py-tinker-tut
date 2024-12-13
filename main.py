from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from openpyxl import load_workbook

root = Tk()


def open_openfilename_dialog(var):
    ftypes = [("xlsx", ".xlsx")]
    var1 = StringVar()
    openfilename = filedialog.askopenfilename(
        parent=root,
        filetypes=ftypes,
        title="New title for open file name dialog box",
        typevariable=var1,
    )
    var.set(f"{openfilename}")


summary_sheet_path_var = StringVar()
summary_sheet_path_var.set("NA")
summary_sheet_label_var = StringVar()
summary_sheet_label_var.set("集采第九批内部统计使用")

frame_summary = Frame(root)
frame_summary.pack(fill=BOTH, expand=True)

btn_choose_summary_sheet = Button(
    frame_summary,
    text="选择汇总表",
    command=lambda: open_openfilename_dialog(summary_sheet_path_var),
    padding=15,
)
btn_choose_summary_sheet.grid(row=0, column=0, padx=5, pady=5)

label_summary_sheet = Label(
    frame_summary,
    textvariable=summary_sheet_path_var,
    padding=(50, 10),
    font="Arial 14 bold",
    background="yellow",
)
label_summary_sheet.grid(row=0, column=1, padx=5, pady=5)

summary_sheet_label = Label(frame_summary, text="工作表名")
summary_sheet_label.grid(row=1, column=0, padx=5, pady=5)
summary_sheet_entry = Entry(
    frame_summary, width=30, textvariable=summary_sheet_label_var
)
summary_sheet_entry.grid(row=1, column=1, padx=5, pady=5)

# 创建一个Frame作为分隔符
separator = Frame(root, height=2, relief=SUNKEN)
separator.pack(fill=X, padx=10, pady=10)  # 水平填充，并设置一些内边距


frame_raw = Frame(root)
frame_raw.pack(fill=BOTH, expand=True)

raw_sheet_path_var = StringVar()
raw_sheet_path_var.set("NA")
raw_sheet_label_var = StringVar()
raw_sheet_label_var.set("Sheet1")

btn_choose_raw_sheet = Button(
    frame_raw,
    text="选择原始数据表",
    command=lambda: open_openfilename_dialog(raw_sheet_path_var),
    padding=15,
)
btn_choose_raw_sheet.grid(row=0, column=0, padx=5, pady=5)
label_raw_sheet = Label(
    frame_raw,
    textvariable=raw_sheet_path_var,
    padding=(50, 10),
    font="Arial 14 bold",
    background="yellow",
)
label_raw_sheet.grid(row=0, column=1, padx=5, pady=5)

raw_sheet_label = Label(frame_raw, text="工作表名")
raw_sheet_label.grid(row=1, column=0, padx=5, pady=5)
raw_sheet_entry = Entry(frame_raw, width=30, textvariable=raw_sheet_label_var)
raw_sheet_entry.grid(row=1, column=1, padx=5, pady=5)


def start_process(raw_sheet_path_var, summary_sheet_path_var):
    summary_sheet_workbook = load_workbook(summary_sheet_path_var.get())
    summary_sheet = summary_sheet_workbook[summary_sheet_label_var.get()]

    summary_item_number_dict = {}
    summary_sheet_item_col = "D"
    summary_sheet_item_row = 4
    cur_row = 0
    for col in summary_sheet[summary_sheet_item_col]:
        cur_row += 1
        if cur_row < summary_sheet_item_row:
            continue

        item_name = col.value
        if item_name is None:
            break
        else:
            if item_name not in summary_item_number_dict:
                summary_item_number_dict[item_name] = 0
            else:
                print(f"!! 汇总表里有重复的 !! {item_name}")

    print("汇总表")
    cur_row = 0
    for key, value in summary_item_number_dict.items():
        cur_row += 1
        print(f"{cur_row}: {key} {value}")
    print("")

    #######
    raw_sheet_workbook = load_workbook(raw_sheet_path_var.get())
    raw_sheet = raw_sheet_workbook[raw_sheet_label_var.get()]

    raw_item_number_dict = {}
    summary_sheet_item_col_num = 3  # "D"
    summary_sheet_number_col_num = 5  # "F"
    summary_sheet_item_row = 5
    cur_row = 0
    for col in raw_sheet:
        cur_row += 1
        if cur_row < summary_sheet_item_row:
            continue

        item_name = col[summary_sheet_item_col_num].value
        if item_name is None:
            break
        else:
            if "非中选" not in item_name:
                if item_name not in raw_item_number_dict:
                    raw_item_number_dict[item_name] = col[
                        summary_sheet_number_col_num
                    ].value
                else:
                    raw_item_number_dict[item_name] = (
                        raw_item_number_dict[item_name]
                        + col[summary_sheet_number_col_num].value
                    )

    print("统计表")
    cur_row = 0
    for key, value in raw_item_number_dict.items():
        cur_row += 1
        print(f"{cur_row}: {key} {value}")


btn_process = Button(
    root,
    text="开始汇总",
    command=lambda: start_process(raw_sheet_path_var, summary_sheet_path_var),
    padding=15,
)
btn_process.pack()

root.mainloop()
