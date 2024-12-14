from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from openpyxl import load_workbook
import os
import csv

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

summary_sheet_label_var = StringVar()
summary_sheet_label_var.set("集采第九批内部统计使用")
summary_sheet_entry = Entry(
    frame_summary, width=30, textvariable=summary_sheet_label_var
)
summary_sheet_entry.grid(row=1, column=1, padx=5, pady=5)

summary_sheet_item_start_row_label = Label(frame_summary, text="数据开始行号")
summary_sheet_item_start_row_label.grid(row=2, column=0, padx=5, pady=5)
summary_sheet_item_start_row_var = StringVar()
summary_sheet_item_start_row_var.set("4")
summary_sheet_item_start_row_entry = Entry(
    frame_summary, width=30, textvariable=summary_sheet_item_start_row_var
)
summary_sheet_item_start_row_entry.grid(row=2, column=1, padx=5, pady=5)

summary_sheet_item_name_col_label = Label(
    frame_summary, text="药品名列号(A为1, B为2...)"
)
summary_sheet_item_name_col_label.grid(row=3, column=0, padx=5, pady=5)
summary_sheet_item_name_col_var = StringVar()
summary_sheet_item_name_col_var.set("4")
summary_sheet_item_name_col_entry = Entry(
    frame_summary, width=30, textvariable=summary_sheet_item_name_col_var
)
summary_sheet_item_name_col_entry.grid(row=3, column=1, padx=5, pady=5)

summary_sheet_item_spec_col_label = Label(frame_summary, text="规格列号(A为1, B为2...)")
summary_sheet_item_spec_col_label.grid(row=4, column=0, padx=5, pady=5)
summary_sheet_item_spec_col_var = StringVar()
summary_sheet_item_spec_col_var.set("6")
summary_sheet_item_spec_col_entry = Entry(
    frame_summary, width=30, textvariable=summary_sheet_item_spec_col_var
)
summary_sheet_item_spec_col_entry.grid(row=4, column=1, padx=5, pady=5)

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


raw_sheet_item_start_row_label = Label(frame_raw, text="数据开始行号")
raw_sheet_item_start_row_label.grid(row=2, column=0, padx=5, pady=5)
raw_sheet_item_start_row_var = StringVar()
raw_sheet_item_start_row_var.set("5")
raw_sheet_item_start_row_entry = Entry(
    frame_raw, width=30, textvariable=raw_sheet_item_start_row_var
)
raw_sheet_item_start_row_entry.grid(row=2, column=1, padx=5, pady=5)

raw_sheet_item_name_col_label = Label(frame_raw, text="药品名列号(A为1, B为2...)")
raw_sheet_item_name_col_label.grid(row=3, column=0, padx=5, pady=5)
raw_sheet_item_name_col_var = StringVar()
raw_sheet_item_name_col_var.set("4")
raw_sheet_item_name_col_entry = Entry(
    frame_raw, width=30, textvariable=raw_sheet_item_name_col_var
)
raw_sheet_item_name_col_entry.grid(row=3, column=1, padx=5, pady=5)

raw_sheet_item_spec_col_label = Label(frame_raw, text="规格列号(A为1, B为2...)")
raw_sheet_item_spec_col_label.grid(row=4, column=0, padx=5, pady=5)
raw_sheet_item_spec_col_var = StringVar()
raw_sheet_item_spec_col_var.set("5")
raw_sheet_item_spec_col_entry = Entry(
    frame_raw, width=30, textvariable=raw_sheet_item_spec_col_var
)
raw_sheet_item_spec_col_entry.grid(row=4, column=1, padx=5, pady=5)

raw_sheet_item_num_col_label = Label(
    frame_raw, text="使用量(包装单位)列号(A为1, B为2...)"
)
raw_sheet_item_num_col_label.grid(row=5, column=0, padx=5, pady=5)
raw_sheet_item_num_col_var = StringVar()
raw_sheet_item_num_col_var.set("6")
raw_sheet_item_num_col_entry = Entry(
    frame_raw, width=30, textvariable=raw_sheet_item_num_col_var
)
raw_sheet_item_num_col_entry.grid(row=5, column=1, padx=5, pady=5)


def start_process(raw_sheet_path_var, summary_sheet_path_var):
    summary_sheet_workbook = load_workbook(summary_sheet_path_var.get())
    summary_sheet = summary_sheet_workbook[summary_sheet_label_var.get()]

    summary_item_number_dict = {}
    item_start_row = int(summary_sheet_item_start_row_var.get())
    item_name_col = int(summary_sheet_item_name_col_var.get()) - 1
    item_spec_col = int(summary_sheet_item_spec_col_var.get()) - 1

    cur_row = 0
    for col in summary_sheet:
        cur_row += 1
        if cur_row < item_start_row:
            continue

        item_name = col[item_name_col].value
        if item_name is None:
            break
        else:
            item_spec = col[item_spec_col].value
            item_tuple = (item_name, item_spec)
            if item_tuple not in summary_item_number_dict:
                summary_item_number_dict[item_tuple] = 0
            else:
                print(f"!! 汇总表里有重复的 !! {item_tuple}")

    #######
    raw_sheet_workbook = load_workbook(raw_sheet_path_var.get())
    raw_sheet = raw_sheet_workbook[raw_sheet_label_var.get()]

    raw_item_number_dict = {}
    item_start_row = int(raw_sheet_item_start_row_var.get())
    item_name_col = int(raw_sheet_item_name_col_var.get()) - 1
    item_spec_col = int(raw_sheet_item_spec_col_var.get()) - 1
    item_num_col = int(raw_sheet_item_num_col_var.get()) - 1

    cur_row = 0
    for col in raw_sheet:
        cur_row += 1
        if cur_row < item_start_row:
            continue

        item_name = col[item_name_col].value
        if item_name is None:
            break
        else:
            if "非中选" in item_name:
                continue

            item_spec = col[item_spec_col].value
            index = item_spec.find("*")
            if index != -1:
                item_spec = item_spec[:index]

            item_tuple = (item_name, item_spec)
            num = col[item_num_col].value
            if item_tuple not in raw_item_number_dict:
                raw_item_number_dict[item_tuple] = num
            else:
                raw_item_number_dict[item_tuple] += num

    print("统计表")
    cur_row = 0
    for key, value in raw_item_number_dict.items():
        cur_row += 1
        print(f"{cur_row}: {key} {value}")
    print("")

    print("汇总表")
    csv_data = []
    csv_data.append(["药品名", "规格", "使用量"])

    cur_row = 0
    for key, _ in summary_item_number_dict.items():
        cur_row += 1
        value = 0
        summary_item_name = key[0]
        summary_item_spec = key[1]

        for raw_key, raw_value in raw_item_number_dict.items():
            raw_item_name = raw_key[0]
            raw_item_spec = raw_key[1]
            if (
                summary_item_name in raw_item_name
                and raw_item_spec in summary_item_spec
            ):
                value = raw_value
                break

        print(f"{cur_row}: {key} {value}")
        csv_data.append([summary_item_name, summary_item_spec, value])

    source_file_path = summary_sheet_path_var.get()
    source_file_name_with_ext = os.path.basename(source_file_path)
    source_file_name_without_ext, _ = os.path.splitext(source_file_name_with_ext)
    generated_file_path = os.path.join(
        os.path.dirname(source_file_path),
        source_file_name_without_ext + "_generated.csv",
    )
    with open(generated_file_path, mode="w", newline="", encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(csv_data)
    print(f"结果已写入到 {generated_file_path}")


btn_process = Button(
    root,
    text="开始汇总",
    command=lambda: start_process(raw_sheet_path_var, summary_sheet_path_var),
    padding=15,
)
btn_process.pack()

root.mainloop()
