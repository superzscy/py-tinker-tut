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


path_summary_sheet_var = StringVar()
path_summary_sheet_var.set("NA")

summary_sheet_label_var = StringVar()
summary_sheet_label_var.set("NA")

frame_summary = Frame(root)
frame_summary.pack(fill=BOTH, expand=True)

btn_choose_summary_sheet = Button(
    frame_summary,
    text="选择汇总表",
    command=lambda: open_openfilename_dialog(path_summary_sheet_var),
    padding=15,
)
btn_choose_summary_sheet.grid(row=0, column=0, padx=5, pady=5)

label_summary_sheet = Label(
    frame_summary,
    textvariable=path_summary_sheet_var,
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

path_raw_sheet_var = StringVar()
path_raw_sheet_var.set("NA")

btn_choose_raw_sheet = Button(
    frame_raw,
    text="选择原始数据表",
    command=lambda: open_openfilename_dialog(path_raw_sheet_var),
    padding=15,
)
btn_choose_raw_sheet.grid(row=0, column=0, padx=5, pady=5)
label_raw_sheet = Label(
    frame_raw,
    textvariable=path_raw_sheet_var,
    padding=(50, 10),
    font="Arial 14 bold",
    background="yellow",
)
label_raw_sheet.grid(row=0, column=1, padx=5, pady=5)

raw_sheet_label = Label(frame_raw, text="工作表名")
raw_sheet_label.grid(row=1, column=0, padx=5, pady=5)
raw_sheet_entry = Entry(frame_raw, width=30)
raw_sheet_entry.grid(row=1, column=1, padx=5, pady=5)


def start_process(path_raw_sheet_var, path_summary_sheet_var):
    # raw_sheet_workbook = load_workbook(path_raw_sheet_var.get())

    # # 选择工作表，这里假设你要读取的是第一个工作表，也可以通过名称选择
    # sheet = raw_sheet_workbook["Sheet1"]

    # # 遍历第一列的单元格（注意，行和列的索引都是从1开始的）
    # for row in sheet.iter_rows(min_row=1, max_col=1, values_only=False):
    #     # row 是一个元组，包含第一列的单元格对象
    #     # row[0] 是第一列的单元格对象
    #     cell_value = row[0].value
    #     print(cell_value)  # 打印单元格的值

    # 如果你只想获取第一列的所有值，可以使用列表推导式
    # first_column_values = [cell.value for cell in sheet['A']]
    # print(first_column_values)

    summary_sheet_workbook = load_workbook(path_summary_sheet_var.get())

    # 选择工作表，这里假设你要读取的是第一个工作表，也可以通过名称选择
    summary_sheet = summary_sheet_workbook[summary_sheet_label_var.get()]

    encountered_valued_data = False
    for col in summary_sheet["D"]:
        d_column_value = col.value
        if d_column_value is None:
            if encountered_valued_data:
                break
        else:
            encountered_valued_data = True
            print(d_column_value)


btn_process = Button(
    root,
    text="开始汇总",
    command=lambda: start_process(path_raw_sheet_var, path_summary_sheet_var),
    padding=15,
)
btn_process.pack()

root.mainloop()
