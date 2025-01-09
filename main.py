from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from openpyxl import load_workbook
from tkinter import messagebox
import os
import errno
import csv

root = Tk()


def open_openfilename_dialog(var):
    ftypes = [("Excel files", ".xlsx")]
    var1 = StringVar()
    openfilename = filedialog.askopenfilename(
        parent=root,
        filetypes=ftypes,
        title="New title for open file name dialog box",
        typevariable=var1,
    )
    var.set(f"{openfilename}")


def allow_only_letters(event):
    """
    只允许在 Entry 小部件中输入字母。
    如果输入的字符不是字母，则将其删除。
    """
    # 获取当前 Entry 小部件中的文本
    current_text = event.widget.get()
    # 获取即将插入的字符（通过 event.char，但需要注意对于删除键等特殊处理）
    new_char = event.char

    # 对于删除键（Backspace）和回车键等特殊情况，允许它们通过
    if (
        new_char == "" or new_char == "\x08" or new_char == "\x0d" or new_char == "\x1b"
    ):  # \x08 是 Backspace，\x0d 是 Enter，\x1b 是 Esc
        return

    if len(current_text) > 0:
        return "break"

    # 检查新字符是否是字母（a-z 或 A-Z）
    if not new_char.isalpha():
        # 如果不是字母，则阻止输入（通过返回 'break'）
        return "break"


def allow_only_numbers(event):
    """
    只允许在 Entry 小部件中输入数字。
    如果输入的字符不是数字，则将其删除。
    """
    # 获取当前 Entry 小部件中的文本
    current_text = event.widget.get()
    # 获取即将插入的字符（通过 event.char）
    new_char = event.char

    # 检查是否按下了退格键（Backspace）或删除键（Delete）
    if (
        new_char == "" or new_char == "\x08" or new_char == "\x7f"
    ):  # \x08 是 Backspace，\x7f 是 Delete
        return

    if not new_char.isdigit():
        return "break"


def convert_letter_to_number(letter):
    if "a" <= letter <= "z":
        return ord(letter) - ord("a") + 1
    elif "A" <= letter <= "Z":
        return ord(letter) - ord("A") + 1
    else:
        # 如果输入不是字母，可以返回一个错误消息或特殊值
        return None  # 或者你可以抛出异常：raise ValueError(f"Input must be a letter: {letter}")


def show_message(title, message):
    # 创建一个根窗口（虽然它不会显示，但创建消息框时需要）
    root = Tk()
    root.withdraw()  # 隐藏根窗口
    if title == "":
        messagebox.showinfo("", message)
    else:
        messagebox.showerror(title, message)


summary_sheet_path_var = StringVar()
summary_sheet_path_var.set("")


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
summary_sheet_entry = Entry(frame_summary, textvariable=summary_sheet_label_var)
summary_sheet_entry.grid(row=1, column=1, padx=5, pady=5)

summary_sheet_item_start_row_label = Label(frame_summary, text="数据开始行号")
summary_sheet_item_start_row_label.grid(row=2, column=0, padx=5, pady=5)
summary_sheet_item_start_row_var = StringVar()
summary_sheet_item_start_row_var.set("4")
summary_sheet_item_start_row_entry = Entry(
    frame_summary, textvariable=summary_sheet_item_start_row_var
)
summary_sheet_item_start_row_entry.bind("<Key>", allow_only_numbers)
summary_sheet_item_start_row_entry.grid(row=2, column=1, padx=5, pady=5)

summary_sheet_item_name_col_label = Label(frame_summary, text="药品名列号")
summary_sheet_item_name_col_label.grid(row=3, column=0, padx=5, pady=5)
summary_sheet_item_name_col_var = StringVar()
summary_sheet_item_name_col_var.set("D")
summary_sheet_item_name_col_entry = Entry(
    frame_summary, textvariable=summary_sheet_item_name_col_var
)
summary_sheet_item_name_col_entry.bind("<Key>", allow_only_letters)
summary_sheet_item_name_col_entry.grid(row=3, column=1, padx=5, pady=5)

summary_sheet_item_spec_col_label = Label(frame_summary, text="规格列号")
summary_sheet_item_spec_col_label.grid(row=4, column=0, padx=5, pady=5)
summary_sheet_item_spec_col_var = StringVar()
summary_sheet_item_spec_col_var.set("F")
summary_sheet_item_spec_col_entry = Entry(
    frame_summary, textvariable=summary_sheet_item_spec_col_var
)
summary_sheet_item_spec_col_entry.bind("<Key>", allow_only_letters)
summary_sheet_item_spec_col_entry.grid(row=4, column=1, padx=5, pady=5)

summary_sheet_code_col_label = Label(frame_summary, text="药品编码")
summary_sheet_code_col_label.grid(row=5, column=0, padx=5, pady=5)
summary_sheet_code_col_var = StringVar()
summary_sheet_code_col_var.set("F")
summary_sheet_code_col_entry = Entry(
    frame_summary, textvariable=summary_sheet_code_col_var
)
summary_sheet_code_col_entry.bind("<Key>", allow_only_letters)
summary_sheet_code_col_entry.grid(row=5, column=1, padx=5, pady=5)

# 创建一个Frame作为分隔符
separator = Frame(root, height=2, relief=SUNKEN)
separator.pack(fill=X, padx=10, pady=10)  # 水平填充，并设置一些内边距


frame_raw = Frame(root)
frame_raw.pack(fill=BOTH, expand=True)

raw_sheet_path_var = StringVar()
raw_sheet_path_var.set("")
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
    frame_raw, textvariable=raw_sheet_item_start_row_var
)
raw_sheet_item_start_row_entry.bind("<Key>", allow_only_numbers)
raw_sheet_item_start_row_entry.grid(row=2, column=1, padx=5, pady=5)

raw_sheet_item_name_col_label = Label(frame_raw, text="药品名列号")
raw_sheet_item_name_col_label.grid(row=3, column=0, padx=5, pady=5)
raw_sheet_item_name_col_var = StringVar()
raw_sheet_item_name_col_var.set("D")
raw_sheet_item_name_col_entry = Entry(
    frame_raw, textvariable=raw_sheet_item_name_col_var
)
raw_sheet_item_name_col_entry.bind("<Key>", allow_only_letters)
raw_sheet_item_name_col_entry.grid(row=3, column=1, padx=5, pady=5)

raw_sheet_item_spec_col_label = Label(frame_raw, text="规格列号")
raw_sheet_item_spec_col_label.grid(row=4, column=0, padx=5, pady=5)
raw_sheet_item_spec_col_var = StringVar()
raw_sheet_item_spec_col_var.set("E")
raw_sheet_item_spec_col_entry = Entry(
    frame_raw, textvariable=raw_sheet_item_spec_col_var
)
raw_sheet_item_spec_col_entry.bind("<Key>", allow_only_letters)
raw_sheet_item_spec_col_entry.grid(row=4, column=1, padx=5, pady=5)

raw_sheet_item_num_col_label = Label(frame_raw, text="使用量(包装单位)列号")
raw_sheet_item_num_col_label.grid(row=5, column=0, padx=5, pady=5)
raw_sheet_item_num_col_var = StringVar()
raw_sheet_item_num_col_var.set("F")
raw_sheet_item_num_col_entry = Entry(frame_raw, textvariable=raw_sheet_item_num_col_var)
raw_sheet_item_num_col_entry.bind("<Key>", allow_only_letters)
raw_sheet_item_num_col_entry.grid(row=5, column=1, padx=5, pady=5)

raw_sheet_code_col_label = Label(frame_raw, text="药品编码")
raw_sheet_code_col_label.grid(row=6, column=0, padx=5, pady=5)
raw_sheet_code_col_var = StringVar()
raw_sheet_code_col_var.set("F")
raw_sheet_code_col_entry = Entry(
    frame_raw, textvariable=raw_sheet_code_col_var
)
raw_sheet_code_col_entry.bind("<Key>", allow_only_letters)
raw_sheet_code_col_entry.grid(row=6, column=1, padx=5, pady=5)

def start_process(raw_sheet_path_var, summary_sheet_path_var):
    # 参数检测
    summary_sheet_path_str = summary_sheet_path_var.get()
    summary_sheet_label_str = summary_sheet_label_var.get()
    summary_sheet_item_start_row_str = summary_sheet_item_start_row_var.get()
    summary_sheet_item_spec_col_str = summary_sheet_item_spec_col_var.get()
    summary_sheet_item_name_col_str = summary_sheet_item_name_col_var.get()
    summary_sheet_code_col_str = summary_sheet_code_col_var.get()

    raw_sheet_path_str = raw_sheet_path_var.get()
    raw_sheet_label_str = raw_sheet_label_var.get()
    raw_sheet_item_start_row_str = raw_sheet_item_start_row_var.get()
    raw_sheet_item_spec_col_str = raw_sheet_item_spec_col_var.get()
    raw_sheet_item_name_col_str = raw_sheet_item_name_col_var.get()
    raw_sheet_item_num_col_str = raw_sheet_item_num_col_var.get()
    raw_sheet_code_col_str = raw_sheet_code_col_var.get()

    args = {
        "汇总表路径": summary_sheet_path_str,
        "汇总表工作表名": summary_sheet_label_str,
        "汇总表数据开始行号": summary_sheet_item_start_row_str,
        "汇总表规格列号": summary_sheet_item_spec_col_str,
        "汇总表药品名列号": summary_sheet_item_name_col_str,
        "原始数据表路径": raw_sheet_path_str,
        "原始数据表工作表名": raw_sheet_label_str,
        "原始数据表数据开始行号": raw_sheet_item_start_row_str,
        "原始数据表规格列号": raw_sheet_item_spec_col_str,
        "原始数据表药品名列号": raw_sheet_item_name_col_str,
    }

    for k, v in args.items():
        if v == "":
            show_message("错误", f"{k} 参数错误, 请检查!")
            return

    summary_sheet_workbook = load_workbook(summary_sheet_path_str)
    if summary_sheet_label_str not in summary_sheet_workbook.sheetnames:
        show_message(
            "错误", f"汇总表工作表名 {summary_sheet_label_str} 不存在, 请检查!"
        )
        return

    raw_sheet_workbook = load_workbook(raw_sheet_path_str)
    if raw_sheet_label_str not in raw_sheet_workbook.sheetnames:
        show_message(
            "错误", f"原始数据表工作表名 {raw_sheet_label_str} 不存在, 请检查!"
        )
        return

    summary_sheet = summary_sheet_workbook[summary_sheet_label_str]
    raw_sheet = raw_sheet_workbook[raw_sheet_label_str]

    summary_item_number_dict = {}
    item_start_row = int(summary_sheet_item_start_row_str)
    item_name_col = convert_letter_to_number(summary_sheet_item_name_col_str) - 1
    item_spec_col = convert_letter_to_number(summary_sheet_item_spec_col_str) - 1
    code_col = convert_letter_to_number(summary_sheet_code_col_str) - 1

    # 结果文件可写状态检测
    source_file_path = summary_sheet_path_var.get()
    source_file_name_with_ext = os.path.basename(source_file_path)
    source_file_name_without_ext, _ = os.path.splitext(source_file_name_with_ext)
    generated_file_path = os.path.join(
        os.path.dirname(source_file_path),
        source_file_name_without_ext + "_generated.csv",
    )

    try:
        with open(generated_file_path, "w") as file:
            pass
    except OSError as e:
        # 捕获 OSError 异常，这通常发生在文件被占用或其他 I/O 错误时
        if e.errno == errno.EACCES:
            show_message(
                "错误",
                f"汇总结果文件 [{generated_file_path}] 无法被写入。可能是文件正在被另一个程序使用。请先关闭.",
            )
        else:
            show_message(
                "错误",
                f"汇总结果文件 [{generated_file_path}] 无法被写入。错误：{e}, 错误码:{e.errno}",
            )
        return
    except Exception as e:
        show_message(
            "错误",
            f"汇总结果文件 [{generated_file_path}] 无法被写入。发生了一个意外错误：{e}",
        )
        return

    # 汇总表据表
    cur_row = 0
    for col in summary_sheet:
        cur_row += 1
        if cur_row < item_start_row:
            continue

        item_name = col[item_name_col].value
        if item_name is None:
            break
        else:
            index = item_name.find("*")
            if index != -1:
                item_name = item_name[:index]
            item_name = item_name.replace("▲", "")
            item_name = item_name.replace("◆", "")
            item_name = item_name.replace("◆", "")
            item_name = item_name.replace("●", "")
            item_name = item_name.split("|")[0]
            item_name = item_name.split(" ")[0]

            # item_spec = col[item_spec_col].value
            # print(item_name, item_spec)
            # item_spec = item_spec.replace("：", ":")

            code = col[code_col].value
            item_tuple = (item_name, code)
            if item_tuple not in summary_item_number_dict:
                summary_item_number_dict[item_tuple] = 0
            else:
                print(f"!! 汇总表里有重复的 !! {item_tuple}")

    # 原始数据表
    raw_item_number_dict = {}
    item_start_row = int(raw_sheet_item_start_row_str)
    item_name_col = convert_letter_to_number(raw_sheet_item_name_col_str) - 1
    item_spec_col = convert_letter_to_number(raw_sheet_item_spec_col_str) - 1
    item_num_col = convert_letter_to_number(raw_sheet_item_num_col_str) - 1
    code_col = convert_letter_to_number(raw_sheet_code_col_str) - 1
    
    # print(item_name_col, item_spec_col, item_num_col)

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

            item_name = item_name.replace("▲", "")
            item_name = item_name.replace("◆", "")
            item_name = item_name.replace("◆", "")
            item_name = item_name.replace("●", "")
            item_name = item_name.split("|")[0]
            item_name = item_name.split(" ")[0]

            # item_spec = col[item_spec_col].value
            # index = item_spec.find("*")
            # if index != -1:
            #     item_spec = item_spec[:index]
            # item_spec = item_spec.replace('：', ':')

            code = col[code_col].value

            # item_tuple = (item_name, code)
            num = col[item_num_col].value
            if code not in raw_item_number_dict:
                raw_item_number_dict[code] = num
            else:
                raw_item_number_dict[code] += num

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
        code = key[1]

        for raw_code, raw_value in raw_item_number_dict.items():
            if code == raw_code:
                value += raw_value
        value = int(value)
        print(f"{cur_row}: {key} {value}")
        csv_data.append([summary_item_name, code, value])

    with open(generated_file_path, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerows(csv_data)

    show_message(
        "",
        f"汇总结果已写入到 [{generated_file_path}], 请用Excel打开查看结果",
    )


btn_process = Button(
    root,
    text="开始汇总",
    command=lambda: start_process(raw_sheet_path_var, summary_sheet_path_var),
    padding=15,
)
btn_process.pack()

root.mainloop()
