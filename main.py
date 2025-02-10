from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
import pandas as pd
import os
import errno
import csv
import sys
import json

# 配置常量
DEFAULT_PADDING = 15
DEFAULT_FONT = "Arial 14 bold"
DEFAULT_BG_COLOR = "yellow"
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

# 默认值配置
DEFAULT_CONFIG = {
    "summary_sheet": {
        "sheet_name": "集采第九批内部统计使用",
        "start_row": "4",
        "name_column": "D",
        "spec_column": "F",
        "code_column": "C",
        "path": ""  # 添加路径配置
    },
    "raw_sheet": {
        "sheet_name": "Sheet1",
        "start_row": "5",
        "name_column": "D",
        "spec_column": "E",
        "num_column": "F",
        "code_column": "C",
        "path": ""  # 添加路径配置
    }
}

class ExcelProcessor:
    """Excel处理器类，处理所有与Excel相关的操作"""
    
    @staticmethod
    def convert_letter_to_number(letter):
        """将Excel列字母转换为数字"""
        if not letter.isalpha():
            return None
        return ord(letter.upper()) - ord('A') + 1

class InputValidator:
    """输入验证器类，处理所有输入验证逻辑"""
    
    @staticmethod
    def allow_only_letters(event):
        """只允许输入字母的验证器"""
        current_text = event.widget.get()
        new_char = event.char

        if new_char in ('', '\x08', '\x0d', '\x1b'):  # 特殊键处理
            return

        if len(current_text) > 0:
            return "break"

        if not new_char.isalpha():
            return "break"

    @staticmethod
    def allow_only_numbers(event):
        """只允许输入数字的验证器"""
        new_char = event.char

        if new_char in ('', '\x08', '\x7f'):  # 特殊键处理
            return

        if not new_char.isdigit():
            return "break"

class ConfigManager:
    """配置管理类，处理配置的保存和加载"""
    
    @staticmethod
    def load_config():
        """加载配置文件"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return DEFAULT_CONFIG
        except Exception as e:
            show_message("警告", f"加载配置文件失败: {str(e)}\n将使用默认配置。")
            return DEFAULT_CONFIG

    @staticmethod
    def save_config(config):
        """保存配置到文件"""
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            show_message("警告", f"保存配置文件失败: {str(e)}")

class GUI:
    """图形界面类，处理所有UI相关操作"""
    
    def __init__(self):
        self.root = Tk()
        self.root.title("Excel数据汇总工具")  # 添加窗口标题
        self.config = ConfigManager.load_config()  # 加载配置
        self.summary_sheet_path_var = StringVar()
        self.raw_sheet_path_var = StringVar()
        self.setup_variables()
        self.create_gui()
        self.setup_window_close()

    def setup_variables(self):
        """初始化所有GUI变量"""
        # 设置路径变量
        self.summary_sheet_path_var.set(self.config["summary_sheet"].get("path", ""))
        self.raw_sheet_path_var.set(self.config["raw_sheet"].get("path", ""))
        
        # 创建并初始化所有输入变量
        self.summary_sheet_label_var = StringVar(value=self.config["summary_sheet"]["sheet_name"])
        self.summary_sheet_start_row_var = StringVar(value=self.config["summary_sheet"]["start_row"])
        self.summary_sheet_name_col_var = StringVar(value=self.config["summary_sheet"]["name_column"])
        self.summary_sheet_spec_col_var = StringVar(value=self.config["summary_sheet"]["spec_column"])
        self.summary_sheet_code_col_var = StringVar(value=self.config["summary_sheet"]["code_column"])
        
        self.raw_sheet_label_var = StringVar(value=self.config["raw_sheet"]["sheet_name"])
        self.raw_sheet_start_row_var = StringVar(value=self.config["raw_sheet"]["start_row"])
        self.raw_sheet_name_col_var = StringVar(value=self.config["raw_sheet"]["name_column"])
        self.raw_sheet_spec_col_var = StringVar(value=self.config["raw_sheet"]["spec_column"])
        self.raw_sheet_num_col_var = StringVar(value=self.config["raw_sheet"]["num_column"])
        self.raw_sheet_code_col_var = StringVar(value=self.config["raw_sheet"]["code_column"])

        # 添加变量跟踪
        self.setup_variable_trace()

    def setup_variable_trace(self):
        """设置变量跟踪，当值改变时保存配置"""
        def save_config(*args):
            self.save_current_config()

        # 跟踪路径变化
        self.summary_sheet_path_var.trace_add("write", save_config)
        self.raw_sheet_path_var.trace_add("write", save_config)

        # 跟踪汇总表配置变化
        self.summary_sheet_label_var.trace_add("write", save_config)
        self.summary_sheet_start_row_var.trace_add("write", save_config)
        self.summary_sheet_name_col_var.trace_add("write", save_config)
        self.summary_sheet_spec_col_var.trace_add("write", save_config)
        self.summary_sheet_code_col_var.trace_add("write", save_config)

        # 跟踪原始数据表配置变化
        self.raw_sheet_label_var.trace_add("write", save_config)
        self.raw_sheet_start_row_var.trace_add("write", save_config)
        self.raw_sheet_name_col_var.trace_add("write", save_config)
        self.raw_sheet_spec_col_var.trace_add("write", save_config)
        self.raw_sheet_num_col_var.trace_add("write", save_config)
        self.raw_sheet_code_col_var.trace_add("write", save_config)

    def save_current_config(self):
        """保存当前配置"""
        current_config = {
            "summary_sheet": {
                "sheet_name": self.summary_sheet_label_var.get(),
                "start_row": self.summary_sheet_start_row_var.get(),
                "name_column": self.summary_sheet_name_col_var.get(),
                "spec_column": self.summary_sheet_spec_col_var.get(),
                "code_column": self.summary_sheet_code_col_var.get(),
                "path": self.summary_sheet_path_var.get()  # 保存路径
            },
            "raw_sheet": {
                "sheet_name": self.raw_sheet_label_var.get(),
                "start_row": self.raw_sheet_start_row_var.get(),
                "name_column": self.raw_sheet_name_col_var.get(),
                "spec_column": self.raw_sheet_spec_col_var.get(),
                "num_column": self.raw_sheet_num_col_var.get(),
                "code_column": self.raw_sheet_code_col_var.get(),
                "path": self.raw_sheet_path_var.get()  # 保存路径
            }
        }
        ConfigManager.save_config(current_config)

    def create_gui(self):
        """创建图形界面"""
        self.create_summary_frame()
        self.create_separator()
        self.create_raw_frame()
        self.create_process_button()

    def create_summary_frame(self):
        """创建汇总表框架"""
        frame = Frame(self.root)
        frame.pack(fill=BOTH, expand=True)

        # 创建选择汇总表按钮和标签
        self.create_file_selector(frame, "选择汇总表", self.summary_sheet_path_var, 0)
        
        # 创建输入字段
        fields = [
            ("工作表名", self.summary_sheet_label_var, None),
            ("数据开始行号", self.summary_sheet_start_row_var, InputValidator.allow_only_numbers),
            ("药品名列号", self.summary_sheet_name_col_var, InputValidator.allow_only_letters),
            ("规格列号", self.summary_sheet_spec_col_var, InputValidator.allow_only_letters),
            ("药品编码", self.summary_sheet_code_col_var, InputValidator.allow_only_letters)
        ]
        
        for i, (label_text, var, validator) in enumerate(fields, 1):
            self.create_input_field(frame, label_text, var, validator, i)

    def create_raw_frame(self):
        """创建原始数据表框架"""
        frame = Frame(self.root)
        frame.pack(fill=BOTH, expand=True)

        # 创建选择原始数据表按钮和标签
        self.create_file_selector(frame, "选择原始数据表", self.raw_sheet_path_var, 0)
        
        # 创建输入字段
        fields = [
            ("工作表名", self.raw_sheet_label_var, None),
            ("数据开始行号", self.raw_sheet_start_row_var, InputValidator.allow_only_numbers),
            ("药品名列号", self.raw_sheet_name_col_var, InputValidator.allow_only_letters),
            ("规格列号", self.raw_sheet_spec_col_var, InputValidator.allow_only_letters),
            ("使用量列号", self.raw_sheet_num_col_var, InputValidator.allow_only_letters),
            ("药品编码", self.raw_sheet_code_col_var, InputValidator.allow_only_letters)
        ]
        
        for i, (label_text, var, validator) in enumerate(fields, 1):
            self.create_input_field(frame, label_text, var, validator, i)

    def create_file_selector(self, parent, button_text, path_var, row):
        """创建文件选择器组件"""
        filetypes = [
            ('Excel files', '*.xlsx;*.xls'),
            ('All files', '*.*')
        ]
        btn = Button(
            parent,
            text=button_text,
            command=lambda: self.open_file_dialog(path_var, filetypes),
            padding=DEFAULT_PADDING
        )
        btn.grid(row=row, column=0, padx=5, pady=5)

        label = Label(
            parent,
            textvariable=path_var,
            padding=(50, 10),
            font=DEFAULT_FONT,
            background=DEFAULT_BG_COLOR
        )
        label.grid(row=row, column=1, padx=5, pady=5)

    def create_input_field(self, parent, label_text, var, validator, row):
        """创建输入字段组件"""
        label = Label(parent, text=label_text)
        label.grid(row=row, column=0, padx=5, pady=5)

        entry = Entry(parent, textvariable=var)
        if validator:
            entry.bind("<Key>", validator)
        entry.grid(row=row, column=1, padx=5, pady=5)

    def create_separator(self):
        """创建分隔符"""
        separator = Frame(self.root, height=2, relief=SUNKEN)
        separator.pack(fill=X, padx=10, pady=10)

    def create_process_button(self):
        """创建处理按钮"""
        btn_process = Button(
            self.root,
            text="开始汇总",
            command=self.start_process,
            padding=DEFAULT_PADDING
        )
        btn_process.pack(pady=20)

    def open_file_dialog(self, var, filetypes):
        """打开文件选择对话框"""
        filename = filedialog.askopenfilename(
            parent=self.root,
            filetypes=filetypes,
            title="选择Excel文件"
        )
        if filename:
            var.set(filename)
            self.save_current_config()

    def setup_window_close(self):
        """设置窗口关闭处理"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def start_process(self):
        """开始处理数据"""
        try:
            summary_path = self.summary_sheet_path_var.get()
            raw_path = self.raw_sheet_path_var.get()

            if not summary_path or not raw_path:
                messagebox.showerror("错误", "请选择汇总表和原始数据表文件")
                return

            # Read summary sheet using pandas
            summary_item_codes_list = []
            try:
                df_summary = pd.read_excel(
                    summary_path, 
                    sheet_name=self.summary_sheet_label_var.get(),
                    header=None,
                    skiprows=int(self.summary_sheet_start_row_var.get())-1
                )
                
                name_col = ExcelProcessor.convert_letter_to_number(self.summary_sheet_name_col_var.get()) - 1
                spec_col = ExcelProcessor.convert_letter_to_number(self.summary_sheet_spec_col_var.get()) - 1
                code_col = ExcelProcessor.convert_letter_to_number(self.summary_sheet_code_col_var.get()) - 1
                
                for idx, row in df_summary.iterrows():
                    code_str = str(row.iloc[code_col])
                    # break if code is empty or contains only whitespace
                    if not code_str or code_str.isspace():
                        break

                    if "," in code_str:
                        codes = code_str.split(",")
                    elif "，" in code_str:
                        codes = code_str.split("，")
                    else:
                        codes = [code_str]

                    summary_item_codes_list.append({
                        # split the code string by comma, strip each code, and only keep non-empty codes
                        'codes': codes,
                        'name': ' '.join(str(row.iloc[name_col]).strip().splitlines()),
                        'spec': str(row.iloc[spec_col]).strip(),
                        'line_number': idx + int(self.summary_sheet_start_row_var.get()),
                        'item_number': 0,
                        })
            except Exception as e:
                messagebox.showerror("错误", f"读取汇总表时出错：{str(e)}")
                return

            # Read and update raw sheet
            try:
                df_raw = pd.read_excel(
                    raw_path,
                    sheet_name=self.raw_sheet_label_var.get(),
                    header=None,
                    skiprows=int(self.raw_sheet_start_row_var.get())-1
                )
                
                name_col = ExcelProcessor.convert_letter_to_number(self.raw_sheet_name_col_var.get()) - 1
                spec_col = ExcelProcessor.convert_letter_to_number(self.raw_sheet_spec_col_var.get()) - 1
                code_col = ExcelProcessor.convert_letter_to_number(self.raw_sheet_code_col_var.get()) - 1
                item_num_col = ExcelProcessor.convert_letter_to_number(self.raw_sheet_num_col_var.get()) - 1

                # Process the data
                for index, row in df_raw.iterrows():
                    code = str(row.iloc[code_col])
                    # break if code is empty or contains only whitespace
                    if not code or code.isspace():
                        break
                    item_num = int(row.iloc[item_num_col])
                    # find item in summary_item_codes_list by code, update its item_number with item_num_col
                    item = next((item for item in summary_item_codes_list if code in item['codes']), None)
                    if item is not None:
                        item['item_number'] += item_num

                print("All summary data:")
                for item in summary_item_codes_list:
                    print(f"Line: {item['line_number']}, Code: {item['codes']}, Name: {item['name']}, Spec: {item['spec']} , Item Number: {item['item_number']}")

                messagebox.showinfo("成功", f"处理完成！")
            except Exception as e:
                messagebox.showerror("错误", f"处理原始数据表时出错：{str(e)}")
                return

        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出现错误：{str(e)}")

    def on_closing(self):
        """窗口关闭处理"""
        self.root.quit()  # 先退出主循环
        self.root.destroy()  # 然后销毁窗口
        sys.exit(0)  # 确保程序完全退出

    def run(self):
        """运行应用程序"""
        self.root.mainloop()

def show_message(title, message):
    """显示消息对话框"""
    root = Tk()
    root.withdraw()
    if not title:
        messagebox.showinfo("", message)
    else:
        messagebox.showerror(title, message)

if __name__ == "__main__":
    app = GUI()
    app.run()
