import json
import os
import sys
from typing import Optional

import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QApplication, QComboBox, QFileDialog, QGridLayout,
                           QGroupBox, QHBoxLayout, QLabel, QLineEdit,
                           QMainWindow, QPushButton, QVBoxLayout, QWidget,
                           QMessageBox)

CONFIG_FILE = "Excel数据汇总工具.json"

class SheetConfig:
    def __init__(self, path: str = "", sheet_name: str = "", start_row: str = "4", name_column: str = "D", code_column: str = "C", num_column: str = ""):  
        self.path = path
        self.sheet_name = sheet_name
        self.start_row = start_row
        self.name_column = name_column
        self.code_column = code_column
        self.num_column = num_column

    def to_dict(self):
        return {
            "path": self.path,
            "sheet_name": self.sheet_name,
            "start_row": self.start_row,
            "name_column": self.name_column,
            "code_column": self.code_column,
            "num_column": self.num_column
        }

    @staticmethod
    def from_dict(data: dict):
        return SheetConfig(
            path=data.get("path", ""),
            sheet_name=data.get("sheet_name", ""),
            start_row=data.get("start_row", "4"),
            name_column=data.get("name_column", "D"),
            code_column=data.get("code_column", "C"),
            num_column=data.get("num_column", "F")
        )

class Config:
    def __init__(self):
        self.summary_sheet = SheetConfig()
        self.raw_sheet = SheetConfig()

    def to_dict(self):
        return {
            "summary_sheet": self.summary_sheet.to_dict(),
            "raw_sheet": self.raw_sheet.to_dict()
        }

    @staticmethod
    def from_dict(data: dict):
        config = Config()
        if data:
            config.summary_sheet = SheetConfig.from_dict(data.get("summary_sheet", {}))
            config.raw_sheet = SheetConfig.from_dict(data.get("raw_sheet", {}))
        return config

def load_config() -> Config:
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return Config.from_dict(json.load(f))
    except Exception as e:
        print(f"Error loading config: {e}")
    return Config()

def save_config(config: Config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config.to_dict(), f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Error saving config: {e}")

def get_sheet_names(file_path: str) -> list[str]:
    """获取Excel文件中的所有工作表名称"""
    if not file_path:
        return []
    try:
        # 使用pandas直接读取Excel文件的工作表名称
        sheet_names = pd.read_excel(file_path, sheet_name=None).keys()
        sheet_names = list(sheet_names)  # 转换为列表
        print(f"Found sheets in {file_path}: {sheet_names}")  # 添加调试信息
        return sheet_names
    except Exception as e:
        print(f"Error getting sheet names from {file_path}: {e}")
        return []

def convert_letter_to_number(letter: str) -> int:
    """将Excel列字母转换为数字索引"""
    if not letter or not letter.isalpha():
        return 0
    result = 0
    for char in letter.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1  # 转换为0-based索引

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = load_config()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Excel数据汇总工具')
        self.setMinimumSize(600, 400)

        # 创建中央部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 创建汇总表配置组
        summary_group = QGroupBox("汇总表配置")
        summary_layout = QGridLayout()
        
        self.summary_path = QLineEdit()
        self.summary_path.setText(self.config.summary_sheet.path)
        self.summary_path.setReadOnly(True)
        
        summary_file_btn = QPushButton("选择文件")
        summary_file_btn.clicked.connect(self.select_summary_file)
        
        self.summary_sheet = QComboBox()
        self.summary_sheet.setEditable(True)  # 允许手动输入
        self.summary_sheet.addItem("请先选择Excel文件")  # 添加默认提示
        self.summary_sheet.setEnabled(False)  # 初始时禁用
        if self.config.summary_sheet.path:
            sheet_names = get_sheet_names(self.config.summary_sheet.path)
            if sheet_names:
                self.summary_sheet.clear()
                self.summary_sheet.setEnabled(True)
                self.summary_sheet.addItems(sheet_names)
                if self.config.summary_sheet.sheet_name in sheet_names:
                    self.summary_sheet.setCurrentText(self.config.summary_sheet.sheet_name)
                else:
                    self.summary_sheet.setCurrentText(sheet_names[0])
        self.summary_sheet.currentTextChanged.connect(self.on_summary_sheet_changed)
        
        # 添加数据开始行号输入框
        self.summary_start_row = QLineEdit()
        self.summary_start_row.setText(self.config.summary_sheet.start_row)
        self.summary_start_row.textChanged.connect(self.on_summary_start_row_changed)
        
        # 添加药品名列号输入框
        self.summary_name_column = QLineEdit()
        self.summary_name_column.setText(self.config.summary_sheet.name_column)
        self.summary_name_column.setMaxLength(1)  # 限制只能输入一个字符
        self.summary_name_column.textChanged.connect(self.on_summary_name_column_changed)
        
        # 添加药品编码列号输入框
        self.summary_code_column = QLineEdit()
        self.summary_code_column.setText(self.config.summary_sheet.code_column)
        self.summary_code_column.setMaxLength(1)  # 限制只能输入一个字符
        self.summary_code_column.textChanged.connect(self.on_summary_code_column_changed)

        summary_layout.addWidget(QLabel("文件路径:"), 0, 0)
        summary_layout.addWidget(self.summary_path, 0, 1)
        summary_layout.addWidget(summary_file_btn, 0, 2)
        summary_layout.addWidget(QLabel("工作表:"), 1, 0)
        summary_layout.addWidget(self.summary_sheet, 1, 1, 1, 2)
        summary_layout.addWidget(QLabel("数据开始行号:"), 2, 0)
        summary_layout.addWidget(self.summary_start_row, 2, 1, 1, 2)
        summary_layout.addWidget(QLabel("药品名列号:"), 3, 0)
        summary_layout.addWidget(self.summary_name_column, 3, 1, 1, 2)
        summary_layout.addWidget(QLabel("药品编码列号:"), 4, 0)
        summary_layout.addWidget(self.summary_code_column, 4, 1, 1, 2)
        summary_group.setLayout(summary_layout)

        # 创建原始表配置组
        raw_group = QGroupBox("原始表配置")
        raw_layout = QGridLayout()
        
        self.raw_path = QLineEdit()
        self.raw_path.setText(self.config.raw_sheet.path)
        self.raw_path.setReadOnly(True)
        
        raw_file_btn = QPushButton("选择文件")
        raw_file_btn.clicked.connect(self.select_raw_file)
        
        self.raw_sheet = QComboBox()
        self.raw_sheet.setEditable(True)  # 允许手动输入
        self.raw_sheet.addItem("请先选择Excel文件")  # 添加默认提示
        self.raw_sheet.setEnabled(False)  # 初始时禁用
        if self.config.raw_sheet.path:
            sheet_names = get_sheet_names(self.config.raw_sheet.path)
            if sheet_names:
                self.raw_sheet.clear()
                self.raw_sheet.setEnabled(True)
                self.raw_sheet.addItems(sheet_names)
                if self.config.raw_sheet.sheet_name in sheet_names:
                    self.raw_sheet.setCurrentText(self.config.raw_sheet.sheet_name)
                else:
                    self.raw_sheet.setCurrentText(sheet_names[0])
        self.raw_sheet.currentTextChanged.connect(self.on_raw_sheet_changed)
        
        # 添加数据开始行号输入框
        self.raw_start_row = QLineEdit()
        self.raw_start_row.setText(self.config.raw_sheet.start_row)
        self.raw_start_row.textChanged.connect(self.on_raw_start_row_changed)
        
        # 添加药品名列号输入框
        self.raw_name_column = QLineEdit()
        self.raw_name_column.setText(self.config.raw_sheet.name_column)
        self.raw_name_column.setMaxLength(1)  # 限制只能输入一个字符
        self.raw_name_column.textChanged.connect(self.on_raw_name_column_changed)
        
        # 添加药品编码列号输入框
        self.raw_code_column = QLineEdit()
        self.raw_code_column.setText(self.config.raw_sheet.code_column)
        self.raw_code_column.setMaxLength(1)  # 限制只能输入一个字符
        self.raw_code_column.textChanged.connect(self.on_raw_code_column_changed)
        
        # 添加使用量列号输入框
        self.raw_num_column = QLineEdit()
        self.raw_num_column.setText(self.config.raw_sheet.num_column)
        self.raw_num_column.setMaxLength(1)  # 限制只能输入一个字符
        self.raw_num_column.textChanged.connect(self.on_raw_num_column_changed)

        raw_layout.addWidget(QLabel("文件路径:"), 0, 0)
        raw_layout.addWidget(self.raw_path, 0, 1)
        raw_layout.addWidget(raw_file_btn, 0, 2)
        raw_layout.addWidget(QLabel("工作表:"), 1, 0)
        raw_layout.addWidget(self.raw_sheet, 1, 1, 1, 2)
        raw_layout.addWidget(QLabel("数据开始行号:"), 2, 0)
        raw_layout.addWidget(self.raw_start_row, 2, 1, 1, 2)
        raw_layout.addWidget(QLabel("药品名列号:"), 3, 0)
        raw_layout.addWidget(self.raw_name_column, 3, 1, 1, 2)
        raw_layout.addWidget(QLabel("药品编码列号:"), 4, 0)
        raw_layout.addWidget(self.raw_code_column, 4, 1, 1, 2)
        raw_layout.addWidget(QLabel("使用量列号:"), 5, 0)
        raw_layout.addWidget(self.raw_num_column, 5, 1, 1, 2)
        raw_group.setLayout(raw_layout)

        # 创建处理按钮
        btn_layout = QHBoxLayout()
        process_btn = QPushButton("开始处理")
        process_btn.setMinimumHeight(40)
        process_btn.clicked.connect(self.process_excel)
        btn_layout.addStretch()
        btn_layout.addWidget(process_btn)
        btn_layout.addStretch()

        # 添加所有组件到主布局
        main_layout.addWidget(summary_group)
        main_layout.addWidget(raw_group)
        main_layout.addLayout(btn_layout)
        main_layout.addStretch()

    def select_summary_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择汇总表文件", "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.summary_path.setText(file_path)
            self.config.summary_sheet.path = file_path
            
            # 更新工作表列表
            sheet_names = get_sheet_names(file_path)
            self.summary_sheet.clear()
            if sheet_names:
                self.summary_sheet.setEnabled(True)
                self.summary_sheet.addItems(sheet_names)
                self.summary_sheet.setCurrentText(sheet_names[0])
                self.config.summary_sheet.sheet_name = sheet_names[0]
            else:
                self.summary_sheet.setEnabled(False)
                self.summary_sheet.addItem("未找到工作表")
            save_config(self.config)

    def select_raw_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择原始表文件", "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.raw_path.setText(file_path)
            self.config.raw_sheet.path = file_path
            
            # 更新工作表列表
            sheet_names = get_sheet_names(file_path)
            self.raw_sheet.clear()
            if sheet_names:
                self.raw_sheet.setEnabled(True)
                self.raw_sheet.addItems(sheet_names)
                self.raw_sheet.setCurrentText(sheet_names[0])
                self.config.raw_sheet.sheet_name = sheet_names[0]
            else:
                self.raw_sheet.setEnabled(False)
                self.raw_sheet.addItem("未找到工作表")
            save_config(self.config)

    def on_summary_sheet_changed(self, sheet_name: str):
        self.config.summary_sheet.sheet_name = sheet_name
        save_config(self.config)
        
    def on_summary_start_row_changed(self, start_row: str):
        if start_row and not start_row.isdigit():
            return
        self.config.summary_sheet.start_row = start_row
        save_config(self.config)
        
    def on_summary_name_column_changed(self, column: str):
        if column and not column.isalpha():
            return
        self.config.summary_sheet.name_column = column.upper()
        save_config(self.config)
        
    def on_summary_code_column_changed(self, column: str):
        if column and not column.isalpha():
            return
        self.config.summary_sheet.code_column = column.upper()
        save_config(self.config)

    def on_raw_sheet_changed(self, sheet_name: str):
        self.config.raw_sheet.sheet_name = sheet_name
        save_config(self.config)
        
    def on_raw_start_row_changed(self, start_row: str):
        if start_row and not start_row.isdigit():
            return
        self.config.raw_sheet.start_row = start_row
        save_config(self.config)
        
    def on_raw_name_column_changed(self, column: str):
        if column and not column.isalpha():
            return
        self.config.raw_sheet.name_column = column.upper()
        save_config(self.config)
        
    def on_raw_code_column_changed(self, column: str):
        if column and not column.isalpha():
            return
        self.config.raw_sheet.code_column = column.upper()
        save_config(self.config)
        
    def on_raw_num_column_changed(self, column: str):
        if column and not column.isalpha():
            return
        self.config.raw_sheet.num_column = column.upper()
        save_config(self.config)

    def process_excel(self):
        # 检查文件和工作表是否已选择
        if not self.config.summary_sheet.path or not self.config.raw_sheet.path:
            QMessageBox.critical(self, "错误", "请选择汇总表和原始数据表文件")
            return

        # 检查工作表名称是否存在
        summary_sheet_name = self.summary_sheet.currentText()
        raw_sheet_name = self.raw_sheet.currentText()
        
        summary_sheets = get_sheet_names(self.config.summary_sheet.path)
        raw_sheets = get_sheet_names(self.config.raw_sheet.path)

        if summary_sheet_name not in summary_sheets:
            QMessageBox.critical(self, "错误", f"汇总表中不存在工作表'{summary_sheet_name}'")
            return

        if raw_sheet_name not in raw_sheets:
            QMessageBox.critical(self, "错误", f"原始表中不存在工作表'{raw_sheet_name}'")
            return
            
        # 检查列号是否有效
        if not self.config.summary_sheet.name_column or not self.config.summary_sheet.code_column:
            QMessageBox.critical(self, "错误", "请填写汇总表的药品名列号和药品编码列号")
            return
            
        if not self.config.raw_sheet.name_column or not self.config.raw_sheet.code_column or not self.config.raw_sheet.num_column:
            QMessageBox.critical(self, "错误", "请填写原始表的药品名列号、药品编码列号和使用量列号")
            return
            
        # 检查开始行号是否有效
        try:
            summary_start_row = int(self.config.summary_sheet.start_row)
            raw_start_row = int(self.config.raw_sheet.start_row)
            if summary_start_row < 1 or raw_start_row < 1:
                raise ValueError("行号必须大于0")
        except ValueError:
            QMessageBox.critical(self, "错误", "数据开始行号必须是有效的正整数")
            return

        try:
            # 获取列索引
            summary_name_col = convert_letter_to_number(self.config.summary_sheet.name_column)
            summary_code_col = convert_letter_to_number(self.config.summary_sheet.code_column)
            raw_name_col = convert_letter_to_number(self.config.raw_sheet.name_column)
            raw_code_col = convert_letter_to_number(self.config.raw_sheet.code_column)
            raw_num_col = convert_letter_to_number(self.config.raw_sheet.num_column)
            
            # 读取汇总表
            df_summary = pd.read_excel(
                self.config.summary_sheet.path,
                sheet_name=summary_sheet_name,  
                header=None,
                skiprows=summary_start_row - 1  # 转换为0-based索引
            )

            # 处理汇总表数据
            summary_item_codes_list = []
            for idx, row in df_summary.iterrows():
                code_str = str(row.iloc[summary_code_col])
                # 如果编码为空或只包含空白字符，跳过该行
                if not code_str or code_str.isspace() or code_str == "nan":
                    continue

                # 处理可能包含多个编码的情况
                if "," in code_str:
                    codes = code_str.split(",")
                elif "，" in code_str:
                    codes = code_str.split("，")
                else:
                    codes = [code_str]

                # 创建药品项目记录
                summary_item_codes_list.append({
                    "codes": [c.strip() for c in codes if c.strip()],  
                    "name": " ".join(str(row.iloc[summary_name_col]).strip().splitlines()),
                    "line_number": idx + summary_start_row,  
                    "item_number": 0
                })

            # 读取原始数据表
            df_raw = pd.read_excel(
                self.config.raw_sheet.path,
                sheet_name=raw_sheet_name,  
                header=None,
                skiprows=raw_start_row - 1  # 转换为0-based索引
            )

            # 处理原始数据
            for index, row in df_raw.iterrows():
                code_str = str(row.iloc[raw_code_col])
                # 如果编码为空或只包含空白字符，跳过该行
                if not code_str or code_str.isspace() or code_str == "nan":
                    continue

                try:
                    item_num = int(row.iloc[raw_num_col])
                except (ValueError, TypeError):
                    print(f"警告：第{index+raw_start_row}行的使用量不是有效数字，已跳过")
                    continue

                # 在汇总表中查找对应的药品并更新使用量
                item = next(
                    (item for item in summary_item_codes_list if code_str in item["codes"]),
                    None
                )
                if item is not None:
                    item["item_number"] += item_num

            # 准备CSV数据
            csv_data = [["药品名", "药品编码", "使用量"]]
            for item in summary_item_codes_list:
                csv_data.append([
                    item["name"],
                    ",".join(item["codes"]),
                    int(item["item_number"])
                ])

            # 生成输出文件路径
            source_file_path = self.config.summary_sheet.path
            source_file_name = os.path.basename(source_file_path)
            base_name, _ = os.path.splitext(source_file_name)
            output_path = os.path.join(
                os.path.dirname(source_file_path),
                f"{base_name}_generated.csv"
            )

            # 保存结果到CSV文件
            df_output = pd.DataFrame(csv_data[1:], columns=csv_data[0])
            df_output.to_csv(output_path, index=False)

            # 将使用量的数据(数字部分, 不含第一行)拷贝到剪切板
            QApplication.clipboard().setText(df_output["使用量"].to_csv(index=False, header=False))

            QMessageBox.information(
                self,
                "成功",
                f"处理完成\n结果保存在：{output_path}\n使用量数据已复制到剪切板, 可以直接粘贴进表格"
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "错误",
                f"处理过程中出现错误：{str(e)}"
            )

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
