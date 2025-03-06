import json
import os
import sys
from typing import Optional

import openpyxl
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QApplication, QComboBox, QFileDialog, QGridLayout,
                           QGroupBox, QHBoxLayout, QLabel, QLineEdit,
                           QMainWindow, QPushButton, QVBoxLayout, QWidget,
                           QMessageBox)

CONFIG_FILE = "config.json"

class SheetConfig:
    def __init__(self, path: str = "", sheet_name: str = ""):
        self.path = path
        self.sheet_name = sheet_name

    def to_dict(self):
        return {
            "path": self.path,
            "sheet_name": self.sheet_name
        }

    @staticmethod
    def from_dict(data: dict):
        return SheetConfig(
            path=data.get("path", ""),
            sheet_name=data.get("sheet_name", "")
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
    if not file_path:
        return []
    try:
        workbook = openpyxl.load_workbook(file_path, read_only=True)
        return workbook.sheetnames
    except Exception as e:
        print(f"Error getting sheet names: {e}")
        return []

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
        if self.config.summary_sheet.path:
            self.summary_sheet.addItems(get_sheet_names(self.config.summary_sheet.path))
            self.summary_sheet.setCurrentText(self.config.summary_sheet.sheet_name)
        self.summary_sheet.currentTextChanged.connect(self.on_summary_sheet_changed)

        summary_layout.addWidget(QLabel("文件路径:"), 0, 0)
        summary_layout.addWidget(self.summary_path, 0, 1)
        summary_layout.addWidget(summary_file_btn, 0, 2)
        summary_layout.addWidget(QLabel("工作表:"), 1, 0)
        summary_layout.addWidget(self.summary_sheet, 1, 1, 1, 2)
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
        if self.config.raw_sheet.path:
            self.raw_sheet.addItems(get_sheet_names(self.config.raw_sheet.path))
            self.raw_sheet.setCurrentText(self.config.raw_sheet.sheet_name)
        self.raw_sheet.currentTextChanged.connect(self.on_raw_sheet_changed)

        raw_layout.addWidget(QLabel("文件路径:"), 0, 0)
        raw_layout.addWidget(self.raw_path, 0, 1)
        raw_layout.addWidget(raw_file_btn, 0, 2)
        raw_layout.addWidget(QLabel("工作表:"), 1, 0)
        raw_layout.addWidget(self.raw_sheet, 1, 1, 1, 2)
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
            "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            self.summary_path.setText(file_path)
            self.config.summary_sheet.path = file_path
            
            sheet_names = get_sheet_names(file_path)
            self.summary_sheet.clear()
            self.summary_sheet.addItems(sheet_names)
            if sheet_names:
                self.summary_sheet.setCurrentText(sheet_names[0])
                self.config.summary_sheet.sheet_name = sheet_names[0]
            
            save_config(self.config)

    def select_raw_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择原始表文件", "",
            "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            self.raw_path.setText(file_path)
            self.config.raw_sheet.path = file_path
            
            sheet_names = get_sheet_names(file_path)
            self.raw_sheet.clear()
            self.raw_sheet.addItems(sheet_names)
            if sheet_names:
                self.raw_sheet.setCurrentText(sheet_names[0])
                self.config.raw_sheet.sheet_name = sheet_names[0]
            
            save_config(self.config)

    def on_summary_sheet_changed(self, sheet_name: str):
        self.config.summary_sheet.sheet_name = sheet_name
        save_config(self.config)

    def on_raw_sheet_changed(self, sheet_name: str):
        self.config.raw_sheet.sheet_name = sheet_name
        save_config(self.config)

    def process_excel(self):
        try:
            # TODO: 在这里添加Excel处理逻辑
            QMessageBox.information(self, "成功", "数据处理完成！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理数据时出错：{str(e)}")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
