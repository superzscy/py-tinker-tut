@echo off
chcp 65001

poetry run pyinstaller --onefile --windowed main_qt.py -i icon.ico -n "Excel数据汇总工具"

pause