@echo off
REM 一键批量打包脚本

REM 打包 excel_to_csv_gui.py
pyinstaller --noconsole --onefile --add-data "config.ini;." --add-data "requirements.txt;." excel_to_csv_gui.py

REM 打包 config_maintainer.py
pyinstaller --noconsole --onefile --add-data "config.ini;." --add-data "requirements.txt;." config_maintainer.py

echo.
echo 打包完成！可执行文件在 dist 目录下。
pause