@echo off
chcp 65001 >nul
cd /d "%~dp0"
start "" "C:\Users\Log\AppData\Local\Programs\Python\Python312\pythonw.exe" "%~dp0locview.py" %*
