@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Compilation en cours...
"C:\Users\Log\AppData\Local\Programs\Python\Python312\python.exe" -m PyInstaller --onefile --windowed --icon="%~dp0locview.ico" --name=LOCVIEW --distpath=. --workpath=build --specpath=build --clean locview.py
if exist build rmdir /s /q build
echo.
if exist LOCVIEW.exe (
    echo OK - LOCVIEW.exe mis a jour
) else (
    echo ERREUR - compilation echouee
)
pause
