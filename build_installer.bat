@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo ── Etape 1 : Compilation LOCVIEW.exe ────────────────────────────────────────
call compile.bat
echo.
echo ── Etape 2 : Creation de l'installeur ───────────────────────────────────────

set NSIS_PATHS=^
"C:\Program Files (x86)\NSIS\makensis.exe" ^
"C:\Program Files\NSIS\makensis.exe" ^
"%LOCALAPPDATA%\Programs\NSIS\makensis.exe"

set MAKENSIS=
for %%P in (%NSIS_PATHS%) do (
    if exist %%P set MAKENSIS=%%P
)

if "%MAKENSIS%"=="" (
    echo NSIS non trouve. Telechargement automatique...
    powershell -Command "& {
        $url = 'https://prdownloads.sourceforge.net/nsis/nsis-3.10-setup.exe'
        $out = \"$env:TEMP\nsis_setup.exe\"
        Invoke-WebRequest -Uri $url -OutFile $out
        Start-Process $out -Wait
    }"
    for %%P in (%NSIS_PATHS%) do (
        if exist %%P set MAKENSIS=%%P
    )
)

if "%MAKENSIS%"=="" (
    echo ERREUR : NSIS introuvable apres installation.
    echo Installez NSIS manuellement : https://nsis.sourceforge.io
    pause & exit /b 1
)

echo Utilisation de : %MAKENSIS%
%MAKENSIS% installer.nsi

if exist "LOCVIEW_Setup.exe" (
    echo.
    echo  OK -- LOCVIEW_Setup.exe cree avec succes !
    echo  Envoyez ce fichier a vos utilisateurs.
) else (
    echo  ERREUR -- La creation de l'installeur a echoue.
)
pause
