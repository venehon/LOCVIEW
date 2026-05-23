; ── LOCVIEW Installer — NSIS Script ──────────────────────────────────────────
; Prérequis : NSIS installé (https://nsis.sourceforge.io)
; Commande  : makensis installer.nsi
; Produit   : LOCVIEW_Setup.exe

!define APP_NAME    "LOCVIEW"
!define APP_VERSION "1.0"
!define APP_EXE     "LOCVIEW.exe"
!define APP_ICON    "locview.ico"
!define INST_DIR    "$PROGRAMFILES\LOCVIEW"
!define UNINST_KEY  "Software\Microsoft\Windows\CurrentVersion\Uninstall\LOCVIEW"

Name "${APP_NAME} ${APP_VERSION}"
OutFile "LOCVIEW_Setup.exe"
InstallDir "${INST_DIR}"
InstallDirRegKey HKCU "Software\LOCVIEW" "InstallPath"
RequestExecutionLevel admin
SetCompressor /SOLID lzma
Icon "${APP_ICON}"

; ── Pages ────────────────────────────────────────────────────────────────────
!include "MUI2.nsh"
!define MUI_ABORTWARNING
!define MUI_ICON "${APP_ICON}"
!define MUI_UNICON "${APP_ICON}"
!define MUI_WELCOMEPAGE_TITLE "Installation de LOCVIEW"
!define MUI_WELCOMEPAGE_TEXT  "LOCVIEW est un visualiseur rapide de PDF et d'images.$\n$\nCliquez sur Suivant pour continuer."
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "French"

; ── Installation ─────────────────────────────────────────────────────────────
Section "LOCVIEW (requis)" SecMain
    SectionIn RO
    SetOutPath "${INST_DIR}"

    ; Fichiers principaux
    File "${APP_EXE}"
    File "${APP_ICON}"
    File /nonfatal "locview_saves.json"

    ; Créer dossier sigs
    CreateDirectory "${INST_DIR}\sigs"

    ; Raccourci bureau
    CreateShortCut "$DESKTOP\LOCVIEW.lnk" "${INST_DIR}\${APP_EXE}" "" "${INST_DIR}\${APP_ICON}"

    ; Raccourci menu démarrer
    CreateDirectory "$SMPROGRAMS\LOCVIEW"
    CreateShortCut  "$SMPROGRAMS\LOCVIEW\LOCVIEW.lnk"         "${INST_DIR}\${APP_EXE}" "" "${INST_DIR}\${APP_ICON}"
    CreateShortCut  "$SMPROGRAMS\LOCVIEW\Désinstaller LOCVIEW.lnk" "$INSTDIR\Uninstall.exe"

    ; Clé de registre pour désinstalleur
    WriteRegStr   HKLM "${UNINST_KEY}" "DisplayName"      "${APP_NAME}"
    WriteRegStr   HKLM "${UNINST_KEY}" "DisplayVersion"   "${APP_VERSION}"
    WriteRegStr   HKLM "${UNINST_KEY}" "UninstallString"  "$INSTDIR\Uninstall.exe"
    WriteRegStr   HKLM "${UNINST_KEY}" "InstallLocation"  "$INSTDIR"
    WriteRegStr   HKLM "${UNINST_KEY}" "DisplayIcon"      "$INSTDIR\${APP_ICON}"
    WriteRegStr   HKLM "${UNINST_KEY}" "Publisher"        "LOCVIEW"
    WriteRegDWORD HKLM "${UNINST_KEY}" "NoModify"         1
    WriteRegDWORD HKLM "${UNINST_KEY}" "NoRepair"         1

    ; Association .pdf (optionnelle — LOCVIEW demande au 1er lancement)
    WriteRegStr HKCU "Software\LOCVIEW" "InstallPath" "$INSTDIR"

    ; Écrire désinstalleur
    WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

; ── Désinstallation ───────────────────────────────────────────────────────────
Section "Uninstall"
    Delete "$INSTDIR\${APP_EXE}"
    Delete "$INSTDIR\${APP_ICON}"
    Delete "$INSTDIR\locview_saves.json"
    Delete "$INSTDIR\Uninstall.exe"
    RMDir  /r "$INSTDIR\sigs"
    RMDir  "$INSTDIR"

    Delete "$DESKTOP\LOCVIEW.lnk"
    Delete "$SMPROGRAMS\LOCVIEW\LOCVIEW.lnk"
    Delete "$SMPROGRAMS\LOCVIEW\Désinstaller LOCVIEW.lnk"
    RMDir  "$SMPROGRAMS\LOCVIEW"

    DeleteRegKey HKLM "${UNINST_KEY}"
    DeleteRegKey HKCU "Software\LOCVIEW"
SectionEnd
