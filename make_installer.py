# -*- coding: utf-8 -*-
import os, sys, subprocess

HERE         = os.path.dirname(os.path.abspath(__file__))
EXE          = os.path.join(HERE, "LOCVIEW.exe")
ICO          = os.path.join(HERE, "locview.ico")
INSTALLER_PY = os.path.join(HERE, "locview_installer.py")

print("Compilation de LOCVIEW_Setup.exe...")
result = subprocess.run([
    sys.executable, "-m", "PyInstaller",
    "--onefile", "--windowed",
    "--name=LOCVIEW_Setup",
    f"--icon={ICO}",
    f"--add-data={EXE};.",
    f"--add-data={ICO};.",
    "--distpath=" + HERE,
    "--workpath=" + os.path.join(HERE, "build"),
    "--specpath=" + os.path.join(HERE, "build"),
    "--clean",
    INSTALLER_PY
], cwd=HERE)

setup_exe = os.path.join(HERE, "LOCVIEW_Setup.exe")
if result.returncode == 0 and os.path.exists(setup_exe):
    size_mb = os.path.getsize(setup_exe) / 1024 / 1024
    print(f"\nOK — LOCVIEW_Setup.exe cree ({size_mb:.1f} MB)")
else:
    print("\nERREUR — PyInstaller a echoue")
    sys.exit(1)
