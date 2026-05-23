# -*- coding: utf-8 -*-
import os, sys, shutil, threading, winreg as _wr
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as mb
from PIL import Image, ImageTk

APP_NAME    = "LOCVIEW"
APP_VER     = "1.0"
INSTALL_DIR = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "LOCVIEW")
EXE_NAME    = "LOCVIEW.exe"

def _resource(name):
    """Find a bundled file (PyInstaller) or local file (dev mode)."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, name)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)

def _make_shortcut(target, link_path, description=""):
    try:
        import subprocess
        ps = (
            f'$ws = New-Object -ComObject WScript.Shell;'
            f'$sc = $ws.CreateShortcut("{link_path}");'
            f'$sc.TargetPath = "{target}";'
            f'$sc.Description = "{description}";'
            f'$sc.WorkingDirectory = "{os.path.dirname(target)}";'
            f'$sc.Save()'
        )
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                       capture_output=True, timeout=10)
    except Exception:
        pass

_FILE_ASSOCS = [
    (".epub", "LOCVIEW.epub", "Fichier EPUB"),
    (".pdf",  "LOCVIEW.pdf",  "Fichier PDF"),
    (".cbz",  "LOCVIEW.cbz",  "Comic Book ZIP"),
    (".fb2",  "LOCVIEW.fb2",  "Fichier FictionBook"),
    (".xps",  "LOCVIEW.xps",  "Fichier XPS"),
    (".oxps", "LOCVIEW.oxps", "Fichier OpenXPS"),
    (".jpg",  "LOCVIEW.jpg",  "Image JPEG"),
    (".jpeg", "LOCVIEW.jpeg", "Image JPEG"),
    (".png",  "LOCVIEW.png",  "Image PNG"),
    (".bmp",  "LOCVIEW.bmp",  "Image BMP"),
    (".gif",  "LOCVIEW.gif",  "Image GIF"),
    (".webp", "LOCVIEW.webp", "Image WebP"),
    (".tiff", "LOCVIEW.tiff", "Image TIFF"),
    (".tif",  "LOCVIEW.tif",  "Image TIFF"),
]

def _register_file_assoc(exe_path):
    for ext, prog_id, desc in _FILE_ASSOCS:
        try:
            with _wr.CreateKey(_wr.HKEY_CURRENT_USER, rf"Software\Classes\{ext}") as k:
                _wr.SetValueEx(k, "", 0, _wr.REG_SZ, prog_id)
            with _wr.CreateKey(_wr.HKEY_CURRENT_USER, rf"Software\Classes\{prog_id}") as k:
                _wr.SetValueEx(k, "", 0, _wr.REG_SZ, desc)
            with _wr.CreateKey(_wr.HKEY_CURRENT_USER, rf"Software\Classes\{prog_id}\DefaultIcon") as k:
                _wr.SetValueEx(k, "", 0, _wr.REG_SZ, f'"{exe_path}",0')
            with _wr.CreateKey(_wr.HKEY_CURRENT_USER, rf"Software\Classes\{prog_id}\shell\open\command") as k:
                _wr.SetValueEx(k, "", 0, _wr.REG_SZ, f'"{exe_path}" "%1"')
        except Exception:
            pass
    try:
        import ctypes
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
    except Exception:
        pass

def _unregister_file_assoc():
    for ext, prog_id, _ in _FILE_ASSOCS:
        for path in [
            rf"Software\Classes\{prog_id}\shell\open\command",
            rf"Software\Classes\{prog_id}\shell\open",
            rf"Software\Classes\{prog_id}\shell",
            rf"Software\Classes\{prog_id}\DefaultIcon",
            rf"Software\Classes\{prog_id}",
            rf"Software\Classes\{ext}",
        ]:
            try: _wr.DeleteKey(_wr.HKEY_CURRENT_USER, path)
            except Exception: pass
    try:
        import ctypes
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
    except Exception:
        pass

def install(progress_cb, status_cb):
    os.makedirs(INSTALL_DIR, exist_ok=True)
    exe_path = os.path.join(INSTALL_DIR, EXE_NAME)

    status_cb("Copie de LOCVIEW.exe...")
    src = _resource(EXE_NAME)
    shutil.copy2(src, exe_path)
    progress_cb(40)

    status_cb("Association des fichiers .epub et .pdf...")
    _register_file_assoc(exe_path)
    progress_cb(55)

    status_cb("Creation du raccourci Bureau...")
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    _make_shortcut(exe_path, os.path.join(desktop, "LOCVIEW.lnk"), "LOCVIEW")
    progress_cb(70)

    status_cb("Creation du raccourci Menu Demarrer...")
    sm = os.path.join(os.environ.get("APPDATA", ""),
                      "Microsoft", "Windows", "Start Menu", "Programs")
    os.makedirs(sm, exist_ok=True)
    _make_shortcut(exe_path, os.path.join(sm, "LOCVIEW.lnk"), "LOCVIEW")
    progress_cb(85)

    status_cb("Enregistrement dans Programmes...")
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\LOCVIEW"
    try:
        k = _wr.CreateKey(_wr.HKEY_CURRENT_USER, key_path)
        _wr.SetValueEx(k, "DisplayName",     0, _wr.REG_SZ,    "LOCVIEW")
        _wr.SetValueEx(k, "DisplayVersion",  0, _wr.REG_SZ,    APP_VER)
        _wr.SetValueEx(k, "Publisher",       0, _wr.REG_SZ,    "LOCVIEW")
        _wr.SetValueEx(k, "InstallLocation", 0, _wr.REG_SZ,    INSTALL_DIR)
        _wr.SetValueEx(k, "UninstallString", 0, _wr.REG_SZ,
                       '"' + exe_path + '" --uninstall')
        _wr.SetValueEx(k, "NoModify",        0, _wr.REG_DWORD, 1)
        _wr.SetValueEx(k, "NoRepair",        0, _wr.REG_DWORD, 1)
        _wr.CloseKey(k)
    except Exception:
        pass

    progress_cb(100)
    status_cb("Installation terminee !")

class InstallerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Installation de LOCVIEW")
        self.root.geometry("480x310")
        self.root.resizable(False, False)
        self.root.configure(bg="#111111")
        ico = _resource("locview.ico")
        if os.path.isfile(ico):
            try: self.root.iconbitmap(ico)
            except Exception: pass
        self._build()

    def _build(self):
        bg = "#111111"; fg = "#e0e0e0"
        acc = "#e8820a"; acc_hov = "#ff9f2e"; acc_act = "#c06a00"
        self.root.configure(bg=bg)
        ico = _resource("locview.ico")
        if os.path.isfile(ico):
            try:
                img = Image.open(ico).resize((64, 64), Image.LANCZOS)
                self._ico_img = ImageTk.PhotoImage(img)
                tk.Label(self.root, image=self._ico_img, bg=bg).pack(pady=(22, 0))
            except Exception:
                pass
        tk.Label(self.root, text="LOCVIEW",
                 font=("Consolas", 26, "bold"), bg=bg, fg=fg).pack(pady=(4, 0))
        tk.Label(self.root, text="Visionneuse PDF, EPUB & Images",
                 font=("Consolas", 9), bg=bg, fg="#777777").pack()
        tk.Label(self.root, text=f"Dossier : {INSTALL_DIR}",
                 font=("Consolas", 8), bg=bg, fg="#444444").pack(pady=(6, 0))

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Loc.Horizontal.TProgressbar",
                        troughcolor="#222222", background=acc,
                        bordercolor="#222222", lightcolor=acc, darkcolor=acc)
        self.prog = ttk.Progressbar(self.root, length=400, mode="determinate",
                                    style="Loc.Horizontal.TProgressbar")
        self.prog.pack(pady=(20, 6))

        self.status_lbl = tk.Label(self.root, text="Pret a installer.",
                                   font=("Consolas", 9), bg=bg, fg="#888888")
        self.status_lbl.pack()

        self.btn = tk.Button(self.root, text="  Installer  ",
                             font=("Consolas", 11, "bold"),
                             bg=acc, fg="#111111", relief=tk.FLAT,
                             padx=20, pady=7, cursor="hand2",
                             activebackground=acc_act, activeforeground="#111111",
                             command=self._start)
        self.btn.pack(pady=18)
        self.btn.bind("<Enter>", lambda e: self.btn.config(bg=acc_hov))
        self.btn.bind("<Leave>", lambda e: self.btn.config(bg=acc))

    def _start(self):
        self.btn.config(state=tk.DISABLED)

        def run():
            try:
                install(
                    lambda v: self.root.after(0, lambda vv=v: self.prog.configure(value=vv)),
                    lambda s: self.root.after(0, lambda ss=s: self.status_lbl.configure(text=ss))
                )
                self.root.after(0, self._done)
            except Exception as ex:
                msg = str(ex)
                self.root.after(0, lambda: (
                    mb.showerror("Erreur", msg, parent=self.root),
                    self.btn.config(state=tk.NORMAL)
                ))

        threading.Thread(target=run, daemon=True).start()

    def _done(self):
        mb.showinfo("Installation reussie",
                    "LOCVIEW a ete installe avec succes !\n\n"
                    "Un raccourci a ete cree sur le Bureau.",
                    parent=self.root)
        self.root.destroy()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    InstallerGUI().run()
