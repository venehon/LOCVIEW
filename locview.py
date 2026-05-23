import sys, os, threading, queue, tkinter as tk, json, ctypes, socket, atexit
import textwrap, html as _html_lib, tempfile
from tkinter import filedialog, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageFont
import fitz
from pathlib import Path
from tkinterdnd2 import TkinterDnD, DND_FILES

try: from docx import Document as _DocxDoc
except ImportError: _DocxDoc = None
try: from pptx import Presentation as _Pptx
except ImportError: _Pptx = None
try:
    from odf.opendocument import load as _OdfLoad
    from odf import text as _OdfText, teletype as _OdfTeletype
except ImportError: _OdfLoad = None
try:
    from svglib.svglib import svg2rlg as _svg2rlg
    from reportlab.graphics import renderPM as _renderPM
except ImportError: _svg2rlg = None
try: import rarfile as _rarfile
except ImportError: _rarfile = None

# ── DPI Awareness (Windows) — doit être avant la création de la fenêtre ───────
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
except Exception:
    pass

# ── Check instance unique AVANT de créer la fenêtre (évite la plume dans la barre) ──
_IPC_PORT_FILE = os.path.join(os.environ.get("TEMP", os.path.expanduser("~")),
                              "locview_ipc.port")

def _early_ipc_send(filepath):
    """Tente d'envoyer le fichier à une instance existante. Retourne True si réussi."""
    try:
        with open(_IPC_PORT_FILE, "r") as f:
            port = int(f.read().strip())
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(1.0)
        s.connect(("127.0.0.1", port))
        s.sendall(filepath.encode("utf-8"))
        s.close()
        return True
    except Exception:
        return False

_file_arg_early = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else None
if _file_arg_early and _early_ipc_send(_file_arg_early):
    sys.exit(0)   # Instance existante trouvée — pas de fenêtre créée, zéro plume

# ── Marque-pages (fichier JSON local) ────────────────────────────────────────
_SAVE_FILE = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "locview_saves.json")

def _load_saves():
    try:
        with open(_SAVE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def _write_saves(data):
    try:
        with open(_SAVE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ── Localisation (EN / FR) ───────────────────────────────────────────────────
STRINGS = {
    "en": {
        "btn_open": "Open  [O]", "btn_print": "Print  [P]",
        "btn_fullscreen": "Full screen  [F11]",
        "btn_prev": "< Prev", "btn_next": "Next >",
        "btn_search": "Search  [Ctrl+F]", "btn_edit": "✏ Edit",
        "btn_undo_toolbar": "↩ Undo", "btn_save": "💾 Save",
        "btn_bookmark": "Bookmark [M]", "lang_btn": "🌐 FR",
        "search_label": "Search:", "search_whole_word": "Whole word",
        "search_no_results": "No results",
        "sel_header": "Selected text", "sel_copy": "⎘  Copy",
        "sel_copied": "✓ Copied", "sel_no_text": "(no text — image PDF?)",
        "hint_drop": "Drop a file here", "hint_open": "Open",
        "tab_untitled": "Untitled",
        "dialog_files": "Files", "dialog_all": "All",
        "dialog_images": "Images",
        "status_pos_saved": "  ✓ Position saved",
        "status_pos_restored": "  ↩ Page {n} restored",
        "status_file_saved": "  ✓ File saved",
        "status_save_error": "  Error: {msg}",
        "status_not_found": "Not found: {path}",
        "status_unreadable": "Error: unreadable file ({ename})",
        "status_unsupported": "Unsupported format: {ext}",
        "status_print_error": "Print error: {msg}",
        "status_no_file": "No file open.",
        "status_undo_diff_doc": "  Undo: different document, skipped",
        "status_undo_done": "  ↩ Edit undone",
        "status_undo_left": "  ({n} left)",
        "status_undo_error": "  Undo error: {msg}",
        "status_open_pdf": "  Open a PDF to edit",
        "status_edit_mode": "  Edit mode — click a word  [Esc: exit]",
        "status_modified": "  ✓ Modified",
        "status_edit_error": "  Error: {msg}",
        "status_default_set": "  ✓ Default app set",
        "status_restart_admin": "  ⚠ Restart as administrator",
        "stamps_up": "▲ Stamps ({n})", "stamps_down": "▼ Stamps ({n})",
        "stamps_none": "No stamps placed", "stamp_delete": "Delete stamp",
        "stamp_paraph": "Paraph",
        "edit_title": "Edit text",
        "edit_instruction": "Select and edit the text:",
        "edit_confirm": "Confirm", "edit_cancel": "Cancel",
        "sig_title": "Signature / Stamp",
        "sig_draw": "✍ Draw", "sig_text_mode": "Aa Text",
        "sig_width": "Width:", "sig_text_label": "Text:",
        "sig_undo": "↩ Undo", "sig_redo": "↪ Redo", "sig_clear": "Clear",
        "sig_saved": "Saved:", "sig_all": "All",
        "sig_no_drawing": "No drawing", "sig_draw_first": "Draw a signature first.",
        "sig_error": "Error", "sig_save_btn": "💾 Save signature",
        "sig_embed": "Embed in document", "sig_cancel": "Cancel",
        "notif_default": "Set LOCVIEW as default app?",
        "notif_dont_ask": "Don't ask again",
        "notif_yes": "Yes", "notif_no": "No",
        "textbox_placeholder": "Text here…",
    },
    "fr": {
        "btn_open": "Ouvrir  [O]", "btn_print": "Imprimer  [P]",
        "btn_fullscreen": "Plein écran  [F11]",
        "btn_prev": "< Préc", "btn_next": "Suiv >",
        "btn_search": "Recherche  [Ctrl+F]", "btn_edit": "✏ Éditer",
        "btn_undo_toolbar": "↩ Annuler", "btn_save": "💾 Sauvegarder",
        "btn_bookmark": "Marque-page [M]", "lang_btn": "🌐 EN",
        "search_label": "Rechercher :", "search_whole_word": "Mot entier",
        "search_no_results": "Aucun résultat",
        "sel_header": "Texte sélectionné", "sel_copy": "⎘  Copier",
        "sel_copied": "✓ Copié", "sel_no_text": "(aucun texte — PDF image ?)",
        "hint_drop": "Glisse un fichier ici", "hint_open": "Ouvrir",
        "tab_untitled": "Nouveau",
        "dialog_files": "Fichiers", "dialog_all": "Tous",
        "dialog_images": "Images",
        "status_pos_saved": "  ✓ Position sauvegardée",
        "status_pos_restored": "  ↩ Page {n} restaurée",
        "status_file_saved": "  ✓ Fichier sauvegardé",
        "status_save_error": "  Erreur : {msg}",
        "status_not_found": "Introuvable : {path}",
        "status_unreadable": "Erreur : fichier non lisible ({ename})",
        "status_unsupported": "Format non supporté : {ext}",
        "status_print_error": "Erreur impression : {msg}",
        "status_no_file": "Aucun fichier ouvert.",
        "status_undo_diff_doc": "  Undo : document différent, ignoré",
        "status_undo_done": "  ↩ Modification annulée",
        "status_undo_left": "  ({n} restante)",
        "status_undo_error": "  Erreur undo : {msg}",
        "status_open_pdf": "  Ouvrez un PDF pour éditer",
        "status_edit_mode": "  Mode édition — cliquez sur un mot  [Échap : quitter]",
        "status_modified": "  ✓ Modifié",
        "status_edit_error": "  Erreur : {msg}",
        "status_default_set": "  ✓ Application par défaut définie",
        "status_restart_admin": "  ⚠ Relancer en administrateur",
        "stamps_up": "▲ Tampons ({n})", "stamps_down": "▼ Tampons ({n})",
        "stamps_none": "Aucun tampon placé", "stamp_delete": "Supprimer ce tampon",
        "stamp_paraph": "Paraphe",
        "edit_title": "Éditer le texte",
        "edit_instruction": "Sélectionnez et modifiez le texte :",
        "edit_confirm": "Confirmer", "edit_cancel": "Annuler",
        "sig_title": "Signature / Paraphe",
        "sig_draw": "✍ Dessin", "sig_text_mode": "Aa Texte",
        "sig_width": "Épaiss. :", "sig_text_label": "Texte :",
        "sig_undo": "↩ Annuler", "sig_redo": "↪ Restaurer", "sig_clear": "Effacer",
        "sig_saved": "Sauvegardées :", "sig_all": "Tous",
        "sig_no_drawing": "Aucun dessin", "sig_draw_first": "Dessinez d'abord une signature.",
        "sig_error": "Erreur", "sig_save_btn": "💾 Sauvegarder la sig.",
        "sig_embed": "Intégrer dans le document", "sig_cancel": "Annuler",
        "notif_default": "Définir LOCVIEW comme appli. par défaut ?",
        "notif_dont_ask": "Ne plus demander",
        "notif_yes": "Oui", "notif_no": "Non",
        "textbox_placeholder": "Texte ici…",
    },
}
LANG = _load_saves().get("lang", "en")
_L   = STRINGS[LANG]

def save_position():
    if S["mode"] == "pdf" and S["doc_path"]:
        data = _load_saves()
        data[S["doc_path"]] = {"page": S["page"], "zoom": round(S["zoom"], 4), "fit": S["fit"]}
        _write_saves(data)
        _flash_save()
    elif S["mode"] == "img" and S["imgs"]:
        data = _load_saves()
        data[S["imgs"][S["img_idx"]]] = {"img_idx": S["img_idx"]}
        _write_saves(data)
        _flash_save()

def _flash_save():
    lbl_save_status.config(text=_L["status_pos_saved"], fg="#7f7")
    root.after(2000, lambda: lbl_save_status.config(text=""))

def _restore_position(path):
    data = _load_saves()
    entry = data.get(path)
    if not entry:
        return False
    if "page" in entry:
        page = entry["page"]
        if 0 <= page < S["total"]:
            S["page"] = page
            S["zoom"] = entry.get("zoom", 1.0)
            S["fit"]  = entry.get("fit", "page")
            lbl_save_status.config(text=_L["status_pos_restored"].format(n=page+1), fg="#aaf")
            root.after(2500, lambda: lbl_save_status.config(text=""))
            return True
    return False

IMG_EXT  = {".jpg",".jpeg",".png",".bmp",".gif",".webp",".tiff",".tif"}
TXT_EXT  = {".txt",".md",".log",".csv",".rst",".ini",".py",".js",".json"}
DOC_EXT  = {".docx",".odt",".pptx",".html",".htm"}
CBR_EXT  = {".cbr"}
SVG_EXT  = {".svg"}

_TXT_PW, _TXT_PH, _TXT_MG = 816, 1056, 52   # page texte (px)
_TXT_BG      = (22, 22, 28)
_TXT_FG      = (215, 215, 215)
_TXT_H_CLR   = {"h1":(190,225,255), "h2":(165,205,245), "h3":(145,190,235)}
_TXT_CODE_BG = (32, 32, 42);  _TXT_CODE_FG = (170, 215, 150)
_TXT_RULE_C  = (75, 75, 95);  _TXT_BULL_C  = (110, 170, 255)
_TXT_LH      = {"body":21,"bullet":21,"body_i":21,"h1":36,"h2":30,"h3":26,"code":18,"hr":18}
_TEMP_FILES  = []
BG       = "#111"
BAR_BG   = "#1c1c1c"
_APP_DIR = Path(os.path.dirname(os.path.abspath(sys.argv[0])))
SIGS_DIR = _APP_DIR / "sigs"
SIGS_DIR.mkdir(exist_ok=True)
TAB_BG  = "#161616"
BTN_BG  = "#2e2e2e"
BTN_FG  = "#e0e0e0"
BTN_HOV = "#505050"
BTN_ACT = "#666"
SBAR_W  = 42
PREV_W  = 180
PREV_H  = 230

root = TkinterDnD.Tk()
root.title("LOCVIEW")
_DPI_SCALE = root.winfo_fpixels("1i") / 96.0
root.geometry(f"{int(1100 * _DPI_SCALE)}x{int(820 * _DPI_SCALE)}")
root.state("zoomed")
root.configure(bg=BG)
try:
    _ico = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "locview.ico")
    if os.path.isfile(_ico):
        root.iconbitmap(_ico)
except Exception:
    pass

# ── État global (onglet actif) ────────────────────────────────────────────────
def _make_state():
    return dict(
        mode=None, doc=None, doc_path=None, page=0, total=0,
        imgs=[], img_idx=0,
        txt_pages=[],
        zoom=1.0, fit="page",
        pan_x=0, pan_y=0,
        ref=None, prev_ref=None, img_item=None,
        pil_cache={},
        cache_lock=threading.Lock(),
        render_id=0,
        scale=1.0, pw=0.0, ph=0.0,
        sel_start=None, sel_text="",
        fullscreen=False,
        title="Untitled",
        drag_start=None,
    )

S      = _make_state()   # état de travail (toujours l'onglet actif)
TABS   = [_make_state()] # états stockés — JAMAIS les mêmes objets que S
ACTIVE = [0]

# ── Thread de rendu unique ────────────────────────────────────────────────────
_CACHE_MAX       = 15          # pages PIL gardées en mémoire
_worker_doc      = [None]      # fitz.Document persistant (thread de rendu uniquement)
_worker_doc_path = [None]

def _get_worker_doc(doc_path):
    """Document fitz réutilisable dans le thread de rendu (jamais ré-ouvert si même chemin)."""
    if _worker_doc_path[0] != doc_path or _worker_doc[0] is None:
        if _worker_doc[0] is not None:
            try: _worker_doc[0].close()
            except: pass
        _worker_doc[0]      = fitz.open(doc_path)
        _worker_doc_path[0] = doc_path
    return _worker_doc[0]

_q = queue.Queue()
def _worker():
    while True:
        fn = _q.get()
        try: fn()
        except Exception: pass
threading.Thread(target=_worker, daemon=True).start()

def _queue_render(fn):
    while not _q.empty():
        try: _q.get_nowait()
        except queue.Empty: break
    _q.put(fn)

# ── Helpers UI ────────────────────────────────────────────────────────────────
def make_btn(parent, text, cmd, **kw):
    b = tk.Button(parent, text=text, command=cmd,
                  bg=BTN_BG, fg=BTN_FG, relief=tk.FLAT,
                  activebackground=BTN_ACT, activeforeground="#fff",
                  font=("Consolas", 9), cursor="hand2", padx=10, pady=2, **kw)
    b.bind("<Enter>", lambda e: b.config(bg=BTN_HOV, fg="#fff"))
    b.bind("<Leave>", lambda e: b.config(bg=BTN_BG,  fg=BTN_FG))
    return b

# ── Signatures ────────────────────────────────────────────────────────────────
_placed_sigs = []   # [{kind, item_id, x, y, page, handle, close_id, ...}]
_sig_drag    = {"on": False, "idx": -1, "ox": 0, "oy": 0, "mode": "body"}
_sig_panel_win = [None]   # panel des tampons places
_text_boxes  = []   # [{win_id, frame, x, y, drag}]

def _sig_hit(ex, ey):
    """Retourne (idx, 'handle'|'close'|'body') ou (-1, None)."""
    HW = 12; CH = 16
    for i, s in enumerate(_placed_sigs):
        bb = canvas.bbox(s["item_id"])
        if not bb: continue
        x1, y1, x2, y2 = bb
        # Bouton fermer (haut-droit)
        if x2-CH <= ex <= x2 and y1 <= ey <= y1+CH:
            return i, "close"
        # Poignee resize (bas-droit)
        if x2-HW <= ex <= x2 and y2-HW <= ey <= y2:
            return i, "handle"
        # Corps
        if x1 <= ex <= x2 and y1 <= ey <= y2:
            return i, "body"
    return -1, None

def _place_sig_at(pil_img, x=None, y=None, orig_pil=None, subkind="image"):
    cw = max(canvas.winfo_width(), 300)
    ch = max(canvas.winfo_height(), 300)
    w, h = pil_img.size
    if x is None: x = cw//2 - w//2
    if y is None: y = ch//2 - h//2
    tk_img  = ImageTk.PhotoImage(pil_img)
    item_id = canvas.create_image(x, y, image=tk_img, anchor="nw", tags="sig")
    px0, py0 = _canvas_to_pdf(x, y) if S.get("doc") else (x, y)
    px1, py1 = _canvas_to_pdf(x+w, y+h) if S.get("doc") else (x+w, y+h)
    sig = {"kind": "image", "subkind": subkind, "orig_pil": orig_pil or pil_img,
           "tk_img": tk_img, "item_id": item_id,
           "x": x, "y": y, "page": S["page"],
           "pdf_rect": (px0, py0, px1, py1),
           "handle": None, "close_id": None}
    _placed_sigs.append(sig)
    canvas.tag_raise("sig")
    _draw_sig_handle(sig)
    _sig_panel_refresh()

def _draw_sig_handle(sig):
    if sig.get("handle"):   canvas.delete(sig["handle"])
    if sig.get("close_id"): canvas.delete(sig["close_id"])
    bb = canvas.bbox(sig["item_id"])
    if not bb: return
    HW = 8
    x1, y1, x2, y2 = bb
    sig["handle"] = canvas.create_rectangle(
        x2-HW, y2-HW, x2, y2,
        fill="#4af", outline="#fff", width=1, tags="sig_handle"
    )
    sig["close_id"] = canvas.create_text(
        x2-1, y1+1, text="x", anchor="ne",
        fill="#ff5555", font=("Consolas", 11, "bold"), tags="sig_close"
    )
def _place_sig(pil_img, subkind="signature"):   # compat
    _place_sig_at(pil_img, subkind=subkind)

def _place_text_stamp(text_str, font_name="Arial", size=24, color="#000000", x=None, y=None):
    cw = max(canvas.winfo_width(), 300)
    ch = max(canvas.winfo_height(), 300)
    if x is None: x = cw // 2
    if y is None: y = ch // 2
    item_id = canvas.create_text(x, y, text=text_str,
                                  font=(font_name, size, "italic"),
                                  fill=color, anchor="center", tags="sig")
    px, py = _canvas_to_pdf(x, y) if S.get("doc") else (x, y)
    sig = {"kind": "text", "subkind": "paraphe", "text": text_str, "font_name": font_name,
           "size": size, "color": color,
           "item_id": item_id, "x": x, "y": y, "page": S["page"],
           "pdf_rect": (px, py, px, py),
           "handle": None, "close_id": None}
    _placed_sigs.append(sig)
    canvas.tag_raise("sig")
    _draw_sig_handle(sig)
    _sig_panel_refresh()


def _del_sig(idx):
    """Supprime le tampon idx et rafraichit le panel."""
    sig = _placed_sigs[idx]
    canvas.delete(sig["item_id"])
    if sig.get("handle"):   canvas.delete(sig["handle"])
    if sig.get("close_id"): canvas.delete(sig["close_id"])
    _placed_sigs.pop(idx)
    _sig_panel_refresh()

# ── Panel tampons (bas, repliable, horizontal) ───────────────────────────────
_sp = {"outer": None, "content": None, "hdr_lbl": None, "expanded": True}

def _sig_panel_refresh():
    """Rafraichit les cartes du panel bas."""
    c = _sp.get("content")
    if not c or not c.winfo_exists(): return
    for w in list(c.winfo_children()): w.destroy()
    lbl = _sp.get("hdr_lbl")
    n = len(_placed_sigs)
    if lbl and lbl.winfo_exists():
        lbl.config(text=_L["stamps_up"].format(n=n) if _sp["expanded"] else _L["stamps_down"].format(n=n))
    if not _placed_sigs:
        tk.Label(c, text=_L["stamps_none"], bg="#141414", fg="#444",
                 font=("Consolas",8)).pack(side=tk.LEFT, padx=10, pady=4)
        return
    _type_counts = {}
    for i, sig in enumerate(_placed_sigs):
        sk = sig.get("subkind", sig.get("kind", "image"))
        _type_counts[sk] = _type_counts.get(sk, 0) + 1
        num = _type_counts[sk]
        labels = {"signature": "Signature", "paraphe": _L["stamp_paraph"],
                  "image": "Image", "text": _L["stamp_paraph"]}
        type_lbl = labels.get(sk, sk.capitalize())
        card = tk.Frame(c, bg="#1e1e1e", bd=0,
                        highlightthickness=1, highlightbackground="#333")
        card.pack(side=tk.LEFT, padx=4, pady=3, ipadx=4, ipady=2)
        page_lbl = f"p.{sig.get('page',0)+1}"
        name = f"{type_lbl} {num}"
        tk.Label(card, text=f"{name}  {page_lbl}", bg="#1e1e1e", fg="#ccc",
                 font=("Consolas",8)).pack(side=tk.LEFT, padx=(4,2))
        def _goto(s=sig):
            target = s.get("page", S["page"])
            if target != S["page"]: go(target - S["page"])
        def _remove(ii=i): _del_sig(ii)
        tk.Button(card, text="→", bg="#1a2a3a", fg="#7af",
                  font=("Consolas",8), relief=tk.FLAT, cursor="hand2",
                  command=_goto).pack(side=tk.LEFT)
        tk.Button(card, text="✕", bg="#2a1a1a", fg="#f66",
                  font=("Consolas",8), relief=tk.FLAT, cursor="hand2",
                  command=_remove).pack(side=tk.LEFT, padx=(0,2))

def _build_sig_panel():
    """Cree le panel bas repliable (appele juste avant canvas.pack)."""
    outer = tk.Frame(root, bg="#111")
    # Sera packe avec before=canvas apres la creation du canvas
    _sp["outer"] = outer

    # En-tete horizontal
    hdr = tk.Frame(outer, bg="#111")
    hdr.pack(fill=tk.X)
    lbl = tk.Label(hdr, text=_L["stamps_up"].format(n=0), bg="#111", fg="#7af",
                   font=("Consolas",8,"bold"), cursor="hand2", padx=8, pady=3)
    lbl.pack(side=tk.LEFT)
    _sp["hdr_lbl"] = lbl

    # Zone cartes
    content = tk.Frame(outer, bg="#141414")
    content.pack(fill=tk.X)
    _sp["content"] = content
    _sp["expanded"] = True

    def _toggle(e=None):
        if _sp["expanded"]:
            content.pack_forget()
            lbl.config(text=_L["stamps_down"].format(n=len(_placed_sigs)))
        else:
            content.pack(fill=tk.X)
            lbl.config(text=_L["stamps_up"].format(n=len(_placed_sigs)))
        _sp["expanded"] = not _sp["expanded"]

    lbl.bind("<Button-1>", _toggle)
    hdr.bind("<Button-1>", _toggle)

def _clear_sigs():
    _placed_sigs.clear()
    canvas.delete("sig")

def _sig_ctx_menu(ex, ey):
    idx, _ = _sig_hit(ex, ey)
    if idx < 0: return False
    m = tk.Menu(root, tearoff=0, bg="#222", fg="#ccc",
                activebackground="#444", activeforeground="#fff")
    def _del(i=idx):
        canvas.delete(_placed_sigs[i]["item_id"])
        if _placed_sigs[i].get("handle"): canvas.delete(_placed_sigs[i]["handle"])
        _placed_sigs.pop(i)
    m.add_command(label=_L["stamp_delete"], command=_del)
    m.tk_popup(root.winfo_pointerx(), root.winfo_pointery())
    return True

# ── Barre d'onglets (haut) ────────────────────────────────────────────────────
tab_bar = tk.Frame(root, bg=TAB_BG, height=30)
tab_bar.pack(side=tk.TOP, fill=tk.X)
tab_bar.pack_propagate(False)

def _rebuild_tabs():
    for w in tab_bar.winfo_children():
        w.destroy()
    for i, t in enumerate(TABS):
        is_active = (i == ACTIVE[0])
        bg  = BTN_BG if is_active else TAB_BG
        fg  = "#fff"  if is_active else "#666"
        # Titre : onglet actif = depuis S, autres = depuis TABS[i]
        src  = S if is_active else t
        name = os.path.basename(src["title"]) if src["title"] != "Untitled" else _L["tab_untitled"]
        if len(name) > 22: name = name[:20] + "…"

        frm = tk.Frame(tab_bar, bg=bg, padx=2)
        frm.pack(side=tk.LEFT, fill=tk.Y, padx=1, pady=2)

        lbl = tk.Label(frm, text=name, bg=bg, fg=fg,
                       font=("Consolas", 8), cursor="hand2", padx=6)
        lbl.pack(side=tk.LEFT, fill=tk.Y)
        lbl.bind("<Button-1>", lambda e, idx=i: switch_tab(idx))
        if not is_active:
            lbl.bind("<Enter>", lambda e, f=frm: f.config(bg=BTN_HOV))
            lbl.bind("<Leave>", lambda e, f=frm: f.config(bg=TAB_BG))

        # Bouton fermer
        if len(TABS) > 1:
            cl = tk.Label(frm, text="×", bg=bg, fg="#444",
                          font=("Consolas", 9), cursor="hand2", padx=2)
            cl.pack(side=tk.LEFT)
            cl.bind("<Button-1>", lambda e, idx=i: close_tab(idx))
            cl.bind("<Enter>", lambda e, c=cl: c.config(fg="#f55"))
            cl.bind("<Leave>", lambda e, c=cl, b=bg: c.config(fg="#444"))

    # Bouton +
    plus = tk.Label(tab_bar, text=" + ", bg=TAB_BG, fg="#555",
                    font=("Consolas", 10), cursor="hand2")
    plus.pack(side=tk.LEFT, padx=4)
    plus.bind("<Button-1>", lambda e: new_tab())
    plus.bind("<Enter>",    lambda e: plus.config(fg="#ccc"))
    plus.bind("<Leave>",    lambda e: plus.config(fg="#555"))

def new_tab(path=None):
    TABS.append(_make_state())
    switch_tab(len(TABS) - 1)
    if path:
        open_file(path)

def close_tab(idx):
    if len(TABS) == 1:
        return
    is_active = (idx == ACTIVE[0])

    if is_active:
        # Fermer le doc live dans S
        if S["doc"]:
            S["doc"].close()
            S["doc"] = None
        # Marquer ACTIVE hors limites pour que switch_tab ne sauvegarde pas
        ACTIVE[0] = len(TABS)
    else:
        # Fermer le doc stocké
        if TABS[idx]["doc"]:
            TABS[idx]["doc"].close()
            TABS[idx]["doc"] = None
        # Corriger l'index actif si l'onglet fermé est avant lui
        if idx < ACTIVE[0]:
            ACTIVE[0] -= 1

    TABS.pop(idx)
    new_idx = min(idx, len(TABS) - 1)
    switch_tab(new_idx)

def switch_tab(idx):
    global S
    cur = ACTIVE[0]

    # 1. Sauvegarder l'état live dans le slot de l'onglet courant
    if cur < len(TABS):
        for k in S:
            TABS[cur][k] = S[k]
        TABS[cur]["img_item"] = None
        TABS[cur]["ref"]      = None
        # Veille : fermer le doc pour libérer RAM
        if TABS[cur]["doc"]:
            TABS[cur]["doc"].close()
            TABS[cur]["doc"] = None

    ACTIVE[0] = idx

    # 2. Charger l'état stocké dans S
    for k in TABS[idx]:
        S[k] = TABS[idx][k]
    S["img_item"] = None
    S["ref"]      = None

    # 3. Réveil : rouvrir le doc PDF si besoin
    if S["mode"] == "pdf" and S["doc_path"] and not S["doc"]:
        try:    S["doc"] = fitz.open(S["doc_path"])
        except: S["mode"] = None

    # Ne PAS delete "img" → pas de flash noir pendant le rendu du nouvel onglet
    canvas.delete("hint")
    canvas.delete("srch")
    canvas.delete("sel")
    clear_search()

    root.title(f"LOCVIEW — {os.path.basename(S['title'])}" if S["title"] != "Untitled" else "LOCVIEW")
    _rebuild_tabs()
    scrollbar.set(0, 1)
    lbl_page.config(text="")
    lbl_zoom.config(text="100%")

    if S["mode"]:
        render(reset_pan=False)
    else:
        # Onglet vide : cacher l'image sans la supprimer, afficher hint
        if canvas.find_withtag("img"):
            canvas.itemconfig(canvas.find_withtag("img")[0], image="")
        S["img_item"] = None
        cw, ch = get_canvas_size()
        canvas.create_text(cw//2, ch//2, text=_L["hint_drop"],
                           fill="#2a2a2a", font=("Consolas", 20), tags="hint")

def _toggle_lang():
    global LANG, _L
    LANG = "fr" if LANG == "en" else "en"
    _L = STRINGS[LANG]
    data = _load_saves(); data["lang"] = LANG; _write_saves(data)
    _apply_lang()

# ── Barre du bas ──────────────────────────────────────────────────────────────
bar = tk.Frame(root, bg=BAR_BG, height=36)
bar.pack(side=tk.BOTTOM, fill=tk.X)
bar.pack_propagate(False)

_bar_left   = tk.Frame(bar, bg=BAR_BG)
_bar_right  = tk.Frame(bar, bg=BAR_BG)
_bar_center = tk.Frame(bar, bg=BAR_BG)
_bar_left.pack(side=tk.LEFT, fill=tk.Y)
_bar_right.pack(side=tk.RIGHT, fill=tk.Y)
_bar_center.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
_bar_ci = tk.Frame(_bar_center, bg=BAR_BG)
_bar_ci.place(relx=0.5, rely=0.5, anchor="center")

# ── Boutons gauche ────────────────────────────────────────────────────────────
_btn_open  = make_btn(_bar_left, _L["btn_open"],       lambda: open_dialog());       _btn_open.pack(side=tk.LEFT, padx=4, pady=4)
_btn_print = make_btn(_bar_left, _L["btn_print"],      lambda: do_print());          _btn_print.pack(side=tk.LEFT, padx=2, pady=4)
_btn_fs    = make_btn(_bar_left, _L["btn_fullscreen"], lambda: toggle_fullscreen()); _btn_fs.pack(side=tk.LEFT, padx=2, pady=4)
tk.Frame(_bar_left, bg="#444", width=1).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=4)
_btn_search = make_btn(_bar_left, _L["btn_search"], lambda: toggle_search()); _btn_search.pack(side=tk.LEFT, padx=2, pady=4)

# ── Boutons centre (navigation + zoom) ───────────────────────────────────────
_btn_prev = make_btn(_bar_ci, _L["btn_prev"], lambda: go(-1)); _btn_prev.pack(side=tk.LEFT, padx=2, pady=4)
_btn_next = make_btn(_bar_ci, _L["btn_next"], lambda: go(+1)); _btn_next.pack(side=tk.LEFT, padx=2, pady=4)
tk.Frame(_bar_ci, bg="#444", width=1).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=4)
make_btn(_bar_ci, "Zoom +", lambda: do_zoom(1.25)).pack(side=tk.LEFT, padx=2, pady=4)
make_btn(_bar_ci, "Zoom -", lambda: do_zoom(0.8)).pack(side=tk.LEFT,  padx=2, pady=4)
lbl_zoom = tk.Label(_bar_ci, text="100%", bg=BAR_BG, fg="#aaa",
                    font=("Consolas", 9, "bold"), width=5)
lbl_zoom.pack(side=tk.LEFT, padx=4)

# ── Labels et boutons droite ──────────────────────────────────────────────────
lbl_page = tk.Label(_bar_right, text="", bg=BAR_BG, fg="#aaa", font=("Consolas", 9, "bold"))
lbl_page.pack(side=tk.RIGHT, padx=10)

lbl_save_status = tk.Label(_bar_right, text="", bg=BAR_BG, fg="#7f7", font=("Consolas", 8))
lbl_save_status.pack(side=tk.RIGHT, padx=2)

def _embed_sigs_to_pdf():
    """Grave les tampons canvas dans les pages PDF correspondantes."""
    import io
    for sig in list(_placed_sigs):
        page_idx = sig.get("page", 0)
        if page_idx >= len(S["doc"]): continue
        page = S["doc"][page_idx]
        r = sig.get("pdf_rect")
        if not r: continue
        if sig["kind"] == "image":
            pil = sig.get("orig_pil")
            if not pil: continue
            buf = io.BytesIO()
            pil.convert("RGBA").save(buf, format="PNG")
            buf.seek(0)
            try:
                page.insert_image(fitz.Rect(r[0], r[1], r[2], r[3]),
                                  stream=buf.read(), overlay=True)
            except Exception: pass
        elif sig["kind"] == "text":
            col_hex = sig.get("color", "#000000").lstrip("#")
            try:
                rc = (int(col_hex[0:2],16)/255,
                      int(col_hex[2:4],16)/255,
                      int(col_hex[4:6],16)/255)
            except Exception:
                rc = (0, 0, 0)
            font_pt = sig.get("size", 24) / max(0.1, S.get("scale", 1.0))
            try:
                page.insert_text((r[0], r[3]), sig["text"],
                                  fontsize=max(6, font_pt),
                                  color=rc, overlay=True)
            except Exception: pass

def _save_file():
    if not S.get("doc") or not S.get("doc_path"): return
    try:
        import shutil
        _embed_sigs_to_pdf()
        tmp = S["doc_path"] + ".save_tmp"
        S["doc"].save(tmp)
        S["doc"].close()
        shutil.move(tmp, S["doc_path"])
        S["doc"] = fitz.open(S["doc_path"])
        S["pil_cache"].clear()
        render()
        lbl_save_status.config(text=_L["status_file_saved"], fg="#7f7")
        root.after(2500, lambda: lbl_save_status.config(text=""))
    except Exception as e:
        lbl_save_status.config(text=_L["status_save_error"].format(msg=str(e)[:55]), fg="#f88")
        root.after(4000, lambda: lbl_save_status.config(text=""))

_btn_lang  = make_btn(_bar_right, _L["lang_btn"],     _toggle_lang);           _btn_lang.pack(side=tk.RIGHT, padx=6, pady=4)
_btn_save  = make_btn(_bar_right, _L["btn_save"],     lambda: _save_file());    _btn_save.pack(side=tk.RIGHT, padx=2, pady=4)
_btn_bmark = make_btn(_bar_right, _L["btn_bookmark"], lambda: save_position()); _btn_bmark.pack(side=tk.RIGHT, padx=6, pady=4)
tk.Frame(_bar_right, bg="#444", width=1).pack(side=tk.RIGHT, fill=tk.Y, padx=6, pady=4)
make_btn(_bar_right, "✍ Signature", lambda: open_sig_editor()).pack(side=tk.RIGHT, padx=2, pady=4)
_btn_edit_mode = make_btn(_bar_right, _L["btn_edit"], lambda: _toggle_edit_mode()); _btn_edit_mode.pack(side=tk.RIGHT, padx=2, pady=4)
btn_undo = tk.Button(_bar_right, text=_L["btn_undo_toolbar"], command=lambda: _undo_edit(),
                     bg=BTN_BG, fg="#555", relief=tk.FLAT, state=tk.DISABLED,
                     font=("Consolas", 9), cursor="arrow", padx=10, pady=2)
btn_undo.pack(side=tk.RIGHT, padx=2, pady=4)
btn_undo.bind("<Enter>", lambda e: btn_undo.config(bg=BTN_HOV, fg="#fff") if _edit_history else None)
btn_undo.bind("<Leave>", lambda e: btn_undo.config(bg=BTN_BG,  fg="#555") if not _edit_history else btn_undo.config(bg=BTN_BG, fg=BTN_FG))

# ── Barre de recherche (masquée par défaut) ───────────────────────────────────
search_frame = tk.Frame(root, bg="#1a1a1a", height=34)
search_frame.pack_propagate(False)
# (pas packée pour l'instant)

_lbl_search = tk.Label(search_frame, text=_L["search_label"], bg="#1a1a1a", fg="#aaa",
                       font=("Consolas", 9))
_lbl_search.pack(side=tk.LEFT, padx=8)
search_var = tk.StringVar()
search_entry = tk.Entry(search_frame, textvariable=search_var,
                        bg="#2e2e2e", fg="#fff", insertbackground="#fff",
                        relief=tk.FLAT, font=("Consolas", 10), width=24)
search_entry.pack(side=tk.LEFT, padx=4, pady=6, ipady=2)

# Options de recherche
search_case_var = tk.BooleanVar(value=False)
search_word_var = tk.BooleanVar(value=False)

def _make_search_chk(parent, text, var):
    chk = tk.Checkbutton(parent, text=text, variable=var,
                         bg="#1a1a1a", fg="#888", selectcolor="#2e2e2e",
                         activebackground="#1a1a1a", activeforeground="#fff",
                         font=("Consolas", 8), cursor="hand2",
                         command=lambda: do_search())
    chk.bind("<Enter>", lambda e: chk.config(fg="#fff"))
    chk.bind("<Leave>", lambda e: chk.config(fg="#888" if not var.get() else "#e8820a"))
    var.trace_add("write", lambda *_: chk.config(fg="#e8820a" if var.get() else "#888"))
    return chk

_make_search_chk(search_frame, "Aa", search_case_var).pack(side=tk.LEFT, padx=2)
_chk_whole_word = _make_search_chk(search_frame, _L["search_whole_word"], search_word_var)
_chk_whole_word.pack(side=tk.LEFT, padx=2)

lbl_search_res = tk.Label(search_frame, text="", bg="#1a1a1a", fg="#888",
                           font=("Consolas", 9))
lbl_search_res.pack(side=tk.LEFT, padx=6)

make_btn(search_frame, "▲", lambda: search_navigate(-1)).pack(side=tk.LEFT, padx=2, pady=4)
make_btn(search_frame, "▼", lambda: search_navigate(+1)).pack(side=tk.LEFT, padx=2, pady=4)
make_btn(search_frame, "✕", lambda: toggle_search()).pack(side=tk.LEFT, padx=6, pady=4)

# ── Panneau de texte sélectionné (fixe, masqué par défaut) ───────────────────
sel_panel = tk.Frame(root, bg="#181818", height=90)
sel_panel.pack_propagate(False)
# pas packée pour l'instant

_sel_hdr = tk.Frame(sel_panel, bg="#212121")
_sel_hdr.pack(fill=tk.X)
_lbl_sel_hdr = tk.Label(_sel_hdr, text=_L["sel_header"], bg="#212121", fg="#666",
                        font=("Consolas", 8), pady=3, padx=8)
_lbl_sel_hdr.pack(side=tk.LEFT)

def _copy_sel_text():
    t = sel_txt.get("1.0", tk.END).strip()
    root.clipboard_clear(); root.clipboard_append(t)
    _btn_copy_lbl.config(text=_L["sel_copied"], fg="#7f7")
    root.after(1400, lambda: _btn_copy_lbl.config(text=_L["sel_copy"], fg="#aaa"))

_btn_copy_lbl = tk.Label(_sel_hdr, text=_L["sel_copy"], bg="#212121", fg="#aaa",
                          font=("Consolas", 8), pady=3, padx=10, cursor="hand2")
_btn_copy_lbl.pack(side=tk.RIGHT, padx=2)
_btn_copy_lbl.bind("<Button-1>", lambda e: _copy_sel_text())
_btn_copy_lbl.bind("<Enter>", lambda e: _btn_copy_lbl.config(fg="#fff"))
_btn_copy_lbl.bind("<Leave>", lambda e: _btn_copy_lbl.config(fg="#aaa"))

_btn_close_sel = tk.Label(_sel_hdr, text=" ✕ ", bg="#212121", fg="#555",
                           font=("Consolas", 9), pady=3, cursor="hand2")
_btn_close_sel.pack(side=tk.RIGHT)
_btn_close_sel.bind("<Button-1>", lambda e: _hide_sel_panel())
_btn_close_sel.bind("<Enter>", lambda e: _btn_close_sel.config(fg="#f55"))
_btn_close_sel.bind("<Leave>", lambda e: _btn_close_sel.config(fg="#555"))

sel_txt = tk.Text(sel_panel, bg="#141414", fg="#e0e0e0", insertbackground="#fff",
                  relief=tk.FLAT, font=("Consolas", 10),
                  wrap=tk.WORD, padx=10, pady=5,
                  selectbackground="#2255aa", selectforeground="#fff",
                  height=4)
sel_txt.pack(fill=tk.BOTH, expand=True, padx=1, pady=(0,1))
sel_txt.bind("<Escape>", lambda e: _hide_sel_panel())

def _show_sel_panel(text):
    sel_txt.config(state=tk.NORMAL)
    sel_txt.delete("1.0", tk.END)
    if text:
        sel_txt.insert("1.0", text)
        sel_txt.config(fg="#e0e0e0")
    else:
        sel_txt.insert("1.0", _L["sel_no_text"])
        sel_txt.config(fg="#555")
    _btn_copy_lbl.config(text=_L["sel_copy"], fg="#aaa")
    if not sel_panel.winfo_ismapped():
        sel_panel.pack(side=tk.BOTTOM, fill=tk.X, before=canvas)

def _hide_sel_panel():
    if sel_panel.winfo_ismapped():
        sel_panel.pack_forget()
    canvas.delete("sel")

SEARCH = dict(results=[], idx=0, active=False, text="")

def toggle_search():
    if SEARCH["active"]:
        search_frame.pack_forget()
        SEARCH["active"] = False
        clear_search()
        canvas.focus_set()
    else:
        search_frame.pack(side=tk.TOP, fill=tk.X, before=canvas)
        SEARCH["active"] = True
        search_entry.focus_set()
        search_entry.select_range(0, tk.END)

def _find_in_page(page, text, case_sensitive, whole_word):
    import re, unicodedata

    def _norm(s):
        s = unicodedata.normalize("NFKC", s)
        s = re.sub(r'[­‐‑‒–—−﹘﹣－]', '-', s)
        s = re.sub(r"[‘’ʼ´]", "'", s)
        return s

    text_n = _norm(text)
    if not text_n:
        return []
    text_cmp = text_n if case_sensitive else text_n.lower()

    if whole_word and ' ' not in text_n:
        results = []
        for w in page.get_text("words"):
            wt_cmp = _norm(w[4]) if case_sensitive else _norm(w[4]).lower()
            if wt_cmp == text_cmp:
                results.append(fitz.Rect(w[0], w[1], w[2], w[3]))
        return results

    rects = page.search_for(text)
    if not rects and text_n != text:
        rects = page.search_for(text_n)
    if not case_sensitive and not whole_word:
        return rects

    flags = 0 if case_sensitive else re.IGNORECASE
    pat = (r'(?<!\w)' + re.escape(text_cmp) + r'(?!\w)') if whole_word else re.escape(text_cmp)
    return [r for r in rects
            if re.search(pat, _norm(page.get_textbox(r)) if case_sensitive
                         else _norm(page.get_textbox(r)).lower(), flags)]

def do_search(*_):
    if S["mode"] != "pdf" or not S["doc"]:
        return
    text = search_var.get().strip()
    if not text:
        clear_search()
        return
    case_sens  = search_case_var.get()
    whole_word = search_word_var.get()
    SEARCH["text"] = text
    SEARCH["results"] = []
    for p in range(S["total"]):
        try:
            rects = _find_in_page(S["doc"][p], text, case_sens, whole_word)
            for r in rects:
                SEARCH["results"].append((p, r))
        except Exception:
            pass
    SEARCH["idx"] = 0
    if SEARCH["results"]:
        page, _ = SEARCH["results"][0]
        S["page"] = page
        render(reset_pan=True)
        lbl_search_res.config(text=f"1 / {len(SEARCH['results'])}", fg="#8f8")
    else:
        lbl_search_res.config(text=_L["search_no_results"], fg="#f88")
        draw_highlights()

def search_navigate(delta):
    if not SEARCH["results"]:
        return
    SEARCH["idx"] = (SEARCH["idx"] + delta) % len(SEARCH["results"])
    page, _ = SEARCH["results"][SEARCH["idx"]]
    lbl_search_res.config(text=f"{SEARCH['idx']+1} / {len(SEARCH['results'])}", fg="#8f8")
    if page != S["page"]:
        S["page"] = page
        render(reset_pan=True)
    else:
        draw_highlights()

def clear_search():
    SEARCH["results"] = []
    SEARCH["idx"]     = 0
    SEARCH["text"]    = ""
    canvas.delete("srch")
    lbl_search_res.config(text="")

def draw_highlights():
    canvas.delete("srch")
    if not SEARCH["results"] or S["mode"] != "pdf":
        return
    cw, ch = get_canvas_size()
    ox = cw//2 + S["pan_x"] - S["pw"] * S["scale"] / 2
    oy = ch//2 + S["pan_y"] - S["ph"] * S["scale"] / 2
    for i, (page, rect) in enumerate(SEARCH["results"]):
        if page != S["page"]:
            continue
        x1 = ox + rect.x0 * S["scale"]
        y1 = oy + rect.y0 * S["scale"]
        x2 = ox + rect.x1 * S["scale"]
        y2 = oy + rect.y1 * S["scale"]
        color = "#ff0" if i == SEARCH["idx"] else "#fa0"
        canvas.create_rectangle(x1, y1, x2, y2,
                                fill=color, stipple="gray50",
                                outline=color, width=1, tags="srch")

_search_debounce = [None]
def _search_delayed(*_):
    if _search_debounce[0]:
        root.after_cancel(_search_debounce[0])
    _search_debounce[0] = root.after(300, do_search)
search_var.trace_add("write", _search_delayed)
search_entry.bind("<Return>",       lambda e: search_navigate(+1))
search_entry.bind("<Shift-Return>", lambda e: search_navigate(-1))

# ── Barre de défilement personnalisée ────────────────────────────────────────
_SBAR_AH = 20  # hauteur des zones flèche haut/bas

class PageScrollbar(tk.Canvas):
    def __init__(self, parent):
        super().__init__(parent, width=SBAR_W, bg="#141414",
                         highlightthickness=0, cursor="arrow")
        self._first = 0.0
        self._last  = 1.0
        self._popup = None
        self._pop_img_lbl = None
        self._pop_num_lbl = None
        self._last_prev_page = -1
        self._dragging    = False
        self._drag_offset = 0
        self._hovered     = False
        self._arrow_hover = None  # "up" | "down" | None
        self._repeat_job  = None
        self.bind("<Configure>",       lambda e: self._draw())
        self.bind("<ButtonPress-1>",   self._press)
        self.bind("<B1-Motion>",       self._drag)
        self.bind("<ButtonRelease-1>", self._release)
        self.bind("<Motion>",          self._motion)
        self.bind("<Enter>",           self._on_enter)
        self.bind("<Leave>",           self._leave)

    def set(self, first, last):
        self._first = float(first)
        self._last  = float(last)
        self._draw()

    def _draw(self):
        self.delete("all")
        h  = max(self.winfo_height(), 1)
        w  = max(self.winfo_width(), 1)
        AH = _SBAR_AH
        track_h   = max(1, h - 2 * AH)
        active    = self._hovered or self._dragging
        bg_col    = "#1a1a1a" if active else "#141414"
        notch_col = "#444"    if active else "#333"
        thumb_col = "#555"    if self._dragging else ("#4a4a4a" if active else "#383838")
        border_col= "#888"    if self._dragging else ("#777"    if active else "#555")

        self.create_rectangle(0, 0, w, h, fill=bg_col, outline="")
        self.create_line(0, 0, 0, h, fill="#333", width=1)

        # Flèche haut
        up_hov = self._arrow_hover == "up"
        up_bg  = "#484848" if up_hov else "#2a2a2a"
        self.create_rectangle(0, 0, w, AH, fill=up_bg, outline="")
        self.create_rectangle(2, 2, w-2, AH-2, fill=up_bg,
                              outline="#666" if up_hov else "#444", width=1)
        self.create_line(0, AH, w, AH, fill="#383838", width=1)
        self.create_text(w//2, AH//2, text="▲",
                         fill="#ddd" if up_hov else "#777",
                         font=("Consolas", 8))

        # Flèche bas
        dn_hov = self._arrow_hover == "down"
        dn_bg  = "#484848" if dn_hov else "#2a2a2a"
        self.create_rectangle(0, h - AH, w, h, fill=dn_bg, outline="")
        self.create_rectangle(2, h - AH + 2, w-2, h-2, fill=dn_bg,
                              outline="#666" if dn_hov else "#444", width=1)
        self.create_line(0, h - AH, w, h - AH, fill="#383838", width=1)
        self.create_text(w//2, h - AH//2, text="▼",
                         fill="#ddd" if dn_hov else "#777",
                         font=("Consolas", 8))

        # Encoches de pages (zone track seulement)
        total = S["total"] if S["mode"] == "pdf" else len(S["imgs"])
        if total > 1:
            for i in range(total):
                fy = AH + int((i / total) * track_h)
                notch_w = 6 if (total <= 60) else 4
                self.create_line(w - notch_w, fy, w, fy, fill=notch_col, width=1)

        # Thumb
        ty = AH + int(self._first * track_h)
        th = max(28, int((self._last - self._first) * track_h))
        if ty + th > h - AH: th = max(0, h - AH - ty)
        if th > 4:
            self.create_rectangle(5, ty+2, w-5, ty+th-2,
                                  fill=thumb_col, outline=border_col, width=1,
                                  tags="thumb")
            if th > 16 and S["total"] > 0:
                txt_col = "#ccc" if active else "#888"
                self.create_text(w//2, ty + th//2, text=str(S["page"]+1),
                                 fill=txt_col, font=("Consolas", 7))

    def _track_frac(self, y):
        AH = _SBAR_AH
        h  = max(self.winfo_height(), 1)
        return max(0.0, min(1.0, (y - AH) / max(1, h - 2 * AH)))

    def _in_arrow(self, y):
        AH = _SBAR_AH
        if y < AH: return "up"
        if y > self.winfo_height() - AH: return "down"
        return None

    def _page_for(self, frac):
        if S["mode"] == "pdf" and S["total"] > 0:
            return max(0, min(int(frac * S["total"]), S["total"] - 1))
        if S["mode"] == "img" and S["imgs"]:
            return max(0, min(int(frac * len(S["imgs"])), len(S["imgs"]) - 1))
        return 0

    def _goto(self, frac):
        if S["mode"] == "pdf":
            S["page"] = self._page_for(frac)
            render(reset_pan=True)
        elif S["mode"] == "img":
            S["img_idx"] = self._page_for(frac)
            render(reset_pan=True)

    def _press(self, e):
        arrow = self._in_arrow(e.y)
        if arrow:
            self._do_arrow(arrow)
            self._repeat_job = root.after(400, lambda: self._repeat(arrow))
            return
        self._dragging = True
        AH = _SBAR_AH
        h  = max(self.winfo_height(), 1)
        track_h = max(1, h - 2 * AH)
        ty = AH + int(self._first * track_h)
        th = max(28, int((self._last - self._first) * track_h))
        if ty <= e.y <= ty + th:
            self._drag_offset = e.y - ty
        else:
            self._drag_offset = th // 2
            self._goto(self._track_frac(e.y))
        self._show_popup(e)

    def _do_arrow(self, direction):
        scroll_by(+80 if direction == "up" else -80)

    def _repeat(self, direction):
        self._do_arrow(direction)
        self._repeat_job = root.after(80, lambda: self._repeat(direction))

    def _stop_repeat(self):
        if self._repeat_job:
            root.after_cancel(self._repeat_job)
            self._repeat_job = None

    def _drag(self, e):
        if not self._dragging: return
        AH = _SBAR_AH
        h  = max(self.winfo_height(), 1)
        track_h = max(1, h - 2 * AH)
        size = self._last - self._first
        frac = max(0.0, min(1.0 - size, (e.y - AH - self._drag_offset) / track_h))
        self._first = frac
        self._last  = frac + size
        self._draw()
        self._goto(frac)
        self._show_popup(e)

    def _on_enter(self, e):
        self._hovered = True
        self._draw()

    def _release(self, e):
        self._stop_repeat()
        self._dragging = False
        self._draw()

    def _motion(self, e):
        new_hover = self._in_arrow(e.y)
        if new_hover != self._arrow_hover:
            self._arrow_hover = new_hover
            self._draw()
        if not self._dragging:
            self._show_popup(e)

    def _leave(self, e):
        self._hovered = False
        self._arrow_hover = None
        self._stop_repeat()
        if not self._dragging:
            self._hide_popup()
        self._draw()

    def _show_popup(self, e):
        if S["mode"] not in ("pdf", "img"):
            return
        page_idx = self._page_for(self._track_frac(e.y))
        if self._popup is None:
            self._popup = tk.Toplevel(root)
            self._popup.overrideredirect(True)
            self._popup.configure(bg="#2a2a2a")
            outer = tk.Frame(self._popup, bg="#555", bd=1)
            outer.pack(fill=tk.BOTH, expand=True)
            inner = tk.Frame(outer, bg="#1e1e1e")
            inner.pack(padx=1, pady=1, fill=tk.BOTH, expand=True)
            self._pop_img_lbl = tk.Label(inner, bg="#1e1e1e", bd=0,
                                          width=PREV_W, height=PREV_H)
            self._pop_img_lbl.pack(padx=6, pady=6)
            self._pop_num_lbl = tk.Label(inner, bg="#252525", fg="#eee",
                                          font=("Consolas", 10, "bold"), pady=5)
            self._pop_num_lbl.pack(fill=tk.X)
        rx = self.winfo_rootx()
        ry = e.y_root
        pw = PREV_W + 20
        ph = PREV_H + 42
        px = max(0, rx - pw - 6)
        py = max(0, ry - ph // 2)
        self._popup.geometry(f"{pw}x{ph}+{px}+{py}")
        if S["mode"] == "pdf":
            self._pop_num_lbl.config(text=f"Page  {page_idx+1} / {S['total']}")
        else:
            self._pop_num_lbl.config(text=f"Image  {page_idx+1} / {len(S['imgs'])}")
        if page_idx != self._last_prev_page:
            self._last_prev_page = page_idx
            threading.Thread(target=self._render_thumb,
                             args=(page_idx,), daemon=True).start()

    def _render_thumb(self, page_idx):
        try:
            if S["mode"] == "pdf" and S["doc_path"]:
                doc  = fitz.open(S["doc_path"])
                page = doc[page_idx]
                pw, ph = page.rect.width, page.rect.height
                scale = min(PREV_W/pw, PREV_H/ph)
                pix  = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                pil  = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                doc.close()
            elif S["mode"] == "img" and S["imgs"]:
                pil = Image.open(S["imgs"][page_idx]).convert("RGB")
                pil.thumbnail((PREV_W, PREV_H), Image.LANCZOS)
            else:
                return
            root.after(0, lambda p=pil, i=page_idx: self._set_thumb(p, i))
        except Exception:
            pass

    def _set_thumb(self, pil, page_idx):
        if page_idx != self._last_prev_page or self._popup is None:
            return
        tk_img = ImageTk.PhotoImage(pil)
        self._pop_img_lbl.config(image=tk_img, text="")
        self._pop_img_lbl._ref = tk_img

    def _hide_popup(self):
        if self._popup:
            self._popup.destroy()
            self._popup = None
        self._last_prev_page = -1

scrollbar = PageScrollbar(root)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

canvas = tk.Canvas(root, bg=BG, highlightthickness=0, cursor="fleur")
canvas.pack(fill=tk.BOTH, expand=True)
canvas.config(takefocus=True)
_build_sig_panel()
_sp["outer"].pack(side=tk.BOTTOM, fill=tk.X, before=canvas)
_hint_txt = canvas.create_text(0, 0, text=_L["hint_drop"],
                               fill="#2a2a2a", font=("Consolas", 28), anchor="center", tags="hint")
_hint_btn = tk.Button(canvas, text=_L["hint_open"], command=lambda: open_dialog(),
                      bg="#2a2a2a", fg="#555", relief=tk.FLAT,
                      font=("Consolas", 14), cursor="hand2", padx=22, pady=8)
_hint_btn.bind("<Enter>", lambda e: _hint_btn.config(bg="#3a3a3a", fg="#aaa"))
_hint_btn.bind("<Leave>", lambda e: _hint_btn.config(bg="#2a2a2a", fg="#555"))
_hint_win = canvas.create_window(0, 0, anchor="center", window=_hint_btn, tags="hint")

def _place_hint():
    if S["mode"]: return
    w = max(canvas.winfo_width(), 300)
    h = max(canvas.winfo_height(), 300)
    cx, cy = w // 2, h // 2
    canvas.coords(_hint_txt, cx, cy - 30)
    canvas.coords(_hint_win, cx, cy + 50)

# ── Rendu ─────────────────────────────────────────────────────────────────────
def get_canvas_size():
    w, h = canvas.winfo_width(), canvas.winfo_height()
    if w <= 1 or h <= 1:          # seulement au démarrage, avant le premier layout
        canvas.update_idletasks()
        w, h = canvas.winfo_width(), canvas.winfo_height()
    return max(w, 300), max(h, 300)

def display(pil_img, rid):
    if rid != S["render_id"]:
        return
    tk_img = ImageTk.PhotoImage(pil_img)
    cw, ch = get_canvas_size()
    edge = _page_edge[0]
    _page_edge[0] = None
    img_h = pil_img.size[1]
    if edge == "top" and img_h > ch:
        S["pan_y"] =  (img_h - ch) // 2   # haut de l'image visible
    elif edge == "bot" and img_h > ch:
        S["pan_y"] = -((img_h - ch) // 2)  # bas de l'image visible
    x = cw//2 + S["pan_x"]
    y = ch//2 + S["pan_y"]
    canvas.delete("hint")
    if S["img_item"] is None:
        existing = canvas.find_withtag("img")
        if existing:
            S["img_item"] = existing[0]
            canvas.itemconfig(S["img_item"], image=tk_img)
            canvas.coords(S["img_item"], x, y)
        else:
            S["img_item"] = canvas.create_image(x, y, anchor="center",
                                                 image=tk_img, tags="img")
    else:
        canvas.itemconfig(S["img_item"], image=tk_img)
        canvas.coords(S["img_item"], x, y)
    # Garder l'ancienne ref vivante jusqu'après itemconfig pour éviter le GC flash
    S["prev_ref"] = S["ref"]
    S["ref"] = tk_img
    if S["mode"] in ("pdf", "txt"):
        lbl_page.config(text=f"{S['page']+1} / {S['total']}")
        lbl_zoom.config(text=f"{int(S['zoom']*100)}%")
        if S["total"] > 0:
            scrollbar.set(S["page"]/S["total"], (S["page"]+1)/S["total"])
    elif S["mode"] == "img":
        n = len(S["imgs"])
        lbl_page.config(text=f"{S['img_idx']+1} / {n}")
        lbl_zoom.config(text=f"{int(S['zoom']*100)}%")
        if n > 0:
            scrollbar.set(S["img_idx"]/n, (S["img_idx"]+1)/n)
    draw_highlights()
    if _placed_sigs:
        canvas.tag_raise("sig")
        for sig in _placed_sigs: _draw_sig_handle(sig)
        canvas.tag_raise("sig_handle")
    # Lancer le pré-rendu des pages adjacentes 300ms après l'affichage
    if S["mode"] == "pdf":
        root.after(300, _preload_neighbors)

def _render_one_page(doc, page_idx, cw, ch):
    """Rastérise une page PDF ; retourne (key, pil_image)."""
    page   = doc[page_idx]
    pw, ph = page.rect.width, page.rect.height
    fit, zoom = S["fit"], S["zoom"]
    if fit == "page":    scale = min(cw/pw, ch/ph) * zoom
    elif fit == "width": scale = (cw/pw) * zoom
    else:                scale = zoom
    scale = max(0.05, min(scale, 8.0))
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    key = (page_idx, round(zoom, 3), cw, ch, fit)
    return key, pil, scale, pw, ph

def _cache_put(page_idx, key, pil):
    """Insère dans le cache PIL avec éviction LRU."""
    with S["cache_lock"]:
        S["pil_cache"].pop(page_idx, None)   # retire pour remettre en fin (LRU)
        S["pil_cache"][page_idx] = (key, pil)
        while len(S["pil_cache"]) > _CACHE_MAX:
            S["pil_cache"].pop(next(iter(S["pil_cache"])))

def _do_pdf_render(page_idx, cw, ch, rid, doc_path):
    if rid != S["render_id"]: return
    with S["cache_lock"]:
        key    = (page_idx, round(S["zoom"],3), cw, ch, S["fit"])
        cached = S["pil_cache"].get(page_idx)
        if cached and cached[0] == key:
            # Remonter en tête LRU
            S["pil_cache"].pop(page_idx)
            S["pil_cache"][page_idx] = cached
            root.after(0, lambda p=cached[1], r=rid: display(p, r))
            _queue_render(lambda: _prefetch_pages(page_idx, cw, ch, rid, doc_path))
            return
    try:
        doc = _get_worker_doc(doc_path)
        key, pil, scale, pw, ph = _render_one_page(doc, page_idx, cw, ch)
        S["scale"] = scale; S["pw"] = pw; S["ph"] = ph
        _cache_put(page_idx, key, pil)
        if rid == S["render_id"]:
            root.after(0, lambda p=pil, r=rid: display(p, r))
            _queue_render(lambda: _prefetch_pages(page_idx, cw, ch, rid, doc_path))
    except Exception:
        pass

def _prefetch_pages(current, cw, ch, rid, doc_path):
    """Pré-charge les pages voisines dans le cache après un rendu réussi."""
    if rid != S["render_id"]: return
    zoom, fit = round(S["zoom"], 3), S["fit"]
    doc = _get_worker_doc(doc_path)
    for delta in (1, -1, 2, -2):
        if rid != S["render_id"]: return
        target = current + delta
        if not (0 <= target < S["total"]): continue
        key = (target, zoom, cw, ch, fit)
        with S["cache_lock"]:
            cached = S["pil_cache"].get(target)
            if cached and cached[0] == key: continue
        try:
            k, pil, _, _, _ = _render_one_page(doc, target, cw, ch)
            _cache_put(target, k, pil)
        except Exception:
            pass

def render_img_bg(idx, cw, ch, rid):
    try:
        img    = Image.open(S["imgs"][idx]).convert("RGB")
        iw, ih = img.size
        if S["fit"] == "page":   scale = min(cw/iw, ch/ih) * S["zoom"]
        elif S["fit"] == "width": scale = (cw/iw) * S["zoom"]
        else:                     scale = S["zoom"]
        scale = max(0.05, scale)
        nw, nh = max(1,int(iw*scale)), max(1,int(ih*scale))
        pil = img.resize((nw,nh), Image.LANCZOS if scale<1 else Image.BILINEAR)
        if rid == S["render_id"]:
            root.after(0, lambda p=pil, r=rid: display(p, r))
    except Exception:
        pass

# ── Rendu texte (TXT / MD / DOCX / ODT / PPTX / HTML) ───────────────────────
_TXT_FONTS = {}
def _get_txt_fonts():
    if _TXT_FONTS: return _TXT_FONTS
    b = "C:/Windows/Fonts/"
    try:
        _TXT_FONTS["body"]   = ImageFont.truetype(b+"calibri.ttf",  14)
        _TXT_FONTS["h1"]     = ImageFont.truetype(b+"calibrib.ttf", 22)
        _TXT_FONTS["h2"]     = ImageFont.truetype(b+"calibrib.ttf", 18)
        _TXT_FONTS["h3"]     = ImageFont.truetype(b+"calibrib.ttf", 16)
        _TXT_FONTS["code"]   = ImageFont.truetype(b+"consola.ttf",  12)
    except Exception:
        d = ImageFont.load_default()
        for k in ("body","h1","h2","h3","code"): _TXT_FONTS[k] = d
    _TXT_FONTS["bullet"] = _TXT_FONTS["body"]
    return _TXT_FONTS

def _parse_md(text):
    """Convertit du texte Markdown en liste de (style, texte)."""
    result = []; in_code = False
    for line in text.splitlines():
        if line.startswith("```"):
            in_code = not in_code; continue
        if in_code:
            result.append(("code", line)); continue
        if line.startswith("### "):   result.append(("h3", line[4:]))
        elif line.startswith("## "): result.append(("h2", line[3:]))
        elif line.startswith("# "):  result.append(("h1", line[2:]))
        elif line.strip() in ("---","===","***"): result.append(("hr", ""))
        elif line.startswith(("- ","* ","+ ")): result.append(("bullet", line[2:]))
        else: result.append(("body", line))
    return result

def _wrap_styled(styled_lines):
    """Enveloppe les lignes et pagine par budget de hauteur."""
    content_h = _TXT_PH - 2*_TXT_MG
    cpl_body  = (_TXT_PW - 2*_TXT_MG) // 8   # ~89 chars par ligne
    cpl_code  = (_TXT_PW - 2*_TXT_MG) // 7
    wrapped = []
    for style, text in styled_lines:
        lh = _TXT_LH.get(style, 21)
        if style == "hr":
            wrapped.append(("hr", "", lh)); continue
        cpl = cpl_code if style == "code" else cpl_body - (2 if style=="bullet" else 0)
        lines = textwrap.wrap(text, cpl) or [""]
        for i, ln in enumerate(lines):
            s = style if (style != "bullet" or i == 0) else "body_i"
            wrapped.append((s, ln, _TXT_LH.get(s, lh)))
    pages = []; page = []; used = 0
    for item in wrapped:
        h = item[2]
        if used + h > content_h and page:
            pages.append(page); page = []; used = 0
        page.append(item); used += h
    if page: pages.append(page)
    return pages or [[("body","",21)]]

def _do_txt_render(page_idx, cw, ch, rid):
    if rid != S["render_id"]: return
    pages = S["txt_pages"]
    if not pages: return
    page_lines = pages[min(page_idx, len(pages)-1)]
    F = _get_txt_fonts()
    img  = Image.new("RGB", (_TXT_PW, _TXT_PH), _TXT_BG)
    draw = ImageDraw.Draw(img)
    y = _TXT_MG
    for style, text, lh in page_lines:
        x = _TXT_MG
        if style == "hr":
            mid = y + lh//2
            draw.line([(x, mid), (_TXT_PW-x, mid)], fill=_TXT_RULE_C, width=1)
        elif style == "code":
            draw.rectangle([x-4, y, _TXT_PW-x+4, y+lh], fill=_TXT_CODE_BG)
            draw.text((x+4, y+2), text, fill=_TXT_CODE_FG, font=F["code"])
        elif style == "bullet":
            draw.text((x, y), "•", fill=_TXT_BULL_C, font=F["body"])
            draw.text((x+16, y), text, fill=_TXT_FG, font=F["body"])
        elif style == "body_i":
            draw.text((x+16, y), text, fill=_TXT_FG, font=F["body"])
        elif style in _TXT_H_CLR:
            draw.text((x, y), text, fill=_TXT_H_CLR[style], font=F[style])
        else:
            draw.text((x, y), text, fill=_TXT_FG, font=F["body"])
        y += lh
    # Mise à l'échelle canvas
    if S["fit"] == "page":   scale = min(cw/_TXT_PW, ch/_TXT_PH) * S["zoom"]
    elif S["fit"] == "width": scale = (cw/_TXT_PW) * S["zoom"]
    else:                     scale = S["zoom"]
    scale = max(0.05, min(scale, 8.0))
    nw, nh = max(1,int(_TXT_PW*scale)), max(1,int(_TXT_PH*scale))
    pil = img.resize((nw, nh), Image.LANCZOS if scale < 1 else Image.BILINEAR)
    if rid == S["render_id"]:
        root.after(0, lambda p=pil, r=rid: display(p, r))

# ── Extraction de texte par format ───────────────────────────────────────────
def _extract_docx(path):
    if _DocxDoc is None: return [("body","python-docx non installé",21)]
    doc = _DocxDoc(path)
    lines = []
    for p in doc.paragraphs:
        style = p.style.name.lower() if p.style else ""
        if "heading 1" in style:   lines.append(("h1", p.text))
        elif "heading 2" in style: lines.append(("h2", p.text))
        elif "heading 3" in style: lines.append(("h3", p.text))
        else:                      lines.append(("body", p.text))
    return lines

def _extract_odt(path):
    if _OdfLoad is None: return [("body","odfpy non installé",21)]
    doc = _OdfLoad(path)
    return [("body", _OdfTeletype.extractText(p))
            for p in doc.getElementsByType(_OdfText.P)]

def _extract_pptx(path):
    if _Pptx is None: return [("body","python-pptx non installé",21)]
    prs = _Pptx(path)
    lines = []
    for i, slide in enumerate(prs.slides, 1):
        lines.append(("h2", f"Slide {i}"))
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                lines.append(("body", shape.text))
    return lines

def _extract_html(path):
    import html.parser
    class _P(html.parser.HTMLParser):
        def __init__(self):
            super().__init__(); self.out=[]; self._skip=False
        def handle_starttag(self, t, a):
            if t in("script","style"): self._skip=True
            if t in("h1","h2","h3","p","div","br","li","tr"): self.out.append("\n")
        def handle_endtag(self, t):
            if t in("script","style"): self._skip=False
        def handle_data(self, d):
            if not self._skip: self.out.append(d)
    with open(path, encoding="utf-8", errors="replace") as f: raw=f.read()
    p=_P(); p.feed(raw)
    return [("body", _html_lib.unescape(l)) for l in "".join(p.out).splitlines()]

def _load_txt_doc(path, ext):
    """Charge un document texte/bureau dans le mode txt."""
    try:
        if ext == ".docx":
            styled = _extract_docx(path)
        elif ext == ".odt":
            styled = _extract_odt(path)
        elif ext == ".pptx":
            styled = _extract_pptx(path)
        elif ext in (".html",".htm"):
            styled = _extract_html(path)
        elif ext == ".md":
            with open(path, encoding="utf-8", errors="replace") as f:
                styled = _parse_md(f.read())
        else:
            with open(path, encoding="utf-8", errors="replace") as f:
                styled = [("body", l) for l in f.read().splitlines()]
        pages = _wrap_styled(styled)
        S["txt_pages"] = pages
        S["total"]     = len(pages)
        S["page"]      = 0
        S["mode"]      = "txt"
        S["title"]     = path
        root.title(f"LOCVIEW — {os.path.basename(path)}")
        _rebuild_tabs()
        _restore_position(path)
        render(reset_pan=True)
    except Exception as e:
        lbl_page.config(text=_L["status_unreadable"].format(ename=str(e)), fg="#f55")

def _open_svg(path):
    """Ouvre un fichier SVG en le rastérisant via svglib."""
    if _svg2rlg is None:
        lbl_page.config(text="svglib non installé", fg="#f55"); return
    try:
        drawing = _svg2rlg(path)
        if drawing is None: raise ValueError("SVG invalide")
        tmp = tempfile.mktemp(suffix=".png")
        _TEMP_FILES.append(tmp)
        _renderPM.drawToFile(drawing, tmp, fmt="PNG")
        S.update(imgs=[tmp], img_idx=0, mode="img", title=path)
        root.title(f"LOCVIEW — {os.path.basename(path)}")
        _rebuild_tabs()
        render(reset_pan=True)
    except Exception as e:
        lbl_page.config(text=f"SVG : {e}", fg="#f55")

def _open_cbr(path):
    """Ouvre un fichier CBR (comic RAR) comme galerie d'images."""
    if _rarfile is None:
        lbl_page.config(text="rarfile non installé", fg="#f55"); return
    try:
        rf = _rarfile.RarFile(path)
        names = sorted(n for n in rf.namelist()
                       if os.path.splitext(n)[1].lower() in IMG_EXT)
        if not names:
            lbl_page.config(text="CBR : aucune image trouvée", fg="#f55"); return
        tmpdir = tempfile.mkdtemp()
        _TEMP_FILES.append(tmpdir)
        rf.extractall(tmpdir)
        imgs = [os.path.join(tmpdir, n) for n in names]
        S.update(imgs=imgs, img_idx=0, mode="img", title=path)
        root.title(f"LOCVIEW — {os.path.basename(path)}")
        _rebuild_tabs()
        render(reset_pan=True)
    except Exception as e:
        lbl_page.config(text=f"CBR : {e}", fg="#f55")

@atexit.register
def _cleanup_temps():
    import shutil
    for p in _TEMP_FILES:
        try:
            if os.path.isdir(p): shutil.rmtree(p, ignore_errors=True)
            elif os.path.isfile(p): os.unlink(p)
        except Exception: pass

_render_job = [None]
_scroll_lock = [None]
_page_edge   = [None]   # None | "top" | "bot" — positionnement au chargement d'une page

_SEP_H     = 18         # hauteur du séparateur entre deux pages
_composite = dict(active=False, pil=None, p1_h=0, p2_h=0, comp_h=0, target=0,
                  direction=None, loading=False, entry_pan_y=0, start_pan_y=0,
                  item2=None, ref2=None, sep_item=None)

# Pré-rendu des pages adjacentes (N-1, N+1) pendant les moments idle
# {page_idx: (zoom, fit, doc_path, pil_img, ImageTk.PhotoImage)}
_pre_photos = {}
_preloading = set()   # pages dont le bg-render est en cours

def render(reset_pan=False, debounce=60):
    S["render_id"] += 1
    rid = S["render_id"]
    if reset_pan:
        S["pan_x"] = S["pan_y"] = 0
        if _page_edge[0] != "bot":   # "bot" déjà positionné par scroll_by
            _page_edge[0] = "top"
        if S["img_item"] is not None:
            cw = max(canvas.winfo_width(), 300)
            ch = max(canvas.winfo_height(), 300)
            canvas.coords(S["img_item"], cw//2, ch//2)
    if _render_job[0]: root.after_cancel(_render_job[0])
    pg  = S["page"]; idx = S["img_idx"]
    dp  = S["doc_path"]; mode = S["mode"]
    cw, ch = get_canvas_size()
    def _fire():
        _render_job[0] = None
        if rid != S["render_id"]: return
        if mode == "pdf":
            _queue_render(lambda: _do_pdf_render(pg, cw, ch, rid, dp))
        elif mode == "img":
            threading.Thread(target=render_img_bg,
                             args=(idx,cw,ch,rid), daemon=True).start()
        elif mode == "txt":
            threading.Thread(target=_do_txt_render,
                             args=(pg,cw,ch,rid), daemon=True).start()
    if debounce:
        _render_job[0] = root.after(debounce, _fire)
    else:
        _fire()

# ── Chargement ────────────────────────────────────────────────────────────────
def open_file(path):
    path = path.strip()
    if path.startswith("{") and "}" in path:
        path = path[1:path.index("}")]
    path = os.path.normpath(path)
    if not os.path.isfile(path):
        lbl_page.config(text=_L["status_not_found"].format(path=path), fg="#f55")
        return
    ext = os.path.splitext(path)[1].lower()
    if _search_debounce[0]:
        root.after_cancel(_search_debounce[0])
        _search_debounce[0] = None
    if SEARCH["active"]:
        search_frame.pack_forget()
        SEARCH["active"] = False
    # Ouvrir dans un nouvel onglet si l'actif a déjà un fichier
    if S["mode"] is not None:
        new_tab(path)
        return
    S["zoom"]=1.0; S["fit"]="page"; S["img_item"]=None
    S["pil_cache"].clear(); _pre_photos.clear(); _preloading.clear()
    canvas.delete("all")
    _clear_sigs(); [b["frame"].destroy() for b in list(_text_boxes)]; _text_boxes.clear()
    lbl_zoom.config(text="100%"); scrollbar.set(0,1)
    clear_search()
    if ext in (".pdf", ".epub", ".cbz", ".fb2", ".xps", ".oxps"):
        try:
            if S["doc"]: S["doc"].close()
            S["doc"]      = fitz.open(path)
            S["doc_path"] = path
            S["total"]    = len(S["doc"])
            S["page"]     = 0
            S["mode"]     = "pdf"
            S["title"]    = path
            root.title(f"LOCVIEW — {os.path.basename(path)}")
            _rebuild_tabs()
            _restore_position(path)
            render(reset_pan=True)
        except Exception as e:
            ename = type(e).__name__
            lbl_page.config(text=_L["status_unreadable"].format(ename=ename), fg="#f55")
    elif ext in IMG_EXT:
        folder = os.path.dirname(path)
        imgs = sorted([os.path.join(folder,f) for f in os.listdir(folder)
                       if os.path.splitext(f)[1].lower() in IMG_EXT])
        S.update(imgs=imgs,
                 img_idx=imgs.index(path) if path in imgs else 0,
                 mode="img", title=path)
        root.title(f"LOCVIEW — {os.path.basename(path)}")
        _rebuild_tabs()
        render(reset_pan=True)
    elif ext in SVG_EXT:
        _open_svg(path)
    elif ext in CBR_EXT:
        _open_cbr(path)
    elif ext in TXT_EXT | DOC_EXT:
        _load_txt_doc(path, ext)
    else:
        lbl_page.config(text=_L["status_unsupported"].format(ext=ext), fg="#f55")

def open_dialog():
    path = filedialog.askopenfilename(
        filetypes=[(_L["dialog_files"],
                    "*.pdf *.epub *.cbz *.cbr *.fb2 *.xps *.oxps *.svg "
                    "*.jpg *.jpeg *.png *.bmp *.gif *.webp *.tiff "
                    "*.txt *.md *.log *.csv *.rst *.docx *.odt *.pptx *.html *.htm"),
                   (_L["dialog_all"],"*.*")])
    if path: open_file(path)
    canvas.focus_set()

# ── Impression ────────────────────────────────────────────────────────────────
def do_print():
    if S["mode"]=="pdf" and S["doc"]:
        try: os.startfile(S["doc"].name, "print")
        except: _print_page()
    elif S["mode"]=="img" and S["imgs"]:
        try: os.startfile(S["imgs"][S["img_idx"]], "print")
        except Exception as e: lbl_page.config(text=_L["status_print_error"].format(msg=str(e)), fg="#f55")
    else: lbl_page.config(text=_L["status_no_file"], fg="#f55")

def _print_page():
    import tempfile
    try:
        page = S["doc"][S["page"]]
        pix  = page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
        tmp  = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        tmp.close(); pix.save(tmp.name)
        os.startfile(tmp.name, "print")
    except Exception as e:
        lbl_page.config(text=_L["status_print_error"].format(msg=str(e)), fg="#f55")

def toggle_fullscreen():
    S["fullscreen"] = not S["fullscreen"]
    root.attributes("-fullscreen", S["fullscreen"])

# ── Scroll continu entre pages ────────────────────────────────────────────────
def _cleanup_composite_extras():
    """Supprime le canvas item de la page cible et le séparateur."""
    if _composite["item2"] is not None:
        canvas.delete(_composite["item2"])
        _composite["item2"] = None
        _composite["ref2"]  = None
    if _composite["sep_item"] is not None:
        canvas.delete(_composite["sep_item"])
        _composite["sep_item"] = None

def _preload_neighbors():
    """Lance le pré-rendu PIL + PhotoImage des pages N-1 et N+1 en tâche de fond."""
    if S["mode"] != "pdf" or not S["doc_path"]: return
    cur = S["page"]
    zoom, fit, doc_path = S["zoom"], S["fit"], S["doc_path"]
    cw, ch = get_canvas_size()
    for pidx in (cur - 1, cur + 1):
        if not (0 <= pidx < S["total"]): continue
        pre = _pre_photos.get(pidx)
        if pre and pre[0] == zoom and pre[1] == fit and pre[2] == doc_path: continue
        if pidx in _preloading: continue
        with S["cache_lock"]:
            pil_cached = S["pil_cache"].get(pidx)
        if pil_cached:
            _, pil_img = pil_cached
            # PIL déjà disponible → créer le PhotoImage dès le prochain idle
            root.after_idle(lambda p=pil_img, i=pidx, z=zoom, f=fit, dp=doc_path:
                            _store_pre_photo(p, i, z, f, dp))
        else:
            _preloading.add(pidx)
            def _bg(i=pidx, z=zoom, f=fit, dp=doc_path):
                try:
                    doc = fitz.open(dp)
                    pg  = doc[i]
                    pw, ph = pg.rect.width, pg.rect.height
                    if f == "page":    sc = min(cw/pw, ch/ph) * z
                    elif f == "width": sc = (cw/pw) * z
                    else:              sc = z
                    sc = max(0.05, min(sc, 8.0))
                    pix = pg.get_pixmap(matrix=fitz.Matrix(sc, sc), alpha=False)
                    pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    doc.close()
                    root.after_idle(lambda p=pil, ii=i, zz=z, ff=f, ddp=dp:
                                    _store_pre_photo(p, ii, zz, ff, ddp))
                except Exception: pass
                finally: _preloading.discard(i)
            threading.Thread(target=_bg, daemon=True).start()

def _store_pre_photo(pil_img, page_idx, zoom, fit, doc_path):
    """Crée et stocke un PhotoImage pré-rendu — appelé en idle, thread principal."""
    if (S["zoom"] != zoom or S["fit"] != fit or S["doc_path"] != doc_path): return
    if S["page"] == page_idx: return       # page devenue courante, inutile
    pre = _pre_photos.get(page_idx)
    if pre and pre[0] == zoom and pre[1] == fit and pre[2] == doc_path: return
    tk_img = ImageTk.PhotoImage(pil_img)
    _pre_photos[page_idx] = (zoom, fit, doc_path, pil_img, tk_img)

def _enter_composite(direction, cw, ch):
    """Lance le rendu composite (page courante + voisine) en arrière-plan."""
    if _composite["active"] or _composite["loading"]: return
    if S["mode"] != "pdf":
        if direction == "down": go(+1)
        else: _page_edge[0] = "bot"; go(-1)
        return
    target = S["page"] + (1 if direction == "down" else -1)
    if not (0 <= target < S["total"]): return
    with S["cache_lock"]:
        cached = S["pil_cache"].get(S["page"])
    if not cached:
        if direction == "down": go(+1)
        else: _page_edge[0] = "bot"; go(-1)
        return
    current_pil = cached[1]
    _composite["loading"]    = True
    _composite["direction"]  = direction
    _composite["target"]     = target
    _composite["entry_pan_y"] = S["pan_y"]
    zoom, fit, doc_path = S["zoom"], S["fit"], S["doc_path"]

    # Raccourci : pré-rendu disponible → composite instantané, zéro blocage
    pre = _pre_photos.get(target)
    if pre and pre[0] == zoom and pre[1] == fit and pre[2] == doc_path:
        _composite["loading"] = False
        _show_composite(direction, current_pil, pre[3], target, pre[4])
        return

    # Sinon : rendu en arrière-plan
    def _bg():
        try:
            doc = fitz.open(doc_path)
            pg  = doc[target]
            pw, ph = pg.rect.width, pg.rect.height
            if fit == "page":    sc = min(cw/pw, ch/ph) * zoom
            elif fit == "width": sc = (cw/pw) * zoom
            else:                sc = zoom
            sc = max(0.05, min(sc, 8.0))
            pix = pg.get_pixmap(matrix=fitz.Matrix(sc, sc), alpha=False)
            tpil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
            root.after(0, lambda: _show_composite(direction, current_pil, tpil, target))
        except Exception:
            _composite["loading"] = False
    threading.Thread(target=_bg, daemon=True).start()

def _show_composite(direction, p_cur, p_tgt, target, pre_tk_img=None):
    """Affiche deux canvas items séparés (page courante + cible) sans composite PIL."""
    _composite["loading"] = False
    if not S["mode"]: return
    cw, ch = get_canvas_size()
    SEP = _SEP_H

    # p1 = page du haut, p2 = page du bas (disposition visuelle)
    if direction == "down":
        p1, p2 = p_cur, p_tgt   # item1=p_cur existant, item2=p_tgt nouveau
    else:
        p1, p2 = p_tgt, p_cur   # item2=p_tgt nouveau (au-dessus), item1=p_cur existant

    comp_h = p1.height + SEP + p2.height
    entry_pan_y = _composite["entry_pan_y"]

    if direction == "down":
        pan_y = comp_h // 2 - p1.height // 2 + entry_pan_y
    else:
        pan_y = comp_h // 2 - p1.height - SEP - p2.height // 2 + entry_pan_y

    _composite.update(active=True, pil=None, p1_h=p1.height, p2_h=p2.height,
                      comp_h=comp_h, target=target, start_pan_y=pan_y)
    S["render_id"] += 1
    S["pan_y"] = pan_y

    # Centres Y des deux items dans le canvas
    cx     = cw // 2 + S["pan_x"]
    base_y = ch // 2 + pan_y - comp_h // 2
    p1_cy  = base_y + p1.height // 2
    p2_cy  = base_y + p1.height + SEP + p2.height // 2

    # PhotoImage de la page cible : pré-rendu si disponible, sinon conversion ici
    _cleanup_composite_extras()
    tk_img2 = pre_tk_img if pre_tk_img is not None else ImageTk.PhotoImage(p_tgt)

    if direction == "down":
        # item1 (S["img_item"]) = p_cur = p1 (en haut)
        canvas.coords(S["img_item"], cx, p1_cy)
        item2 = canvas.create_image(cx, p2_cy, anchor="center", image=tk_img2, tags="img")
    else:
        # item1 (S["img_item"]) = p_cur = p2 (en bas)
        canvas.coords(S["img_item"], cx, p2_cy)
        item2 = canvas.create_image(cx, p1_cy, anchor="center", image=tk_img2, tags="img")

    # Séparateur léger via rectangle canvas (pas de PIL draw)
    sep_top = base_y + p1.height
    sep_item = canvas.create_rectangle(0, sep_top, cw, sep_top + SEP,
                                       fill="#252525", outline="", tags="composite_sep")
    canvas.tag_lower("composite_sep", "img")

    _composite["item2"]   = item2
    _composite["ref2"]    = tk_img2
    _composite["sep_item"] = sep_item
    S["prev_ref"] = S["ref"]

def _scroll_composite(dy, cw, ch):
    """Défile dans le composite ; commit quand l'ancienne page est hors écran."""
    p1_h      = _composite["p1_h"]
    p2_h      = _composite["p2_h"]
    comp_h    = _composite["comp_h"]
    SEP       = _SEP_H
    direction = _composite["direction"]
    new_y     = S["pan_y"] + dy
    # Sortie vers l'arrière (retour à la page d'origine)
    start = _composite["start_pan_y"]
    if direction == "down" and new_y > start:
        _exit_composite_backward(); return
    if direction == "up"   and new_y < start:
        _exit_composite_backward(); return
    # Commit : ancienne page entièrement hors champ
    top_pixel = comp_h // 2 - ch // 2 - new_y
    bot_pixel = comp_h // 2 + ch // 2 - new_y
    if direction == "down" and top_pixel >= p1_h + SEP:
        _commit_composite("down", cw, ch, p1_h, SEP, p2_h); return
    if direction == "up"   and bot_pixel < p1_h:
        _commit_composite("up",   cw, ch, p1_h, SEP, p2_h); return
    # Déplacement : déplacer les deux items séparément (canvas.coords = instant)
    S["pan_y"] = new_y
    cx     = cw // 2 + S["pan_x"]
    base_y = ch // 2 + new_y - comp_h // 2
    p1_cy  = base_y + p1_h // 2
    p2_cy  = base_y + p1_h + SEP + p2_h // 2
    if direction == "down":
        canvas.coords(S["img_item"],          cx, p1_cy)
        canvas.coords(_composite["item2"],    cx, p2_cy)
    else:
        canvas.coords(S["img_item"],          cx, p2_cy)
        canvas.coords(_composite["item2"],    cx, p1_cy)
    if _composite["sep_item"] is not None:
        sep_top = base_y + p1_h
        canvas.coords(_composite["sep_item"], 0, sep_top, cw, sep_top + SEP)

def _exit_composite_backward():
    """Quitte le composite en restaurant la position d'entrée (page d'origine)."""
    _composite["active"] = False
    _composite["pil"]    = None
    _cleanup_composite_extras()
    S["pan_y"] = _composite["entry_pan_y"]
    render(reset_pan=False)

def _commit_composite(direction, cw, ch, p1_h, SEP, p2_h):
    """Valide le changement de page, calcule pan_y cohérent et reprend le mode normal."""
    comp_h    = p1_h + SEP + p2_h
    old_pan_y = S["pan_y"]
    if direction == "down":
        new_pan_y = old_pan_y - comp_h // 2 + p1_h + SEP + p2_h // 2
    else:
        new_pan_y = old_pan_y - comp_h // 2 + p1_h // 2
    _composite["active"] = False
    _composite["pil"]    = None
    _cleanup_composite_extras()
    S["page"]  = _composite["target"]
    S["pan_y"] = new_pan_y
    render(reset_pan=False)
    lbl_page.config(text=f"{S['page']+1} / {S['total']}")
    scrollbar.set(S["page"] / S["total"], (S["page"] + 1) / S["total"])

# ── Navigation ────────────────────────────────────────────────────────────────
def go(delta):
    _cleanup_composite_extras()
    _composite["active"] = False; _composite["pil"] = None
    db = int(60 + max(0, S["zoom"] - 1.0) * 80)  # 60ms à 1x, ~140ms à 2x
    if S["mode"] in ("pdf", "txt"):
        S["page"] = max(0, min(S["page"]+delta, S["total"]-1))
        render(reset_pan=True, debounce=db)
    elif S["mode"]=="img":
        S["img_idx"] = (S["img_idx"]+delta) % len(S["imgs"])
        root.title(f"LOCVIEW — {os.path.basename(S['imgs'][S['img_idx']])}")
        render(reset_pan=True, debounce=db)

def do_zoom(factor, cx=None, cy=None):
    cw, ch = get_canvas_size()
    if cx is None: cx, cy = cw//2, ch//2
    img_cx = cw//2 + S["pan_x"]; img_cy = ch//2 + S["pan_y"]
    dx, dy = cx-img_cx, cy-img_cy
    old = S["zoom"]
    S["zoom"] = max(0.05, min(S["zoom"]*factor, 8.0))
    S["fit"]  = "manual"
    ratio = S["zoom"]/old
    S["pan_x"] = int(img_cx+dx*ratio-cw//2)
    S["pan_y"] = int(img_cy+dy*ratio-ch//2)
    lbl_zoom.config(text=f"{int(S['zoom']*100)}%")
    if S["img_item"] is not None:
        canvas.coords(S["img_item"], cw//2+S["pan_x"], ch//2+S["pan_y"])
    _pre_photos.clear(); _preloading.clear()
    render()

def toggle_fit():
    S["fit"]="width" if S["fit"]=="page" else "page"
    S["zoom"]=1.0
    _pre_photos.clear(); _preloading.clear()
    render(reset_pan=True)

# ── Sélection texte (clic droit) ──────────────────────────────────────────────
def _canvas_to_pdf(cx, cy):
    cw, ch = get_canvas_size()
    ox = cw//2+S["pan_x"]-S["pw"]*S["scale"]/2
    oy = ch//2+S["pan_y"]-S["ph"]*S["scale"]/2
    return (cx-ox)/S["scale"], (cy-oy)/S["scale"]

def on_sel_press(e):
    if _sig_ctx_menu(e.x, e.y): return  # clic droit sur un tampon
    if S["mode"]!="pdf": return
    S["sel_start"]=(e.x,e.y); canvas.delete("sel")

def on_sel_drag(e):
    if not S["sel_start"]: return
    x0,y0=S["sel_start"]
    canvas.delete("sel")
    canvas.create_rectangle(x0,y0,e.x,e.y, outline="#4af", width=1,
                            fill="#2288ff", stipple="gray25", tags="sel")

def on_sel_release(e):
    if not S["sel_start"] or S["mode"]!="pdf": return
    x0,y0=S["sel_start"]; x1,y1=e.x,e.y; S["sel_start"]=None
    if abs(x1-x0)<5 and abs(y1-y0)<5:
        canvas.delete("sel"); _hide_sel_panel(); return
    px0,py0=_canvas_to_pdf(min(x0,x1),min(y0,y1))
    px1,py1=_canvas_to_pdf(max(x0,x1),max(y0,y1))
    pw,ph=S["pw"],S["ph"]
    px0,py0=max(0,px0),max(0,py0)
    px1,py1=min(pw,px1),min(ph,py1)
    try:
        text=S["doc"][S["page"]].get_text("text", clip=fitz.Rect(px0,py0,px1,py1)).strip()
        S["sel_text"]=text
    except Exception:
        text=""
    _show_sel_panel(text)

# ── Pan + drag de tampon ──────────────────────────────────────────────────────
def on_press(e):
    canvas.focus_set()
    if _edit_mode[0]:
        _edit_sel_press(e.x, e.y)
        return
    idx, part = _sig_hit(e.x, e.y)
    if idx >= 0:
        if part == "close":
            _del_sig(idx)
            return
        _sig_drag.update(on=True, idx=idx, mode=part,
                         ox=e.x-_placed_sigs[idx]["x"],
                         oy=e.y-_placed_sigs[idx]["y"])
        return
    S["drag_start"]=(e.x,e.y,S["pan_x"],S["pan_y"])

def on_drag(e):
    if _edit_mode[0]:
        _edit_sel_drag(e.x, e.y)
        return
    if _sig_drag["on"]:
        idx = _sig_drag["idx"]; sig = _placed_sigs[idx]
        if _sig_drag["mode"] == "body":
            nx = e.x-_sig_drag["ox"]; ny = e.y-_sig_drag["oy"]
            sig["x"] = nx; sig["y"] = ny
            canvas.coords(sig["item_id"], nx, ny)
            # Mettre a jour les coords PDF
            bb = canvas.bbox(sig["item_id"])
            if bb and S.get("doc") and S["page"] == sig.get("page"):
                px0,py0 = _canvas_to_pdf(bb[0],bb[1])
                px1,py1 = _canvas_to_pdf(bb[2],bb[3])
                sig["pdf_rect"] = (px0, py0, px1, py1)
        else:  # handle → resize par diagonale uniquement
            if sig.get("kind") == "text":
                # Texte vectoriel: changer la taille de police
                bb = canvas.bbox(sig["item_id"])
                if bb:
                    orig_diag = ((bb[2]-bb[0])**2 + (bb[3]-bb[1])**2) ** 0.5 or 1
                    dx = e.x - sig["x"]; dy = e.y - sig["y"]
                    cur_diag = (dx**2 + dy**2) ** 0.5
                    raw_scale = cur_diag / orig_diag
                    new_size = max(6, int(sig["size"] * raw_scale))
                    sig["size"] = new_size
                    canvas.itemconfig(sig["item_id"],
                                      font=(sig["font_name"], new_size, "italic"))
            else:
                orig_w, orig_h = sig["orig_pil"].size
                orig_diag = (orig_w**2 + orig_h**2) ** 0.5
                dx = e.x - sig["x"]; dy = e.y - sig["y"]
                cur_diag = (dx**2 + dy**2) ** 0.5
                scale = max(0.05, cur_diag / orig_diag)
                nw = max(10, int(orig_w * scale))
                nh = max(10, int(orig_h * scale))
                resized = sig["orig_pil"].resize((nw, nh), Image.LANCZOS)
                new_tk  = ImageTk.PhotoImage(resized)
                sig["tk_img"] = new_tk
                canvas.itemconfig(sig["item_id"], image=new_tk)
        _draw_sig_handle(sig)
        canvas.tag_raise("sig"); canvas.tag_raise("sig_handle")
        return
    if not S["drag_start"] or not S["mode"]: return
    sx,sy,px0,py0=S["drag_start"]
    S["pan_x"]=px0+(e.x-sx); S["pan_y"]=py0+(e.y-sy)
    if S["img_item"] is not None:
        cw,ch=get_canvas_size()
        canvas.coords(S["img_item"], cw//2+S["pan_x"], ch//2+S["pan_y"])

def on_release(e):
    if _edit_mode[0]:
        _edit_sel_release(e.x, e.y)
        return
    _sig_drag["on"]=False
    S["drag_start"]=None

# ── Clavier ───────────────────────────────────────────────────────────────────
def on_key(e):
    # Ne pas intercepter les touches quand un champ texte a le focus
    fw = root.focus_get()
    if isinstance(fw, (tk.Entry, tk.Text)):
        if e.keysym.lower() == "escape":
            if SEARCH["active"]: toggle_search()
            else: _hide_sel_panel()
        return
    k=e.keysym.lower()
    ctrl=(e.state&0x4)
    if   ctrl and k=="f":                   toggle_search()
    elif ctrl and k=="c" and S["sel_text"]: root.clipboard_clear(); root.clipboard_append(S["sel_text"])
    elif ctrl and k=="t":                   new_tab()
    elif ctrl and k=="w":                   close_tab(ACTIVE[0])
    elif k in("down","space","next"):  scroll_by(-80)
    elif k in("up","prior"):           scroll_by(+80)
    elif k=="right":                   go(+1)
    elif k=="left":                    go(-1)
    elif k=="home":                         S["page"]=0; render(reset_pan=True)
    elif k=="end":                          S["page"]=S["total"]-1; render(reset_pan=True)
    elif k in("plus","equal","kp_add"):     do_zoom(1.25)
    elif k in("minus","kp_subtract"):       do_zoom(0.8)
    elif k=="f":                            toggle_fit()
    elif k=="r":                            S.update(zoom=1.0,fit="page"); render(reset_pan=True)
    elif k=="f11":                          toggle_fullscreen()
    elif k=="m":                            save_position()
    elif k=="o":                            open_dialog()
    elif k=="p":                            do_print()
    elif k=="escape":
        if SEARCH["active"]:               toggle_search()
        elif S["fullscreen"]:              S["fullscreen"]=False; root.attributes("-fullscreen",False)
        else:                              root.destroy()
    elif k=="q":                            root.destroy()

def scroll_by(dy):
    """Défile de dy px. Aux bords, lance le scroll continu composite."""
    if not S["mode"]: return
    cw, ch = get_canvas_size()
    if _composite["active"]:
        _scroll_composite(dy, cw, ch)
        return
    bb = None
    if S["img_item"] is not None:
        try: bb = canvas.bbox(S["img_item"])
        except: pass
    if bb:
        img_h = bb[3] - bb[1]
        if img_h <= ch:
            if dy < 0: _enter_composite("down", cw, ch)
            else:       _enter_composite("up",   cw, ch)
            return
        pan_min = -((img_h - ch + 1) // 2)
        pan_max =   (img_h - ch + 1) // 2
        new_y   = S["pan_y"] + dy
        if dy < 0 and new_y < pan_min:
            _enter_composite("down", cw, ch); return
        if dy > 0 and new_y > pan_max:
            _enter_composite("up",   cw, ch); return
        S["pan_y"] = max(pan_min, min(pan_max, new_y))
    else:
        S["pan_y"] += dy
    if S["img_item"] is not None:
        canvas.coords(S["img_item"], cw//2 + S["pan_x"], ch//2 + S["pan_y"])

def on_scroll(e):
    if root.focus_get() is None:   # fenêtre sans focus → ignorer (scroll inactive windows)
        return
    if (e.state & 0x4):  # Ctrl+molette → changer de page
        go(-1 if (e.delta > 0 or e.num == 4) else 1)
        return
    scroll_by(80 if (e.delta > 0 or e.num == 4) else -80)

def on_resize(e):
    if S["mode"]: root.after(80, render)
    else: _place_hint()

# ── Bindings ──────────────────────────────────────────────────────────────────
root.bind_all("<Key>",           on_key)
root.bind("<Configure>",         on_resize)
# ── Mode edition de texte PDF ────────────────────────────────────────────────
_edit_mode    = [False]
_edit_ctx     = {"widget": None, "win_id": None, "word": None}
_edit_history = []   # [(doc_path, bytes_snapshot), ...]

def _push_edit_snapshot():
    if S.get("doc") and S.get("doc_path"):
        try:
            _edit_history.append((S["doc_path"], S["doc"].tobytes()))
            if len(_edit_history) > 20:
                _edit_history.pop(0)
            root.after(0, _refresh_undo_btn)
        except Exception:
            pass

def _refresh_undo_btn():
    if _edit_history:
        btn_undo.config(state=tk.NORMAL, fg=BTN_FG, cursor="hand2")
    else:
        btn_undo.config(state=tk.DISABLED, fg="#555", cursor="arrow")

def _undo_edit(*_):
    if not _edit_history:
        return
    doc_path, snapshot = _edit_history.pop()
    _refresh_undo_btn()
    if doc_path != S.get("doc_path"):
        lbl_save_status.config(text=_L["status_undo_diff_doc"], fg="#fa0")
        root.after(2500, lambda: lbl_save_status.config(text=""))
        return
    try:
        import shutil
        tmp = doc_path + ".undo_tmp"
        with open(tmp, "wb") as f:
            f.write(snapshot)
        if S["doc"]:
            S["doc"].close()
        shutil.move(tmp, doc_path)
        S["doc"] = fitz.open(doc_path)
        S["total"] = len(S["doc"])
        S["pil_cache"].clear()
        n = len(_edit_history)
        lbl_save_status.config(
            text=_L["status_undo_done"] + (_L["status_undo_left"].format(n=n) if n else ""),
            fg="#aaf")
        root.after(3000, lambda: lbl_save_status.config(text=""))
        render()
    except Exception as ex:
        lbl_save_status.config(text=_L["status_undo_error"].format(msg=str(ex)[:60]), fg="#f88")
        root.after(4000, lambda: lbl_save_status.config(text=""))

def _pdf_to_canvas(px, py):
    cw, ch = get_canvas_size()
    ox = cw//2 + S["pan_x"] - S["pw"]*S["scale"]/2
    oy = ch//2 + S["pan_y"] - S["ph"]*S["scale"]/2
    return int(ox + px * S["scale"]), int(oy + py * S["scale"])

def _cancel_inline_edit():
    if _edit_ctx["widget"] and _edit_ctx["widget"].winfo_exists():
        _edit_ctx["widget"].destroy()
    _edit_ctx["widget"] = None
    if _edit_ctx["win_id"]:
        canvas.delete(_edit_ctx["win_id"])
    _edit_ctx["win_id"] = None

def _toggle_edit_mode():
    if not S.get("doc") or S["mode"] != "pdf":
        lbl_save_status.config(text=_L["status_open_pdf"], fg="#fa0")
        root.after(2500, lambda: lbl_save_status.config(text=""))
        return
    _edit_mode[0] = not _edit_mode[0]
    if _edit_mode[0]:
        canvas.configure(cursor="xterm")
        lbl_save_status.config(text=_L["status_edit_mode"], fg="#8af")
    else:
        canvas.configure(cursor="")
        lbl_save_status.config(text="")
        _cancel_inline_edit()

def _get_span_at(page, x, y):
    try:
        for block in page.get_text("dict")["blocks"]:
            if block.get("type") != 0: continue
            for line in block["lines"]:
                for span in line["spans"]:
                    r = span["bbox"]
                    if r[0]-1 <= x <= r[2]+1 and r[1]-1 <= y <= r[3]+1:
                        return span
    except Exception: pass
    return {}

def _try_edit_word(cx, cy):
    """En mode edition: debut de selection de zone (drag pour definir le bloc a editer)."""
    pass  # La selection se fait via on_press/drag — voir _edit_sel_*

_edit_sel = {"start": None}

def _edit_sel_press(cx, cy):
    _edit_sel["start"] = (cx, cy)
    canvas.delete("edit_sel")

def _edit_sel_drag(cx, cy):
    s = _edit_sel.get("start")
    if not s: return
    canvas.delete("edit_sel")
    canvas.create_rectangle(s[0], s[1], cx, cy,
        outline="#f90", width=2, dash=(4,3), fill="", tags="edit_sel")

def _edit_sel_release(cx, cy):
    s = _edit_sel.get("start")
    canvas.delete("edit_sel")
    _edit_sel["start"] = None
    if not s: return
    if abs(cx-s[0]) < 6 and abs(cy-s[1]) < 6: return
    if not S.get("doc") or S["mode"] != "pdf": return
    x0p, y0p = _canvas_to_pdf(min(s[0],cx), min(s[1],cy))
    x1p, y1p = _canvas_to_pdf(max(s[0],cx), max(s[1],cy))
    rect = fitz.Rect(x0p, y0p, x1p, y1p)
    page = S["doc"][S["page"]]
    # Extraire tout le texte du bloc selectionne
    block_text = page.get_text("text", clip=rect).strip()
    if not block_text: return
    # Detecter infos typo du premier span dans la zone
    span = _get_span_at(page, x0p+1, y0p+1)
    fsize = span.get("size", 10)
    color_int = span.get("color", 0)
    if isinstance(color_int, int):
        fcol = (((color_int>>16)&0xff)/255, ((color_int>>8)&0xff)/255, (color_int&0xff)/255)
    else:
        fcol = (0, 0, 0)
    _open_block_editor(rect, block_text, fsize, fcol, page)

def _open_block_editor(rect, orig_text, fsize, fcol, page):
    """Ouvre une fenetre d'edition pour le bloc de texte selectionne."""
    win = tk.Toplevel(root)
    win.title(_L["edit_title"]); win.configure(bg="#1a1a1a")
    win.attributes("-topmost", True)
    win.resizable(True, True)
    win.geometry("500x340")
    win.minsize(400, 280)

    # Boutons EN HAUT pour qu'ils soient toujours visibles
    btn_row = tk.Frame(win, bg="#1a1a1a")
    btn_row.pack(fill=tk.X, padx=10, pady=(10,4))

    tk.Label(win, text=_L["edit_instruction"],
             bg="#1a1a1a", fg="#888", font=("Consolas",8)).pack(padx=10,pady=(0,2),anchor="w")

    txt = tk.Text(win, bg="#111", fg="#eee", font=("Consolas",10),
                  insertbackground="#fff", relief=tk.FLAT,
                  highlightthickness=1, highlightbackground="#333",
                  wrap=tk.WORD, undo=True)
    txt.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,10))
    txt.insert("1.0", orig_text)
    txt.focus_set()

    def _apply():
        new_text = txt.get("1.0", "end-1c").strip()
        win.destroy()
        if not new_text or new_text == orig_text: return
        try:
            _push_edit_snapshot()
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
            # Réduire la taille si le texte déborde du rect
            fs = max(6, fsize)
            rc = -1
            while fs >= 4 and rc < 0:
                rc = page.insert_textbox(rect, new_text, fontsize=fs,
                                         color=fcol, align=0, fontname="helv")
                fs -= 1
            if rc < 0:
                # Dernier recours : insert_text sans contrainte de boîte
                page.insert_text(rect.tl + fitz.Point(2, fsize + 2),
                                 new_text, fontsize=max(6, fsize), color=fcol,
                                 fontname="helv")
            import tempfile, shutil
            doc_path = S["doc_path"]
            try:
                S["doc"].save(doc_path, incremental=True,
                              encryption=fitz.PDF_ENCRYPT_KEEP)
            except Exception:
                tmp = doc_path + ".edit_tmp"
                S["doc"].save(tmp)
                S["doc"].close()
                shutil.move(tmp, doc_path)
                S["doc"] = fitz.open(doc_path)
            S["pil_cache"].clear()
            lbl_save_status.config(text=_L["status_modified"], fg="#7f7")
            root.after(4000, lambda: lbl_save_status.config(text=""))
            render()
        except Exception as ex:
            lbl_save_status.config(text=_L["status_edit_error"].format(msg=str(ex)[:60]), fg="#f88")
            root.after(5000, lambda: lbl_save_status.config(text=""))

    def _cancel():
        win.destroy()

    tk.Button(btn_row, text=_L["edit_confirm"], bg="#1a2a3a", fg="#7af",
              relief=tk.FLAT, font=("Consolas",9), cursor="hand2", padx=10,
              command=_apply).pack(side=tk.LEFT, padx=4)
    tk.Button(btn_row, text=_L["edit_cancel"], bg="#2a1a1a", fg="#a77",
              relief=tk.FLAT, font=("Consolas",9), cursor="hand2", padx=10,
              command=_cancel).pack(side=tk.LEFT)
    win.bind("<Control-Return>", lambda e: _apply())
    txt.bind("<Control-Return>", lambda e: _apply())

# Echap pour quitter le mode edition, Ctrl+Z pour annuler
def _on_edit_key(e):
    if _edit_mode[0] and e.keysym.lower() == "escape":
        _toggle_edit_mode()
root.bind("<Escape>", _on_edit_key, add="+")
root.bind_all("<Control-z>", _undo_edit)
root.bind_all("<Control-Z>", _undo_edit)

canvas.bind("<MouseWheel>",      on_scroll)
canvas.bind("<Button-4>",        on_scroll)
canvas.bind("<Button-5>",        on_scroll)
canvas.bind("<ButtonPress-1>",   on_press)
canvas.bind("<B1-Motion>",       on_drag)
canvas.bind("<ButtonRelease-1>", on_release)
canvas.bind("<ButtonPress-3>",   on_sel_press)
canvas.bind("<B3-Motion>",       on_sel_drag)
canvas.bind("<ButtonRelease-3>", on_sel_release)

# ── Éditeur de signature ──────────────────────────────────────────────────────

def open_sig_editor():
    win = tk.Toplevel(root)
    win.title(_L["sig_title"])
    win.configure(bg="#1a1a1a")
    win.resizable(False, False)
    win.grab_set()

    CWIDTH, CHEIGHT = 380, 130
    mode       = tk.StringVar(value="dessin")
    draw_color = tk.StringVar(value="#000000")
    pen_size    = tk.IntVar(value=3)
    stabilizer  = tk.BooleanVar(value=True)   # lissage trajectoire

    draw_canvas = tk.Canvas(win, width=CWIDTH, height=CHEIGHT,
                            bg="white", cursor="pencil",
                            highlightthickness=1, highlightbackground="#444")
    draw_canvas.grid(row=1, column=0, columnspan=4, padx=12, pady=(4,4))

    def _draw_guides():
        draw_canvas.delete("guide")
        base_y = int(CHEIGHT * 0.72)
        xh_y   = int(CHEIGHT * 0.30)
        draw_canvas.create_line(8, base_y, CWIDTH-8, base_y,
            fill="#88ccff", dash=(4,6), width=1, tags="guide")
        draw_canvas.create_line(8, xh_y, CWIDTH-8, xh_y,
            fill="#88ccff", dash=(2,8), width=1, tags="guide")
    _draw_guides()

    _stroke     = {"pts": [], "cur_id": None}
    _history    = []   # [{"id": item_id, "pts": [...], "color": str, "width": int}]
    _redo_stack = []

    def _refresh_undo_btns():
        btn_undo.config(state=tk.NORMAL if _history    else tk.DISABLED)
        btn_redo.config(state=tk.NORMAL if _redo_stack else tk.DISABLED)

    # Lazy rope : le stylo suit le curseur avec inertie
    _rope = {"x": 0.0, "y": 0.0}
    LAZY_R = 2
    LAZY_F = 0.15

    def _dc_press(e):
        _stroke["pts"] = [(e.x, e.y)]
        _stroke["cur_id"] = None
        _rope["x"] = float(e.x); _rope["y"] = float(e.y)
        draw_canvas.delete("rope_dot")
        draw_canvas.create_oval(e.x-4, e.y-4, e.x+4, e.y+4,
            fill=draw_color.get(), outline="", tags="rope_dot")

    def _dc_drag(e):
        pts = _stroke["pts"]
        rx, ry = _rope["x"], _rope["y"]
        cx, cy = float(e.x), float(e.y)
        dx = cx - rx; dy = cy - ry
        dist = (dx*dx + dy*dy) ** 0.5
        draw_canvas.delete("rope_dot")
        R = LAZY_R
        draw_canvas.create_oval(cx-R, cy-R, cx+R, cy+R,
            outline=draw_color.get(), width=1, tags="rope_dot")
        if dist < LAZY_R: return
        step = dist * (1.0 - LAZY_F)
        nx = rx + dx / dist * step
        ny = ry + dy / dist * step
        _rope["x"] = nx; _rope["y"] = ny
        pts.append((int(nx), int(ny)))
        if len(pts) < 2: return
        if _stroke["cur_id"]: draw_canvas.delete(_stroke["cur_id"])
        flat = [c for p in pts for c in p]
        _stroke["cur_id"] = draw_canvas.create_line(
            *flat, smooth=True, splinesteps=48,
            fill=draw_color.get(), width=pen_size.get(),
            capstyle=tk.ROUND, joinstyle=tk.ROUND)
        draw_canvas.tag_raise("guide")
        draw_canvas.tag_raise("rope_dot")

    def _dc_release(e):
        draw_canvas.delete("rope_dot")
        if _stroke["cur_id"]:
            pts = list(_stroke["pts"])
            if pts and (abs(e.x - pts[-1][0]) > 2 or abs(e.y - pts[-1][1]) > 2):
                pts.append((e.x, e.y))
            if len(pts) >= 4:
                for _ in range(4):
                    new_pts = [pts[0]]
                    for j in range(len(pts)-1):
                        p0, p1 = pts[j], pts[j+1]
                        new_pts.append((0.75*p0[0]+0.25*p1[0], 0.75*p0[1]+0.25*p1[1]))
                        new_pts.append((0.25*p0[0]+0.75*p1[0], 0.25*p0[1]+0.75*p1[1]))
                    new_pts.append(pts[-1])
                    pts = new_pts
                draw_canvas.delete(_stroke["cur_id"])
                flat = [c for p in pts for c in p]
                _stroke["cur_id"] = draw_canvas.create_line(
                    *flat, smooth=True, splinesteps=48,
                    fill=draw_color.get(), width=pen_size.get(),
                    capstyle=tk.ROUND, joinstyle=tk.ROUND)
                draw_canvas.tag_raise("guide")
            _history.append({
                "id":    _stroke["cur_id"],
                "pts":   pts,
                "color": draw_color.get(),
                "width": pen_size.get(),
            })
            _redo_stack.clear()
            _refresh_undo_btns()
        _stroke["pts"] = []; _stroke["cur_id"] = None

    def _undo():
        if not _history: return
        entry = _history.pop()
        draw_canvas.delete(entry["id"])
        entry["id"] = None
        _redo_stack.append(entry)
        _refresh_undo_btns()

    def _redo():
        if not _redo_stack: return
        entry = _redo_stack.pop()
        flat = [c for p in entry["pts"] for c in p]
        entry["id"] = draw_canvas.create_line(
            *flat, smooth=True, splinesteps=48,
            fill=entry["color"], width=entry["width"],
            capstyle=tk.ROUND, joinstyle=tk.ROUND)
        _history.append(entry)
        _refresh_undo_btns()

    txt_var  = tk.StringVar(value="Signature")
    txt_size = tk.IntVar(value=32)
    _loaded_img = {"pil": None, "tk": None}

    def _render_text(*_):
        if mode.get() != "texte": return
        draw_canvas.delete("all")
        try:
            draw_canvas.create_text(CWIDTH//2, CHEIGHT//2, text=txt_var.get(),
                                    font=("Arial", txt_size.get(), "italic"),
                                    fill=draw_color.get(), anchor="center")
        except Exception: pass

    for v in (txt_var, txt_size, draw_color):
        v.trace_add("write", _render_text)

    def set_mode(m):
        mode.set(m)
        draw_canvas.delete("all")
        if m == "dessin":
            draw_canvas.configure(cursor="pencil")
            draw_canvas.bind("<ButtonPress-1>",   _dc_press)
            draw_canvas.bind("<B1-Motion>",       _dc_drag)
            draw_canvas.bind("<ButtonRelease-1>", _dc_release)
            _draw_guides()
        elif m == "texte":
            draw_canvas.unbind("<ButtonPress-1>")
            draw_canvas.unbind("<B1-Motion>")
            draw_canvas.configure(cursor="arrow")
            _render_text()
            win.after(50, lambda: (txt_entry.focus_set(),
                                   txt_entry.selection_range(0, tk.END)))
        elif m == "image":
            draw_canvas.unbind("<ButtonPress-1>")
            draw_canvas.unbind("<B1-Motion>")
            draw_canvas.configure(cursor="arrow")
            _load_img()

    def _load_img():
        p = filedialog.askopenfilename(parent=win,
            filetypes=[(_L["dialog_images"],"*.png *.jpg *.jpeg *.bmp *.gif *.webp"),(_L["sig_all"],"*.*")])
        if not p: set_mode("dessin"); return
        try:
            img = Image.open(p).convert("RGBA")
            img.thumbnail((CWIDTH, CHEIGHT), Image.LANCZOS)
            _loaded_img["pil"] = img
            bg2 = Image.new("RGB", img.size, (255,255,255))
            bg2.paste(img, mask=img.split()[3])
            _loaded_img["tk"] = ImageTk.PhotoImage(bg2)
            draw_canvas.delete("all")
            draw_canvas.create_image(CWIDTH//2, CHEIGHT//2,
                                     image=_loaded_img["tk"], anchor="center")
        except Exception: set_mode("dessin")

    # Barre de modes
    btn_f = tk.Frame(win, bg="#1a1a1a")
    btn_f.grid(row=0, column=0, columnspan=4, padx=12, pady=(10,2), sticky="w")
    for lbl, m in [(_L["sig_draw"],"dessin"),(_L["sig_text_mode"],"texte"),("🖼 Image","image")]:
        tk.Radiobutton(btn_f, text=lbl, variable=mode, value=m, indicatoron=False,
                       bg="#222", fg="#aaa", selectcolor="#1a3a5a", activebackground="#333",
                       font=("Consolas",9), relief=tk.FLAT, padx=8, pady=3, cursor="hand2",
                       command=lambda mv=m: set_mode(mv)).pack(side=tk.LEFT, padx=2)

    # Contrôles
    ctrl = tk.Frame(win, bg="#1a1a1a")
    ctrl.grid(row=2, column=0, columnspan=4, padx=12, pady=(2,4), sticky="w")

    def pick_color():
        c = colorchooser.askcolor(color=draw_color.get(), parent=win, title="Couleur")[1]
        if c: draw_color.set(c)

    tk.Button(ctrl, text="Couleur", bg="#222", fg="#aaa", relief=tk.FLAT,
              font=("Consolas",8), cursor="hand2", command=pick_color).pack(side=tk.LEFT, padx=2)
    tk.Label(ctrl, text=_L["sig_width"], bg="#1a1a1a", fg="#888", font=("Consolas",8)).pack(side=tk.LEFT, padx=(6,0))
    tk.Scale(ctrl, from_=1, to=12, variable=pen_size, orient=tk.HORIZONTAL,
             bg="#1a1a1a", fg="#aaa", troughcolor="#333", highlightthickness=0,
             length=70, width=12, showvalue=False).pack(side=tk.LEFT)
    tk.Label(ctrl, text=_L["sig_text_label"], bg="#1a1a1a", fg="#888", font=("Consolas",8)).pack(side=tk.LEFT, padx=(8,0))
    txt_entry = tk.Entry(ctrl, textvariable=txt_var, width=14, bg="#222", fg="#eee",
                         insertbackground="white", font=("Consolas",9), relief=tk.FLAT)
    txt_entry.pack(side=tk.LEFT, padx=4)
    tk.Spinbox(ctrl, from_=8, to=80, textvariable=txt_size, width=4,
               bg="#222", fg="#eee", buttonbackground="#333",
               font=("Consolas",9), relief=tk.FLAT).pack(side=tk.LEFT, padx=2)
    btn_undo = tk.Button(ctrl, text=_L["sig_undo"], bg="#222", fg="#aaa", relief=tk.FLAT,
                         font=("Consolas",8), cursor="hand2", state=tk.DISABLED,
                         command=_undo)
    btn_undo.pack(side=tk.LEFT, padx=2)
    btn_redo = tk.Button(ctrl, text=_L["sig_redo"], bg="#222", fg="#aaa", relief=tk.FLAT,
                         font=("Consolas",8), cursor="hand2", state=tk.DISABLED,
                         command=_redo)
    btn_redo.pack(side=tk.LEFT, padx=2)

    def _effacer():
        draw_canvas.delete("all")
        _history.clear(); _redo_stack.clear(); _refresh_undo_btns()
        _draw_guides()
    tk.Button(ctrl, text=_L["sig_clear"], bg="#222", fg="#888", relief=tk.FLAT,
              font=("Consolas",8), cursor="hand2",
              command=_effacer).pack(side=tk.LEFT, padx=6)
    tk.Checkbutton(ctrl, text="Lissage", variable=stabilizer,
                   bg="#1a1a1a", fg="#888", selectcolor="#1a3a5a",
                   activebackground="#1a1a1a", font=("Consolas",8),
                   cursor="hand2").pack(side=tk.LEFT, padx=4)

    # Raccourcis clavier Ctrl+Z / Ctrl+Y
    win.bind("<Control-z>", lambda e: _undo())
    win.bind("<Control-y>", lambda e: _redo())

    # Sigs sauvegardées
    saved_row = tk.Frame(win, bg="#1a1a1a")
    saved_row.grid(row=3, column=0, columnspan=4, padx=12, pady=(4,2), sticky="w")
    tk.Label(saved_row, text=_L["sig_saved"], bg="#1a1a1a", fg="#555",
             font=("Consolas",8)).pack(side=tk.LEFT)
    _saved_tk = []

    def _load_saved():
        for w in list(saved_row.winfo_children())[1:]: w.destroy()
        _saved_tk.clear()
        for p in sorted(SIGS_DIR.glob("*.png")):
            try:
                img = Image.open(p).convert("RGBA")
                img.thumbnail((60,30), Image.LANCZOS)
                bg2 = Image.new("RGB", img.size, (255,255,255))
                bg2.paste(img, mask=img.split()[3])
                tki = ImageTk.PhotoImage(bg2); _saved_tk.append(tki)
                b = tk.Button(saved_row, image=tki, bg="#222", relief=tk.FLAT,
                              cursor="hand2", command=lambda pp=p: _apply_saved(pp))
                b.pack(side=tk.LEFT, padx=2)
                b.bind("<ButtonPress-3>", lambda e, pp=p: (os.remove(pp), _load_saved()))
            except Exception: pass

    def _apply_saved(path):
        try:
            img = Image.open(path).convert("RGBA")
            _place_sig(img); win.destroy()
        except Exception: pass

    _load_saved()

    def _capture():
        m = mode.get()
        if m == "image" and _loaded_img["pil"]: return _loaded_img["pil"].copy()
        if m == "texte":
            txt = txt_var.get().strip()
            if not txt: return None
            try:    fnt = ImageFont.truetype("arial.ttf", txt_size.get())
            except: fnt = ImageFont.load_default()
            dummy = Image.new("RGBA",(1,1)); dd = ImageDraw.Draw(dummy)
            bb = dd.textbbox((0,0), txt, font=fnt)
            w = bb[2]-bb[0]+20; h = bb[3]-bb[1]+20
            img = Image.new("RGBA",(w,h),(255,255,255,0))
            ImageDraw.Draw(img).text((10,10), txt, font=fnt, fill=draw_color.get())
            return img
        # Rendu vectoriel haute résolution depuis l'historique des traits
        if not _history:
            return None
        SCALE = 4
        w = CWIDTH * SCALE; h = CHEIGHT * SCALE
        img = Image.new("RGBA", (w, h), (255, 255, 255, 0))
        d   = ImageDraw.Draw(img)
        for entry in _history:
            pts = entry["pts"]
            if len(pts) < 2:
                continue
            scaled = [(int(x * SCALE), int(y * SCALE)) for x, y in pts]
            base_w = max(2, entry["width"] * SCALE)
            n = len(scaled)
            if n < 3:
                d.line(scaled, fill=entry["color"], width=base_w)
                continue
            # Dessin segment par segment : variation de largeur selon vitesse + effilage
            for j in range(n - 1):
                p0, p1 = scaled[j], scaled[j+1]
                dx = p1[0] - p0[0]; dy = p1[1] - p0[1]
                speed = max(1.0, (dx*dx + dy*dy) ** 0.5)
                # Pression : lent = epais (0.6-1.4x), rapide = fin
                press = max(0.4, min(1.4, 6.0 * SCALE / speed))
                # Effilage : debut et fin du trait
                t = j / max(1, n - 2)
                taper = min(t / 0.12, 1.0, (1 - t) / 0.12)
                lw = max(1, int(base_w * press * (0.25 + 0.75 * taper)))
                d.line([p0, p1], fill=entry["color"], width=lw)
        # Recadrer sur le contenu réel + marge
        bbox = img.getbbox()
        if not bbox:
            return None
        pad  = SCALE * 6
        bbox = (max(0, bbox[0]-pad), max(0, bbox[1]-pad),
                min(w, bbox[2]+pad), min(h, bbox[3]+pad))
        return img.crop(bbox)

    bot = tk.Frame(win, bg="#1a1a1a")
    bot.grid(row=4, column=0, columnspan=4, padx=12, pady=(4,10))

    def _do_save():
        sig = _capture()
        if not sig: return
        i = 1
        while (SIGS_DIR/f"sig_{i:02d}.png").exists(): i += 1
        sig.save(SIGS_DIR/f"sig_{i:02d}.png")
        _load_saved()

    def _do_apply():
        try:
            if mode.get() == "texte":
                t = txt_var.get().strip()
                if t:
                    _place_text_stamp(t, "Arial", txt_size.get(), draw_color.get())
                    win.destroy()
            elif mode.get() == "dessin":
                hires = _capture()
                if hires:
                    w, h = hires.size
                    display = hires.resize((max(1, w//4), max(1, h//4)), Image.LANCZOS)
                    _place_sig_at(display, orig_pil=hires, subkind="signature")
                    win.destroy()
                else:
                    import tkinter.messagebox as mb
                    mb.showwarning(_L["sig_no_drawing"], _L["sig_draw_first"], parent=win)
            else:
                sig = _capture()
                if sig: _place_sig(sig, subkind="image"); win.destroy()
        except Exception as _e:
            import tkinter.messagebox as mb
            mb.showerror(_L["sig_error"], str(_e), parent=win)

    tk.Button(bot, text=_L["sig_save_btn"], bg="#1a3a1a", fg="#8f8",
              relief=tk.FLAT, font=("Consolas",9), cursor="hand2", padx=8,
              command=_do_save).pack(side=tk.LEFT, padx=4)
    tk.Button(bot, text=_L["sig_embed"], bg="#1a2a3a", fg="#7af",
              relief=tk.FLAT, font=("Consolas",9), cursor="hand2", padx=8,
              command=_do_apply).pack(side=tk.LEFT, padx=4)
    tk.Button(bot, text=_L["sig_cancel"], bg="#2a1a1a", fg="#a77",
              relief=tk.FLAT, font=("Consolas",9), cursor="hand2", padx=8,
              command=win.destroy).pack(side=tk.LEFT, padx=4)

    set_mode("dessin")


# ── Boîte de texte draggable ──────────────────────────────────────────────────

def add_text_box():
    cw=max(canvas.winfo_width(),400); ch=max(canvas.winfo_height(),300)
    box={"drag":{"on":False,"ox":0,"oy":0}}

    frm=tk.Frame(canvas,bg="#fffde7",bd=0,
                 highlightthickness=1,highlightbackground="#e0c060",cursor="fleur")
    title_bar=tk.Frame(frm,bg="#e8d44d",height=16,cursor="fleur")
    title_bar.pack(fill=tk.X,side=tk.TOP)
    close_lbl=tk.Label(title_bar,text="✕",bg="#e8d44d",fg="#333",
                       font=("Consolas",7),cursor="hand2",padx=4)
    close_lbl.pack(side=tk.RIGHT)

    txt=tk.Text(frm,width=22,height=4,wrap=tk.WORD,
                bg="#fffde7",fg="#111",font=("Consolas",10),
                relief=tk.FLAT,insertbackground="#333",
                highlightthickness=0,borderwidth=0)
    txt.pack(padx=3,pady=(0,3))
    txt.insert("1.0", _L["textbox_placeholder"])
    txt.bind("<FocusIn>",lambda e:(
        txt.tag_add("sel","1.0","end") if txt.get("1.0","end-1c")==_L["textbox_placeholder"] else None))

    box["frame"]=frm; _text_boxes.append(box)

    x0=cw//2-90; y0=ch//2-40
    win_id=canvas.create_window(x0,y0,window=frm,anchor="nw",tags="textbox")
    box.update(win_id=win_id,x=x0,y=y0)

    def _tb_press(e):
        box["drag"]["on"]=True
        box["drag"]["ox"]=e.x_root-box["x"]
        box["drag"]["oy"]=e.y_root-box["y"]
    def _tb_drag(e):
        if not box["drag"]["on"]: return
        nx=e.x_root-box["drag"]["ox"]; ny=e.y_root-box["drag"]["oy"]
        box["x"]=nx; box["y"]=ny
        canvas.coords(win_id,nx,ny)
    def _tb_release(e): box["drag"]["on"]=False

    title_bar.bind("<ButtonPress-1>",_tb_press)
    title_bar.bind("<B1-Motion>",_tb_drag)
    title_bar.bind("<ButtonRelease-1>",_tb_release)

    def _close(e=None):
        canvas.delete(win_id)
        frm.destroy()
        if box in _text_boxes: _text_boxes.remove(box)
    close_lbl.bind("<Button-1>",_close)


# ── Notification "application par défaut" ─────────────────────────────────────

def _show_default_prompt():
    prefs=_load_saves()
    if prefs.get("default_app_asked"): return

    import winreg as _wr
    try:
        k = _wr.OpenKey(_wr.HKEY_CURRENT_USER, r"Software\Classes\.pdf")
        val, _ = _wr.QueryValueEx(k, "")
        _wr.CloseKey(k)
        if val == "LOCVIEW.File":
            p = _load_saves(); p["default_app_asked"] = True; _write_saves(p)
            return
    except Exception:
        pass
    notif=tk.Frame(root,bg=BAR_BG,bd=0,
                   highlightthickness=1,highlightbackground="#444")
    notif.place(x=12,y=38,width=310)

    bar_f=tk.Frame(notif,bg="#333",height=3); bar_f.pack(fill=tk.X,side=tk.TOP)
    timer_bar=tk.Frame(bar_f,bg="#666",height=3)
    timer_bar.place(x=0,y=0,relwidth=1.0,height=3)

    tk.Label(notif,text=_L["notif_default"],
             bg=BAR_BG,fg=BTN_FG,font=("Consolas",9),
             wraplength=270,justify="left").pack(padx=10,pady=(8,4),anchor="w")

    chk=tk.BooleanVar(value=False)
    tk.Checkbutton(notif,text=_L["notif_dont_ask"],variable=chk,
                   bg=BAR_BG,fg="#888",activebackground=BAR_BG,
                   selectcolor=BTN_BG,font=("Consolas",8)).pack(padx=10,anchor="w")

    btn_row=tk.Frame(notif,bg=BAR_BG); btn_row.pack(padx=10,pady=(4,8),anchor="e")

    def _dismiss(do_set=False):
        p=_load_saves()
        if chk.get() or do_set: p["default_app_asked"]=True
        if do_set:
            try:
                exe=str(_APP_DIR/"LOCVIEW.exe")
                for ext in (".pdf",".png",".jpg",".jpeg"):
                    k=_wr.CreateKey(_wr.HKEY_CURRENT_USER,rf"Software\Classes\{ext}")
                    _wr.SetValue(k,"",_wr.REG_SZ,"LOCVIEW.File"); _wr.CloseKey(k)
                prog=_wr.CreateKey(_wr.HKEY_CURRENT_USER,r"Software\Classes\LOCVIEW.File")
                _wr.SetValue(prog,"",_wr.REG_SZ,"Document LOCVIEW")
                cmd=_wr.CreateKey(prog,r"shell\open\command")
                _wr.SetValue(cmd,"",_wr.REG_SZ,f'"{exe}" "%1"'); _wr.CloseKey(cmd); _wr.CloseKey(prog)
                ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
                lbl_save_status.config(text=_L["status_default_set"],fg="#aaa")
                root.after(3000,lambda: lbl_save_status.config(text=""))
            except Exception:
                lbl_save_status.config(text=_L["status_restart_admin"],fg="#fa0")
                root.after(4000,lambda: lbl_save_status.config(text=""))
        _write_saves(p); notif.destroy()

    b_oui=tk.Button(btn_row,text=_L["notif_yes"],bg=BTN_BG,fg=BTN_FG,relief=tk.FLAT,
                    font=("Consolas",9),cursor="hand2",padx=8,
                    command=lambda:_dismiss(True))
    b_oui.bind("<Enter>",lambda e:b_oui.config(bg=BTN_HOV,fg="#fff"))
    b_oui.bind("<Leave>",lambda e:b_oui.config(bg=BTN_BG,fg=BTN_FG))
    b_oui.pack(side=tk.LEFT,padx=2)
    b_non=tk.Button(btn_row,text=_L["notif_no"],bg=BTN_BG,fg=BTN_FG,relief=tk.FLAT,
                    font=("Consolas",9),cursor="hand2",padx=8,
                    command=lambda:_dismiss(False))
    b_non.bind("<Enter>",lambda e:b_non.config(bg=BTN_HOV,fg="#fff"))
    b_non.bind("<Leave>",lambda e:b_non.config(bg=BTN_BG,fg=BTN_FG))
    b_non.pack(side=tk.LEFT,padx=2)

    DURATION=12000
    start=int(root.tk.call("clock","milliseconds"))
    def _tick():
        if not notif.winfo_exists(): return
        frac=max(0.0,1.0-( int(root.tk.call("clock","milliseconds"))-start)/DURATION)
        timer_bar.place(x=0,y=0,relwidth=frac,height=3)
        if frac<=0: _dismiss(False)
        else: root.after(80,_tick)
    root.after(80,_tick)


root.drop_target_register(DND_FILES)
root.dnd_bind("<<Drop>>", lambda e: open_file(e.data))

_rebuild_tabs()
root.after(200, canvas.focus_set)
root.after(210, _place_hint)
root.after(900, _show_default_prompt)

def _apply_lang():
    """Met à jour tous les textes de l'interface après un changement de langue."""
    _btn_open.config(text=_L["btn_open"])
    _btn_print.config(text=_L["btn_print"])
    _btn_fs.config(text=_L["btn_fullscreen"])
    _btn_prev.config(text=_L["btn_prev"])
    _btn_next.config(text=_L["btn_next"])
    _btn_search.config(text=_L["btn_search"])
    _btn_edit_mode.config(text=_L["btn_edit"])
    btn_undo.config(text=_L["btn_undo_toolbar"])
    _btn_lang.config(text=_L["lang_btn"])
    _btn_save.config(text=_L["btn_save"])
    _btn_bmark.config(text=_L["btn_bookmark"])
    _lbl_search.config(text=_L["search_label"])
    _chk_whole_word.config(text=_L["search_whole_word"])
    _lbl_sel_hdr.config(text=_L["sel_header"])
    _btn_copy_lbl.config(text=_L["sel_copy"])
    canvas.itemconfig(_hint_txt, text=_L["hint_drop"])
    _hint_btn.config(text=_L["hint_open"])
    _sig_panel_refresh()
    _rebuild_tabs()

# ── Instance unique — serveur IPC ────────────────────────────────────────────
def _ipc_serve():
    """Démarre le serveur IPC ; écrit le port dans le fichier partagé."""
    try:
        srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        srv.bind(("127.0.0.1", 0))
        port = srv.getsockname()[1]
        with open(_IPC_PORT_FILE, "w") as f:
            f.write(str(port))
        atexit.register(lambda: os.path.exists(_IPC_PORT_FILE) and os.remove(_IPC_PORT_FILE))
        srv.listen(5)
        def _loop():
            while True:
                try:
                    conn, _ = srv.accept()
                    data = conn.recv(8192).decode("utf-8", errors="replace").strip()
                    conn.close()
                    if data:
                        root.after(0, lambda p=data: _ipc_open(p))
                except Exception:
                    break
        threading.Thread(target=_loop, daemon=True).start()
    except Exception:
        pass

def _ipc_open(path):
    """Ouvre le fichier dans un nouvel onglet et met la fenêtre au premier plan."""
    root.lift()
    root.attributes("-topmost", True)
    root.attributes("-topmost", False)
    root.focus_force()
    new_tab(path)   # nouvel onglet au lieu d'écraser l'onglet courant

_ipc_serve()

if _file_arg_early:
    root.after(150, lambda: open_file(_file_arg_early))

root.mainloop()
