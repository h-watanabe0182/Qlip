import tkinter as tk
from tkinter import messagebox, filedialog
import os
import sys
import subprocess
import webbrowser
import logging
import json
import uuid

# --- Constants ---
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE   = os.path.join(APP_DIR, "Qlip.log")
DATA_FILE  = os.path.join(APP_DIR, "qlip_data.json")
META_FILE  = os.path.join(APP_DIR, "metadata.json")       # 旧ファイル（移行用）
IMG_DIR    = os.path.join(APP_DIR, "images")
FONT = "Meiryo"
COLS = 8
ROWS = 5

# --- Colors ---
C_BG              = "#f2f3f5"
C_HEADER_BG       = "#ffffff"
C_HEADER_BORDER   = "#d0d5dd"
C_SIDEBAR_BG      = "#f7f8fa"
C_SIDEBAR_BORDER  = "#e0e3e8"
C_SIDEBAR_ACTIVE_BG = "#e0eaff"
C_SIDEBAR_ACTIVE_FG = "#1a56db"
C_SIDEBAR_FG      = "#444444"
C_TEXT            = "#1f2937"
C_TEXT_SUB        = "#6b7280"
C_PREVIEW_BG      = "#ffffff"
C_PREVIEW_BORDER  = "#d0d5dd"
C_ADD_FG          = "#16a34a"
C_EDIT_FG         = "#d97706"
C_BTN_BG          = "#2563eb"
C_BTN_FG          = "#ffffff"

DEFAULT_CAT_COLORS = {
    "0": "#90caf9",
    "1": "#a5d6a7",
    "2": "#ffe082",
    "3": "#f48fb1",
    "4": "#ce93d8",
    "5": "#80cbc4",
}
DEFAULT_CARD_COLOR = "#90caf9"

# --- Logging ---
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format="%(asctime)s | %(message)s", encoding="utf-8")
log = logging.info

# ---------------------------------------------------------------------------
# qlip_data.json スキーマ
# {
#   "categories": [{"id": "0", "name": "お気に入り"}, ...],
#   "items": [
#     {"id": "<uuid>", "cat": "0", "name": "memo", "path": "...",
#      "description": "", "image": "", "color": "", "order": 0}, ...
#   ],
#   "settings": {"show_card_image": true, "web_urls": [...]}
# }
# ---------------------------------------------------------------------------

def _default_data():
    return {"categories": [], "items": [], "settings": {"show_card_image": True, "web_urls": ["http://localhost/"]}}

def load_data():
    """qlip_data.json を読み込む。なければ旧ファイルから移行して生成する。"""
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    # --- 旧ファイルから移行 ---
    data = _default_data()

    # アクセスリスト.txt を探す
    txt_path = None
    for f in os.scandir(APP_DIR):
        if f.is_file() and f.name.lower().endswith(".txt"):
            txt_path = f.path
            break

    if txt_path:
        cats_raw, items_raw = _parse_txt(txt_path)
        data["categories"] = [{"id": cid, "name": cname} for cid, cname in cats_raw]
        # metadata.json があれば読み込む
        old_meta = {}
        if os.path.exists(META_FILE):
            try:
                with open(META_FILE, "r", encoding="utf-8") as f:
                    old_meta = json.load(f)
            except Exception:
                pass
        for order, (cat, name, path) in enumerate(items_raw):
            m = old_meta.get(name, {})
            data["items"].append({
                "id":          str(uuid.uuid4()),
                "cat":         cat,
                "name":        name,
                "path":        path,
                "description": m.get("description", ""),
                "image":       m.get("image", ""),
                "color":       m.get("color", ""),
                "order":       order,
            })
        # settings
        s = old_meta.get("__settings__", {})
        if "show_card_image" in s:
            data["settings"]["show_card_image"] = s["show_card_image"]
        if "web_urls" in s:
            data["settings"]["web_urls"] = s["web_urls"]
        log(f"Migrated from {txt_path}: {len(data['categories'])} cats, {len(data['items'])} items")

    save_data(data)
    return data

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def _parse_txt(path):
    """アクセスリスト.txt をパースして (categories, items) を返す。"""
    categories, items = [], []
    for enc in ("cp932", "utf-8"):
        try:
            with open(path, "r", encoding=enc) as f:
                lines = f.readlines()
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    else:
        return categories, items
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith("#"):
            parts = line.split(" ", 3)
            if len(parts) >= 4 and parts[1].upper() == "CATEGORY":
                categories.append((parts[2], parts[3]))
        else:
            parts = line.split(" ", 2)
            if len(parts) >= 3:
                items.append((parts[0], parts[1], parts[2]))
    return categories, items

# --- Launch ---
def launch_path(path_str):
    path_str = path_str.strip()
    if not path_str:
        return
    log(f"Launch: {path_str}")
    try:
        if path_str.lower().startswith(("http://", "https://")):
            webbrowser.open(path_str)
        elif path_str.lower().endswith(".ps1"):
            basedir = os.path.dirname(path_str)
            subprocess.Popen(
                ["powershell.exe", "-ExecutionPolicy", "Bypass", "-NoProfile",
                 "-NoExit", "-File", path_str],
                cwd=basedir or None,
                creationflags=subprocess.CREATE_NEW_CONSOLE)
        elif path_str.lower().endswith((".bat", ".cmd")):
            basedir = os.path.dirname(path_str)
            subprocess.Popen(["cmd", "/c", path_str],
                             cwd=basedir or None,
                             creationflags=subprocess.CREATE_NEW_CONSOLE)
        else:
            os.startfile(path_str)
    except Exception as e:
        log(f"ERROR Launch: {e}")
        messagebox.showerror("エラー", f"開けませんでした:\n{path_str}\n\n{e}")


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Qlip")
        self.root.state("zoomed")
        self.root.configure(bg=C_BG)

        ico_path = os.path.join(APP_DIR, "Qlip.ico")
        if os.path.exists(ico_path):
            try:
                self.root.iconbitmap(ico_path)
            except Exception:
                pass

        # --- データ読み込み ---
        self.data = load_data()
        self.current_cat = "all"
        self.card_widgets = []   # (item_id, frame_widget)
        self.card_images  = []   # PIL GC 防止

        # settings
        self.show_card_image = tk.BooleanVar(
            value=self.data["settings"].get("show_card_image", True))
        self._web_panel_open  = False
        self._web_urls        = self.data["settings"].get("web_urls", ["http://localhost/"])
        self._web_current_idx = 0

        # D&D 状態
        self._drag_item_id   = None   # ドラッグ中の item id
        self._drag_ghost     = None   # ゴーストラベル
        self._drag_origin    = None   # (row, col) ドラッグ元グリッド位置
        self._drag_armed     = False  # 長押し後にD&D有効になったか
        self._drag_after_id  = None  # after() のID（キャンセル用）

        if not os.path.exists(IMG_DIR):
            os.makedirs(IMG_DIR, exist_ok=True)

        self._build_ui()

        cats = self.data["categories"]
        if cats:
            self._select_cat(cats[0]["id"])
        else:
            self.render_grid("all")

        self.search_entry.focus_set()

    # ------------------------------------------------------------------
    # ヘルパー: データアクセス
    # ------------------------------------------------------------------
    def _visible_items(self, cat, filter_str=""):
        """表示対象アイテムを order 順で返す。"""
        items = sorted(self.data["items"], key=lambda x: x.get("order", 0))
        result = []
        for it in items:
            if cat != "all" and it["cat"] != cat:
                continue
            if filter_str and not self._match_filter(it, filter_str):
                continue
            result.append(it)
            if len(result) >= COLS * ROWS:
                break
        return result

    def _match_filter(self, it, q):
        q = q.lower()
        return (q in it["name"].lower() or
                q in it["path"].lower() or
                q in it.get("description", "").lower())

    def _save(self):
        save_data(self.data)

    def _save_settings(self):
        self.data["settings"]["show_card_image"] = self.show_card_image.get()
        self.data["settings"]["web_urls"]        = self._web_urls
        self._save()

    # ------------------------------------------------------------------
    # UI 構築
    # ------------------------------------------------------------------
    def _build_ui(self):
        # === Header ===
        header = tk.Frame(self.root, bg=C_HEADER_BG, height=70,
                          highlightbackground=C_HEADER_BORDER, highlightthickness=1)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        self._header_image = None
        for img_path in [os.path.join(APP_DIR, "参考", "Qlipイメージ画像.png"),
                         os.path.join(APP_DIR, "Qlipイメージ画像.png")]:
            if os.path.exists(img_path):
                try:
                    from PIL import Image, ImageTk
                    pil_img = Image.open(img_path)
                    h = 70
                    w = int(pil_img.width * h / pil_img.height)
                    pil_img = pil_img.resize((w, h), Image.LANCZOS)
                    self._header_image = ImageTk.PhotoImage(pil_img)
                    tk.Label(header, image=self._header_image, bg=C_HEADER_BG,
                             bd=0).pack(side=tk.RIGHT)
                except Exception:
                    pass
                break

        tk.Label(header, text="Qlip", bg=C_HEADER_BG, fg=C_SIDEBAR_ACTIVE_FG,
                 font=(FONT, 16, "bold")).pack(side=tk.LEFT, padx=(16, 4))
        tk.Label(header, text="Your Quick Link Organizer", bg=C_HEADER_BG, fg=C_TEXT_SUB,
                 font=(FONT, 9)).pack(side=tk.LEFT, padx=(0, 12), pady=(14, 0))

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(header, textvariable=self.search_var,
                                     bg="#ffffff", fg=C_TEXT, insertbackground=C_TEXT,
                                     font=(FONT, 22), relief=tk.SOLID, bd=1,
                                     highlightbackground=C_HEADER_BORDER, highlightthickness=1)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=12, pady=10)
        self.search_entry.bind("<Return>", self._on_search)

        tk.Label(header, text="Enter: 起動  |  複数: フィルタ", bg=C_HEADER_BG, fg=C_TEXT_SUB,
                 font=(FONT, 9)).pack(side=tk.RIGHT, padx=(8, 16))

        # Web パネル トグルボタン（将来活用予定・現在非表示）
        self._web_toggle_btn = tk.Label(
            header, text="🌐 Web ▶", bg="#e0eaff", fg=C_SIDEBAR_ACTIVE_FG,
            font=(FONT, 9, "bold"), padx=10, pady=4, cursor="hand2", relief=tk.FLAT)
        # self._web_toggle_btn.pack(side=tk.RIGHT, padx=(0, 8), pady=16)
        self._web_toggle_btn.bind("<Button-1>", lambda e: self._toggle_web_panel())

        # === Body ===
        body = tk.Frame(self.root, bg=C_BG)
        body.pack(fill=tk.BOTH, expand=True)

        # === Left Sidebar ===
        sidebar = tk.Frame(body, bg=C_SIDEBAR_BG, width=140,
                           highlightbackground=C_SIDEBAR_BORDER, highlightthickness=1)
        sidebar.pack(fill=tk.Y, side=tk.LEFT)
        sidebar.pack_propagate(False)

        self.cat_buttons = {}
        self._add_cat_btn(sidebar, "all", "All")
        for cat in self.data["categories"]:
            self._add_cat_btn(sidebar, cat["id"], cat["name"])

        tk.Frame(sidebar, bg=C_SIDEBAR_BG, height=20).pack(fill=tk.X)

        add_btn = tk.Label(sidebar, text="＋ 追加", bg=C_SIDEBAR_BG, fg=C_ADD_FG,
                           font=(FONT, 10, "bold"), anchor="w", padx=12, pady=7, cursor="hand2")
        add_btn.pack(fill=tk.X)
        add_btn.bind("<Button-1>", lambda e: self._item_dialog(mode="add"))
        add_btn.bind("<Enter>", lambda e: add_btn.configure(bg="#e8fce8"))
        add_btn.bind("<Leave>", lambda e: add_btn.configure(bg=C_SIDEBAR_BG))

        edit_btn = tk.Label(sidebar, text="✎ 編集", bg=C_SIDEBAR_BG, fg=C_EDIT_FG,
                            font=(FONT, 10, "bold"), anchor="w", padx=12, pady=7, cursor="hand2")
        edit_btn.pack(fill=tk.X)
        edit_btn.bind("<Button-1>", lambda e: self._item_dialog(mode="edit"))
        edit_btn.bind("<Enter>", lambda e: edit_btn.configure(bg="#fff5e0"))
        edit_btn.bind("<Leave>", lambda e: edit_btn.configure(bg=C_SIDEBAR_BG))

        # Settings
        tk.Frame(sidebar, bg=C_SIDEBAR_BORDER, height=1).pack(fill=tk.X, padx=8, pady=(10, 4))
        tk.Label(sidebar, text="設定", bg=C_SIDEBAR_BG, fg=C_TEXT_SUB,
                 font=(FONT, 8), anchor="w", padx=14).pack(fill=tk.X)

        img_toggle_frame = tk.Frame(sidebar, bg=C_SIDEBAR_BG, cursor="hand2")
        img_toggle_frame.pack(fill=tk.X, padx=8, pady=3)
        self._img_toggle_canvas = tk.Canvas(img_toggle_frame, width=34, height=18,
                                            bg=C_SIDEBAR_BG, highlightthickness=0)
        self._img_toggle_canvas.pack(side=tk.LEFT)
        tk.Label(img_toggle_frame, text="カード画像", bg=C_SIDEBAR_BG, fg=C_SIDEBAR_FG,
                 font=(FONT, 9), cursor="hand2").pack(side=tk.LEFT, padx=(4, 0))

        def _draw_toggle():
            c = self._img_toggle_canvas
            c.delete("all")
            on = self.show_card_image.get()
            bg_col = C_SIDEBAR_ACTIVE_FG if on else "#adb5bd"
            c.create_polygon(
                7, 2, 27, 2, 34, 9, 27, 16, 7, 16, 0, 9,
                smooth=True, fill=bg_col, outline="")
            knob_x = 22 if on else 10
            c.create_oval(knob_x-6, 3, knob_x+6, 15, fill="#ffffff", outline="")

        def _toggle_card_image(e=None):
            self.show_card_image.set(not self.show_card_image.get())
            _draw_toggle()
            self._save_settings()
            self.render_grid(self.current_cat)

        _draw_toggle()
        self._img_toggle_canvas.bind("<Button-1>", _toggle_card_image)
        img_toggle_frame.bind("<Button-1>", _toggle_card_image)
        for child in img_toggle_frame.winfo_children():
            child.bind("<Button-1>", _toggle_card_image)

        # === Preview pane ===
        self.preview_frame = tk.Frame(body, bg=C_PREVIEW_BG, width=260,
                                      highlightbackground=C_PREVIEW_BORDER, highlightthickness=1)
        self.preview_frame.pack(fill=tk.Y, side=tk.LEFT)
        self.preview_frame.pack_propagate(False)

        self.preview_title = tk.Label(self.preview_frame, text="", bg=C_PREVIEW_BG, fg=C_TEXT,
                                      font=(FONT, 12, "bold"), wraplength=240, justify=tk.CENTER)
        self.preview_title.pack(pady=(12, 4), padx=8)
        self.preview_img_label = tk.Label(self.preview_frame, bg=C_PREVIEW_BG)
        self.preview_path_label = tk.Label(self.preview_frame, text="", bg=C_PREVIEW_BG, fg=C_TEXT_SUB,
                                           font=(FONT, 8), wraplength=240, justify=tk.CENTER)
        self.preview_path_label.pack(pady=(0, 4), padx=8)
        self.preview_desc_label = tk.Label(self.preview_frame, text="", bg=C_PREVIEW_BG, fg=C_TEXT,
                                           font=(FONT, 9), wraplength=240, justify=tk.LEFT, anchor="nw")
        self.preview_desc_label.pack(fill=tk.BOTH, expand=True, padx=12, pady=(4, 12))
        self._clear_preview()

        # === Grid area ===
        self.grid_frame = tk.Frame(body, bg=C_BG)
        self.grid_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        for c in range(COLS):
            self.grid_frame.columnconfigure(c, weight=1, uniform="col")
        for r in range(ROWS):
            self.grid_frame.rowconfigure(r, weight=1, uniform="row")

        # === Web Panel（将来活用予定・現在非表示） ===
        self.web_panel = tk.Frame(body, bg=C_SIDEBAR_BG, width=420,
                                  highlightbackground=C_SIDEBAR_BORDER, highlightthickness=1)
        self.web_panel.pack_propagate(False)

        wp_header = tk.Frame(self.web_panel, bg=C_HEADER_BG,
                             highlightbackground=C_HEADER_BORDER, highlightthickness=1)
        wp_header.pack(fill=tk.X)
        tk.Label(wp_header, text="🌐", bg=C_HEADER_BG, font=(FONT, 11)).pack(side=tk.LEFT, padx=(8,2), pady=6)
        self._web_url_var = tk.StringVar(value=self._web_urls[self._web_current_idx])
        url_entry = tk.Entry(wp_header, textvariable=self._web_url_var,
                             font=(FONT, 9), relief=tk.SOLID, bd=1)
        url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4, pady=6)
        url_entry.bind("<Return>", lambda e: self._web_load(self._web_url_var.get()))
        tk.Button(wp_header, text="→", font=(FONT, 9, "bold"), bg=C_BTN_BG, fg=C_BTN_FG,
                  relief=tk.FLAT, padx=6,
                  command=lambda: self._web_load(self._web_url_var.get())).pack(side=tk.LEFT, padx=(0,2), pady=6)
        tk.Button(wp_header, text="↺", font=(FONT, 9, "bold"), bg="#f0f0f0", relief=tk.FLAT, padx=6,
                  command=lambda: self._web_load(self._web_url_var.get())).pack(side=tk.LEFT, padx=(0,4), pady=6)
        tk.Button(wp_header, text="✕", font=(FONT, 9), bg="#f0f0f0", relief=tk.FLAT, padx=6,
                  command=self._toggle_web_panel).pack(side=tk.RIGHT, padx=(0,4), pady=6)

        self._wp_tab_frame = tk.Frame(self.web_panel, bg=C_SIDEBAR_BG)
        self._wp_tab_frame.pack(fill=tk.X)
        self._rebuild_web_tabs()

        try:
            from tkinterweb import HtmlFrame
            self.web_view = HtmlFrame(self.web_panel, messages_enabled=False)
            self.web_view.pack(fill=tk.BOTH, expand=True)
            self._web_available = True
        except Exception:
            self.web_view = tk.Label(self.web_panel, text="tkinterweb が利用できません",
                                     bg=C_SIDEBAR_BG, fg=C_TEXT_SUB)
            self.web_view.pack(fill=tk.BOTH, expand=True)
            self._web_available = False

    # ------------------------------------------------------------------
    # Sidebar category buttons
    # ------------------------------------------------------------------
    def _add_cat_btn(self, parent, cat_id, label):
        btn = tk.Label(parent, text=label, bg=C_SIDEBAR_BG, fg=C_SIDEBAR_FG,
                       font=(FONT, 10), anchor="w", padx=14, pady=8, cursor="hand2")
        btn.pack(fill=tk.X)
        btn.bind("<Button-1>", lambda e, c=cat_id: self._select_cat(c))
        btn.bind("<Enter>",
                 lambda e, b=btn, c=cat_id: b.configure(bg="#e8ecf0") if c != self.current_cat else None)
        btn.bind("<Leave>",
                 lambda e, b=btn, c=cat_id: b.configure(bg=C_SIDEBAR_BG) if c != self.current_cat else None)
        self.cat_buttons[cat_id] = btn

    def _select_cat(self, cat):
        for cid, btn in self.cat_buttons.items():
            btn.configure(bg=C_SIDEBAR_BG, fg=C_SIDEBAR_FG, font=(FONT, 10))
        self.cat_buttons[cat].configure(bg=C_SIDEBAR_ACTIVE_BG, fg=C_SIDEBAR_ACTIVE_FG,
                                        font=(FONT, 10, "bold"))
        self.current_cat = cat
        self.render_grid(cat)

    # ------------------------------------------------------------------
    # Grid render
    # ------------------------------------------------------------------
    def render_grid(self, cat, filter_str=""):
        for _, w in self.card_widgets:
            w.destroy()
        self.card_widgets = []
        self.card_images  = []

        visible = self._visible_items(cat, filter_str)

        idx = 0
        for r in range(ROWS):
            for c in range(COLS):
                if idx < len(visible):
                    it = visible[idx]
                    card = self._make_card(self.grid_frame, it, r, c)
                else:
                    card = self._make_empty(self.grid_frame)
                card.grid(row=r, column=c, sticky="nsew", padx=2, pady=2)
                if idx < len(visible):
                    self.card_widgets.append((visible[idx]["id"], card))
                idx += 1

    def _get_card_color(self, it):
        if it.get("color"):
            return it["color"]
        return DEFAULT_CAT_COLORS.get(it["cat"], DEFAULT_CARD_COLOR)

    def _make_card(self, parent, it, grid_row, grid_col):
        name  = it["name"]
        path  = it["path"]
        accent    = self._get_card_color(it)
        border    = self._darken_color(accent, 0.35)
        hover_bg  = border

        frame = tk.Frame(parent, bg="#ffffff", cursor="hand2",
                         highlightbackground=border, highlightthickness=2)

        # カード内画像
        card_img_label = None
        if self.show_card_image.get():
            img_path = it.get("image", "")
            if img_path and os.path.exists(img_path):
                try:
                    from PIL import Image, ImageTk
                    pil_img = Image.open(img_path)
                    pil_img.thumbnail((96, 54), Image.LANCZOS)
                    tk_img = ImageTk.PhotoImage(pil_img)
                    self.card_images.append(tk_img)
                    card_img_label = tk.Label(frame, image=tk_img, bg="#ffffff",
                                             cursor="hand2", bd=0)
                    card_img_label.pack(pady=(4, 0))
                    card_img_label.bind("<Button-1>", lambda e, p=path: launch_path(p))
                except Exception:
                    pass

        btn = tk.Button(frame, text=name, bg="#ffffff", fg=border,
                        activebackground=hover_bg, activeforeground="#ffffff",
                        font=(FONT, 9, "bold"), relief=tk.FLAT, bd=0, cursor="hand2",
                        wraplength=130, padx=4, pady=3,
                        command=lambda p=path: launch_path(p))
        btn.pack(fill=tk.X, padx=4, pady=(2 if card_img_label else 5, 1))

        if card_img_label is None:
            basename  = os.path.basename(path)
            parentdir = os.path.basename(os.path.dirname(path))
            CPL = 13
            nl = max(1, (len(name)      + CPL - 1) // CPL)
            pl = max(1, (len(parentdir) + CPL - 1) // CPL)
            fl = max(1, (len(basename)  + CPL - 1) // CPL)
            disp = basename if nl + pl + fl > 4 else ((parentdir + "\n" + basename) if parentdir else basename)
            path_lbl = tk.Label(frame, text=disp, bg="#ffffff", fg="#888888",
                                font=(FONT, 8), wraplength=130, justify=tk.CENTER)
            path_lbl.pack(fill=tk.BOTH, expand=True, padx=3, pady=(0, 4))
            path_lbl.bind("<Button-1>", lambda e, p=path: launch_path(p))
        else:
            path_lbl = tk.Label(frame, text="", bg="#ffffff")
            path_lbl.pack(pady=(0, 2))

        hover_widgets = [frame, btn, path_lbl] + ([card_img_label] if card_img_label else [])

        def on_enter(e):
            frame.configure(bg=hover_bg, highlightbackground=hover_bg)
            btn.configure(bg=hover_bg, fg="#ffffff")
            path_lbl.configure(bg=hover_bg)
            if card_img_label:
                card_img_label.configure(bg=hover_bg)
            self._show_preview(it)

        def on_leave(e):
            frame.configure(bg="#ffffff", highlightbackground=border)
            btn.configure(bg="#ffffff", fg=border)
            path_lbl.configure(bg="#ffffff")
            if card_img_label:
                card_img_label.configure(bg="#ffffff")

        for w in hover_widgets:
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)

        # --- D&D バインド ---
        item_id = it["id"]
        for w in hover_widgets:
            w.bind("<ButtonPress-1>",   lambda e, iid=item_id, r=grid_row, c=grid_col:
                                            self._dnd_start(e, iid, r, c))
            w.bind("<B1-Motion>",       lambda e: self._dnd_motion(e))
            w.bind("<ButtonRelease-1>", lambda e: self._dnd_drop(e))

        # btn は command があるので Press/Release で launch と D&D を両立
        # → 短いクリック（移動距離小）は launch、長い移動は D&D とする（_dnd_start で判定）
        btn.bind("<ButtonPress-1>",   lambda e, iid=item_id, r=grid_row, c=grid_col:
                                          self._dnd_start(e, iid, r, c), add="+")
        btn.bind("<B1-Motion>",       lambda e: self._dnd_motion(e),   add="+")
        btn.bind("<ButtonRelease-1>", lambda e: self._dnd_drop(e),     add="+")

        return frame

    @staticmethod
    def _darken_color(hex_color, amount=0.08):
        try:
            hex_color = hex_color.lstrip("#")
            r = max(0, int(int(hex_color[0:2], 16) * (1 - amount)))
            g = max(0, int(int(hex_color[2:4], 16) * (1 - amount)))
            b = max(0, int(int(hex_color[4:6], 16) * (1 - amount)))
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return "#dde4ee"

    def _make_empty(self, parent):
        return tk.Frame(parent, bg="#fafbfc", bd=0,
                        highlightbackground="#e8eaed", highlightthickness=1)

    # ------------------------------------------------------------------
    # Preview
    # ------------------------------------------------------------------
    def _clear_preview(self):
        self.preview_title.configure(text="ホバーで詳細表示")
        self.preview_img_label.configure(image="")
        self.preview_path_label.configure(text="")
        self.preview_desc_label.configure(text="カードにカーソルを\n合わせると、\n説明が表示されます。")

    def _show_preview(self, it):
        self.preview_title.configure(text=it["name"])
        self.preview_img_label.configure(image="")
        folder = os.path.dirname(it["path"])
        self.preview_path_label.configure(text=folder)
        desc = it.get("description", "")
        self.preview_desc_label.configure(text=desc if desc else "(説明なし)")

    # ------------------------------------------------------------------
    # Search
    # ------------------------------------------------------------------
    def _on_search(self, event=None):
        s = self.search_var.get().strip()
        if not s:
            return
        log(f"Search: {s}")
        exact = [it["path"] for it in self.data["items"] if it["name"].lower() == s.lower()]
        if len(exact) == 1:
            self.search_var.set("")
            launch_path(exact[0])
            return
        matches = [it for it in self.data["items"] if self._match_filter(it, s)]
        if not matches:
            self.search_var.set("")
            messagebox.showinfo("検索", f"「{s}」に一致するショートカットが見つかりません")
        elif len(matches) == 1:
            self.search_var.set("")
            launch_path(matches[0]["path"])
        else:
            self._select_cat("all")
            self.render_grid("all", s)

    # ------------------------------------------------------------------
    # Drag & Drop
    # ------------------------------------------------------------------
    def _dnd_start(self, e, item_id, grid_row, grid_col):
        """ボタン押下時: 0.5秒後にD&Dを有効化するタイマーをセット。"""
        self._drag_item_id   = item_id
        self._drag_origin    = (grid_row, grid_col)
        self._drag_start_xy  = (e.x_root, e.y_root)
        self._drag_armed     = False

        # 既存タイマーをキャンセル
        if self._drag_after_id:
            self.root.after_cancel(self._drag_after_id)

        # 0.5秒後にD&D有効化 → ゴースト表示
        def arm_dnd():
            self._drag_armed = True
            if self._drag_ghost:
                self._drag_ghost.destroy()
            it = next((x for x in self.data["items"] if x["id"] == self._drag_item_id), None)
            if it:
                self._drag_ghost = tk.Label(self.root, text=f"  ✥ {it['name']}  ",
                                            bg="#1a56db", fg="#ffffff",
                                            font=(FONT, 9, "bold"), padx=6, pady=4,
                                            relief=tk.FLAT)
                self._drag_ghost.place(x=e.x_root - self.root.winfo_rootx(),
                                       y=e.y_root - self.root.winfo_rooty())
                self._drag_ghost.lift()

        self._drag_after_id = self.root.after(500, arm_dnd)

    def _dnd_motion(self, e):
        if not self._drag_armed:
            return
        if self._drag_ghost:
            self._drag_ghost.place(x=e.x_root - self.root.winfo_rootx() + 4,
                                   y=e.y_root - self.root.winfo_rooty() + 4)

    def _dnd_drop(self, e):
        # タイマーキャンセル
        if self._drag_after_id:
            self.root.after_cancel(self._drag_after_id)
            self._drag_after_id = None

        if self._drag_ghost:
            self._drag_ghost.destroy()
            self._drag_ghost = None

        armed    = self._drag_armed
        item_id  = self._drag_item_id
        self._drag_item_id = None
        self._drag_armed   = False

        if not armed or item_id is None:
            return  # 短いクリック → ボタンの command に任せる

        # ドロップ先のグリッドセルを計算
        gf = self.grid_frame
        gx = e.x_root - gf.winfo_rootx()
        gy = e.y_root - gf.winfo_rooty()
        gw = gf.winfo_width()
        gh = gf.winfo_height()
        col = max(0, min(COLS - 1, int(gx / gw * COLS)))
        row = max(0, min(ROWS - 1, int(gy / gh * ROWS)))

        # 表示中アイテム（現在のカテゴリ＆フィルタ）の order を並び替える
        visible = self._visible_items(self.current_cat)
        src_pos = next((i for i, it in enumerate(visible) if it["id"] == item_id), None)
        dst_pos = row * COLS + col

        if src_pos is None or src_pos == dst_pos:
            return

        # visible リスト内で位置を入れ替え
        item = visible.pop(src_pos)
        dst_pos = min(dst_pos, len(visible))
        visible.insert(dst_pos, item)

        # order を振り直す（表示中アイテムの order のみ更新）
        # 非表示アイテムは order をずらさないよう、全アイテムの元の order を保持して隙間を作る
        all_items = sorted(self.data["items"], key=lambda x: x.get("order", 0))
        visible_ids = [it["id"] for it in visible]
        non_visible = [it for it in all_items if it["id"] not in visible_ids]

        # visible の order を、元の visible が持っていた order 値に割り当て
        old_orders = sorted([it.get("order", 0) for it in all_items if it["id"] in visible_ids])
        id_to_item = {it["id"]: it for it in self.data["items"]}
        for new_idx, it in enumerate(visible):
            id_to_item[it["id"]]["order"] = old_orders[new_idx]

        self._save()
        log(f"DnD move: {item['name']} → row={row} col={col}")
        self.render_grid(self.current_cat)

    # ------------------------------------------------------------------
    # Add / Edit dialog
    # ------------------------------------------------------------------
    def _item_dialog(self, mode="add"):
        dlg = tk.Toplevel(self.root)
        dlg.title("ショートカット追加" if mode == "add" else "ショートカット編集")
        dlg.geometry("560x490")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.configure(bg="#fafafa")

        ft  = (FONT, 11)
        pad = {"padx": 14, "pady": 6}
        row = 0

        # 編集対象選択
        item_var = tk.StringVar()
        if mode == "edit":
            tk.Label(dlg, text="編集対象:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
            items_sorted = sorted(self.data["items"], key=lambda x: x.get("order", 0))
            item_labels  = [f"{it['name']}  ({it['path'][:40]}...)" if len(it['path']) > 40
                            else f"{it['name']}  ({it['path']})" for it in items_sorted]
            if item_labels:
                item_var.set(item_labels[0])
            item_menu = tk.OptionMenu(dlg, item_var, *item_labels,
                                      command=lambda v: fill_from_selection(v))
            item_menu.configure(font=(FONT, 9), width=40)
            item_menu.grid(row=row, column=1, sticky="ew", **pad)
            row += 1

        # カテゴリ
        tk.Label(dlg, text="カテゴリ:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
        cat_var     = tk.StringVar()
        cat_options = [f"{c['id']} - {c['name']}" for c in self.data["categories"]]
        if self.current_cat != "all":
            for i, c in enumerate(self.data["categories"]):
                if c["id"] == self.current_cat:
                    cat_var.set(cat_options[i]); break
        elif cat_options:
            cat_var.set(cat_options[0])
        tk.OptionMenu(dlg, cat_var, *cat_options).grid(row=row, column=1, sticky="ew", **pad)
        row += 1

        # 名前
        tk.Label(dlg, text="名前:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
        name_var = tk.StringVar()
        tk.Entry(dlg, textvariable=name_var, font=ft, width=38).grid(row=row, column=1, sticky="ew", **pad)
        row += 1

        # パス
        tk.Label(dlg, text="パス:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
        pf = tk.Frame(dlg, bg="#fafafa"); pf.grid(row=row, column=1, sticky="ew", **pad)
        path_var = tk.StringVar()
        tk.Entry(pf, textvariable=path_var, font=ft, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(pf, text="参照", font=(FONT, 9),
                  command=lambda: self._browse_path(path_var)).pack(side=tk.RIGHT, padx=(4, 0))
        row += 1

        # 説明
        tk.Label(dlg, text="説明:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="nw", **pad)
        desc_text = tk.Text(dlg, font=(FONT, 10), width=38, height=3, wrap=tk.WORD)
        desc_text.grid(row=row, column=1, sticky="ew", **pad)
        row += 1

        # 画像
        tk.Label(dlg, text="画像:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
        imgf = tk.Frame(dlg, bg="#fafafa"); imgf.grid(row=row, column=1, sticky="ew", **pad)
        img_var = tk.StringVar()
        tk.Entry(imgf, textvariable=img_var, font=(FONT, 9), width=28).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(imgf, text="選択", font=(FONT, 9),
                  command=lambda: img_var.set(
                      filedialog.askopenfilename(
                          title="画像を選択",
                          filetypes=[("画像", "*.png *.gif *.pgm *.ppm"), ("All", "*.*")]
                      ).replace("/", "\\") or img_var.get()
                  )).pack(side=tk.RIGHT, padx=(4, 0))
        row += 1

        # 枠色
        tk.Label(dlg, text="枠色:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
        colorf = tk.Frame(dlg, bg="#fafafa"); colorf.grid(row=row, column=1, sticky="ew", **pad)
        color_var = tk.StringVar()
        tk.Entry(colorf, textvariable=color_var, font=(FONT, 9), width=12).pack(side=tk.LEFT)
        color_preview = tk.Label(colorf, text="  ", bg="#ffffff", width=4, relief=tk.SOLID, bd=1)
        color_preview.pack(side=tk.LEFT, padx=(6, 4))
        def update_color_preview(*_):
            try:
                color_preview.configure(bg=color_var.get().strip())
            except Exception:
                color_preview.configure(bg="#ffffff")
        color_var.trace_add("write", update_color_preview)
        for pc in ["#e8f4f8","#e8f5e9","#fff8e1","#fce4ec","#ede7f6",
                   "#e0f2f1","#fff3e0","#f3e5f5","#e1f5fe","#ffffff"]:
            cb = tk.Label(colorf, text="", bg=pc, width=2, height=1, relief=tk.RAISED, bd=1, cursor="hand2")
            cb.pack(side=tk.LEFT, padx=1)
            cb.bind("<Button-1>", lambda e, c=pc: color_var.set(c))
        tk.Label(colorf, text="(空=カテゴリ色)", bg="#fafafa", fg=C_TEXT_SUB,
                 font=(FONT, 8)).pack(side=tk.LEFT, padx=(6, 0))
        row += 1

        # 編集時のフォーム埋め込み
        orig_id = [None]
        if mode == "edit":
            items_sorted = sorted(self.data["items"], key=lambda x: x.get("order", 0))

        def fill_from_selection(sel):
            for it in items_sorted:
                lbl = f"{it['name']}  ({it['path'][:40]}...)" if len(it['path']) > 40 else f"{it['name']}  ({it['path']})"
                if lbl == sel:
                    orig_id[0] = it["id"]
                    name_var.set(it["name"])
                    path_var.set(it["path"])
                    for i, c in enumerate(self.data["categories"]):
                        if c["id"] == it["cat"]:
                            cat_var.set(cat_options[i]); break
                    desc_text.delete("1.0", tk.END)
                    desc_text.insert("1.0", it.get("description", ""))
                    img_var.set(it.get("image", ""))
                    color_var.set(it.get("color", ""))
                    break

        if mode == "edit" and self.data["items"]:
            fill_from_selection(item_var.get())

        # ボタン
        bf = tk.Frame(dlg, bg="#fafafa")
        bf.grid(row=row, column=0, columnspan=2, pady=14)

        def do_save():
            cat_sel = cat_var.get()
            name    = name_var.get().strip()
            path    = path_var.get().strip()
            if not cat_sel or not name or not path:
                messagebox.showwarning("入力不足", "カテゴリ・名前・パスは必須です", parent=dlg)
                return
            cat_id = cat_sel.split(" - ")[0]
            desc   = desc_text.get("1.0", tk.END).strip()
            img    = img_var.get().strip()
            color  = color_var.get().strip()

            if mode == "edit" and orig_id[0]:
                for it in self.data["items"]:
                    if it["id"] == orig_id[0]:
                        it.update({"cat": cat_id, "name": name, "path": path,
                                   "description": desc, "image": img, "color": color})
                        break
            else:
                max_order = max((x.get("order", 0) for x in self.data["items"]), default=-1)
                self.data["items"].append({
                    "id": str(uuid.uuid4()), "cat": cat_id, "name": name, "path": path,
                    "description": desc, "image": img, "color": color,
                    "order": max_order + 1,
                })
            self._save()
            log(f"{'Edit' if mode == 'edit' else 'Add'}: {cat_id} {name} {path}")
            self.render_grid(self.current_cat)
            dlg.destroy()
            messagebox.showinfo("完了", f"「{name}」を{'更新' if mode == 'edit' else '追加'}しました")

        tk.Button(bf, text="更新" if mode == "edit" else "追加",
                  bg=C_BTN_BG, fg=C_BTN_FG, font=(FONT, 11, "bold"), padx=24, pady=4,
                  command=do_save).pack(side=tk.LEFT, padx=8)

        if mode == "edit":
            def do_delete():
                if messagebox.askyesno("削除確認", "このショートカットを削除しますか？", parent=dlg):
                    self.data["items"] = [it for it in self.data["items"] if it["id"] != orig_id[0]]
                    self._save()
                    log(f"Delete: {orig_id[0]}")
                    self.render_grid(self.current_cat)
                    dlg.destroy()
            tk.Button(bf, text="削除", bg="#dc2626", fg="#ffffff",
                      font=(FONT, 11), padx=24, pady=4,
                      command=do_delete).pack(side=tk.LEFT, padx=8)

        tk.Button(bf, text="キャンセル", font=(FONT, 11), padx=24, pady=4,
                  command=dlg.destroy).pack(side=tk.LEFT, padx=8)
        dlg.columnconfigure(1, weight=1)

    def _browse_path(self, path_var):
        choice = messagebox.askyesnocancel("参照",
                                           "ファイルを選択しますか？\n\n[はい] ファイル\n[いいえ] フォルダ")
        if choice is None:
            return
        p = filedialog.askopenfilename(title="ファイルを選択") if choice \
            else filedialog.askdirectory(title="フォルダを選択")
        if p:
            path_var.set(p.replace("/", "\\"))

    # ------------------------------------------------------------------
    # Web Panel（将来活用予定）
    # ------------------------------------------------------------------
    def _toggle_web_panel(self):
        self._web_panel_open = not self._web_panel_open
        if self._web_panel_open:
            self.web_panel.pack(fill=tk.Y, side=tk.RIGHT, before=self.grid_frame)
            self._web_toggle_btn.configure(text="🌐 Web ◀", bg="#c7d9ff")
            self._web_load(self._web_url_var.get())
        else:
            self.web_panel.pack_forget()
            self._web_toggle_btn.configure(text="🌐 Web ▶", bg="#e0eaff")

    def _web_load(self, url):
        url = url.strip()
        if not url:
            return
        if not url.startswith(("http://", "https://")):
            url = "http://" + url
        self._web_url_var.set(url)
        if self._web_available:
            try:
                self.web_view.load_url(url)
            except Exception as ex:
                log(f"WebPanel load error: {ex}")
        self._web_urls[self._web_current_idx] = url
        self._save_settings()

    def _rebuild_web_tabs(self):
        for w in self._wp_tab_frame.winfo_children():
            w.destroy()
        for i, u in enumerate(self._web_urls):
            label = u.replace("https://", "").replace("http://", "").split("/")[0] or f"Tab{i+1}"
            is_active = (i == self._web_current_idx)
            btn = tk.Label(self._wp_tab_frame, text=label,
                           bg=C_SIDEBAR_ACTIVE_BG if is_active else C_SIDEBAR_BG,
                           fg=C_SIDEBAR_ACTIVE_FG if is_active else C_SIDEBAR_FG,
                           font=(FONT, 8, "bold" if is_active else "normal"),
                           padx=8, pady=4, cursor="hand2")
            btn.pack(side=tk.LEFT)
            btn.bind("<Button-1>", lambda e, idx=i: self._switch_web_tab(idx))
        add = tk.Label(self._wp_tab_frame, text="＋", bg=C_SIDEBAR_BG, fg=C_ADD_FG,
                       font=(FONT, 9, "bold"), padx=6, pady=4, cursor="hand2")
        add.pack(side=tk.LEFT)
        add.bind("<Button-1>", lambda e: self._add_web_tab())
        if len(self._web_urls) > 1:
            rm = tk.Label(self._wp_tab_frame, text="✕", bg=C_SIDEBAR_BG, fg="#dc2626",
                          font=(FONT, 9), padx=6, pady=4, cursor="hand2")
            rm.pack(side=tk.RIGHT)
            rm.bind("<Button-1>", lambda e: self._remove_web_tab())

    def _switch_web_tab(self, idx):
        self._web_current_idx = idx
        self._web_url_var.set(self._web_urls[idx])
        self._rebuild_web_tabs()
        self._web_load(self._web_urls[idx])

    def _add_web_tab(self):
        self._web_urls.append("http://localhost/")
        self._web_current_idx = len(self._web_urls) - 1
        self._web_url_var.set(self._web_urls[self._web_current_idx])
        self._rebuild_web_tabs()
        self._save_settings()

    def _remove_web_tab(self):
        if len(self._web_urls) <= 1:
            return
        del self._web_urls[self._web_current_idx]
        self._web_current_idx = max(0, self._web_current_idx - 1)
        self._web_url_var.set(self._web_urls[self._web_current_idx])
        self._rebuild_web_tabs()
        self._web_load(self._web_urls[self._web_current_idx])
        self._save_settings()


if __name__ == "__main__":
    log("=== Application Start ===")
    root = tk.Tk()
    app = App(root)
    root.mainloop()
