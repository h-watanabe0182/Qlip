import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import os
import sys
import subprocess
import webbrowser
import logging
import json

# --- Constants ---
if getattr(sys, 'frozen', False):
    # PyInstaller でビルドされた exe の場合
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(APP_DIR, "Qlip.log")
META_FILE = os.path.join(APP_DIR, "metadata.json")
IMG_DIR = os.path.join(APP_DIR, "images")
FONT = "Meiryo"
COLS = 8
ROWS = 5

# --- Colors (Light mode) ---
C_BG = "#f2f3f5"
C_HEADER_BG = "#ffffff"
C_HEADER_BORDER = "#d0d5dd"
C_SIDEBAR_BG = "#f7f8fa"
C_SIDEBAR_BORDER = "#e0e3e8"
C_SIDEBAR_ACTIVE_BG = "#e0eaff"
C_SIDEBAR_ACTIVE_FG = "#1a56db"
C_SIDEBAR_FG = "#444444"
C_CARD_BG = "#ffffff"
C_CARD_HOVER = "#edf2ff"
C_CARD_BORDER = "#d0d5dd"
C_BTN_BG = "#2563eb"
C_BTN_FG = "#ffffff"
C_BTN_HOVER = "#1d4ed8"
C_TEXT = "#1f2937"
C_TEXT_SUB = "#6b7280"
C_PREVIEW_BG = "#ffffff"
C_PREVIEW_BORDER = "#d0d5dd"
C_ADD_FG = "#16a34a"
C_EDIT_FG = "#d97706"

# Default gentle card colors per category
DEFAULT_CAT_COLORS = {
    "0": "#90caf9",  # sky blue (favorites)
    "1": "#a5d6a7",  # green (projects)
    "2": "#ffe082",  # yellow (tools)
    "3": "#f48fb1",  # pink (documents)
    "4": "#ce93d8",  # lavender (sales)
    "5": "#80cbc4",  # teal (misc)
}
DEFAULT_CARD_COLOR = "#90caf9"  # blue fallback

# --- Logging ---
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format="%(asctime)s | %(message)s", encoding="utf-8")
log = logging.info

# --- Metadata (image path + description per item) ---
def load_metadata():
    if os.path.exists(META_FILE):
        try:
            with open(META_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_metadata(meta):
    with open(META_FILE, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

# --- Config ---
def find_config():
    # Same folder as the exe/script
    for f in os.scandir(APP_DIR):
        if f.is_file() and f.name.lower().endswith(".txt"):
            return f.path
    return None

def load_config(path):
    categories = []
    items = []
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

def save_config(path, categories, items):
    lines = []
    for num, name in categories:
        lines.append(f"# CATEGORY {num} {name}")
    lines.append("")
    for cat, name, p in items:
        lines.append(f"{cat} {name} {p}")
    lines.append("")
    with open(path, "w", encoding="cp932", errors="replace") as f:
        f.write("\n".join(lines))

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
            # CREATE_NEW_CONSOLE: 新しいコンソール窓を開いて実行（対話型スクリプト対応）
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

        # Window icon
        ico_path = os.path.join(APP_DIR, "Qlip.ico")
        if os.path.exists(ico_path):
            try:
                self.root.iconbitmap(ico_path)
            except Exception:
                pass

        self.categories = []
        self.items = []
        self.current_cat = "all"
        self.card_widgets = []
        self.card_images = []  # keep references to prevent GC
        self.metadata = load_metadata()
        self.preview_image = None  # keep reference to prevent GC
        self.show_card_image = tk.BooleanVar(
            value=self.metadata.get("__settings__", {}).get("show_card_image", True)
        )
        self._web_panel_open = False
        self._web_urls = self.metadata.get("__settings__", {}).get(
            "web_urls", ["http://localhost/"]
        )
        self._web_current_idx = 0

        self.cfg_path = find_config()
        if self.cfg_path:
            log(f"Config: {self.cfg_path}")
            self.categories, self.items = load_config(self.cfg_path)
            log(f"Loaded {len(self.categories)} categories, {len(self.items)} items")

        if not os.path.exists(IMG_DIR):
            os.makedirs(IMG_DIR, exist_ok=True)

        self._build_ui()

        # Initial: first category (favorites)
        if self.categories:
            self._select_cat(self.categories[0][0])
        else:
            self.render_grid("all")

        # Focus search
        self.search_entry.focus_set()

    def _build_ui(self):
        # === Header ===
        header = tk.Frame(self.root, bg=C_HEADER_BG, height=70,
                          highlightbackground=C_HEADER_BORDER, highlightthickness=1)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        # Header image (right side, cropped to header height)
        self._header_image = None
        img_path = os.path.join(APP_DIR, "参考", "Qlipイメージ画像.png")
        if not os.path.exists(img_path):
            # fallback: same folder
            img_path = os.path.join(APP_DIR, "Qlipイメージ画像.png")
        if os.path.exists(img_path):
            try:
                from PIL import Image, ImageTk
                pil_img = Image.open(img_path)
                # Scale to header height keeping aspect ratio
                h = 70
                w = int(pil_img.width * h / pil_img.height)
                pil_img = pil_img.resize((w, h), Image.LANCZOS)
                self._header_image = ImageTk.PhotoImage(pil_img)
                tk.Label(header, image=self._header_image, bg=C_HEADER_BG,
                         bd=0).pack(side=tk.RIGHT, padx=(0, 0))
            except Exception:
                pass

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

        # Web パネル トグルボタン
        self._web_toggle_btn = tk.Label(
            header, text="🌐 Web ▶", bg="#e0eaff", fg=C_SIDEBAR_ACTIVE_FG,
            font=(FONT, 9, "bold"), padx=10, pady=4, cursor="hand2",
            relief=tk.FLAT, bd=0)
        self._web_toggle_btn.pack(side=tk.RIGHT, padx=(0, 8), pady=16)
        self._web_toggle_btn.bind("<Button-1>", lambda e: self._toggle_web_panel())
        self._web_toggle_btn.bind("<Enter>", lambda e: self._web_toggle_btn.configure(bg="#c7d9ff"))
        self._web_toggle_btn.bind("<Leave>", lambda e: self._web_toggle_btn.configure(
            bg="#e0eaff" if not self._web_panel_open else "#c7d9ff"))

        # === Body ===
        body = tk.Frame(self.root, bg=C_BG)
        body.pack(fill=tk.BOTH, expand=True)

        # === Sidebar ===
        sidebar = tk.Frame(body, bg=C_SIDEBAR_BG, width=140,
                           highlightbackground=C_SIDEBAR_BORDER, highlightthickness=1)
        sidebar.pack(fill=tk.Y, side=tk.LEFT)
        sidebar.pack_propagate(False)

        self.cat_buttons = {}
        self._add_cat_btn(sidebar, "all", "All")
        for num, name in self.categories:
            self._add_cat_btn(sidebar, num, name)

        # Spacer
        tk.Frame(sidebar, bg=C_SIDEBAR_BG, height=20).pack(fill=tk.X)

        # Add button
        add_btn = tk.Label(sidebar, text="＋ 追加", bg=C_SIDEBAR_BG, fg=C_ADD_FG,
                           font=(FONT, 10, "bold"), anchor="w", padx=12, pady=7, cursor="hand2")
        add_btn.pack(fill=tk.X)
        add_btn.bind("<Button-1>", lambda e: self._item_dialog(mode="add"))
        add_btn.bind("<Enter>", lambda e: add_btn.configure(bg="#e8fce8"))
        add_btn.bind("<Leave>", lambda e: add_btn.configure(bg=C_SIDEBAR_BG))

        # Edit button
        edit_btn = tk.Label(sidebar, text="✎ 編集", bg=C_SIDEBAR_BG, fg=C_EDIT_FG,
                            font=(FONT, 10, "bold"), anchor="w", padx=12, pady=7, cursor="hand2")
        edit_btn.pack(fill=tk.X)
        edit_btn.bind("<Button-1>", lambda e: self._item_dialog(mode="edit"))
        edit_btn.bind("<Enter>", lambda e: edit_btn.configure(bg="#fff5e0"))
        edit_btn.bind("<Leave>", lambda e: edit_btn.configure(bg=C_SIDEBAR_BG))

        # Separator
        tk.Frame(sidebar, bg=C_SIDEBAR_BORDER, height=1).pack(fill=tk.X, padx=8, pady=(10, 4))

        # Settings label
        tk.Label(sidebar, text="設定", bg=C_SIDEBAR_BG, fg=C_TEXT_SUB,
                 font=(FONT, 8), anchor="w", padx=14).pack(fill=tk.X)

        # Card image toggle
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
            c.create_rounded_rect = lambda x1,y1,x2,y2,r,**kw: c.create_polygon(
                x1+r,y1, x2-r,y1, x2,y1, x2,y1+r, x2,y2-r, x2,y2,
                x2-r,y2, x1+r,y2, x1,y2, x1,y2-r, x1,y1+r, x1,y1,
                smooth=True, **kw)
            c.create_rounded_rect(0, 2, 34, 16, 7, fill=bg_col, outline="")
            knob_x = 22 if on else 10
            c.create_oval(knob_x-6, 3, knob_x+6, 15, fill="#ffffff", outline="")

        def _toggle_card_image(e=None):
            self.show_card_image.set(not self.show_card_image.get())
            _draw_toggle()
            # 設定を保存
            settings = self.metadata.get("__settings__", {})
            settings["show_card_image"] = self.show_card_image.get()
            self.metadata["__settings__"] = settings
            save_metadata(self.metadata)
            self.render_grid(self.current_cat)

        _draw_toggle()
        self._img_toggle_canvas.bind("<Button-1>", _toggle_card_image)
        img_toggle_frame.bind("<Button-1>", _toggle_card_image)
        for child in img_toggle_frame.winfo_children():
            child.bind("<Button-1>", _toggle_card_image)

        # === Preview pane (2 columns wide) ===
        self.preview_frame = tk.Frame(body, bg=C_PREVIEW_BG, width=260,
                                      highlightbackground=C_PREVIEW_BORDER, highlightthickness=1)
        self.preview_frame.pack(fill=tk.Y, side=tk.LEFT)
        self.preview_frame.pack_propagate(False)

        self.preview_title = tk.Label(self.preview_frame, text="", bg=C_PREVIEW_BG, fg=C_TEXT,
                                      font=(FONT, 12, "bold"), wraplength=240, justify=tk.CENTER)
        self.preview_title.pack(pady=(12, 4), padx=8)

        self.preview_img_label = tk.Label(self.preview_frame, bg=C_PREVIEW_BG)
        # 画像表示は廃止（カード内画像のみ）

        self.preview_path_label = tk.Label(self.preview_frame, text="", bg=C_PREVIEW_BG, fg=C_TEXT_SUB,
                                           font=(FONT, 8), wraplength=240, justify=tk.CENTER)
        self.preview_path_label.pack(pady=(0, 4), padx=8)

        self.preview_desc_label = tk.Label(self.preview_frame, text="", bg=C_PREVIEW_BG, fg=C_TEXT,
                                           font=(FONT, 9), wraplength=240, justify=tk.LEFT,
                                           anchor="nw")
        self.preview_desc_label.pack(fill=tk.BOTH, expand=True, padx=12, pady=(4, 12))

        self._clear_preview()

        # === Grid area ===
        self.grid_frame = tk.Frame(body, bg=C_BG)
        self.grid_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        for c in range(COLS):
            self.grid_frame.columnconfigure(c, weight=1, uniform="col")
        for r in range(ROWS):
            self.grid_frame.rowconfigure(r, weight=1, uniform="row")

        # === Web Panel (右サイドバー、初期非表示) ===
        self.web_panel = tk.Frame(body, bg=C_SIDEBAR_BG, width=420,
                                  highlightbackground=C_SIDEBAR_BORDER, highlightthickness=1)
        self.web_panel.pack_propagate(False)
        # pack しない（toggle で制御）

        # Webパネル ヘッダー行
        wp_header = tk.Frame(self.web_panel, bg=C_HEADER_BG,
                             highlightbackground=C_HEADER_BORDER, highlightthickness=1)
        wp_header.pack(fill=tk.X)

        tk.Label(wp_header, text="🌐", bg=C_HEADER_BG, font=(FONT, 11)).pack(side=tk.LEFT, padx=(8,2), pady=6)
        self._web_url_var = tk.StringVar(value=self._web_urls[self._web_current_idx])
        url_entry = tk.Entry(wp_header, textvariable=self._web_url_var,
                             font=(FONT, 9), relief=tk.SOLID, bd=1)
        url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4, pady=6)
        url_entry.bind("<Return>", lambda e: self._web_load(self._web_url_var.get()))

        tk.Button(wp_header, text="→", font=(FONT, 9, "bold"),
                  bg=C_BTN_BG, fg=C_BTN_FG, relief=tk.FLAT, padx=6,
                  command=lambda: self._web_load(self._web_url_var.get())
                  ).pack(side=tk.LEFT, padx=(0,2), pady=6)
        tk.Button(wp_header, text="↺", font=(FONT, 9, "bold"),
                  bg="#f0f0f0", relief=tk.FLAT, padx=6,
                  command=lambda: self._web_load(self._web_url_var.get())
                  ).pack(side=tk.LEFT, padx=(0,4), pady=6)
        tk.Button(wp_header, text="✕", font=(FONT, 9),
                  bg="#f0f0f0", relief=tk.FLAT, padx=6,
                  command=self._toggle_web_panel
                  ).pack(side=tk.RIGHT, padx=(0,4), pady=6)

        # URLタブ行
        self._wp_tab_frame = tk.Frame(self.web_panel, bg=C_SIDEBAR_BG)
        self._wp_tab_frame.pack(fill=tk.X)
        self._rebuild_web_tabs()

        # Webビュー
        try:
            from tkinterweb import HtmlFrame
            self.web_view = HtmlFrame(self.web_panel, messages_enabled=False)
            self.web_view.pack(fill=tk.BOTH, expand=True)
            self._web_available = True
        except Exception:
            self.web_view = tk.Label(self.web_panel,
                                     text="tkinterweb が利用できません",
                                     bg=C_SIDEBAR_BG, fg=C_TEXT_SUB)
            self.web_view.pack(fill=tk.BOTH, expand=True)
            self._web_available = False

    def _add_cat_btn(self, parent, cat_id, label):
        btn = tk.Label(parent, text=label, bg=C_SIDEBAR_BG, fg=C_SIDEBAR_FG,
                       font=(FONT, 10), anchor="w", padx=14, pady=8, cursor="hand2")
        btn.pack(fill=tk.X)
        btn.bind("<Button-1>", lambda e, c=cat_id: self._select_cat(c))
        btn.bind("<Enter>", lambda e, b=btn, c=cat_id: b.configure(bg="#e8ecf0") if c != self.current_cat else None)
        btn.bind("<Leave>", lambda e, b=btn, c=cat_id: b.configure(bg=C_SIDEBAR_BG) if c != self.current_cat else None)
        self.cat_buttons[cat_id] = btn

    def _select_cat(self, cat):
        for cid, btn in self.cat_buttons.items():
            btn.configure(bg=C_SIDEBAR_BG, fg=C_SIDEBAR_FG, font=(FONT, 10))
        self.cat_buttons[cat].configure(bg=C_SIDEBAR_ACTIVE_BG, fg=C_SIDEBAR_ACTIVE_FG,
                                        font=(FONT, 10, "bold"))
        self.current_cat = cat
        self.render_grid(cat)

    def _match_filter(self, name, path, filter_str):
        """Check if filter_str matches name, path, or description."""
        q = filter_str.lower()
        if q in name.lower():
            return True
        if q in path.lower():
            return True
        meta = self.metadata.get(name, {})
        desc = meta.get("description", "")
        if desc and q in desc.lower():
            return True
        return False

    def render_grid(self, cat, filter_str=""):
        for w in self.card_widgets:
            w.destroy()
        self.card_widgets = []
        self.card_images = []

        visible = []
        for ic, name, path in self.items:
            if cat != "all" and ic != cat:
                continue
            if filter_str and not self._match_filter(name, path, filter_str):
                continue
            visible.append((ic, name, path))
            if len(visible) >= COLS * ROWS:
                break

        idx = 0
        for r in range(ROWS):
            for c in range(COLS):
                if idx < len(visible):
                    card = self._make_card(self.grid_frame, *visible[idx])
                else:
                    card = self._make_empty(self.grid_frame)
                card.grid(row=r, column=c, sticky="nsew", padx=2, pady=2)
                self.card_widgets.append(card)
                idx += 1

    def _get_card_color(self, name, item_cat):
        """Get card color: metadata > category default > fallback."""
        meta = self.metadata.get(name, {})
        if meta.get("color"):
            return meta["color"]
        return DEFAULT_CAT_COLORS.get(item_cat, DEFAULT_CARD_COLOR)

    def _make_card(self, parent, item_cat, name, path):
        accent = self._get_card_color(name, item_cat)   # カテゴリカラー（淡い）
        border = self._darken_color(accent, 0.35)        # ボーダー・テキスト用（濃い）
        hover_bg = border                                # ホバー時背景 = ボーダー色

        # カード外枠: 白背景 + アクセントカラーのボーダー（2px）
        frame = tk.Frame(parent, bg="#ffffff", cursor="hand2",
                         highlightbackground=border, highlightthickness=2)

        # カード内画像（show_card_image=True かつ 画像が設定されている場合）
        card_img_label = None
        if self.show_card_image.get():
            meta = self.metadata.get(name, {})
            img_path = meta.get("image", "")
            if img_path and os.path.exists(img_path):
                try:
                    from PIL import Image, ImageTk
                    pil_img = Image.open(img_path)
                    # カード内に収まるよう縮小（最大 96x54 の横長）
                    max_w, max_h = 96, 54
                    pil_img.thumbnail((max_w, max_h), Image.LANCZOS)
                    tk_img = ImageTk.PhotoImage(pil_img)
                    self.card_images.append(tk_img)
                    card_img_label = tk.Label(frame, image=tk_img, bg="#ffffff",
                                             cursor="hand2", bd=0)
                    card_img_label.pack(pady=(4, 0))
                    card_img_label.bind("<Button-1>", lambda e, p=path: launch_path(p))
                except Exception:
                    pass

        # ボタン: 白背景・アクセントカラーのテキスト、ホバーで反転
        btn = tk.Button(frame, text=name, bg="#ffffff", fg=border,
                        activebackground=hover_bg, activeforeground="#ffffff",
                        font=(FONT, 9, "bold"), relief=tk.FLAT, bd=0, cursor="hand2",
                        wraplength=130, padx=4, pady=3,
                        command=lambda p=path: launch_path(p))
        btn.pack(fill=tk.X, padx=4, pady=(2 if card_img_label else 5, 1))

        # パス表示: 画像あり時は省略、なし時は従来通り
        if card_img_label is None:
            basename = os.path.basename(path)
            parentdir = os.path.basename(os.path.dirname(path))
            CHARS_PER_LINE = 13
            name_lines   = max(1, (len(name)      + CHARS_PER_LINE - 1) // CHARS_PER_LINE)
            parent_lines = max(1, (len(parentdir)  + CHARS_PER_LINE - 1) // CHARS_PER_LINE)
            file_lines   = max(1, (len(basename)   + CHARS_PER_LINE - 1) // CHARS_PER_LINE)
            if name_lines + parent_lines + file_lines > 4:
                disp = basename
            else:
                disp = (parentdir + "\n" + basename) if parentdir else basename
            path_lbl = tk.Label(frame, text=disp, bg="#ffffff", fg="#888888",
                                font=(FONT, 8), wraplength=130, justify=tk.CENTER)
            path_lbl.pack(fill=tk.BOTH, expand=True, padx=3, pady=(0, 4))
            path_lbl.bind("<Button-1>", lambda e, p=path: launch_path(p))
        else:
            path_lbl = tk.Label(frame, text="", bg="#ffffff")
            path_lbl.pack(pady=(0, 2))

        # ホバー: 全体をアクセント色に塗りつぶし・テキスト白に反転
        hover_widgets = [frame, btn, path_lbl]
        if card_img_label:
            hover_widgets.append(card_img_label)

        def on_enter(e):
            frame.configure(bg=hover_bg, highlightbackground=hover_bg)
            btn.configure(bg=hover_bg, fg="#ffffff")
            path_lbl.configure(bg=hover_bg)
            if card_img_label:
                card_img_label.configure(bg=hover_bg)
            self._show_preview(name, path)
        def on_leave(e):
            frame.configure(bg="#ffffff", highlightbackground=border)
            btn.configure(bg="#ffffff", fg=border)
            path_lbl.configure(bg="#ffffff")
            if card_img_label:
                card_img_label.configure(bg="#ffffff")
        for w in hover_widgets:
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)

        return frame

    @staticmethod
    def _darken_color(hex_color, amount=0.08):
        """Darken a hex color by a fraction."""
        try:
            hex_color = hex_color.lstrip("#")
            r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
            r = max(0, int(r * (1 - amount)))
            g = max(0, int(g * (1 - amount)))
            b = max(0, int(b * (1 - amount)))
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return "#dde4ee"

    def _make_empty(self, parent):
        return tk.Frame(parent, bg="#fafbfc", bd=0,
                        highlightbackground="#e8eaed", highlightthickness=1)

    # === Preview ===
    def _clear_preview(self):
        self.preview_title.configure(text="ホバーで詳細表示")
        self.preview_img_label.configure(image="")
        self.preview_path_label.configure(text="")
        self.preview_desc_label.configure(text="カードにカーソルを\n合わせると、\n説明が表示されます。")
        self.preview_image = None

    def _show_preview(self, name, path):
        self.preview_title.configure(text=name)
        self.preview_img_label.configure(image="")
        self.preview_image = None

        # フォルダ表示（親フォルダ名）
        folder = os.path.dirname(path)
        self.preview_path_label.configure(text=folder)

        meta = self.metadata.get(name, {})
        desc = meta.get("description", "")
        self.preview_desc_label.configure(text=desc if desc else "(説明なし)")

    # === Search ===
    def _on_search(self, event=None):
        s = self.search_var.get().strip()
        if not s:
            return
        log(f"Search: {s}")

        # Exact name match → launch immediately
        exact = [p for c, n, p in self.items if n.lower() == s.lower()]
        if len(exact) == 1:
            self.search_var.set("")
            launch_path(exact[0])
            return

        # Broad match (name, path, description)
        matches = [(c, n, p) for c, n, p in self.items if self._match_filter(n, p, s)]

        if not matches:
            self.search_var.set("")
            messagebox.showinfo("検索", f"「{s}」に一致するショートカットが見つかりません")
        elif len(matches) == 1:
            self.search_var.set("")
            launch_path(matches[0][2])
        else:
            # Filter display (keep search text visible for reference)
            self._select_cat("all")
            self.render_grid("all", s)

    # === Add / Edit dialog ===
    def _item_dialog(self, mode="add", prefill_name=""):
        dlg = tk.Toplevel(self.root)
        dlg.title("ショートカット追加" if mode == "add" else "ショートカット編集")
        dlg.geometry("560x490")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.configure(bg="#fafafa")

        ft = (FONT, 11)
        pad = {"padx": 14, "pady": 6}

        row = 0

        # For edit mode: item selector
        item_var = tk.StringVar()
        if mode == "edit":
            tk.Label(dlg, text="編集対象:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
            item_names = [f"{n}  ({p[:40]}...)" if len(p) > 40 else f"{n}  ({p})"
                          for c, n, p in self.items]
            if item_names:
                item_var.set(item_names[0])
            item_menu = tk.OptionMenu(dlg, item_var, *item_names,
                                      command=lambda v: fill_from_selection(v))
            item_menu.configure(font=(FONT, 9), width=40)
            item_menu.grid(row=row, column=1, sticky="ew", **pad)
            row += 1

        # Category
        tk.Label(dlg, text="カテゴリ:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
        cat_var = tk.StringVar()
        cat_options = [f"{num} - {name}" for num, name in self.categories]
        if self.current_cat != "all":
            for i, (num, name) in enumerate(self.categories):
                if num == self.current_cat:
                    cat_var.set(cat_options[i])
                    break
        elif cat_options:
            cat_var.set(cat_options[0])
        cat_menu = tk.OptionMenu(dlg, cat_var, *cat_options)
        cat_menu.configure(font=ft, width=35)
        cat_menu.grid(row=row, column=1, sticky="ew", **pad)
        row += 1

        # Name
        tk.Label(dlg, text="名前:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
        name_var = tk.StringVar()
        tk.Entry(dlg, textvariable=name_var, font=ft, width=38).grid(row=row, column=1, sticky="ew", **pad)
        row += 1

        # Path
        tk.Label(dlg, text="パス:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
        pf = tk.Frame(dlg, bg="#fafafa")
        pf.grid(row=row, column=1, sticky="ew", **pad)
        path_var = tk.StringVar()
        tk.Entry(pf, textvariable=path_var, font=ft, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(pf, text="参照", font=(FONT, 9),
                  command=lambda: self._browse_path(path_var)).pack(side=tk.RIGHT, padx=(4, 0))
        row += 1

        # Description
        tk.Label(dlg, text="説明:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="nw", **pad)
        desc_text = tk.Text(dlg, font=(FONT, 10), width=38, height=3, wrap=tk.WORD)
        desc_text.grid(row=row, column=1, sticky="ew", **pad)
        row += 1

        # Image
        tk.Label(dlg, text="画像:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
        imgf = tk.Frame(dlg, bg="#fafafa")
        imgf.grid(row=row, column=1, sticky="ew", **pad)
        img_var = tk.StringVar()
        tk.Entry(imgf, textvariable=img_var, font=(FONT, 9), width=28).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(imgf, text="選択", font=(FONT, 9),
                  command=lambda: img_var.set(
                      filedialog.askopenfilename(title="画像を選択",
                                                 filetypes=[("画像", "*.png *.gif *.pgm *.ppm"), ("All", "*.*")])
                      .replace("/", "\\") or img_var.get()
                  )).pack(side=tk.RIGHT, padx=(4, 0))
        row += 1

        # Color
        tk.Label(dlg, text="枠色:", bg="#fafafa", font=ft).grid(row=row, column=0, sticky="w", **pad)
        colorf = tk.Frame(dlg, bg="#fafafa")
        colorf.grid(row=row, column=1, sticky="ew", **pad)
        color_var = tk.StringVar()
        color_entry = tk.Entry(colorf, textvariable=color_var, font=(FONT, 9), width=12)
        color_entry.pack(side=tk.LEFT)
        color_preview = tk.Label(colorf, text="  ", bg="#ffffff", width=4, relief=tk.SOLID, bd=1)
        color_preview.pack(side=tk.LEFT, padx=(6, 4))
        def update_color_preview(*_):
            c = color_var.get().strip()
            try:
                color_preview.configure(bg=c)
            except Exception:
                color_preview.configure(bg="#ffffff")
        color_var.trace_add("write", update_color_preview)
        # Preset color buttons
        preset_colors = ["#e8f4f8", "#e8f5e9", "#fff8e1", "#fce4ec", "#ede7f6",
                         "#e0f2f1", "#fff3e0", "#f3e5f5", "#e1f5fe", "#ffffff"]
        for pc in preset_colors:
            cb = tk.Label(colorf, text="", bg=pc, width=2, height=1, relief=tk.RAISED, bd=1, cursor="hand2")
            cb.pack(side=tk.LEFT, padx=1)
            cb.bind("<Button-1>", lambda e, c=pc: color_var.set(c))
        tk.Label(colorf, text="(空=カテゴリ色)", bg="#fafafa", fg=C_TEXT_SUB,
                 font=(FONT, 8)).pack(side=tk.LEFT, padx=(6, 0))
        row += 1

        # Fill function for edit mode
        orig_idx = [0]

        def fill_from_selection(sel):
            for i, (c, n, p) in enumerate(self.items):
                label = f"{n}  ({p[:40]}...)" if len(p) > 40 else f"{n}  ({p})"
                if label == sel:
                    orig_idx[0] = i
                    name_var.set(n)
                    path_var.set(p)
                    for j, (cn, cl) in enumerate(self.categories):
                        if cn == c:
                            cat_var.set(cat_options[j])
                            break
                    meta = self.metadata.get(n, {})
                    desc_text.delete("1.0", tk.END)
                    desc_text.insert("1.0", meta.get("description", ""))
                    img_var.set(meta.get("image", ""))
                    color_var.set(meta.get("color", ""))
                    break

        if mode == "edit" and self.items:
            fill_from_selection(item_var.get())

        # Buttons
        bf = tk.Frame(dlg, bg="#fafafa")
        bf.grid(row=row, column=0, columnspan=2, pady=14)

        def do_save():
            cat_sel = cat_var.get()
            name = name_var.get().strip()
            path = path_var.get().strip()
            if not cat_sel or not name or not path:
                messagebox.showwarning("入力不足", "カテゴリ・名前・パスは必須です", parent=dlg)
                return
            cat_num = cat_sel.split(" - ")[0]
            desc = desc_text.get("1.0", tk.END).strip()
            img = img_var.get().strip()
            color = color_var.get().strip()

            if mode == "edit":
                old_name = self.items[orig_idx[0]][1]
                self.items[orig_idx[0]] = (cat_num, name, path)
                if old_name != name and old_name in self.metadata:
                    self.metadata[name] = self.metadata.pop(old_name)
            else:
                self.items.append((cat_num, name, path))

            # Save metadata
            meta_entry = {"description": desc, "image": img, "color": color}
            self.metadata[name] = meta_entry
            save_metadata(self.metadata)

            # Save config
            save_config(self.cfg_path, self.categories, self.items)
            log(f"{'Edit' if mode == 'edit' else 'Add'}: {cat_num} {name} {path}")

            self.render_grid(self.current_cat)
            dlg.destroy()
            messagebox.showinfo("完了", f"「{name}」を{'更新' if mode == 'edit' else '追加'}しました")

        action_text = "更新" if mode == "edit" else "追加"
        tk.Button(bf, text=action_text, bg=C_BTN_BG, fg=C_BTN_FG,
                  font=(FONT, 11, "bold"), padx=24, pady=4,
                  command=do_save).pack(side=tk.LEFT, padx=8)

        if mode == "edit":
            def do_delete():
                if messagebox.askyesno("削除確認", "このショートカットを削除しますか？", parent=dlg):
                    old_name = self.items[orig_idx[0]][1]
                    del self.items[orig_idx[0]]
                    if old_name in self.metadata:
                        del self.metadata[old_name]
                    save_metadata(self.metadata)
                    save_config(self.cfg_path, self.categories, self.items)
                    log(f"Delete: {old_name}")
                    self.render_grid(self.current_cat)
                    dlg.destroy()

            tk.Button(bf, text="削除", bg="#dc2626", fg="#ffffff",
                      font=(FONT, 11), padx=24, pady=4,
                      command=do_delete).pack(side=tk.LEFT, padx=8)

        tk.Button(bf, text="キャンセル", font=(FONT, 11),
                  padx=24, pady=4, command=dlg.destroy).pack(side=tk.LEFT, padx=8)

        dlg.columnconfigure(1, weight=1)

    # === Web Panel ===
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
        if not url.strip():
            return
        url = url.strip()
        if not url.startswith(("http://", "https://")):
            url = "http://" + url
        self._web_url_var.set(url)
        if self._web_available:
            try:
                self.web_view.load_url(url)
            except Exception as ex:
                log(f"WebPanel load error: {ex}")
        # 現在のタブURLを更新して保存
        self._web_urls[self._web_current_idx] = url
        self._save_web_settings()

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
                           padx=8, pady=4, cursor="hand2",
                           relief=tk.FLAT)
            btn.pack(side=tk.LEFT)
            btn.bind("<Button-1>", lambda e, idx=i: self._switch_web_tab(idx))
        # ＋ボタン
        add = tk.Label(self._wp_tab_frame, text="＋", bg=C_SIDEBAR_BG, fg=C_ADD_FG,
                       font=(FONT, 9, "bold"), padx=6, pady=4, cursor="hand2")
        add.pack(side=tk.LEFT)
        add.bind("<Button-1>", lambda e: self._add_web_tab())
        # 削除ボタン（タブ2枚以上の時）
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
        self._save_web_settings()

    def _remove_web_tab(self):
        if len(self._web_urls) <= 1:
            return
        del self._web_urls[self._web_current_idx]
        self._web_current_idx = max(0, self._web_current_idx - 1)
        self._web_url_var.set(self._web_urls[self._web_current_idx])
        self._rebuild_web_tabs()
        self._web_load(self._web_urls[self._web_current_idx])
        self._save_web_settings()

    def _save_web_settings(self):
        settings = self.metadata.get("__settings__", {})
        settings["web_urls"] = self._web_urls
        self.metadata["__settings__"] = settings
        save_metadata(self.metadata)

    def _browse_path(self, path_var):
        choice = messagebox.askyesnocancel("参照",
                                           "ファイルを選択しますか？\n\n[はい] ファイル\n[いいえ] フォルダ")
        if choice is None:
            return
        if choice:
            p = filedialog.askopenfilename(title="ファイルを選択")
        else:
            p = filedialog.askdirectory(title="フォルダを選択")
        if p:
            path_var.set(p.replace("/", "\\"))


if __name__ == "__main__":
    log("=== Application Start ===")
    root = tk.Tk()
    app = App(root)
    root.mainloop()
