# -*- coding: utf-8 -*-
"""桌面整理工具：自动识别文件类型并归类，预览确认后一键整理，支持撤销。"""

import ctypes
import json
import math
import os
import shutil
import sys
import time
import tkinter as tk
import tkinter.font as tkfont
from collections import Counter
from ctypes import wintypes
from tkinter import filedialog, messagebox, ttk

APP_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
RES_DIR = getattr(sys, "_MEIPASS", APP_DIR)
LOG_PATH = os.path.join(APP_DIR, "整理记录.json")
SETTINGS_PATH = os.path.join(APP_DIR, "设置.json")
BANNER_PATH = os.path.join(RES_DIR, "assets", "banner.jpg")
MAX_BATCHES = 20
BANNER_H = 150

BG = "#FFF6FA"
CARD = "#FFFFFF"
TEXT = "#4A3B4E"
MUTED = "#A0919F"
PINK = "#FF77AA"
PINK_DARK = "#F2568F"
PINK_SOFT = "#FFE3EF"
LINE = "#F6DCE8"
FONT_FAMILY = "Microsoft YaHei UI"

CATEGORIES = [
    ("文档", (".doc", ".docx", ".pdf", ".txt", ".md", ".rtf", ".wps", ".odt",
             ".pages", ".epub", ".mobi", ".xmind", ".chm")),
    ("表格", (".xls", ".xlsx", ".xlsm", ".csv", ".et", ".numbers")),
    ("演示文稿", (".ppt", ".pptx", ".pps", ".ppsx", ".dps", ".key")),
    ("图片", (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".svg", ".ico",
             ".tif", ".tiff", ".heic", ".psd", ".ai", ".raw", ".cr2", ".nef")),
    ("视频", (".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm", ".m4v",
             ".rmvb", ".ts", ".mpg", ".mpeg")),
    ("音频", (".mp3", ".wav", ".flac", ".aac", ".m4a", ".ogg", ".wma", ".ape",
             ".mid", ".midi")),
    ("压缩包", (".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".iso", ".cab")),
    ("安装包", (".exe", ".msi", ".dmg", ".pkg", ".apk", ".deb", ".rpm")),
    ("代码", (".py", ".js", ".ts", ".jsx", ".tsx", ".java", ".c", ".cpp", ".h",
             ".hpp", ".cs", ".go", ".rs", ".php", ".rb", ".html", ".htm", ".css",
             ".scss", ".json", ".xml", ".yml", ".yaml", ".sql", ".sh", ".bat",
             ".ps1", ".ipynb", ".vue")),
    ("字体", (".ttf", ".otf", ".woff", ".woff2", ".fon")),
]

# 快捷方式按「用途」细分，而不是全部塞进一个「软件图标」文件夹
SOFTWARE_CATEGORIES = [
    ("开发与AI", ()),
    ("办公与学习", ()),
    ("影音与游戏", ()),
    ("社交与通讯", ()),
    ("网盘与存储", ()),
    ("系统与工具", ()),
    ("其他软件", ()),
]

SOFTWARE_RULES = [
    ("系统与工具", ("此电脑", "回收站", "控制面板", "设置", "任务管理器", "计算器",
                  "画图", "截图", "火绒", "360", "geek", "everything", "7-zip",
                  "winrar", "bandizip", "驱动", "清理", "壁纸", "wallpaper",
                  "输入法", "搜狗", "文件夹", "todesk", "远程", "watt")),
    ("开发与AI", ("deepseek", "trae", "豆包", "千问", "通义", "chatgpt", "claude",
                "copilot", "cursor", "code", "pycharm", "visual studio", "vs code",
                "vscode", "idea", "intellij", "git", "docker", "anaconda", "python",
                "node", "hbuilder", "dev-c++", "keil", "arduino", "matlab", "xshell",
                "mobaxterm", "postman", "navicat", "mysql")),
    ("办公与学习", ("wps", "office", "word", "excel", "powerpoint", "edge", "chrome",
                 "firefox", "acrobat", "adobe", "xmind", "notion", "有道", "词典",
                 "百度翻译", "腾讯会议", "zoom", "钉钉", "企业微信", "学习通",
                 "雨课堂", "中国大学", "知网", "caj", "考试", "oms", "题库",
                 "workbuddy")),
    ("影音与游戏", ("音乐", "music", "网易云", "酷狗", "波点", "视频", "爱奇艺",
                 "腾讯视频", "哔哩", "bilibili", "potplayer", "vlc", "steam",
                 "游戏", "epic", "wegame", "抖音", "快手", "虎牙", "斗鱼",
                 "网易云游戏", "双影奇境", "无畏契约", "valorant", "uu加速器",
                 "加速器")),
    ("社交与通讯", ("微信", "wechat", "qq", "飞书", "lark", "telegram", "discord",
                 "微博", "小红书", "知乎")),
    ("网盘与存储", ("云盘", "网盘", "drive", "onedrive", "阿里云盘", "百度网盘",
                 "夸克", "坚果云", "移动云盘", "天翼云盘")),
]

CAT_ORDER = ([name for name, _ in CATEGORIES]
             + [name for name, _ in SOFTWARE_CATEGORIES]
             + ["其他文件"])
EXT2CAT = {ext: name for name, exts in CATEGORIES for ext in exts}
SOFT_EXTS = (".lnk", ".url", ".website")
LEGACY_SOFTWARE_DIR = "软件图标"


def classify_software(name, path=None):
    lower = name.lower()
    if path and lower.endswith(".url"):
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as fh:
                content = fh.read().lower()
            if "steam://" in content:
                # 壁纸引擎这类美化工具归系统，其余 Steam 项目归游戏
                if "壁纸" in lower or "wallpaper" in lower:
                    return "系统与工具"
                return "影音与游戏"
        except OSError:
            pass
    for cat, keys in SOFTWARE_RULES:
        if any(key in lower for key in keys):
            return cat
    return "其他软件"

# 扩展名判断不出来时，再用文件名关键词兜底
KEYWORD_RULES = [
    ("截图", "图片"), ("screenshot", "图片"), ("screen shot", "图片"),
    ("录屏", "视频"), ("screen record", "视频"),
]


def classify(name, path=None):
    ext = os.path.splitext(name)[1].lower()
    if ext in SOFT_EXTS:
        return classify_software(name, path)
    cat = EXT2CAT.get(ext)
    if cat:
        return cat
    lower = name.lower()
    for keyword, c in KEYWORD_RULES:
        if keyword in lower:
            return c
    return "其他文件"


def desktop_path():
    try:
        buf = ctypes.create_unicode_buffer(260)
        ctypes.windll.shell32.SHGetFolderPathW(None, 0x10, None, 0, buf)
        if buf.value:
            return buf.value
    except Exception:
        pass
    return os.path.join(os.path.expanduser("~"), "Desktop")


def public_desktop_path():
    base = os.environ.get("PUBLIC") or r"C:\Users\Public"
    return os.path.join(base, "Desktop")


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _desktop_listview():
    """找到桌面图标所在的 SysListView32 窗口句柄。"""
    user32 = ctypes.windll.user32

    def find_view(parent):
        return user32.FindWindowExW(parent, 0, "SHELLDLL_DefView", None)

    prog = user32.FindWindowW("Progman", None)
    view = find_view(prog) if prog else None
    if not view:
        found = []
        proto = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

        def callback(hwnd, _lparam):
            child = find_view(hwnd)
            if child:
                found.append(child)
                return False
            return True

        user32.EnumWindows(proto(callback), 0)
        view = found[0] if found else None
    if not view:
        return None
    return user32.FindWindowExW(view, 0, "SysListView32", None)


def arrange_desktop_icons():
    """让桌面图标按网格自动紧密排列，并补齐整理后留下的空位。"""
    if not _desktop_listview():
        return False
    try:
        user32 = ctypes.windll.user32
        lv = _desktop_listview()
        style = user32.GetWindowLongW(lv, -16)
        user32.SetWindowLongW(lv, -16, style | 0x0100)  # LVS_AUTOARRANGE
        user32.SendMessageW(lv, 0x101A, 0, 0)            # LVM_ARRANGE
        user32.InvalidateRect(lv, None, True)
    except Exception:
        pass
    return compact_desktop_icons()


def compact_desktop_icons():
    """按图标间距重算网格坐标并逐个落位，确保没有空位。"""
    lv = _desktop_listview()
    if not lv:
        return False
    user32 = ctypes.windll.user32
    gwl_style = -16
    lvs_autoarrange = 0x0100
    lvm_get_count = 0x1004
    lvm_get_spacing = 0x1033
    lvm_set_position = 0x100F
    wm_setredraw = 0x000B
    try:
        style = user32.GetWindowLongW(lv, gwl_style)
        count = int(user32.SendMessageW(lv, lvm_get_count, 0, 0))
        if count <= 0:
            return False

        spacing = int(user32.SendMessageW(lv, lvm_get_spacing, 0, 0))
        cell_w, cell_h = spacing & 0xFFFF, (spacing >> 16) & 0xFFFF
        if cell_w <= 0 or cell_h <= 0:
            cell_w, cell_h = 115, 154

        rect = wintypes.RECT()
        user32.GetClientRect(lv, ctypes.byref(rect))
        rows = max(1, rect.bottom // cell_h)

        # 临时关闭自动排列，才能按自定义坐标放置
        user32.SetWindowLongW(lv, gwl_style, style & ~lvs_autoarrange)
        user32.SendMessageW(lv, wm_setredraw, 0, 0)
        for index in range(count):
            col, row = divmod(index, rows)
            x, y = col * cell_w, row * cell_h
            user32.SendMessageW(lv, lvm_set_position, index, ((y & 0xFFFF) << 16) | (x & 0xFFFF))
        user32.SendMessageW(lv, wm_setredraw, 1, 0)
        user32.InvalidateRect(lv, None, True)
        user32.SetWindowLongW(lv, gwl_style, style)  # 恢复自动排列
        return True
    except Exception:
        try:
            user32.SetWindowLongW(lv, gwl_style, style)
        except Exception:
            pass
        return False


def human_size(size):
    units = ["B", "KB", "MB", "GB", "TB"]
    value = float(size)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return ("%d %s" % (value, unit)) if unit == "B" else ("%.1f %s" % (value, unit))
        value /= 1024.0


def scan(root):
    """扫描目录顶层的文件，自动分类。桌面还会同时扫描公共桌面。"""
    rows = []
    sources = [(root, "本机")]
    if os.path.normcase(os.path.abspath(root)) == os.path.normcase(os.path.abspath(desktop_path())):
        pub = public_desktop_path()
        if os.path.isdir(pub):
            sources.append((pub, "公共"))

    for base, origin in sources:
        try:
            names = os.listdir(base)
        except OSError:
            continue
        for name in names:
            path = os.path.join(base, name)
            try:
                st = os.stat(path)
            except OSError:
                continue
            if not os.path.isfile(path):
                continue
            if name.lower() in ("desktop.ini", "桌面整理.lnk"):
                continue
            attrs = getattr(st, "st_file_attributes", 0)
            if attrs & 0x2 or attrs & 0x4:  # 隐藏 / 系统文件
                continue
            rows.append({
                "name": name,
                "path": path,
                "ext": os.path.splitext(name)[1].lower(),
                "size": st.st_size,
                "cat": classify(name, path),
                "origin": origin,
                "checked": True,
            })

    # 把以前「全部塞进一个文件夹」的旧软件目录也纳入，按用途重新拆分
    if os.path.normcase(os.path.abspath(root)) == os.path.normcase(os.path.abspath(desktop_path())):
        legacy = os.path.join(root, LEGACY_SOFTWARE_DIR)
        if os.path.isdir(legacy):
            try:
                names = os.listdir(legacy)
            except OSError:
                names = []
            for name in names:
                path = os.path.join(legacy, name)
                try:
                    st = os.stat(path)
                except OSError:
                    continue
                if not os.path.isfile(path) or name.lower() == "desktop.ini":
                    continue
                rows.append({
                    "name": name,
                    "path": path,
                    "ext": os.path.splitext(name)[1].lower(),
                    "size": st.st_size,
                    "cat": classify(name, path),
                    "origin": LEGACY_SOFTWARE_DIR,
                    "nested": legacy,
                    "checked": True,
                })

    rows.sort(key=lambda r: (CAT_ORDER.index(r["cat"]), r["name"].lower()))
    return rows


def unique_path(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    index = 1
    while True:
        candidate = "%s (%d)%s" % (base, index, ext)
        if not os.path.exists(candidate):
            return candidate
        index += 1


def move_rows(rows, root):
    batch = {
        "time": time.strftime("%Y-%m-%d %H:%M:%S"),
        "root": root,
        "moves": [],
        "created": [],
        "removed_dirs": [],
    }
    errors = []
    for row in rows:
        folder = os.path.join(root, row["cat"])
        if not os.path.isdir(folder):
            try:
                os.makedirs(folder)
                batch["created"].append(folder)
            except OSError as exc:
                errors.append("%s：无法创建文件夹（%s）" % (row["name"], exc))
                continue
        target = unique_path(os.path.join(folder, row["name"]))
        try:
            shutil.move(row["path"], target)
            batch["moves"].append({"src": row["path"], "dst": target})
        except OSError as exc:
            errors.append("%s：%s" % (row["name"], exc))

    # 拆空旧「软件图标」文件夹（撤销时会重建）
    nested_dirs = {row.get("nested") for row in rows if row.get("nested")}
    for folder in nested_dirs:
        try:
            if os.path.isdir(folder) and not os.listdir(folder):
                os.rmdir(folder)
                batch["removed_dirs"].append(folder)
        except OSError:
            pass
    return batch, errors


def undo_batch(batch):
    restored = 0
    skipped = 0
    errors = []
    for folder in batch.get("removed_dirs", []):
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError:
            pass
    for move in reversed(batch["moves"]):
        src, dst = move["src"], move["dst"]
        if not os.path.exists(dst):
            continue
        if os.path.exists(src):
            skipped += 1
            continue
        try:
            os.makedirs(os.path.dirname(src), exist_ok=True)
            shutil.move(dst, src)
            restored += 1
        except OSError as exc:
            errors.append("%s：%s" % (os.path.basename(dst), exc))
    for folder in reversed(batch.get("created", [])):
        try:
            if os.path.isdir(folder) and not os.listdir(folder):
                os.rmdir(folder)
        except OSError:
            pass
    return restored, skipped, errors


def load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return default


def save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
    except OSError:
        pass


def load_batches():
    return load_json(LOG_PATH, [])


def save_batch(batch):
    batches = load_batches()
    batches.append(batch)
    save_json(LOG_PATH, batches[-MAX_BATCHES:])


def short_path(path, limit=56):
    if len(path) <= limit:
        return path
    return path[:limit // 2 - 2] + "…" + path[-(limit // 2 - 1):]


def round_rect(canvas, x1, y1, x2, y2, radius=16, **kwargs):
    points = [
        x1 + radius, y1, x2 - radius, y1, x2, y1,
        x2, y1 + radius, x2, y2 - radius, x2, y2,
        x2 - radius, y2, x1 + radius, y2, x1, y2,
        x1, y2 - radius, x1, y1 + radius, x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kwargs)


class PillButton(tk.Canvas):
    """圆角胶囊按钮。"""

    def __init__(self, parent, text, command, kind="ghost", bg=BG,
                 height=32, padx=16, font_size=10):
        self.font = tkfont.Font(family=FONT_FAMILY, size=font_size,
                                weight="bold" if kind == "primary" else "normal")
        self.padx = padx
        self.text_value = text
        width = self.font.measure(text) + padx * 2
        super().__init__(parent, width=width, height=height, bg=bg,
                         highlightthickness=0, bd=0, cursor="hand2")
        self.command = command
        self.kind = kind
        self._enabled = True
        self._bw, self._bh = width, height

        if kind == "primary":
            self.fill, self.hover, self.outline, self.fg = PINK, PINK_DARK, PINK, "#FFFFFF"
        elif kind == "soft":
            self.fill, self.hover, self.outline, self.fg = "#FFE9F2", "#FFD8E8", "#FFE9F2", PINK_DARK
        else:
            self.fill, self.hover, self.outline, self.fg = "#FFFFFF", "#FFF3F8", LINE, "#7B6678"

        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)
        self._draw()

    def _capsule(self):
        w, h = self._bw, self._bh
        x1, y1, x2, y2 = 1, 1, w - 1, h - 1
        r = (y2 - y1) / 2.0
        cy = (y1 + y2) / 2.0
        right_cx, left_cx = x2 - r, x1 + r
        points = []
        for i in range(15):
            angle = -math.pi / 2 + math.pi * i / 14.0
            points += [right_cx + r * math.cos(angle), cy + r * math.sin(angle)]
        for i in range(15):
            angle = math.pi / 2 + math.pi * i / 14.0
            points += [left_cx + r * math.cos(angle), cy + r * math.sin(angle)]
        return self.create_polygon(points, fill=self.fill, outline=self.outline, width=1)

    def _draw(self):
        self.delete("all")
        self._shape = self._capsule()
        self._text = self.create_text(self._bw / 2, self._bh / 2, text=self.text_value,
                                      font=self.font, fill=self.fg)
        if not self._enabled:
            self._paint_disabled()

    def _paint_disabled(self):
        self.itemconfigure(self._shape, fill="#F7F0F4", outline="#F1E4EC")
        self.itemconfigure(self._text, fill="#C6B9C3")

    def _on_enter(self, _event=None):
        if self._enabled:
            self.itemconfigure(self._shape, fill=self.hover)

    def _on_leave(self, _event=None):
        if self._enabled:
            self.itemconfigure(self._shape, fill=self.fill)

    def _on_click(self, _event=None):
        if self._enabled and self.command:
            self.command()

    def set_text(self, text):
        if text == self.text_value:
            return
        self.text_value = text
        width = self.font.measure(text) + self.padx * 2
        if width != self._bw:
            self._bw = width
            self.configure(width=width)
        self._draw()

    def set_enabled(self, enabled):
        if enabled == self._enabled and not enabled:
            return
        self._enabled = enabled
        self.configure(cursor="hand2" if enabled else "arrow")
        if enabled:
            self.itemconfigure(self._shape, fill=self.fill, outline=self.outline)
            self.itemconfigure(self._text, fill=self.fg)
        else:
            self._paint_disabled()


class OrganizerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("桌面整理")
        self.configure(bg=BG)
        self.minsize(800, 480)
        self._center(1100, 760)

        self.admin = is_admin()
        settings = load_json(SETTINGS_PATH, {})
        root = settings.get("folder") or desktop_path()
        if not os.path.isdir(root):
            root = desktop_path()
        self.root_dir = root
        self.rows = []
        self._banner_photo = None
        self._banner_width = 0
        self._petals = []
        self._after_id = None
        self._load_banner()

        self._init_style()
        self._build_ui()
        self.refresh()
        self._animate_petals()

    # ---------- 外观 ----------

    def _center(self, width, height):
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(width, screen_w - 40)
        height = min(height, screen_h - 70)
        x = (screen_w - width) // 2
        y = max((screen_h - height) // 2 - 10, 0)
        self.geometry("%dx%d+%d+%d" % (width, height, x, y))

    def _load_banner(self):
        try:
            from PIL import Image
            self._banner_src = Image.open(BANNER_PATH)
        except Exception:
            self._banner_src = None

    def _init_style(self):
        global FONT_FAMILY
        available = tkfont.families(self)
        if FONT_FAMILY not in available:
            FONT_FAMILY = "Microsoft YaHei" if "Microsoft YaHei" in available else "TkDefaultFont"

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
                        rowheight=29, borderwidth=0, font=(FONT_FAMILY, 10))
        style.configure("Treeview.Heading", font=(FONT_FAMILY, 9), background="#FFF1F7",
                        foreground="#B08A9C", relief="flat", padding=(8, 7))
        style.map("Treeview.Heading", background=[("active", "#FFE9F2")])
        style.map("Treeview", background=[("selected", "#FFE3EF")],
                  foreground=[("selected", TEXT)])
        style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        style.configure("Cute.Vertical.TScrollbar", background="#FFD3E4", troughcolor=CARD,
                        bordercolor=CARD, arrowcolor=PINK, lightcolor="#FFD3E4",
                        darkcolor="#FFD3E4", relief="flat", width=9)

    def _build_ui(self):
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True, padx=14, pady=12)

        # 顶部插画卡
        header = tk.Frame(outer, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        header.pack(fill="x")

        self.banner = tk.Canvas(header, height=BANNER_H, bg="#FFEBF4",
                                highlightthickness=0, bd=0)
        self.banner.pack(fill="x")
        self.banner.bind("<Configure>", self._draw_banner)

        info = tk.Frame(header, bg=CARD)
        info.pack(fill="x", padx=16, pady=(9, 11))
        tk.Label(info, text="整理位置", bg=CARD, fg=MUTED,
                 font=(FONT_FAMILY, 9)).pack(side="left")
        self.path_label = tk.Label(info, text=short_path(self.root_dir), bg=CARD, fg=TEXT,
                                   font=(FONT_FAMILY, 10, "bold"))
        self.path_label.pack(side="left", padx=(10, 12))
        PillButton(info, "更换", self.choose_folder, kind="ghost", bg=CARD,
                   height=28, padx=14).pack(side="right")

        # 工具条
        bar = tk.Frame(outer, bg=BG)
        bar.pack(fill="x", pady=(11, 9))
        for index, (text, command) in enumerate([
                ("全选", lambda: self.set_all(True)),
                ("全不选", lambda: self.set_all(False)),
                ("重新扫描", self.refresh),
                ("排列图标", self.arrange_icons),
                ("打开文件夹", self.open_folder)]):
            PillButton(bar, text, command, bg=BG, height=30,
                       padx=14, font_size=9).pack(side="left", padx=(0 if index == 0 else 7, 0))

        self.admin_btn = PillButton(bar, "管理员模式", self.elevate, kind="soft",
                                    bg=BG, height=30, padx=14, font_size=9)
        self.admin_tip = tk.Label(bar, text="灰色图标在公共桌面，需管理员", bg=BG, fg=MUTED,
                                  font=(FONT_FAMILY, 8))

        # 底部操作区（先占位，保证列表自适应）
        footer = tk.Frame(outer, bg=BG)
        footer.pack(side="bottom", fill="x", pady=(9, 0))
        self.status = tk.Label(footer, text="", bg=BG, fg=MUTED, font=(FONT_FAMILY, 9))
        self.status.pack(side="left")
        self.start_btn = PillButton(footer, "一键整理", self.start, kind="primary",
                                    bg=BG, height=38, padx=26, font_size=10)
        self.start_btn.pack(side="right")
        self.undo_btn = PillButton(footer, "撤销上次整理", self.undo, kind="ghost",
                                   bg=BG, height=34, padx=16, font_size=9)
        self.undo_btn.pack(side="right", padx=(0, 10))

        # 文件列表
        tree_card = tk.Frame(outer, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        tree_card.pack(fill="both", expand=True)

        columns = ("check", "name", "cat", "size", "source")
        self.tree = ttk.Treeview(tree_card, columns=columns, show="headings",
                                 selectmode="browse")
        headings = {"check": "", "name": "文件名", "cat": "将归入",
                    "size": "大小", "source": "来源"}
        widths = {"check": 46, "name": 440, "cat": 160, "size": 90, "source": 110}
        anchors = {"check": "center", "name": "w", "cat": "w", "size": "e", "source": "w"}
        for key in columns:
            self.tree.heading(key, text=headings[key], anchor=anchors[key])
            self.tree.column(key, width=widths[key], anchor=anchors[key],
                             stretch=(key == "name"))
        self.tree.tag_configure("odd", background=CARD)
        self.tree.tag_configure("even", background="#FFF9FC")
        self.tree.tag_configure("locked", foreground="#C9BCC9")

        scroll = ttk.Scrollbar(tree_card, orient="vertical", command=self.tree.yview,
                               style="Cute.Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=6)
        scroll.pack(side="right", fill="y", pady=6, padx=(2, 4))

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<space>", self.toggle_selected_row)
        self.bind("<Control-a>", lambda _e: self.set_all(True))

    def _draw_banner(self, event=None):
        width = self.banner.winfo_width()
        if width < 20:
            return
        height = BANNER_H
        self.banner.delete("bgline")
        gradient = ["#FFE2F0", "#FFE9F4", "#F3E9FF", "#E3F1FF"]
        for line in range(height):
            ratio = line / float(height - 1)
            index = ratio * (len(gradient) - 1)
            low = int(index)
            high = min(low + 1, len(gradient) - 1)
            frac = index - low
            c1 = tuple(int(gradient[low][i:i + 2], 16) for i in (1, 3, 5))
            c2 = tuple(int(gradient[high][i:i + 2], 16) for i in (1, 3, 5))
            color = "#%02x%02x%02x" % tuple(int(c1[k] + (c2[k] - c1[k]) * frac) for k in range(3))
            self.banner.create_line(0, line, width, line, fill=color, tags="bgline")

        if self._banner_src and width != self._banner_width:
            from PIL import Image, ImageTk
            image = self._banner_src
            scale = max(width / float(image.width), height / float(image.height))
            new_size = (max(int(image.width * scale + 1), width),
                        max(int(image.height * scale + 1), height))
            resized = image.resize(new_size, Image.LANCZOS)
            left = (new_size[0] - width) // 2
            top = max(new_size[1] - height - 12, 0)
            cropped = resized.crop((left, top, left + width, top + height))
            self._banner_photo = ImageTk.PhotoImage(cropped)
            self.banner.delete("bannerimg")
            self.banner.create_image(0, 0, anchor="nw", image=self._banner_photo,
                                     tags="bannerimg")
            self._banner_width = width

        if not self.banner.find_withtag("bannercard"):
            round_rect(self.banner, 22, 34, 300, 118, 20,
                       fill="#FFFFFF", outline="#FFDCEB", width=1, tags="bannercard")
            self.banner.create_text(42, 65, text="桌面整理", anchor="w",
                                    font=(FONT_FAMILY, 20, "bold"), fill=PINK_DARK,
                                    tags="bannertext")
            self.banner.create_text(43, 98, text="自动识别类型 · 一键归类", anchor="w",
                                    font=(FONT_FAMILY, 9), fill="#A88C9C",
                                    tags="bannertext")
            self.banner.create_text(272, 60, text="♥", anchor="e",
                                    font=(FONT_FAMILY, 13), fill="#FFB3CE",
                                    tags="bannertext")

        self.banner.tag_lower("bannerimg")
        self.banner.tag_lower("bgline")
        self.banner.tag_raise("bannercard")
        self.banner.tag_raise("bannertext")
        self._spawn_petals()
        for petal in self._petals:
            self.banner.tag_lower(petal[0], "bannercard")

    def _spawn_petals(self):
        if self._petals:
            return
        colors = ["#FFC4D9", "#FFDDEA", "#FFB7D2", "#FFE7F1"]
        for index in range(8):
            x = 320 + (index * 89) % 600
            y = 12 + (index * 41) % 100
            size = 4 + index % 3
            shape = self.banner.create_oval(x, y, x + size * 1.7, y + size,
                                            fill=colors[index % len(colors)],
                                            outline="", tags="petal")
            self._petals.append([shape, x, y, 0.55 + (index % 4) * 0.18,
                                 0.35 + (index % 3) * 0.16, index])

    def _animate_petals(self):
        try:
            width = self.banner.winfo_width()
            height = BANNER_H
        except Exception:
            return
        for petal in self._petals:
            shape, x, y, speed, sway, index = petal
            y += speed
            x -= speed * 0.7 + math.sin((y + index * 24) / 16.0) * 0.5 * sway * 10
            if y > height + 8 or x < -12:
                y = -8
                x = 240 + (index * 137 + int(abs(math.sin(y)) * 61)) % max(width - 280, 260)
            self.banner.coords(shape, x, y, x + 7, y + 4)
            petal[1], petal[2] = x, y
        self._after_id = self.after(60, self._animate_petals)

    def destroy(self):
        if self._after_id:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        super().destroy()

    # ---------- 数据 ----------

    def refresh(self):
        self.rows = scan(self.root_dir)
        self.admin = is_admin()
        has_public = any(r["origin"] == "公共" for r in self.rows)
        if self.admin or not has_public:
            self.admin_btn.pack_forget()
            self.admin_tip.pack_forget()
        else:
            self.admin_btn.pack(side="right")
            self.admin_tip.pack(side="right", padx=(0, 8))
        self.render()

    def render(self):
        self.tree.delete(*self.tree.get_children())
        for index, row in enumerate(self.rows):
            locked = row["origin"] == "公共" and not self.admin
            row["locked"] = locked
            if locked:
                row["checked"] = False
            tags = ["locked"] if locked else ["even" if index % 2 else "odd"]
            self.tree.insert(
                "", "end", iid=str(index),
                values=(
                    "♥" if row["checked"] else "♡",
                    row["name"],
                    row["cat"],
                    human_size(row["size"]),
                    {"公共": "公共桌面", LEGACY_SOFTWARE_DIR: "旧软件文件夹"}.get(
                        row["origin"], "本机"),
                ),
                tags=tags,
            )
        self.update_status()

    def update_status(self):
        total = len(self.rows)
        selected = sum(1 for r in self.rows if r["checked"])
        locked = sum(1 for r in self.rows if r.get("locked"))
        text = "共 %d 个文件 · 已选 %d 个" % (total, selected)
        if locked:
            text += " · %d 个需管理员权限" % locked
        self.status.configure(text=text)

        batches = load_batches()
        if batches:
            last = batches[-1]
            self.undo_btn.set_enabled(True)
            self.undo_btn.set_text("撤销上次整理 · %s · %d 个" % (
                last["time"][5:16], len(last["moves"])))
        else:
            self.undo_btn.set_enabled(False)
            self.undo_btn.set_text("撤销上次整理")

    # ---------- 交互 ----------

    def on_tree_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#1":
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        self.toggle_row(iid)

    def toggle_selected_row(self, _event=None):
        for iid in self.tree.selection():
            self.toggle_row(iid)

    def toggle_row(self, iid):
        row = self.rows[int(iid)]
        if row.get("locked"):
            return
        row["checked"] = not row["checked"]
        self.tree.set(iid, "check", "♥" if row["checked"] else "♡")
        self.update_status()

    def set_all(self, value):
        for index, row in enumerate(self.rows):
            if row.get("locked"):
                continue
            row["checked"] = value
            self.tree.set(str(index), "check", "♥" if value else "♡")
        self.update_status()

    def choose_folder(self):
        folder = filedialog.askdirectory(title="选择要整理的文件夹", initialdir=self.root_dir)
        if not folder:
            return
        self.root_dir = os.path.normpath(folder)
        self.path_label.configure(text=short_path(self.root_dir))
        save_json(SETTINGS_PATH, {"folder": self.root_dir})
        self.refresh()

    def open_folder(self):
        try:
            os.startfile(self.root_dir)
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def arrange_icons(self):
        if arrange_desktop_icons():
            self.update_idletasks()
            messagebox.showinfo("排列完成", "已开启自动排列，桌面图标已按网格紧密排好。")
        else:
            messagebox.showwarning(
                "排列失败",
                "没有找到桌面图标层。\n请右键桌面 → 查看 → 勾选「自动排列图标」。")

    def elevate(self):
        if self.admin:
            return
        script = os.path.abspath(sys.argv[0])
        try:
            result = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, '"%s"' % script, APP_DIR, 1)
        except Exception as exc:
            messagebox.showerror("提权失败", "无法以管理员身份启动：%s" % exc)
            return
        if result <= 32:
            messagebox.showerror("提权失败", "系统拒绝了管理员权限请求。")
            return
        self.destroy()

    # ---------- 整理 / 撤销 ----------

    def start(self):
        selected = [row for row in self.rows if row["checked"] and not row.get("locked")]
        if not selected:
            messagebox.showinfo("没有可整理的文件", "当前没有勾选任何文件。")
            return
        counts = Counter(row["cat"] for row in selected)
        detail = "、".join("%s %d 个" % (cat, counts[cat])
                          for cat in CAT_ORDER if counts.get(cat))
        message = "将把 %d 个文件移入对应的类型文件夹：\n\n%s\n\n整理后可以在本窗口点「撤销上次整理」一键还原。\n是否继续？" % (
            len(selected), detail)
        if not messagebox.askyesno("确认整理", message):
            return

        self.config(cursor="watch")
        self.update()
        batch, errors = move_rows(selected, self.root_dir)
        self.config(cursor="")
        if batch["moves"]:
            save_batch(batch)
        self.refresh()

        arranged = arrange_desktop_icons() if batch["moves"] else False

        if errors:
            head = "成功整理 %d 个文件，%d 个失败：" % (len(batch["moves"]), len(errors))
            messagebox.showwarning("整理完成（部分失败）",
                                   head + "\n\n" + "\n".join(errors[:6]))
        else:
            messagebox.showinfo("整理完成",
                                "已整理 %d 个文件到 %d 个文件夹。%s" % (
                                    len(batch["moves"]), len(set(m["dst"].rsplit(os.sep, 1)[0]
                                                                 for m in batch["moves"])),
                                    "桌面图标已自动排列。" if arranged else ""))

    def undo(self):
        batches = load_batches()
        if not batches:
            return
        last = batches[-1]
        if not messagebox.askyesno(
                "撤销上次整理",
                "将把 %s 整理的 %d 个文件移回原位置，是否继续？" % (
                    last["time"], len(last["moves"]))):
            return

        restored, skipped, errors = undo_batch(last)
        batches.pop()
        save_json(LOG_PATH, batches)
        self.refresh()

        message = "已还原 %d 个文件。" % restored
        if skipped:
            message += "\n%d 个文件因原位置已有同名文件而未还原。" % skipped
        if errors:
            message += "\n%d 个文件还原失败（可能是公共桌面需要管理员权限）。" % len(errors)
        messagebox.showinfo("撤销完成", message)


if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    OrganizerApp().mainloop()
