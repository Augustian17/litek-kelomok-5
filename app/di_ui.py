"""Halaman data DI dan pembuat Excel untuk Dockit."""
import datetime
import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

import customtkinter as ctk
from PIL import Image, ImageDraw, ImageTk

import di_data as data
import dockit_icons as icons
import dockit_motion as motion

PAGE_DASHBOARD = "Dasbor"
PAGE_DASHBOARD_PO = "Dasbor PO"
PAGE_CALENDAR = "Kalender ETD"
PAGE_CALENDAR_PO = "Kalender ETD PO"
PAGE_MASTER = "Master DI"
PAGE_MASTER_PO = "Master PO"
PAGE_BUILD_PO = "Buat Excel PO"
PAGE_GOODS = "Goods Description"
PAGE_WEIGHT = "CBM & Berat"
PAGE_PRICE = "Harga"
PAGE_BUILD = "Buat Excel"
PAGE_PDF = "PDF"
PAGES = [PAGE_DASHBOARD, PAGE_CALENDAR, PAGE_MASTER, PAGE_GOODS, PAGE_WEIGHT, PAGE_PRICE, PAGE_BUILD, PAGE_PDF]
PO_LABELS = [PAGE_DASHBOARD, PAGE_CALENDAR, PAGE_MASTER_PO, PAGE_GOODS, PAGE_WEIGHT, PAGE_PRICE, PAGE_BUILD, PAGE_PDF]


def page_key(tab, label):
    """Tab PO memakai halaman Buat Excel sendiri walaupun labelnya sama."""
    if tab == "PO":
        return {PAGE_BUILD: PAGE_BUILD_PO, PAGE_DASHBOARD: PAGE_DASHBOARD_PO, PAGE_CALENDAR: PAGE_CALENDAR_PO}.get(label, label)
    return label
FILTER_ALL = "Semua"
FILTER_FLAG = "Perlu diperbaiki"
FILTER_DIRTY = "Belum disimpan"
FILTER_TALLY = "CBM tidak tally"
FILTER_SPECIAL = "Khusus"
EMPTY_VALUE = "(Kosong)"
POD_COLORS = {
    "LOS ANGELES": ("#BBD7F7", "#0B3A6E"), "LONG BEACH": ("#B9EBF2", "#0A4F5C"), "NEW YORK": ("#FBE3A1", "#6B4A00"),
    "NEWARK": ("#FBCFA8", "#7A3300"), "NORFOLK": ("#DCCCF7", "#42218A"), "SAVANNAH": ("#C4EDC2", "#1D5A1B"),
    "TACOMA": ("#B5EADB", "#07574A"), "HOUSTON": ("#F9C8DC", "#7D1848"),
}
EXTRA_POD_COLORS = [("#E5E0A8", "#555000"), ("#D3D8E8", "#2B3A67"), ("#F3CFC6", "#6E2A1A"), ("#CFE3B4", "#3B5A12")]
BL_COLORS = [
    ("#FFD6E0", "#7A1F3D"), ("#D6E4FF", "#1D3B8A"), ("#FFE8C2", "#7A4B00"), ("#D9F2D0", "#245A17"),
    ("#E8DAFF", "#4A2391"), ("#C9F0F0", "#0D5257"),
]


def pod_color(name):
    name = " ".join(str(name).upper().split())
    if name in POD_COLORS:
        return POD_COLORS[name]
    return EXTRA_POD_COLORS[sum(map(ord, name)) % len(EXTRA_POD_COLORS)]


def shorten_name(path):
    name = os.path.basename(path)
    return name if len(name) <= 34 else name[:18] + "…" + name[-14:]


def row_colors():
    dark = ctk.get_appearance_mode() == "Dark"
    return ("#4A2320", "#F97066", "#4A3B1A") if dark else ("#FDE7E4", "#B42318", "#FFF4D6")


def is_dark():
    return ctk.get_appearance_mode() == "Dark"


def pick_color(color):
    """Warna tema customtkinter bisa berupa pasangan (terang, gelap)."""
    if isinstance(color, (tuple, list)):
        return color[1 if is_dark() else 0]
    return color


def rounded_points(x1, y1, x2, y2, radius):
    return (
        x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius, x2, y2 - radius, x2, y2,
        x2 - radius, y2, x1 + radius, y2, x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
    )


class Grid(ctk.CTkFrame):
    """Tabel gaya baris datar: kolom bisa disembunyikan, dikunci di kiri, dipindah, dan diubah lebarnya.
    Hanya baris yang terlihat yang digambar, jadi data besar tetap ringan. Semua ini hanya tampilan,
    urutan baris dan isi data tidak berubah."""

    DENSITY = {"Lega": (54, 38), "Normal": (44, 32), "Rapat": (34, 26)}
    HEAD = 44
    INDEX = 52
    GAP = 0
    PAD = 12
    ICON = 22
    ICON_GAP = 6
    SAMPLE = 300
    MAX_FIT = 480
    FLAG_KEY = "__flag"

    def __init__(self, parent, columns, on_edit, flag_column=True, on_filter=None, name=None, rounded=False, toolbar=True, movable=True, pins=None):
        super().__init__(parent, fg_color="transparent")
        self.rounded = rounded
        self.movable = movable
        self.default_pins = list(pins) if pins is not None else [0]
        if rounded:
            # Gaya lama: sel bulat dan berjarak dengan tinggi baris tetap.
            self.DENSITY = {"Normal": (46, 36)}
        self.columns = columns
        self.on_edit = on_edit
        self.on_filter = on_filter
        self.on_select = None
        self.flag_column = flag_column
        self.name = name
        self.headers = [column.label for column in columns] + (["⚠ Perlu diperbaiki"] if flag_column else [])
        self.keys = [column.key for column in columns] + ([self.FLAG_KEY] if flag_column else [])
        self.kinds = [column.kind for column in columns] + (["text"] if flag_column else [])
        self.scale = self._get_widget_scaling()
        family = "Segoe UI" if os.name == "nt" else "DejaVu Sans"
        self.font = tkfont.Font(family=family, size=-round(14 * self.scale))
        self.bold = tkfont.Font(family=family, size=-round(14 * self.scale), weight="bold")
        self.head_font = tkfont.Font(family=family, size=-round(13 * self.scale), weight="bold")
        self.small = tkfont.Font(family=family, size=-round(12 * self.scale))
        self.fit_cache = {}
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.build_toolbar(family)
        if not toolbar:
            self.toolbar.grid_remove()
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, takefocus=1, bg=self.color("card"))
        self.canvas.grid(row=1, column=0, sticky="nsew")
        self.vbar = ctk.CTkScrollbar(self, command=self.scroll_y)
        self.vbar.grid(row=1, column=1, sticky="ns")
        self.hbar = ctk.CTkScrollbar(self, orientation="horizontal", command=self.scroll_x)
        self.hbar.grid(row=2, column=0, sticky="ew")
        self.widths = [self.px(self.title_width(label, i)) for i, label in enumerate(self.headers)]
        self.saved_widths = {}
        self.resized = set()
        self.order = list(range(len(self.headers)))
        self.hidden = set()
        self.pinned = [i for i in self.default_pins if i < len(self.headers)]
        self.density = "Normal"
        self.load_view()
        self.rows, self.tags, self.styles = [], [], {}
        self.marked = set()
        self.x0 = self.y0 = 0
        self.anchor = self.active = None
        self.picked_rows = set()
        self.hover = None
        self.hover_col = None
        self.editor = None
        self.edit_cell = None
        self.drag_width = None
        self.drag_mode = None
        self.head_drag = None
        self.drop_at = None
        self.draw_job = None
        self.menu = None
        c = self.canvas
        c.bind("<Configure>", lambda _: self.schedule())
        c.bind("<Motion>", self.on_motion)
        c.bind("<Leave>", lambda _: (self.set_hover(None), self.set_hover_col(None)))
        c.bind("<ButtonPress-1>", self.on_press)
        c.bind("<B1-Motion>", self.on_drag)
        c.bind("<ButtonRelease-1>", self.on_release)
        c.bind("<Double-Button-1>", self.on_double)
        c.bind("<MouseWheel>", self.on_wheel)
        c.bind("<Shift-MouseWheel>", lambda e: self.move_x(-e.delta))
        c.bind("<Button-4>", lambda _: self.move_y(-3 * self.row_h()))
        c.bind("<Button-5>", lambda _: self.move_y(3 * self.row_h()))
        for key, step in (("Up", (-1, 0)), ("Down", (1, 0)), ("Left", (0, -1)), ("Right", (0, 1))):
            c.bind(f"<{key}>", lambda _, s=step: self.move_active(*s))
            c.bind(f"<Shift-{key}>", lambda _, s=step: self.move_active(*s, extend=True))
        c.bind("<Prior>", lambda _: self.move_active(-self.page_rows(), 0))
        c.bind("<Next>", lambda _: self.move_active(self.page_rows(), 0))
        c.bind("<Return>", lambda _: self.begin_edit())
        c.bind("<F2>", lambda _: self.begin_edit())
        c.bind("<Control-c>", self.copy)
        c.bind("<Control-C>", self.copy)
        c.bind("<Control-a>", self.select_all)
        c.bind("<Control-A>", self.select_all)
        c.bind("<Key>", self.on_key)
        self.update_toolbar()

    # ---------- tampilan kolom ----------

    def build_toolbar(self, family):
        bar = ctk.CTkFrame(self, fg_color="transparent", height=40)
        self.toolbar = bar
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=4, pady=(2, 6))
        muted = ("#5E6675", "#A0A8B8")
        self.info_label = ctk.CTkLabel(bar, text="", font=ctk.CTkFont(family=family, size=13), text_color=muted)
        self.info_label.pack(side="left", padx=(6, 0))
        ctk.CTkLabel(bar, text="·  Tarik judul: pindah kolom  ·  tarik garis judul: ubah lebar  ·  dobel klik garis: pas otomatis",
                     font=ctk.CTkFont(family=family, size=12), text_color=muted).pack(side="left", padx=8)
        line = ("#E0D2BF", "#323B52")
        self.reset_button = ctk.CTkButton(
            bar, text="↺  Reset", width=86, height=34, corner_radius=9, font=ctk.CTkFont(family=family, size=13),
            fg_color=("#FFFFFF", "#1A2030"), hover_color=("#F6EEE3", "#2A3D63"), border_width=1, border_color=line,
            text_color=("#16202E", "#EEF0F5"), command=self.reset_view,
        )
        self.reset_button.pack(side="right")
        self.density_button = ctk.CTkSegmentedButton(
            bar, values=list(self.DENSITY), command=self.set_density, height=34, font=ctk.CTkFont(family=family, size=13),
            fg_color=("#F3ECE2", "#232B3E"), selected_color=("#FFFFFF", "#0F2C59"), selected_hover_color=("#FFFFFF", "#0F2C59"),
            unselected_color=("#F3ECE2", "#232B3E"), unselected_hover_color=("#EEE4D6", "#2D3750"), text_color=("#16202E", "#EEF0F5"),
        )
        self.density_button.pack(side="right", padx=8)
        self.column_button = ctk.CTkButton(
            bar, text="Kolom", width=110, height=34, corner_radius=9, font=ctk.CTkFont(family=family, size=13, weight="bold"),
            fg_color=("#FFFFFF", "#1A2030"), hover_color=("#F6EEE3", "#2A3D63"), border_width=1, border_color=line,
            text_color=("#16202E", "#EEF0F5"), command=self.toggle_menu,
        )
        self.column_button.pack(side="right")

    @property
    def view(self):
        """Urutan kolom yang tampil: kolom dikunci dulu, lalu sisanya."""
        pinned = [i for i in self.pinned if i not in self.hidden]
        return pinned + [i for i in self.order if i not in self.hidden and i not in self.pinned]

    def pin_count(self):
        return sum(1 for i in self.pinned if i not in self.hidden)

    def row_h(self):
        return self.px(self.DENSITY[self.density][0])

    def cell_h(self):
        return self.px(self.DENSITY[self.density][1])

    def is_default(self):
        return (self.order == list(range(len(self.headers))) and not self.hidden and self.pinned == [i for i in self.default_pins if i < len(self.headers)]
                and self.density == "Normal" and not self.resized)

    def update_toolbar(self):
        total = len(self.headers)
        shown = total - len(self.hidden)
        self.column_button.configure(text=f"▥  Kolom  {shown}/{total}  ▾")
        self.density_button.set(self.density)
        self.reset_button.configure(text="↺  Reset" if self.is_default() else "↺  Reset  ●")
        self.info_label.configure(text=f"{shown} dari {total} kolom · {self.pin_count()} dikunci · baris {self.density.lower()}")

    def load_view(self):
        if not self.name:
            return
        saved = data.load_settings().get("table_view", {}).get(self.name)
        if not saved:
            return
        index = {key: i for i, key in enumerate(self.keys)}
        order = [index[k] for k in saved.get("order", []) if k in index]
        self.order = order + [i for i in range(len(self.keys)) if i not in order]
        self.hidden = {index[k] for k in saved.get("hidden", []) if k in index}
        if len(self.hidden) >= len(self.keys):
            self.hidden = set()
        self.pinned = [index[k] for k in saved.get("pinned", []) if k in index]
        if saved.get("density") in self.DENSITY:
            self.density = saved["density"]
        self.saved_widths = {index[k]: w for k, w in saved.get("widths", {}).items() if k in index}
        for i, width in self.saved_widths.items():
            self.widths[i] = max(self.px(self.title_width(self.headers[i], i)), self.px(width))
            self.resized.add(i)

    def save_view(self):
        self.update_toolbar()
        if not self.name:
            return
        settings = data.load_settings()
        views = settings.setdefault("table_view", {})
        if self.is_default():
            views.pop(self.name, None)
        else:
            views[self.name] = {
                "order": [self.keys[i] for i in self.order], "hidden": [self.keys[i] for i in sorted(self.hidden)],
                "pinned": [self.keys[i] for i in self.pinned], "density": self.density,
                "widths": {self.keys[i]: round(self.widths[i] / self.scale) for i in self.resized},
            }
        data.save_settings(settings)

    def set_density(self, value):
        self.commit_edit()
        self.density = value
        self.save_view()
        self.draw()

    def reset_view(self):
        self.commit_edit()
        self.order = list(range(len(self.headers)))
        self.hidden = set()
        self.pinned = [i for i in self.default_pins if i < len(self.headers)]
        self.density = "Normal"
        self.resized = set()
        self.widths = [self.px(self.title_width(label, i)) for i, label in enumerate(self.headers)]
        self.fit_columns()
        self.x0 = 0
        self.anchor = self.active = None
        self.save_view()
        self.draw()

    def toggle_pin(self, real):
        if real in self.pinned:
            self.pinned.remove(real)
        else:
            self.pinned.append(real)
        self.anchor = self.active = None
        self.save_view()
        self.draw()

    def set_hidden(self, real, hide):
        if hide and len(self.headers) - len(self.hidden) <= 1:
            return False
        (self.hidden.add if hide else self.hidden.discard)(real)
        self.anchor = self.active = None
        self.commit_edit()
        self.save_view()
        self.draw()
        return True

    def toggle_menu(self):
        if self.menu is not None and self.menu.winfo_exists():
            self.menu.destroy()
            self.menu = None
            return
        self.menu = ColumnMenu(self)

    # ---------- ukuran dan warna ----------

    def px(self, value):
        return round(value * self.scale)

    def color(self, name):
        if self.rounded and name in ("head", "head_text", "line", "flag"):
            name = {"head": "head_round", "head_text": "head_round_text", "line": "line_round", "flag": "flag_round"}[name]
        palette = {
            "head_round": ("#F1E6D7", "#1B2A47"), "head_round_text": ("#0A1F40", "#FFFFFF"), "line_round": ("#EEE4D6", "#283044"),
            "flag_round": ("#FDEDEA", "#4A2320"), "cell_line": ("#E0D2BF", "#323B52"),
            "card": ("#FFFFFF", "#1A2030"), "text": ("#16202E", "#EEF0F5"), "muted": ("#5E6675", "#A0A8B8"),
            "faint": ("#9AA0AB", "#737B8C"), "line": ("#F3ECE2", "#283044"), "head_line": ("#E8DCCB", "#323B52"),
            "head": ("#FBF6EF", "#151B29"), "head_text": ("#5E6675", "#E0D2BF"), "link": ("#0F2C59", "#C9A97F"),
            "hover": ("#F3E8DA", "#2A3D63"), "active": ("#F6EEE3", "#22304A"), "range": ("#E6D5BF", "#2B3F66"),
            "accent": ("#0F2C59", "#C9A97F"), "flag": ("#FDF1EF", "#4A2320"), "dirty": ("#FFF6D9", "#4A3B1A"),
            "danger": ("#8C2015", "#F97066"), "shadow": ("#EEE4D6", "#0E1220"),
        }
        return palette[name][1 if is_dark() else 0]

    def title_width(self, label, index):
        icons = (self.ICON * 2 + self.ICON_GAP + 6) if index < len(self.columns) else (self.ICON + 6)
        return (self.head_font.measure(label) / self.scale) + self.PAD * 2 + icons

    def set_marks(self, marked):
        self.marked = set(marked)
        self.schedule()

    def fit_columns(self, only=None):
        """Lebar kolom mengikuti judul dan isinya, ditambah sedikit jarak."""
        sample = self.rows[:self.SAMPLE // 2] + self.rows[-self.SAMPLE // 2:] if len(self.rows) > self.SAMPLE else self.rows
        for index, label in enumerate(self.headers):
            if (only is not None and index != only) or (only is None and index in self.resized):
                continue
            content = max((self.font.measure(self.clean(row[index])) for row in sample if index < len(row)), default=0)
            wanted = max(self.px(self.title_width(label, index)), min(content + self.px(self.PAD * 2 + 20), self.px(self.MAX_FIT)))
            self.widths[index] = wanted if only is not None else max(self.widths[index], wanted)

    @staticmethod
    def clean(value):
        return " ".join(str(value).split()) if value is not None else ""

    def fill(self, rows, tags, cell_styles=None):
        self.commit_edit()
        self.rows = rows
        self.tags = list(tags)
        self.styles = {}
        for style, cells in (cell_styles or {}).items():
            for cell in cells:
                self.styles[tuple(cell)] = style
        self.fit_columns()
        self.picked_rows = {r for r in self.picked_rows if r < len(rows)}
        if self.active and self.active[0] >= len(rows):
            self.anchor = self.active = None
        self.y0 = min(self.y0, self.max_y())
        self.draw()

    def schedule(self):
        if self.draw_job is None:
            self.draw_job = self.after_idle(self.draw)

    def view_size(self):
        return max(1, self.canvas.winfo_width()), max(1, self.canvas.winfo_height())

    def content_width(self):
        return sum(self.widths[i] for i in self.view) + self.px(self.INDEX)

    def pin_right(self):
        return self.px(self.INDEX) + sum(self.widths[i] for i in self.view[:self.pin_count()])

    def max_y(self):
        return max(0, len(self.rows) * self.row_h() - (self.view_size()[1] - self.px(self.HEAD)))

    def max_x(self):
        return max(0, self.content_width() - self.view_size()[0])

    def page_rows(self):
        return max(1, (self.view_size()[1] - self.px(self.HEAD)) // self.row_h() - 1)

    def column_left(self, position):
        view = self.view
        left = self.px(self.INDEX) + sum(self.widths[i] for i in view[:position])
        return left if position < self.pin_count() else left - self.x0

    def fit_text(self, text, width, font):
        key = (text, width, str(font))
        cached = self.fit_cache.get(key)
        if cached is not None:
            return cached
        if font.measure(text) <= width:
            result = text
        else:
            low, high = 0, len(text)
            while low < high:
                middle = (low + high + 1) // 2
                if font.measure(text[:middle] + "…") <= width:
                    low = middle
                else:
                    high = middle - 1
            result = text[:low] + "…"
        if len(self.fit_cache) > 20000:
            self.fit_cache.clear()
        self.fit_cache[key] = result
        return result

    def selection_box(self):
        if not self.anchor or not self.active:
            return None
        return (min(self.anchor[0], self.active[0]), max(self.anchor[0], self.active[0]),
                min(self.anchor[1], self.active[1]), max(self.anchor[1], self.active[1]))

    def band_color(self, row):
        if row == self.hover:
            return self.color("hover")
        if row in self.picked_rows or (self.active and self.active[0] == row):
            return self.color("active")
        tag = self.tags[row] if row < len(self.tags) else ""
        return self.color("flag") if tag == "flag" else self.color("dirty") if tag == "dirty" else self.color("card")

    # ---------- menggambar ----------

    def draw(self):
        self.draw_job = None
        c = self.canvas
        c.delete("all")
        c.configure(bg=self.color("card"))
        width, height = self.view_size()
        row_h, head_h, index_w = self.row_h(), self.px(self.HEAD), self.px(self.INDEX)
        self.x0 = min(self.x0, self.max_x())
        self.y0 = min(self.y0, self.max_y())
        view = self.view
        pins = self.pin_count()
        lefts = [self.column_left(p) for p in range(len(view))]
        pin_right = self.pin_right()
        loose = [p for p in range(pins, len(view)) if lefts[p] + self.widths[view[p]] > pin_right and lefts[p] < width]
        first = self.y0 // row_h
        last = min(len(self.rows), (self.y0 + height - head_h) // row_h + 2)
        line = self.color("line")
        for row in range(first, last):
            top = head_h + row * row_h - self.y0
            c.create_rectangle(0, top, width, top + row_h, fill=self.band_color(row), outline="", tags=(f"band{row}",))
            c.create_line(0, top + row_h - 1, width, top + row_h - 1, fill=line)
            for p in loose:
                self.draw_cell(row, p, view[p], lefts[p], top)
        for row in range(first, last):
            top = head_h + row * row_h - self.y0
            c.create_rectangle(0, top, pin_right, top + row_h, fill=self.band_color(row), outline="", tags=(f"band{row}",))
            c.create_line(0, top + row_h - 1, pin_right, top + row_h - 1, fill=line)
            for p in range(pins):
                self.draw_cell(row, p, view[p], lefts[p], top)
            c.create_text(index_w // 2, top + row_h // 2, text=str(row + 1), fill=self.color("faint"), font=self.small)
        self.draw_header(view, pins, lefts, loose, pin_right, width)
        if pins and self.x0 > 0:
            for step, shade in enumerate((self.color("head_line"), self.color("shadow"), self.color("line"))):
                c.create_line(pin_right + step, 0, pin_right + step, height, fill=shade)
        elif pins:
            c.create_line(pin_right, 0, pin_right, height, fill=self.color("head_line"))
        if self.drop_at is not None:
            x = self.px(self.INDEX) + sum(self.widths[i] for i in view[:self.drop_at]) - (0 if self.drop_at < pins else self.x0)
            c.create_rectangle(x - 1, 0, x + 2, height, fill=self.color("accent"), outline="")
        content = max(1, len(self.rows) * row_h)
        view_h = max(1, height - head_h)
        self.vbar.set(self.y0 / content, min(1, (self.y0 + view_h) / content))
        total = max(1, self.content_width())
        self.hbar.set(self.x0 / total, min(1, (self.x0 + width) / total))
        self.place_editor()

    def draw_cell(self, row, position, real, left, top):
        c = self.canvas
        row_h, pad = self.row_h(), self.px(self.PAD)
        right = left + self.widths[real]
        values = self.rows[row]
        value = self.clean(values[real]) if real < len(values) else ""
        box = self.selection_box()
        in_range = bool(box and box[0] <= row <= box[1] and box[2] <= position <= box[3] and (box[0] != box[1] or box[2] != box[3]))
        if self.rounded:
            self.draw_round_cell(row, position, real, left, right, top, value, in_range)
            return
        if in_range:
            c.create_rectangle(left, top, right, top + row_h - 1, fill=self.color("range"), outline="")
        if self.active == (row, position):
            c.create_rectangle(left + 1, top + 1, right - 1, top + row_h - 2, outline=self.color("accent"), width=2)
        if not value:
            return
        style = self.styles.get((row, real))
        fg, font = self.color("text"), self.font
        if real == 0 and real < len(self.columns):
            fg, font = self.color("link"), self.bold
        if style:
            fg, font = style[1], self.bold
        elif self.flag_column and real == len(self.columns):
            fg, font = self.color("danger"), self.bold
        numeric = self.kinds[real] == "num"
        room = self.widths[real] - pad * 2
        middle = top + row_h // 2
        if style and style[0]:
            text = self.fit_text(value, room - self.px(16), font)
            size = font.measure(text) + self.px(18)
            height = min(self.cell_h(), self.px(26))
            x1 = right - pad - size if numeric else left + pad - self.px(6)
            c.create_polygon(rounded_points(x1, middle - height // 2, x1 + size, middle + height // 2, height // 2), smooth=True, fill=style[0], outline="")
            c.create_text(x1 + size // 2, middle, text=text, fill=fg, font=font)
            return
        text = self.fit_text(value, room, font)
        if numeric:
            c.create_text(right - pad, middle, text=text, anchor="e", fill=fg, font=font)
        else:
            c.create_text(left + pad, middle, text=text, anchor="w", fill=fg, font=font)

    def draw_round_cell(self, row, position, real, left, right, top, value, in_range):
        c = self.canvas
        inset, cell_h, pad = self.px(3), self.cell_h(), self.px(self.PAD)
        cell_top = top + (self.row_h() - cell_h) // 2
        style = self.styles.get((row, real))
        bg, fg, font = self.color("card"), self.color("text"), self.font
        if real == 0 and real < len(self.columns):
            fg, font = self.color("link"), self.bold
        if style:
            bg, fg, font = style[0] or bg, style[1], self.bold
        elif self.flag_column and real == len(self.columns) and value:
            fg, font = self.color("danger"), self.bold
        if in_range and not (style and style[0]):
            bg = self.color("range")
        active = self.active == (row, position)
        c.create_polygon(
            rounded_points(left + inset, cell_top, right - inset, cell_top + cell_h, self.px(8)), smooth=True, fill=bg,
            outline=self.color("accent") if active else self.color("cell_line"), width=2 if active else 1,
        )
        if not value:
            return
        text = self.fit_text(value, self.widths[real] - pad * 2, font)
        if self.kinds[real] == "num":
            c.create_text(right - pad, cell_top + cell_h // 2, text=text, anchor="e", fill=fg, font=font)
        else:
            c.create_text(left + pad, cell_top + cell_h // 2, text=text, anchor="w", fill=fg, font=font)

    def icon_boxes(self, position):
        """Posisi ikon kunci dan filter di judul kolom."""
        real = self.view[position]
        right = self.column_left(position) + self.widths[real]
        icon, gap = self.px(self.ICON), self.px(self.ICON_GAP)
        top = (self.px(self.HEAD) - icon) // 2
        has_filter = self.on_filter is not None and real < len(self.columns)
        filter_box = (right - icon - self.px(6), top, right - self.px(6), top + icon) if has_filter else None
        pin_right = filter_box[0] - gap if filter_box else right - self.px(6)
        return (pin_right - icon, top, pin_right, top + icon), filter_box

    def draw_header(self, view, pins, lefts, loose, pin_right, width):
        c = self.canvas
        head_h = self.px(self.HEAD)
        c.create_rectangle(0, 0, width, head_h, fill=self.color("head"), outline="")
        for p in loose:
            self.draw_head_cell(p, view[p], lefts[p])
        c.create_rectangle(0, 0, pin_right, head_h, fill=self.color("head"), outline="")
        for p in range(pins):
            self.draw_head_cell(p, view[p], lefts[p])
        c.create_line(0, head_h - 1, width, head_h - 1, fill=self.color("head_line"))

    def draw_head_cell(self, position, real, left):
        c = self.canvas
        head_h, pad = self.px(self.HEAD), self.px(self.PAD)
        right = left + self.widths[real]
        pin_box, filter_box = self.icon_boxes(position)
        numeric = self.kinds[real] == "num"
        room = pin_box[0] - left - pad - self.px(4)
        text = self.fit_text(self.headers[real], room, self.head_font)
        if numeric:
            c.create_text(pin_box[0] - self.px(6), head_h // 2, text=text, anchor="e", fill=self.color("head_text"), font=self.head_font)
        else:
            c.create_text(left + pad, head_h // 2, text=text, anchor="w", fill=self.color("head_text"), font=self.head_font)
        c.create_line(right - 1, head_h * 0.3, right - 1, head_h * 0.7, fill=self.color("head_line"))
        pinned = real in self.pinned
        if pinned or self.hover_col == position:
            self.draw_pin(pin_box, pinned)
        if filter_box:
            self.draw_filter_icon(filter_box, self.keys[real] in self.marked)

    def draw_pin(self, box, active):
        c = self.canvas
        x, y, size = box[0], box[1], box[2] - box[0]
        if icons.available("pin"):
            if active:
                c.create_polygon(rounded_points(x, y, x + size, y + size, self.px(5)), smooth=True, fill=self.color("active"), outline="")
            image = icons.photo("pin", self.color("accent") if active else self.color("faint"), max(8, round(size * 0.7)))
            c.create_image(x + size // 2, y + size // 2, image=image)
            return
        u = size / 22
        color = self.color("accent") if active else self.color("faint")
        if active:
            c.create_polygon(rounded_points(x, y, x + size, y + size, self.px(5)), smooth=True, fill=self.color("active"), outline="")
        c.create_polygon(x + 7 * u, y + 4 * u, x + 15 * u, y + 4 * u, x + 14 * u, y + 9 * u, x + 17 * u, y + 13 * u,
                         x + 5 * u, y + 13 * u, x + 8 * u, y + 9 * u, fill=color, outline="")
        c.create_line(x + 11 * u, y + 13 * u, x + 11 * u, y + 19 * u, fill=color, width=max(1, round(2 * u)))

    def draw_filter_icon(self, box, active):
        c = self.canvas
        x, y, size = box[0], box[1], box[2] - box[0]
        accent = self.color("accent")
        c.create_polygon(rounded_points(x, y, x + size, y + size, self.px(6)), smooth=True, fill=accent if active else self.color("card"), outline=accent if active else self.color("head_line"))
        mark = "#FFFFFF" if active else accent
        if icons.available("funnel"):
            c.create_image(x + size // 2, y + size // 2, image=icons.photo("funnel", mark, max(8, round(size * 0.62))))
            return
        u = size / 24
        c.create_polygon(x + 6 * u, y + 7 * u, x + 18 * u, y + 7 * u, x + 13.5 * u, y + 12.5 * u, x + 13.5 * u, y + 17 * u, x + 10.5 * u, y + 18 * u, x + 10.5 * u, y + 12.5 * u, fill=mark, outline="")

    # ---------- gulir ----------

    def scroll_y(self, *args):
        self.scroll(args, "y")

    def scroll_x(self, *args):
        self.scroll(args, "x")

    def scroll(self, args, axis):
        total = len(self.rows) * self.row_h() if axis == "y" else self.content_width()
        if args[0] == "moveto":
            target = float(args[1]) * total
        else:
            step = self.row_h() if axis == "y" else self.px(60)
            target = (self.y0 if axis == "y" else self.x0) + int(args[1]) * step
        if axis == "y":
            self.move_y(target - self.y0)
        else:
            self.move_x(target - self.x0)

    def move_y(self, delta):
        self.commit_edit()
        self.y0 = int(max(0, min(self.max_y(), self.y0 + delta)))
        self.schedule()

    def move_x(self, delta):
        self.commit_edit()
        self.x0 = int(max(0, min(self.max_x(), self.x0 + delta)))
        self.schedule()

    def on_wheel(self, event):
        self.move_y(-3 * self.row_h() if event.delta > 0 else 3 * self.row_h())

    # ---------- posisi mouse ----------

    def row_at(self, y):
        head = self.px(self.HEAD)
        if y < head:
            return None
        row = (y - head + self.y0) // self.row_h()
        return row if row < len(self.rows) else None

    def column_at(self, x):
        view = self.view
        pins = self.pin_count()
        for position in list(range(pins)) + list(range(pins, len(view))):
            left = self.column_left(position)
            if position >= pins and left + self.widths[view[position]] <= self.pin_right():
                continue
            if position >= pins and x < self.pin_right():
                continue
            if left <= x < left + self.widths[view[position]]:
                return position
        return None

    def resize_edge(self, x, y):
        if y >= self.px(self.HEAD):
            return None
        view = self.view
        for position, real in enumerate(view):
            edge = self.column_left(position) + self.widths[real]
            if position >= self.pin_count() and edge < self.pin_right():
                continue
            if abs(edge - x) <= self.px(5):
                return position
        return None

    @staticmethod
    def inside(box, x, y):
        return box is not None and box[0] <= x <= box[2] and box[1] <= y <= box[3]

    def set_hover(self, row):
        if row == self.hover:
            return
        old, self.hover = self.hover, row
        for item in (old, row):
            if item is not None:
                self.canvas.itemconfigure(f"band{item}", fill=self.band_color(item))

    def set_hover_col(self, position):
        if position != self.hover_col:
            self.hover_col = position
            self.schedule()

    def on_motion(self, event):
        edge = self.resize_edge(event.x, event.y)
        head = event.y < self.px(self.HEAD)
        cursor = "sb_h_double_arrow" if edge is not None else ("hand2" if head and event.x > self.px(self.INDEX) else "")
        self.canvas.configure(cursor=cursor)
        self.set_hover(self.row_at(event.y))
        self.set_hover_col(self.column_at(event.x) if head else None)

    def on_press(self, event):
        self.commit_edit()
        self.canvas.focus_set()
        edge = self.resize_edge(event.x, event.y)
        if edge is not None:
            real = self.view[edge]
            self.drag_width = (real, event.x, self.widths[real])
            return
        if event.y < self.px(self.HEAD):
            self.drag_mode = None
            position = self.column_at(event.x)
            if position is None:
                return
            pin_box, filter_box = self.icon_boxes(position)
            real = self.view[position]
            if self.inside(filter_box, event.x, event.y):
                self.on_filter(self.columns[real])
            elif self.inside(pin_box, event.x, event.y):
                self.toggle_pin(real)
            elif self.movable:
                self.head_drag = [position, event.x, False]
            elif self.on_filter and real < len(self.columns):
                self.on_filter(self.columns[real])
            return
        row = self.row_at(event.y)
        if row is None:
            return
        shift, ctrl = event.state & 0x1, event.state & 0x4
        if event.x < self.px(self.INDEX):
            if shift and self.active:
                self.picked_rows |= set(range(min(self.active[0], row), max(self.active[0], row) + 1))
            elif ctrl:
                self.picked_rows ^= {row}
            else:
                self.picked_rows = {row}
            self.anchor = (row, 0)
            self.active = (row, 0)
            self.drag_mode = "rows"
            self.draw()
            return
        position = self.column_at(event.x)
        if position is None:
            return
        if not ctrl:
            self.picked_rows = set()
        self.active = (row, position)
        if not shift or not self.anchor:
            self.anchor = (row, position)
        self.drag_mode = "cells"
        self.draw()
        self.report_selection()

    def drop_index(self, x, moving):
        """Posisi sisip kolom yang sedang ditarik, dihitung dari titik tengah kolom lain."""
        view = self.view
        count = 0
        for position, real in enumerate(view):
            if position == moving:
                continue
            middle = self.column_left(position) + self.widths[real] / 2
            if middle < x:
                count = position + (0 if position < moving else 1)
        return min(count, len(view))

    def on_drag(self, event):
        if self.drag_width:
            real, start, width = self.drag_width
            self.widths[real] = max(self.px(self.title_width(self.headers[real], real)), width + event.x - start)
            self.resized.add(real)
            self.schedule()
            return
        if self.head_drag:
            position, start, moved = self.head_drag
            if moved or abs(event.x - start) > self.px(6):
                self.head_drag[2] = True
                self.canvas.configure(cursor="fleur")
                target = self.drop_index(event.x, position)
                self.drop_at = target if target not in (position, position + 1) else None
                width = self.view_size()[0]
                if event.x > width - self.px(30):
                    self.x0 = min(self.max_x(), self.x0 + self.px(20))
                elif event.x < self.pin_right() + self.px(30) and self.x0 > 0:
                    self.x0 = max(0, self.x0 - self.px(20))
                self.schedule()
            return
        row = self.row_at(event.y)
        if row is None:
            return
        if self.drag_mode == "rows":
            self.picked_rows |= set(range(min(self.anchor[0], row), max(self.anchor[0], row) + 1))
            self.schedule()
        elif self.drag_mode == "cells":
            position = self.column_at(event.x)
            if position is not None and (row, position) != self.active:
                self.active = (row, position)
                self.schedule()

    def move_column(self, position, target):
        """Memindah kolom di tampilan. Masuk ke area kunci berarti ikut dikunci, keluar berarti dilepas."""
        view = self.view
        pins = self.pin_count()
        real = view.pop(position)
        if target > position:
            target -= 1
        view.insert(target, real)
        was_pinned = position < pins
        now_pins = pins - (1 if was_pinned else 0)
        pinned = set(i for i in self.pinned if i not in self.hidden)
        if was_pinned and target >= now_pins:
            pinned.discard(real)
        elif not was_pinned and target < pins:
            pinned.add(real)
        hidden_tail = [i for i in self.order if i in self.hidden]
        self.pinned = [i for i in view if i in pinned]
        self.order = view + hidden_tail
        self.anchor = self.active = None
        self.save_view()

    def on_release(self, event):
        if self.drag_width:
            self.drag_width = None
            self.save_view()
            return
        if self.head_drag:
            position, _, moved = self.head_drag
            self.head_drag = None
            target, self.drop_at = self.drop_at, None
            self.canvas.configure(cursor="")
            if moved:
                if target is not None:
                    self.move_column(position, target)
                self.draw()
            else:
                real = self.view[position]
                if self.on_filter and real < len(self.columns):
                    self.on_filter(self.columns[real])
            return
        self.drag_mode = None

    def on_double(self, event):
        edge = self.resize_edge(event.x, event.y)
        if edge is not None:
            real = self.view[edge]
            self.resized.discard(real)
            self.fit_columns(only=real)
            self.save_view()
            self.draw()
            return
        if event.y >= self.px(self.HEAD) and event.x >= self.px(self.INDEX) and self.row_at(event.y) is not None:
            self.begin_edit()

    # ---------- keyboard dan edit ----------

    def on_key(self, event):
        if event.char and event.char.isprintable() and not event.state & 0x4:
            self.begin_edit(event.char)
            return "break"
        return None

    def move_active(self, rows, columns, extend=False):
        if not self.rows:
            return "break"
        row, column = self.active or (0, 0)
        row = max(0, min(len(self.rows) - 1, row + rows))
        column = max(0, min(len(self.view) - 1, column + columns))
        self.active = (row, column)
        if not extend or not self.anchor:
            self.anchor = self.active
            self.picked_rows = set()
        self.reveal(row, column)
        self.draw()
        self.report_selection()
        return "break"

    def report_selection(self):
        if self.on_select and self.active and self.active[1] < len(self.view):
            self.on_select(self.active[0], self.view[self.active[1]])

    def reveal(self, row, position):
        width, height = self.view_size()
        row_h, head = self.row_h(), self.px(self.HEAD)
        top = row * row_h
        if top < self.y0:
            self.y0 = top
        elif top + row_h > self.y0 + height - head:
            self.y0 = top + row_h - (height - head)
        if position >= self.pin_count():
            view = self.view
            left = self.px(self.INDEX) + sum(self.widths[i] for i in view[:position])
            right = left + self.widths[view[position]]
            visible_left = self.pin_right() - self.px(self.INDEX) + self.px(self.INDEX)
            if left - self.x0 < visible_left:
                self.x0 = max(0, left - visible_left)
            elif right - self.x0 > width:
                self.x0 = right - width
        self.y0 = max(0, min(self.max_y(), self.y0))
        self.x0 = max(0, min(self.max_x(), self.x0))

    def can_edit(self, row, real):
        return real < len(self.columns) and self.on_edit(str(row), self.columns[real], "", check=True)

    def begin_edit(self, first=None):
        if not self.active or self.active[0] >= len(self.rows) or self.active[1] >= len(self.view):
            return "break"
        row, position = self.active
        real = self.view[position]
        if not self.can_edit(row, real):
            return "break"
        self.reveal(row, position)
        self.commit_edit()
        self.edit_cell = (row, real)
        self.editor = tk.Entry(
            self.canvas, font=self.font, bd=0, relief="flat", highlightthickness=2, highlightcolor=self.color("accent"),
            highlightbackground=self.color("accent"), bg=self.color("card"), fg=self.color("text"), insertbackground=self.color("text"),
        )
        value = self.rows[row][real] if real < len(self.rows[row]) else ""
        self.editor.insert(0, first if first is not None else ("" if value is None else str(value)))
        self.editor.bind("<Return>", lambda _: self.finish_edit(1))
        self.editor.bind("<Tab>", lambda _: self.finish_edit(0, 1))
        self.editor.bind("<Escape>", lambda _: self.cancel_edit())
        self.editor.bind("<FocusOut>", lambda _: self.commit_edit())
        self.draw()
        self.editor.focus_set()
        if first is None:
            self.editor.select_range(0, "end")
        return "break"

    def place_editor(self):
        if not self.editor or not self.edit_cell:
            return
        row, real = self.edit_cell
        if real not in self.view:
            return
        position = self.view.index(real)
        top = self.px(self.HEAD) + row * self.row_h() - self.y0 + 2
        self.editor.place(x=self.column_left(position) + 1, y=top, width=max(self.widths[real], self.px(120)) - 2, height=self.row_h() - 5)

    def close_editor(self):
        editor, self.editor = self.editor, None
        cell, self.edit_cell = self.edit_cell, None
        if editor is not None:
            editor.unbind("<FocusOut>")
            text = editor.get()
            editor.destroy()
            return cell, text
        return None, None

    def cancel_edit(self):
        self.close_editor()
        self.canvas.focus_set()
        return "break"

    def commit_edit(self):
        cell, text = self.close_editor()
        if cell is None:
            return
        row, real = cell
        old = self.rows[row][real] if row < len(self.rows) and real < len(self.rows[row]) else ""
        if text != ("" if old is None else str(old)):
            self.after_idle(lambda: self.on_edit(str(row), self.columns[real], text))

    def finish_edit(self, rows=0, columns=0):
        self.commit_edit()
        self.canvas.focus_set()
        self.move_active(rows, columns)
        return "break"

    def copy(self, _event=None):
        rows = self.selected_rows()
        if not rows:
            return "break"
        view = self.view
        if self.picked_rows:
            reals = [i for i in view if i < len(self.columns)]
        else:
            box = self.selection_box()
            reals = view[box[2]:box[3] + 1]
        lines = ["\t".join(self.clean(self.rows[r][c]) if c < len(self.rows[r]) else "" for c in reals) for r in rows]
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        return "break"

    def select_all(self, _event=None):
        if self.rows:
            self.anchor = (0, 0)
            self.active = (len(self.rows) - 1, len(self.view) - 1)
            self.picked_rows = set()
            self.draw()
        return "break"

    def selected_rows(self):
        rows = set(self.picked_rows)
        box = self.selection_box()
        if box:
            rows |= set(range(box[0], box[1] + 1))
        return sorted(r for r in rows if r < len(self.rows))

    def see(self, row):
        self.picked_rows = {row}
        self.anchor = self.active = (row, 0)
        self.reveal(row, 0)
        self.draw()

    def header_anchor(self, column_key):
        """Posisi layar di bawah judul kolom, tempat munculnya kotak filter."""
        for position, real in enumerate(self.view):
            if self.keys[real] == column_key:
                left = max(self.px(self.INDEX), self.column_left(position))
                return self.canvas.winfo_rootx() + left, self.canvas.winfo_rooty() + self.px(self.HEAD)
        return self.canvas.winfo_rootx(), self.canvas.winfo_rooty()


class ColumnMenu(ctk.CTkToplevel):
    """Daftar kolom untuk ditampilkan atau disembunyikan, beserta lebarnya."""

    def __init__(self, grid):
        super().__init__(grid, fg_color=("#FFFFFF", "#1A2030"))
        self.grid_view = grid
        self.title("Pilih kolom")
        self.resizable(False, False)
        button = grid.column_button
        rows = len(grid.headers)
        width, height = 320, min(560, 40 * rows + 70)
        x = button.winfo_rootx() + button.winfo_width() - width
        y = button.winfo_rooty() + button.winfo_height() + 4
        self.geometry(f"{width}x{height}+{max(0, x)}+{y}")
        self.transient(grid.winfo_toplevel())
        family = "Segoe UI" if os.name == "nt" else "DejaVu Sans"
        body = ctk.CTkScrollableFrame(self, fg_color="transparent") if rows > 12 else ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=8, pady=(8, 0))
        self.boxes = []
        for real, label in enumerate(grid.headers):
            line = ctk.CTkFrame(body, fg_color="transparent", height=38)
            line.pack(fill="x")
            var = tk.BooleanVar(value=real not in grid.hidden)
            suffix = "  (dikunci)" if real in grid.pinned else ""
            box = ctk.CTkCheckBox(line, text=label + suffix, variable=var, font=ctk.CTkFont(family=family, size=14), checkbox_width=18, checkbox_height=18,
                                  corner_radius=5, fg_color="#0F2C59", hover_color="#0A1F40", command=lambda r=real, v=var: self.toggle(r, v))
            box.pack(side="left", padx=6, pady=6)
            info = ctk.CTkLabel(line, text="", font=ctk.CTkFont(family="Consolas", size=12), text_color=("#9AA0AB", "#737B8C"))
            info.pack(side="right", padx=8)
            self.boxes.append((real, var, info))
        foot = ctk.CTkFrame(self, fg_color="transparent")
        foot.pack(fill="x", padx=14, pady=8)
        self.count = ctk.CTkLabel(foot, text="", font=ctk.CTkFont(family=family, size=13), text_color=("#5E6675", "#A0A8B8"))
        self.count.pack(side="left")
        ctk.CTkButton(foot, text="Tampilkan semua", width=120, height=30, fg_color="transparent", hover_color=("#F6EEE3", "#2A3D63"),
                      text_color=("#0F2C59", "#C9A97F"), font=ctk.CTkFont(family=family, size=13, weight="bold"), command=self.show_all).pack(side="right")
        self.bind("<Escape>", lambda _: self.close())
        self.bind("<FocusOut>", lambda _: self.after(150, self.check_focus))
        self.refresh()
        self.after(80, self.focus_set)

    def check_focus(self):
        if not self.winfo_exists():
            return
        widget = self.focus_get()
        if widget is None or widget.winfo_toplevel() is not self:
            self.close()

    def close(self):
        if self.winfo_exists():
            self.destroy()
        self.grid_view.menu = None

    def refresh(self):
        grid = self.grid_view
        for real, var, info in self.boxes:
            shown = real not in grid.hidden
            var.set(shown)
            info.configure(text=f"{round(grid.widths[real] / grid.scale)} px" if shown else "tersembunyi")
        total = len(grid.headers)
        self.count.configure(text=f"{total - len(grid.hidden)} dari {total} terlihat")

    def toggle(self, real, var):
        if not self.grid_view.set_hidden(real, not var.get()):
            var.set(True)
        self.refresh()

    def show_all(self):
        grid = self.grid_view
        grid.hidden = set()
        grid.save_view()
        grid.draw()
        self.refresh()


DATE_VALUE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})")
MONTHS = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
MONTHS_SHORT = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]


class FilterDialog(ctk.CTkToplevel):
    """Filter dengan ringkasan di kiri dan pohon centang di kanan. Tanggal tersusun Tahun, Bulan, Tanggal,
    tiap kelompok punya badge "x dari y", dan pencarian bisa langsung mencentang semua hasil."""

    def __init__(self, page, column):
        super().__init__(page, fg_color=page.manager.ui.CARD)
        ui = page.manager.ui
        self.ui = ui
        self.page, self.target = page, column
        self.title(f"Filter {column.label}")
        self.counts = page.column_counts(column)
        self.values = set(self.counts)
        self.total_rows = sum(self.counts.values())
        current = set(page.filters.get(column.key, ())) & self.values
        self.base = current or set(self.values)
        self.chosen = set(self.base)
        self.order_by_count = False
        self.last_words = []
        self.nodes = {}
        self.shown = set(self.values)
        self.is_dates = any(DATE_VALUE.match(v) for v in self.values)
        self.unit = page.profile.name if page.table in ("master", "po_master") and page.profile else "baris"
        scale = page.grid_widget.scale
        family = "Segoe UI" if os.name == "nt" else "DejaVu Sans"
        x, y = page.grid_widget.header_anchor(column.key)
        top = page.winfo_toplevel()
        width, height = 760, 640
        x = max(top.winfo_rootx(), min(x, top.winfo_rootx() + top.winfo_width() - width - 20))
        y = max(top.winfo_rooty(), min(y, top.winfo_rooty() + top.winfo_height() - height - 20))
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.transient(top)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        soft = ("#FBF6EF", "#151B29")
        side = ctk.CTkFrame(self, fg_color=soft, corner_radius=0, width=260)
        side.grid(row=0, column=0, sticky="nsew")
        side.grid_propagate(False)
        side.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(side, text="FILTER KOLOM", font=ui.font(12, "bold"), text_color=ui.ACCENT, anchor="w").grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 0))
        ctk.CTkLabel(side, text=column.label, font=ui.font(22, "bold"), text_color=ui.TEXT, anchor="w").grid(row=1, column=0, sticky="ew", padx=20)
        ctk.CTkLabel(side, text="Centang nilai yang mau ditampilkan. Yang tidak dicentang hanya disembunyikan, datanya tetap ada.",
                     font=ui.font(12), text_color=ui.MUTED, anchor="w", justify="left", wraplength=220).grid(row=2, column=0, sticky="ew", padx=20, pady=(4, 12))
        number = ctk.CTkFrame(side, fg_color="transparent")
        number.grid(row=3, column=0, sticky="ew", padx=20)
        self.big = ctk.CTkLabel(number, text="0", font=ui.font(46, "bold"), text_color=ui.ACCENT)
        self.big.pack(side="left")
        ctk.CTkLabel(number, text=f"{self.unit} tampil", font=ui.font(13), text_color=ui.MUTED).pack(side="left", padx=8, pady=(16, 0))
        self.sub = ctk.CTkLabel(side, text="", font=ui.font(12), text_color=ui.MUTED, anchor="w")
        self.sub.grid(row=4, column=0, sticky="ew", padx=20)
        self.bar = ctk.CTkProgressBar(side, height=6, progress_color=ui.ACCENT, fg_color=("#E8DCCB", "#283044"))
        self.bar.grid(row=5, column=0, sticky="ew", padx=20, pady=(8, 14))
        ctk.CTkLabel(side, text="YANG DIPILIH", font=ui.font(12, "bold"), text_color=ui.MUTED, anchor="w").grid(row=6, column=0, sticky="ew", padx=20)
        self.chips = ctk.CTkFrame(side, fg_color="transparent")
        self.chips.grid(row=7, column=0, sticky="new", padx=20, pady=(6, 0))
        side.grid_rowconfigure(8, weight=1)
        up, down = ("Terlama dulu", "Terbaru dulu") if self.is_dates else (("Kecil dulu", "Besar dulu") if column.kind == "num" else ("A ke Z", "Z ke A"))
        sort = ctk.CTkFrame(side, fg_color="transparent")
        sort.grid(row=9, column=0, sticky="ew", padx=20, pady=(0, 8))
        sort.grid_columnconfigure((0, 1), weight=1, uniform="sort")
        ctk.CTkLabel(sort, text="Urutkan tampilan saja", font=ui.font(11), text_color=ui.MUTED, anchor="w").grid(row=0, column=0, columnspan=2, sticky="ew")
        ui.button(sort, up, lambda: self.sort(False), height=32).grid(row=1, column=0, sticky="ew", padx=(0, 3))
        ui.button(sort, down, lambda: self.sort(True), height=32).grid(row=1, column=1, sticky="ew", padx=(3, 0))
        actions = ctk.CTkFrame(side, fg_color="transparent")
        actions.grid(row=10, column=0, sticky="ew", padx=20, pady=(0, 20))
        actions.grid_columnconfigure(1, weight=1)
        ui.button(actions, "Reset", self.reset, width=74, height=42).grid(row=0, column=0)
        ui.button(actions, "Terapkan", self.apply, primary=True, height=42).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=16, pady=16)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)
        search = ctk.CTkFrame(main, fg_color=ui.CARD, border_color=ui.BORDER, border_width=1, corner_radius=11, height=44)
        search.grid(row=0, column=0, sticky="ew")
        search.grid_columnconfigure(1, weight=1)
        self.search_box = search
        ctk.CTkLabel(search, text="Cari", font=ui.font(12, "bold"), text_color=ui.MUTED, width=36).grid(row=0, column=0, padx=(10, 0))
        self.entry = ctk.CTkEntry(search, font=ui.font(14), height=38, border_width=0, fg_color="transparent", placeholder_text="Cari nilai, misal 9 nov atau november" if self.is_dates else "Cari nilai")
        self.entry.grid(row=0, column=1, sticky="ew", pady=2)
        self.entry.bind("<KeyRelease>", lambda _: self.on_search())
        self.entry.bind("<Return>", lambda _: self.apply())
        self.hit_label = ctk.CTkLabel(search, text="", font=ui.font(12, "bold"), text_color=ui.ACCENT, fg_color=ui.ACCENT_SOFT, corner_radius=6, height=24)
        self.clear_button = ctk.CTkButton(search, text="×", width=28, height=28, fg_color="transparent", hover_color=ui.ACCENT_SOFT, text_color=ui.MUTED, font=ui.font(16), command=self.clear_search)
        self.quick = ctk.CTkFrame(main, fg_color="transparent")
        self.quick.grid(row=1, column=0, sticky="ew", pady=(8, 6))
        self.build_quick()
        self.images = self.make_marks(round(18 * scale))
        dark = is_dark()
        style = ttk.Style(self)
        style.configure(
            "Filter.Treeview", rowheight=round(36 * scale), font=(family, -round(14 * scale)), borderwidth=0,
            background="#1A2030" if dark else "#FFFFFF", fieldbackground="#1A2030" if dark else "#FFFFFF", foreground="#EEF0F5" if dark else "#16202E",
        )
        style.layout("Filter.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        style.map("Filter.Treeview", background=[("selected", "#22304A" if dark else "#F1E6D7")], foreground=[("selected", "#EEF0F5" if dark else "#0F2C59")])
        holder = ctk.CTkFrame(main, fg_color="transparent")
        holder.grid(row=2, column=0, sticky="nsew")
        holder.grid_columnconfigure(0, weight=1)
        holder.grid_rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(holder, columns=("badge", "count"), show="tree", style="Filter.Treeview", selectmode="none")
        self.tree.column("#0", stretch=True, width=260)
        self.tree.column("badge", anchor="center", width=round(84 * scale), stretch=False)
        self.tree.column("count", anchor="e", width=round(62 * scale), stretch=False)
        self.tree.tag_configure("group", font=(family, -round(14 * scale), "bold"))
        self.tree.tag_configure("picked", background="#22304A" if dark else "#F6EEE3")
        self.tree.tag_configure("some", foreground="#FDBA74" if dark else "#9A3412")
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll = ctk.CTkScrollbar(holder, command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<Button-1>", self.on_click)
        tools = ctk.CTkFrame(main, fg_color="transparent")
        tools.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        link = {"fg_color": "transparent", "hover_color": ui.ACCENT_SOFT, "text_color": ui.ACCENT, "font": ui.font(13, "bold"), "height": 32, "width": 80}
        ctk.CTkButton(tools, text="Pilih semua", command=self.select_all, **link).pack(side="left")
        ctk.CTkButton(tools, text="Balik pilihan", command=self.invert, **link).pack(side="left", padx=2)
        self.count_button = ctk.CTkButton(tools, text="Jumlah", command=lambda: self.set_order(True), fg_color="transparent", hover_color=ui.ACCENT_SOFT, text_color=ui.TEXT, font=ui.font(12), height=30, width=56)
        self.name_button = ctk.CTkButton(tools, text="Nama", command=lambda: self.set_order(False), fg_color="transparent", hover_color=ui.ACCENT_SOFT, text_color=ui.TEXT, font=ui.font(12), height=30, width=50)
        self.mark_button = ctk.CTkButton(tools, text="", command=self.mark_hits, fg_color="#0F2C59", hover_color="#0A1F40", text_color="#FFFFFF", font=ui.font(13, "bold"), height=34, corner_radius=17)
        self.count_button.pack(side="right")
        self.name_button.pack(side="right", padx=2)
        self.bind("<Escape>", lambda _: self.destroy())
        self.rebuild()
        self.after(100, self.bring_front)

    def bring_front(self):
        self.lift()
        self.entry.focus_set()

    def build_quick(self):
        """Tombol cepat: bulan-bulan terakhir yang ada datanya."""
        if not self.is_dates:
            self.quick.grid_remove()
            return
        months = sorted({v[:7] for v in self.values if DATE_VALUE.match(v)}, reverse=True)[:4]
        for month in months:
            name = MONTHS[int(month[5:7]) - 1]
            label = name if month[:4] == str(datetime.date.today().year) or len({m[:4] for m in months}) == 1 else f"{name} {month[:4]}"
            ctk.CTkButton(
                self.quick, text=label, height=28, width=40, corner_radius=14, font=self.ui.font(12), fg_color=self.ui.CARD,
                border_width=1, border_color=self.ui.BORDER, hover_color=self.ui.ACCENT_SOFT, text_color=self.ui.TEXT,
                command=lambda m=month: self.quick_search(m),
            ).pack(side="left", padx=(0, 6))

    def quick_search(self, month):
        self.entry.delete(0, "end")
        self.entry.insert(0, f"{month}-")
        self.on_search()

    def sorted_values(self, values):
        if self.order_by_count:
            return sorted(values, key=lambda v: (-self.counts[v], v))
        if self.target.kind == "num":
            def number(value):
                try:
                    return (0, float(value.replace(",", "")), value)
                except ValueError:
                    return (1, 0, value)
            return sorted(values, key=number)
        return sorted(values)

    def add_node(self, parent, iid, label, values, group):
        self.nodes[iid] = (label, list(values), group)
        self.tree.insert(parent, "end", iid=iid, text=f"  {label}", image=self.node_image(iid), values=self.node_values(iid), tags=self.node_tags(iid))

    def node_state(self, values):
        count = sum(1 for v in values if v in self.chosen)
        return "off" if count == 0 else ("on" if count == len(values) else "some")

    def node_image(self, iid):
        return self.images[self.node_state(self.nodes[iid][1])]

    def node_values(self, iid):
        _, values, group = self.nodes[iid]
        rows = f"{sum(self.counts[v] for v in values):,}"
        picked = sum(1 for v in values if v in self.chosen)
        badge = f"{picked} dari {len(values)}" if group and picked else ""
        return (badge, rows)

    def node_tags(self, iid):
        _, values, group = self.nodes[iid]
        state = self.node_state(values)
        tags = ["group"] if group else []
        if state != "off":
            tags.append("picked")
        if group and state == "some":
            tags.append("some")
        return tuple(tags)

    @staticmethod
    def make_marks(size):
        """Kotak centang digambar sendiri supaya tampil sama di semua komputer."""
        scale = 4
        full = size * scale
        images = {}
        for state in ("on", "off", "some"):
            image = Image.new("RGBA", (full, full), (0, 0, 0, 0))
            draw = ImageDraw.Draw(image)
            filled = state != "off"
            draw.rounded_rectangle(
                (scale, scale, full - scale, full - scale), radius=full // 4, fill="#0F2C59" if filled else "#FFFFFF",
                outline="#0F2C59" if filled else "#C9C3B8", width=max(scale, full // 12),
            )
            stroke = max(scale, full // 8)
            if state == "on":
                draw.line([(full * 0.25, full * 0.52), (full * 0.43, full * 0.7), (full * 0.76, full * 0.3)], fill="#FFFFFF", width=stroke, joint="curve")
            elif state == "some":
                draw.line([(full * 0.27, full * 0.5), (full * 0.73, full * 0.5)], fill="#FFFFFF", width=stroke)
            images[state] = ImageTk.PhotoImage(image.resize((size, size), Image.LANCZOS))
        return images

    def rebuild(self):
        opened = {iid for iid in self.nodes if self.tree.exists(iid) and self.tree.item(iid, "open")}
        searching = bool(self.entry.get().strip())
        self.tree.delete(*self.tree.get_children())
        self.nodes = {}
        dated = {}
        others = []
        for value in self.shown:
            match = DATE_VALUE.match(value)
            if match:
                dated.setdefault(match.group(1), {}).setdefault(match.group(2), []).append((match.group(3), value))
            else:
                others.append(value)
        total = lambda group: sum(self.counts[v] for v in group)
        years = sorted(dated, key=lambda y: -total([v for m in dated[y].values() for _, v in m]) if self.order_by_count else y, reverse=not self.order_by_count)
        selected = self.chosen if self.chosen != self.values else set()
        for year in years:
            year_values = [v for days in dated[year].values() for _, v in days]
            iid = f"y{year}"
            self.add_node("", iid, year, year_values, True)
            months = sorted(dated[year], key=lambda m: -total([v for _, v in dated[year][m]]) if self.order_by_count else m)
            for month in months:
                days = dated[year][month]
                month_id = f"m{year}{month}"
                self.add_node(iid, month_id, MONTHS[int(month) - 1], [v for _, v in days], True)
                ordered = sorted(days, key=lambda d: (-self.counts[d[1]], d[0]) if self.order_by_count else d[0])
                for day, value in ordered:
                    self.add_node(month_id, f"d{value}", f"{int(day)} {MONTHS_SHORT[int(month) - 1]}", [value], False)
                wanted = month_id in opened or searching or any(v in selected for _, v in days)
                self.tree.item(month_id, open=wanted)
            self.tree.item(iid, open=iid in opened or searching or year == years[0] or any(v in selected for v in year_values))
        for index, value in enumerate(self.sorted_values(others)):
            self.add_node("", f"v{index}", value, [value], False)
        self.update_info()

    def update_info(self):
        chosen = self.chosen & self.shown
        rows = sum(self.counts[v] for v in chosen)
        self.big.configure(text=f"{rows:,}")
        share = rows / self.total_rows if self.total_rows else 0
        self.sub.configure(text=f"{len(chosen):,} nilai dari {len(self.values):,} · {share * 100:.1f}% dari {self.total_rows:,}")
        self.bar.set(share)
        active = self.ui.ACCENT_SOFT
        self.name_button.configure(fg_color="transparent" if self.order_by_count else active)
        self.count_button.configure(fg_color=active if self.order_by_count else "transparent")
        searching = bool(self.entry.get().strip())
        self.search_box.configure(border_color=self.ui.ACCENT if searching else self.ui.BORDER)
        if searching:
            self.hit_label.configure(text=f"  {len(self.shown):,} hasil  ")
            self.hit_label.grid(row=0, column=2, padx=4)
            self.clear_button.grid(row=0, column=3, padx=(0, 6))
            self.mark_button.configure(text=f"  Centang {len(self.shown):,} hasil  ")
            self.mark_button.pack(side="right", padx=8)
        else:
            self.hit_label.grid_remove()
            self.clear_button.grid_remove()
            self.mark_button.pack_forget()
        self.update_chips(chosen)

    def update_chips(self, chosen):
        for child in self.chips.winfo_children():
            child.destroy()
        items = []
        if not chosen:
            items.append(("Belum ada yang dipilih", False))
        elif chosen == self.values:
            items.append(("Semua nilai", True))
        elif self.is_dates:
            months = {}
            for value in chosen:
                if DATE_VALUE.match(value):
                    months.setdefault(value[:7], set()).add(value)
            for month in sorted(months):
                full = {v for v in self.values if v.startswith(month)}
                name = f"{MONTHS[int(month[5:7]) - 1]} {month[:4]}"
                items.append((f"{name}: semua tanggal" if months[month] == full else f"{name}: {len(months[month])} tanggal", months[month] == full))
            rest = [v for v in chosen if not DATE_VALUE.match(v)]
            if rest:
                items.append((f"{len(rest)} nilai lain", False))
        else:
            for value in sorted(chosen)[:6]:
                items.append((value if len(value) <= 26 else value[:25] + "…", False))
            if len(chosen) > 6:
                items.append((f"+{len(chosen) - 6} lainnya", False))
        for text, strong in items[:7]:
            ctk.CTkLabel(
                self.chips, text=f"  {text}  ", font=self.ui.font(12, "bold"), height=26, corner_radius=13,
                fg_color=self.ui.ACCENT if strong else self.ui.CARD, text_color="#FFFFFF" if strong else self.ui.ACCENT, anchor="w",
            ).pack(anchor="w", pady=2)
        if len(items) > 7:
            ctk.CTkLabel(self.chips, text=f"+{len(items) - 7} lagi", font=self.ui.font(12), text_color=self.ui.MUTED).pack(anchor="w")

    def refresh_marks(self, iids=None):
        for iid in iids if iids is not None else self.nodes:
            self.tree.item(iid, image=self.node_image(iid), values=self.node_values(iid), tags=self.node_tags(iid))
        self.update_info()

    def on_click(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return None
        if "indicator" in str(self.tree.identify_element(event.x, event.y)):
            return None
        values = self.nodes[iid][1]
        if self.node_state(values) == "on":
            self.chosen.difference_update(values)
        else:
            self.chosen.update(values)
        branch = [iid]
        parent = self.tree.parent(iid)
        while parent:
            branch.append(parent)
            parent = self.tree.parent(parent)
        stack = list(self.tree.get_children(iid))
        while stack:
            child = stack.pop()
            branch.append(child)
            stack.extend(self.tree.get_children(child))
        self.refresh_marks(branch)
        return "break"

    def on_search(self):
        words = self.entry.get().lower().split()
        if words == self.last_words:
            return
        self.last_words = words
        if words:
            self.shown = {v for v in self.values if all(w in self.search_text(v) for w in words)}
        else:
            self.shown = set(self.values)
        self.rebuild()

    def clear_search(self):
        self.entry.delete(0, "end")
        self.on_search()

    @staticmethod
    def search_text(value):
        match = DATE_VALUE.match(value)
        if match:
            month = int(match.group(2)) - 1
            day = int(match.group(3))
            return f"{value} {day} {MONTHS[month]} {MONTHS_SHORT[month]} {day} {MONTHS_SHORT[month]}".lower()
        return value.lower()

    def mark_hits(self):
        """Seperti di WPS: hasil pencarian menjadi pilihan."""
        self.chosen = set(self.shown)
        self.refresh_marks()

    def select_all(self):
        self.chosen.update(self.shown)
        self.refresh_marks()

    def invert(self):
        for value in self.shown:
            (self.chosen.discard if value in self.chosen else self.chosen.add)(value)
        self.refresh_marks()

    def reset(self):
        self.entry.delete(0, "end")
        self.last_words = None
        self.chosen = set(self.values)
        self.on_search()

    def set_order(self, by_count):
        self.order_by_count = by_count
        self.rebuild()

    def result(self):
        return set(self.chosen)

    def apply(self):
        chosen = self.result()
        if not chosen:
            messagebox.showinfo("Dockit", "Pilih minimal satu nilai.", parent=self)
            return
        self.page.set_filter(self.target.key, set() if chosen >= self.values else chosen)
        self.destroy()

    def sort(self, reverse):
        """Hanya urutan tampilan. Urutan data di file tetap sesuai paste."""
        chosen = self.result()
        if chosen:
            self.page.filters.pop(self.target.key, None)
            if not chosen >= self.values:
                self.page.filters[self.target.key] = chosen
        self.page.set_sort(self.target, reverse)
        self.destroy()


class WeightDialog(ctk.CTkToplevel):
    """Pratinjau update NW/GW dari file timbang, dengan pilihan tanggal timbang."""

    def __init__(self, manager, name, weights):
        ui = manager.ui
        super().__init__(ui, fg_color=ui.CARD)
        self.manager, self.weights = manager, weights
        self.title("Update NW/GW dari file timbang")
        self.geometry("760x620")
        self.transient(ui)
        store = manager.store
        known = store.item_skus()
        missing = [s for s in weights if s not in known]
        every = store.weight_fixes(weights)
        dates = {}
        changed = {}
        for fix in every:
            changed.setdefault(fix["sku"], fix["date"])
        for day in changed.values():
            dates[day] = dates.get(day, 0) + 1
        ordered = sorted((d for d in dates if d is not None), reverse=True)
        self.options = {}
        running = 0
        for day in ordered:
            running += dates[day]
            self.options[f"Sejak {day.isoformat()}  ({running:,} SKU)"] = day
        self.options[f"Semua tanggal  ({len(changed):,} SKU)"] = None
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        ctk.CTkLabel(self, text="Update NW/GW dari file timbang", font=ui.font(20, "bold"), text_color=ui.TEXT, anchor="w").grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 0))
        info = (f"{name}: {len(weights):,} SKU. Per SKU dipakai timbangan terbaru, baris dengan berat paling kecil (NW dan GW di baris yang sama). "
                f"{len(weights) - len(missing):,} SKU ada di data, {len(missing):,} tidak ada sehingga dilewati.")
        ctk.CTkLabel(self, text=info, font=ui.font(12), text_color=ui.MUTED, anchor="w", justify="left", wraplength=710).grid(row=1, column=0, sticky="ew", padx=20, pady=(4, 10))
        pick = ctk.CTkFrame(self, fg_color="transparent")
        pick.grid(row=2, column=0, sticky="ew", padx=20)
        ctk.CTkLabel(pick, text="Update SKU yang ditimbang", font=ui.font(13, "bold"), text_color=ui.TEXT).pack(side="left")
        names = list(self.options)
        self.choice = ctk.StringVar(value=names[0] if names else "")
        ctk.CTkOptionMenu(pick, values=names or ["Tidak ada perubahan"], variable=self.choice, command=lambda _: self.show(), width=300,
                          fg_color=ui.NEUTRAL, button_color=ui.ACCENT, button_hover_color=ui.ACCENT_HOVER, text_color=ui.TEXT, font=ui.font(13)).pack(side="left", padx=10)
        self.preview = ctk.CTkTextbox(self, font=ui.font(12), fg_color=ui.NEUTRAL, text_color=ui.TEXT, corner_radius=9, wrap="none")
        self.preview.grid(row=3, column=0, sticky="nsew", padx=20, pady=12)
        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 18))
        self.count = ctk.CTkLabel(actions, text="", font=ui.font(13), text_color=ui.MUTED)
        self.count.pack(side="left")
        self.apply_button = ui.button(actions, "Update & Simpan", self.apply, primary=True, height=40, plain=True)
        self.apply_button.pack(side="right")
        ui.button(actions, "Batal", self.destroy, height=40, width=80).pack(side="right", padx=8)
        self.fixes = []
        self.show()
        self.after(100, self.lift)

    def show(self):
        since = self.options.get(self.choice.get())
        self.fixes = self.manager.store.weight_fixes(self.weights, since)
        by_sku = {}
        for fix in self.fixes:
            by_sku.setdefault(fix["sku"], {})[fix["key"]] = fix
        lines = []
        for sku, pair in by_sku.items():
            parts = []
            for key, label in (("E", "NW"), ("F", "GW")):
                fix = pair.get(key)
                if fix:
                    old = "kosong" if not isinstance(fix["old"], (int, float)) else f"{fix['old']:g}"
                    parts.append(f"{label} {old} → {fix['new']:g}")
            day = next(iter(pair.values()))["date"]
            lines.append(f"{sku}   {' · '.join(parts)}   ({day.isoformat() if day else 'tanpa tanggal'})")
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", "\n".join(lines) if lines else "Tidak ada NW/GW yang berbeda untuk pilihan ini.")
        self.preview.configure(state="disabled")
        self.count.configure(text=f"{len(by_sku):,} SKU akan diupdate. Backup dibuat otomatis.")
        self.apply_button.configure(state="normal" if self.fixes else "disabled")

    def apply(self):
        if not self.fixes:
            return
        self.manager.store.apply_fixes(self.fixes)
        data.clear_estimates({f["sku"] for f in self.fixes})
        self.destroy()
        self.manager.save("items")


class EstimateDialog(ctk.CTkToplevel):
    """Perkiraan NW/GW untuk SKU yang belum punya berat, langsung dipakai dengan tanda perkiraan."""

    def __init__(self, manager):
        ui = manager.ui
        if manager.running:
            messagebox.showinfo("Dockit", "Proses masih berjalan. Coba lagi sebentar.")
            return
        super().__init__(ui, fg_color=ui.CARD)
        self.manager = manager
        self.title("Perkirakan NW/GW")
        self.geometry("820x620")
        self.transient(ui)
        self.all_fixes, self.all_marks, skipped = manager.store.estimate_weights()
        high = len({f["sku"] for f in self.all_fixes if f["level"] == "Tinggi"})
        mid = len({f["sku"] for f in self.all_fixes if f["level"] == "Sedang"})
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        ctk.CTkLabel(self, text="Perkirakan NW/GW", font=ui.font(20, "bold"), text_color=ui.TEXT, anchor="w").grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 0))
        shown = ", ".join(skipped[:6]) + (f", dan {len(skipped) - 6} lainnya" if len(skipped) > 6 else "")
        info = ("Dihitung dari berat SKU yang sudah ada. Tinggi: dari SKU satu model dengan ukuran lain (uji: rata-rata meleset 3.5%). "
                "Sedang: dari rata-rata kategori foam/spring/hybrid (uji: rata-rata meleset 6.5%). "
                "Hasilnya langsung dipakai di Buat Excel dengan tanda perkiraan, sampai diganti hasil timbang atau diedit manual."
                + (f"\n\n{len(skipped)} SKU dilewati karena type-nya tidak dikenali dari SKU: {shown}." if skipped else ""))
        ctk.CTkLabel(self, text=info, font=ui.font(12), text_color=ui.MUTED, anchor="w", justify="left", wraplength=770).grid(row=1, column=0, sticky="ew", padx=20, pady=(4, 10))
        pick = ctk.CTkFrame(self, fg_color="transparent")
        pick.grid(row=2, column=0, sticky="ew", padx=20)
        ctk.CTkLabel(pick, text="Pakai perkiraan", font=ui.font(13, "bold"), text_color=ui.TEXT).pack(side="left")
        self.options = {f"Tinggi dan sedang  ({high + mid:,} SKU)": ("Tinggi", "Sedang"), f"Keyakinan tinggi saja  ({high:,} SKU)": ("Tinggi",)}
        self.choice = ctk.StringVar(value=next(iter(self.options)))
        ctk.CTkOptionMenu(pick, values=list(self.options), variable=self.choice, command=lambda _: self.show(), width=280,
                          fg_color=ui.NEUTRAL, button_color=ui.ACCENT, button_hover_color=ui.ACCENT_HOVER, text_color=ui.TEXT, font=ui.font(13)).pack(side="left", padx=10)
        self.preview = ctk.CTkTextbox(self, font=ui.font(12), fg_color=ui.NEUTRAL, text_color=ui.TEXT, corner_radius=9, wrap="none")
        self.preview.grid(row=3, column=0, sticky="nsew", padx=20, pady=12)
        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 18))
        self.count = ctk.CTkLabel(actions, text="", font=ui.font(13), text_color=ui.MUTED)
        self.count.pack(side="left")
        self.apply_button = ui.button(actions, "Pakai & Simpan", self.apply, primary=True, height=40, plain=True)
        self.apply_button.pack(side="right")
        ui.button(actions, "Batal", self.destroy, height=40, width=80).pack(side="right", padx=8)
        self.show()
        self.after(100, self.lift)

    def show(self):
        levels = self.options.get(self.choice.get(), ("Tinggi",))
        self.fixes = [f for f in self.all_fixes if f["level"] in levels]
        lines = []
        for sku, mark in self.all_marks.items():
            if mark["level"] in levels:
                lines.append(f"{sku}   NW {mark['nw']:g} · GW {mark['gw']:g}   {mark['level']}, {mark['source']}")
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", "\n".join(lines) if lines else "Semua SKU yang dikenali sudah punya NW/GW.")
        self.preview.configure(state="disabled")
        self.count.configure(text=f"{len(lines):,} SKU akan diisi perkiraan. Backup dibuat otomatis.")
        self.apply_button.configure(state="normal" if self.fixes else "disabled")

    def apply(self):
        if not self.fixes:
            return
        levels = self.options.get(self.choice.get(), ("Tinggi",))
        marks = dict(data.estimate_marks())
        marks.update({sku: mark for sku, mark in self.all_marks.items() if mark["level"] in levels})
        data.save_estimate_marks(marks)
        self.manager.store.apply_fixes(self.fixes)
        self.destroy()
        self.manager.save("items")


class DashboardPage(ctk.CTkFrame):
    """Jumlah PO atau DI per ETD: kalender setahun dan grafik per bulan serta per tanggal."""

    ETD_KEY = {"DI": "L", "PO": "D"}
    TABLE = {"DI": "master", "PO": "po_master"}

    def __init__(self, parent, manager, profile):
        super().__init__(parent, fg_color="transparent")
        self.manager, self.profile = manager, profile
        ui = manager.ui
        self.counts = {}
        self.labels = {}
        self.selected = None
        self.month = None
        self.year = max(datetime.date.today().year, 2026)
        self.hits = {}
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        ctk.CTkLabel(top, text=f"Dasbor {profile.name}", font=ui.font(26, "bold"), text_color=ui.TEXT).pack(side="left")
        ctk.CTkLabel(top, text=f"Berdasarkan ETD {self.year}. Klik tanggal atau bulan.", font=ui.font(12), text_color=ui.MUTED).pack(side="left", padx=14)
        self.open_button = ui.button(top, f"Lihat di Master {profile.name}", self.open_master, primary=True, width=170, height=42)
        self.open_button.pack(side="right")
        cards = ctk.CTkFrame(self, fg_color="transparent")
        cards.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        white = (ui.CARD, ui.BORDER, ui.MUTED, ui.TEXT)
        soft = (ui.ACCENT_SOFT, ui.ACCENT, ui.ACCENT, ui.ACCENT)
        self.cards = []
        for index, colors in enumerate((soft, white, white, white)):
            card = SummaryCard(cards, ui, "", lambda: None, colors, icon=("calendar-days", "calendar-days", "chart-column-increasing", "chart-column-increasing")[index])
            cards.grid_columnconfigure(index, weight=1, uniform="card")
            card.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 12, 0))
            self.cards.append(card)
        self.calendar_box = ctk.CTkFrame(self, fg_color=ui.CARD, border_color=ui.BORDER, border_width=1, corner_radius=12)
        self.calendar_box.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        head = ctk.CTkFrame(self.calendar_box, fg_color="transparent")
        head.pack(fill="x", padx=16, pady=(10, 2))
        ctk.CTkLabel(head, text=f"Kalender ETD, jumlah {profile.name} per tanggal", font=ui.font(15, "bold"), text_color=ui.TEXT).pack(side="left")
        self.scale = self._get_widget_scaling()
        self.calendar = tk.Canvas(self.calendar_box, highlightthickness=0, bd=0, height=round(166 * self.scale))
        self.calendar.pack(fill="x", padx=16, pady=(0, 10))
        charts = ctk.CTkFrame(self, fg_color="transparent")
        charts.grid(row=3, column=0, sticky="nsew")
        charts.grid_rowconfigure(0, weight=1)
        charts.grid_columnconfigure(0, weight=2, uniform="chart")
        charts.grid_columnconfigure(1, weight=3, uniform="chart")
        self.month_title, self.month_chart = self.chart_card(charts, 0)
        self.day_title, self.day_chart = self.chart_card(charts, 1)
        family = "Segoe UI" if os.name == "nt" else "DejaVu Sans"
        self.font = tkfont.Font(family=family, size=-round(12 * self.scale))
        self.bold = tkfont.Font(family=family, size=-round(12 * self.scale), weight="bold")
        self.tiny = tkfont.Font(family=family, size=-round(10 * self.scale), weight="bold")
        for canvas in (self.calendar, self.month_chart, self.day_chart):
            canvas.bind("<Configure>", lambda _: self.draw())
            canvas.bind("<Button-1>", self.on_click)
            canvas.bind("<Motion>", self.on_motion)

    def chart_card(self, parent, column):
        ui = self.manager.ui
        box = ctk.CTkFrame(parent, fg_color=ui.CARD, border_color=ui.BORDER, border_width=1, corner_radius=12)
        box.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 12, 0), pady=(0, 8))
        title = ctk.CTkLabel(box, text="", font=ui.font(15, "bold"), text_color=ui.TEXT, anchor="w")
        title.pack(fill="x", padx=16, pady=(8, 0))
        canvas = tk.Canvas(box, highlightthickness=0, bd=0, height=10)
        canvas.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        return title, canvas

    def px(self, value):
        return round(value * self.scale)

    def load(self):
        store = self.manager.store
        counts, labels, seen = {}, {}, set()
        key = self.ETD_KEY[self.profile.name]
        column = next((c for c in store_columns(self.profile) if c.key == key), None)
        kind = column.kind if column else "date"
        for row in store.rows_of(self.TABLE[self.profile.name]):
            code = data.unique_key(row.get("A"))
            if not code or code in seen:
                continue
            seen.add(code)
            text = data.format_value(row.get(key), kind)
            match = DATE_VALUE.match(text)
            if not match or int(match.group(1)) != self.year:
                continue
            day = datetime.date(self.year, int(match.group(2)), int(match.group(3)))
            counts[day] = counts.get(day, 0) + 1
            labels.setdefault(day, set()).add(text)
        self.counts, self.labels = counts, labels

    def refresh(self):
        self.load()
        if self.selected not in self.counts:
            today = datetime.date.today()
            ahead = sorted(d for d in self.counts if d >= today)
            self.selected = ahead[0] if ahead else (max(self.counts) if self.counts else None)
            self.month = self.selected.month if self.selected else today.month
        self.draw()

    def monthly(self):
        totals = [0] * 12
        for day, value in self.counts.items():
            totals[day.month - 1] += value
        return totals

    def draw(self):
        name = self.profile.name
        totals = self.monthly()
        selected = self.selected
        count = self.counts.get(selected, 0) if selected else 0
        busiest = max(self.counts, key=self.counts.get) if self.counts else None
        month_name = MONTHS[(self.month or 1) - 1]
        texts = [
            (f"ETD {selected.day} {MONTHS_SHORT[selected.month - 1]} {selected.year}" if selected else "ETD", count),
            (f"Bulan {month_name}", totals[(self.month or 1) - 1]),
            (f"Total ETD {self.year}", sum(totals)),
            (f"ETD terpadat: {busiest.day} {MONTHS_SHORT[busiest.month - 1]}" if busiest else "ETD terpadat", self.counts.get(busiest, 0)),
        ]
        for card, (title, value) in zip(self.cards, texts):
            card.title_label.configure(text=title)
            card.value.configure(text=f"{value:,} {name}")
        self.month_title.configure(text=f"Jumlah {name} per bulan")
        self.day_title.configure(text=f"Jumlah {name} per ETD, {month_name} {self.year}")
        self.open_button.configure(state="normal" if count else "disabled")
        self.draw_calendar()
        self.draw_bars(self.month_chart, [(MONTHS_SHORT[i], v, i + 1) for i, v in enumerate(totals)], self.month)
        days = sorted(d for d in self.counts if d.month == self.month)
        self.draw_bars(self.day_chart, [(str(d.day), self.counts[d], d) for d in days], selected)

    @staticmethod
    def level(value):
        dark = is_dark()
        if not value:
            return ("#283044", "#A0A8B8") if dark else ("#F3ECE2", "#5E6675")
        if value <= 5:
            return "#EADBC8", "#0A1F40"
        if value <= 15:
            return "#D2B48E", "#0A1F40"
        if value <= 25:
            return "#3B5F9E", "#FFFFFF"
        return "#0A1F40", "#FFFFFF"

    def draw_calendar(self):
        c = self.calendar
        c.delete("all")
        c.configure(bg=pick_color(self.manager.ui.CARD))
        hits = []
        first = datetime.date(self.year, 1, 1)
        offset = first.weekday()
        days = (datetime.date(self.year + 1, 1, 1) - first).days
        weeks = (offset + days - 1) // 7 + 1
        left, top = self.px(34), self.px(20)
        step = max(self.px(14), min(self.px(22), (max(1, c.winfo_width()) - left) // weeks))
        size = step - max(1, self.px(1))
        needed = top + 7 * step
        if int(c.cget("height")) != needed:
            c.configure(height=needed)
        muted = pick_color(self.manager.ui.MUTED)
        for row, label in ((0, "Sen"), (2, "Rab"), (4, "Jum"), (6, "Min")):
            c.create_text(0, top + row * step + size // 2, text=label, anchor="w", fill=muted, font=self.font)
        for number in range(days):
            day = first + datetime.timedelta(days=number)
            week, weekday = (offset + number) // 7, (offset + number) % 7
            x, y = left + week * step, top + weekday * step
            if day.day == 1:
                c.create_text(x, top - self.px(10), text=MONTHS_SHORT[day.month - 1], anchor="w", fill=muted, font=self.font)
            value = self.counts.get(day, 0)
            bg, fg = self.level(value)
            active = day == self.selected
            c.create_polygon(rounded_points(x, y, x + size, y + size, self.px(3)), smooth=True, fill=bg, outline="#C77700" if active else "", width=2 if active else 1)
            if value:
                c.create_text(x + size // 2, y + size // 2, text=str(value), fill=fg, font=self.tiny)
            hits.append((x, y, x + size, y + size, day))
        self.hits[c] = hits

    def draw_bars(self, canvas, items, active):
        canvas.delete("all")
        canvas.configure(bg=pick_color(self.manager.ui.CARD))
        hits = []
        width, height = max(1, canvas.winfo_width()), max(1, canvas.winfo_height())
        muted, text = pick_color(self.manager.ui.MUTED), pick_color(self.manager.ui.TEXT)
        if not items:
            canvas.create_text(width // 2, height // 2, text="Belum ada ETD di bulan ini.", fill=muted, font=self.font)
            self.hits[canvas] = hits
            return
        foot, head = self.px(22), self.px(22)
        top_value = max(value for _, value, _ in items) or 1
        gap = self.px(4)
        slot = min(self.px(56), max(self.px(10), width // len(items)))
        for index, (label, value, key) in enumerate(items):
            x = index * slot
            bar = round((height - foot - head) * value / top_value) if value else 0
            y = height - foot - max(bar, self.px(3) if value else 0)
            is_active = key == active
            if value:
                canvas.create_rectangle(x + gap // 2, y, x + slot - gap // 2, height - foot, fill="#0F2C59" if is_active else "#D2B48E", outline="")
                canvas.create_text(x + slot // 2, y - self.px(4), text=str(value), anchor="s", fill=text, font=self.bold if is_active else self.font)
            canvas.create_text(x + slot // 2, height - foot // 2, text=label, fill=text if is_active else muted, font=self.bold if is_active else self.font)
            hits.append((x, 0, x + slot, height, key))
        self.hits[canvas] = hits

    def hit(self, event):
        for x1, y1, x2, y2, key in self.hits.get(event.widget, ()):
            if x1 <= event.x <= x2 and y1 <= event.y <= y2:
                return key
        return None

    def on_motion(self, event):
        event.widget.configure(cursor="hand2" if self.hit(event) is not None else "")

    def on_click(self, event):
        key = self.hit(event)
        if key is None:
            return
        if isinstance(key, datetime.date):
            self.selected = key
            self.month = key.month
        else:
            self.month = key
        self.draw()

    def open_master(self):
        labels = self.labels.get(self.selected)
        if not labels:
            return
        name = PAGE_MASTER_PO if self.profile.name == "PO" else PAGE_MASTER
        self.manager.ui.show_page(name)
        self.manager.pages[name].set_filter(self.ETD_KEY[self.profile.name], set(labels))


def store_columns(profile):
    return data.MASTER_PO_COLUMNS if profile.name == "PO" else data.MASTER_COLUMNS


class SummaryCard(ctk.CTkFrame):
    """Kartu angka yang bisa diklik untuk memfilter tabel."""

    def __init__(self, parent, ui, title, command, colors, icon=None):
        super().__init__(parent, fg_color=colors[0], border_color=colors[1], border_width=1, corner_radius=12, height=78)
        self.base_border = colors[1]
        self.ui = ui
        icon = icon or icons.card_icon(title)
        widgets = [self]
        if icon:
            self.icon_label = ctk.CTkLabel(self, text="", width=40, height=40, corner_radius=10, fg_color=("#FFFFFF", "#22304A"))
            self.icon_label.pack(side="left", padx=(14, 0), pady=12)
            self.icon_animator = icons.Animator(self.icon_label, icon, colors[2], 22, hover_widgets=(self,))
            widgets.append(self.icon_label)
        self.title_label = ctk.CTkLabel(self, text=title, font=ui.font(13), text_color=colors[2], anchor="w", height=20)
        self.title_label.pack(fill="x", padx=16 if not icon else 12, pady=(10, 0))
        self.value = ctk.CTkLabel(self, text="0", font=ui.font(24, "bold"), text_color=colors[3], anchor="w", height=32)
        self.value.pack(fill="x", padx=16 if not icon else 12, pady=(0, 8))
        widgets += [self.title_label, self.value]
        for widget in widgets:
            widget.bind("<Button-1>", lambda _: command())

    def show(self, number, active):
        self.value.configure(text=f"{number:,}")
        self.configure(border_color=self.ui.ACCENT if active else self.base_border, border_width=2 if active else 1)


SEARCH = {
    "master": ("Cari DI, BL, kontainer, invoice", [("A", "DI"), ("N", "BL"), ("G", "CN"), ("E", "INV")], "K"),
    "po_master": ("Cari PO, SI, BL, kontainer", [("A", "PO"), ("B", "SI"), ("F", "BL"), ("H", "CN")], "D"),
    "goods": ("Cari SKU, nama barang, atau HS code", [("A", "SKU"), ("B", "Nama"), ("C", "HS")], None),
    "weight": ("Cari SKU atau nama barang", [("A", "SKU"), ("B", "Nama")], None),
    "prices": ("Cari DI/PO atau SKU", [("G", "DI/PO"), ("H", "SKU")], None),
}


class DataPage(ctk.CTkFrame):
    SPECIAL = {"master": ("Double", "Double"), "po_master": ("Double", "Double"), "prices": ("Harga ganda", "harga berbeda")}

    def __init__(self, parent, manager, title, hint, columns, table, flag, paste_keys, extra_modes=(), profile=None):
        super().__init__(parent, fg_color="transparent")
        self.manager, self.title_text, self.columns, self.table = manager, title, columns, table
        self.flag, self.paste_keys = flag, paste_keys
        self.profile = profile
        ui = manager.ui
        self.search = ctk.StringVar()
        self.mode = ctk.StringVar(value=FILTER_ALL)
        self.visible = []
        self.search_job = None
        self.filters = {}
        self.sort_by = None
        self.filter_dialog = None
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        ctk.CTkLabel(top, text=title, font=ui.font(26, "bold"), text_color=ui.TEXT).pack(side="left")
        ctk.CTkLabel(top, text=hint, font=ui.font(12), text_color=ui.MUTED, wraplength=420, justify="left").pack(side="left", padx=14)
        self.save_button = ui.button(top, "Simpan", self.save, primary=True, width=120, height=42)
        self.save_button.pack(side="right")
        self.menu_button = ui.button(top, "Menu  ▾", self.open_menu, width=96, height=42, plain=True)
        self.menu_button.pack(side="right", padx=(8, 0))
        self.menu = motion.StaggerMenu(self.menu_button, self.menu_items(), width=300)
        self.refresh_button = ctk.CTkButton(top, text="", width=42, height=42, corner_radius=9, fg_color=ui.CARD, border_width=1,
                                            border_color=ui.BORDER, hover_color=ui.NEUTRAL, command=self.reload_spin)
        self.refresh_button.pack(side="right", padx=(8, 0))
        self.spin = motion.SyncSpin(self.refresh_button, lambda: self.manager.running)
        cards = ctk.CTkFrame(self, fg_color="transparent")
        cards.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        white = (ui.CARD, ui.BORDER, ui.MUTED, ui.TEXT)
        special = "CBM tidak tally" if extra_modes else self.SPECIAL.get(table, ("Typo", ""))[0]
        self.special_text = {"CBM tidak tally": "tally", "Typo": "cocok"}.get(special, self.SPECIAL.get(table, ("", ""))[1])
        self.special_mode = FILTER_TALLY if extra_modes else FILTER_SPECIAL
        self.cards = {
            FILTER_ALL: SummaryCard(cards, ui, "Total", lambda: self.set_mode(FILTER_ALL), white),
            FILTER_FLAG: SummaryCard(cards, ui, "Perlu dicek", lambda: self.set_mode(FILTER_FLAG), (("#FDEDEA", "#3A1E1B"), ("#F3C7C0", "#5C2B27"), ("#8C2015", "#F97066"), ("#8C2015", "#F97066"))),
            FILTER_DIRTY: SummaryCard(cards, ui, "Belum disimpan", lambda: self.set_mode(FILTER_DIRTY), (("#FFF6D9", "#3A3118"), ("#EFDFA8", "#5C4C24"), ("#6B4A00", "#F2C879"), ("#6B4A00", "#F2C879"))),
            self.special_mode: SummaryCard(cards, ui, special, lambda: self.set_mode(self.special_mode), white),
        }
        for index, card in enumerate(self.cards.values()):
            cards.grid_columnconfigure(index, weight=1, uniform="card")
            card.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 12, 0))
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        self.search_key = "weight" if (table == "items" and extra_modes) else ("goods" if table == "items" else table)
        placeholder = SEARCH.get(self.search_key, ("Cari...", [], None))[0]
        self.search_bar = motion.SearchBar(bar, self.search, placeholder, self.suggest, width=380, height=42)
        self.search_bar.pack(side="left")
        self.search.trace_add("write", self.on_search)
        self.legend = ctk.CTkFrame(bar, fg_color="transparent", width=1, height=1)
        self.legend.pack(side="right")
        self.legend_names = []
        strip = ctk.CTkFrame(self, fg_color="transparent", height=36)
        strip.grid(row=3, column=0, sticky="ew", pady=(0, 6))
        self.chips = ctk.CTkFrame(strip, fg_color="transparent", width=1, height=1)
        self.chips.pack(side="left")
        self.clear_filter_button = ctk.CTkButton(
            strip, text="Hapus semua", font=ui.font(13, "bold"), height=32, width=100, fg_color="transparent",
            hover_color=ui.ACCENT_SOFT, text_color=ui.ACCENT, command=self.clear_filters, state="disabled",
        )
        self.clear_filter_button.pack(side="left", padx=6)
        self.count = ctk.CTkLabel(strip, text="", font=ui.font(14, "bold"), text_color=ui.ACCENT, height=32)
        self.count.pack(side="right")
        holder = ctk.CTkFrame(self, fg_color=ui.CARD, border_color=ui.BORDER, border_width=1, corner_radius=12)
        holder.grid(row=4, column=0, sticky="nsew")
        holder.grid_columnconfigure(0, weight=1)
        holder.grid_rowconfigure(0, weight=1)
        master = table in ("master", "po_master")
        self.grid_widget = Grid(
            holder, columns, self.edit, on_filter=self.open_filter, name=table if master else title,
            rounded=not master, toolbar=master, movable=master, pins=[0, 1] if table == "prices" else [0],
        )
        self.grid_widget.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.grid_widget.on_select = self.show_etd_change
        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        ui.button(actions, "+ Baris", self.add_row, width=90, height=40).pack(side="left")
        ui.button(actions, "Tempel", self.paste, width=90, height=40).pack(side="left", padx=8)
        ui.button(actions, "Hapus baris", self.delete_rows, width=110, height=40, danger=True).pack(side="left")
        if self.etd_key():
            ui.button(actions, "Hapus tanda ETD", self.clear_etd_marks, width=130, height=40).pack(side="left", padx=8)
        if table == "items" and "CBM" in title:
            ui.button(actions, "Update NW/GW dari file", manager.import_weights, width=180, height=40).pack(side="left")
            ui.button(actions, "Perkirakan NW/GW", lambda: EstimateDialog(manager), width=150, height=40).pack(side="left", padx=(8, 0))
        if table == "items":
            kinds = ("cbm",) if "CBM" in title else ("typo",)
            text = "Tally otomatis CBM" if kinds == ("cbm",) else "Perbaiki typo otomatis"
            ui.button(actions, text, lambda k=kinds: manager.auto_fix(None, k), width=180, height=40, primary=True).pack(side="left", padx=8)
        source = "Input File Master" if table in ("master", "po_master") else "Input File Data"
        ui.button(actions, source, lambda: manager.choose_file(table), width=140, height=40).pack(side="right")

    def menu_items(self):
        manager, table = self.manager, self.table
        master = table in ("master", "po_master")
        items = [("Upload file Master..." if master else "Upload file data...", "Excel", lambda: manager.choose_file(table))]
        if table == "items" and "CBM" in self.title_text:
            items += [("Update NW/GW dari file timbang...", "", manager.import_weights), ("Perkirakan NW/GW", "", lambda: EstimateDialog(manager))]
        items += [("Upload file loading...", "装柜明细", lambda: manager.choose_file("loading")),
                  ("Muat ulang data", "", self.reload_spin), ("Ganti semua file sumber", "", manager.choose_sources), None]
        if self.etd_key():
            items.append(("Hapus tanda ETD", "", self.clear_etd_marks))
        items.append(("!Hapus semua filter", "", self.clear_filters))
        return items

    def open_menu(self):
        self.menu.toggle()

    def reload_spin(self):
        if self.manager.running:
            return
        self.manager.reload(notify=True)
        if self.manager.running:
            self.spin.start()

    def suggest(self, query):
        """Saran pencarian dari kolom utama halaman ini."""
        ql = query.lower()
        _, columns, side_key = SEARCH.get(self.search_key, ("", [], None))
        side_column = next((c for c in self.columns if c.key == side_key), None)
        found, seen = [], set()
        for row in self.rows():
            for key, tag in columns:
                column = next((c for c in self.columns if c.key == key), None)
                if column is None:
                    continue
                value = data.format_value(row.get(key), column.kind)
                if value and ql in value.lower() and value not in seen:
                    seen.add(value)
                    side = ""
                    if tag in ("Nama", "HS", "SKU") and self.table != "prices" and key != "A":
                        side = data.format_value(row.get("A"), "text")
                    elif self.table == "prices" and key == "H":
                        side = data.format_value(row.get("G"), "text")
                    elif side_column is not None:
                        side = "ETD " + data.format_value(row.get(side_key), side_column.kind)
                    found.append((tag, value, side if side.strip() not in ("ETD", "") else ""))
                    if len(found) >= 6:
                        return found
        return found

    def search_styles(self, styles, words):
        """Sel yang cocok dengan pencarian diberi latar kuning lembut."""
        if not words:
            return styles
        styles = dict(styles or {})
        taken = {tuple(cell) for cells in styles.values() for cell in cells}
        bg = "#4A3B12" if is_dark() else "#F6E3B4"
        fg = "#F8F0E5" if is_dark() else "#16202E"
        cells = []
        for number, row in enumerate(self.visible):
            for index, column in enumerate(self.columns):
                if (number, index) in taken:
                    continue
                text = data.format_value(row.get(column.key), column.kind).lower()
                if text and any(word in text for word in words):
                    cells.append((number, index))
        if cells:
            styles[(bg, fg, "search")] = cells
        return styles

    def set_mode(self, mode):
        self.mode.set(FILTER_ALL if self.mode.get() == mode else mode)
        self.refresh()

    def show_chips(self):
        for child in self.chips.winfo_children():
            child.destroy()
        ui = self.manager.ui
        labels = {c.key: c.label for c in self.columns}
        for key, allowed in self.filters.items():
            if not allowed:
                continue
            value = next(iter(allowed)) if len(allowed) == 1 else f"{len(allowed)} nilai"
            ctk.CTkButton(
                self.chips, text=f"{labels[key]}: {value}   ✕", font=ui.font(13, "bold"), height=32, width=30, corner_radius=16,
                fg_color=ui.ACCENT, hover_color=ui.ACCENT_HOVER, text_color="#FFFFFF", command=lambda k=key: self.set_filter(k, set()),
            ).pack(side="left", padx=(0, 6))
        if self.sort_by:
            column, reverse = self.sort_by
            ctk.CTkButton(
                self.chips, text=f"Urut {column.label} {'↓' if reverse else '↑'} (tampilan saja)   ✕", font=ui.font(13, "bold"), height=32, width=30, corner_radius=16,
                fg_color=ui.ACCENT, hover_color=ui.ACCENT_HOVER, text_color="#FFFFFF", command=lambda: self.set_sort(None),
            ).pack(side="left", padx=(0, 6))

    def rows(self):
        return self.manager.store.rows_of(self.table)

    def open_filter(self, column):
        if self.filter_dialog is not None and self.filter_dialog.winfo_exists():
            self.filter_dialog.destroy()
        self.filter_dialog = FilterDialog(self, column)

    def etd_key(self):
        return DashboardPage.ETD_KEY[self.profile.name] if self.profile else None

    def etd_marks(self):
        """ETD awal per PO/DI. Disimpan di pengaturan supaya tandanya tetap ada setelah app ditutup."""
        return data.load_settings().get("etd_changes", {}).get(self.profile.name, {})

    def save_etd_marks(self, marks):
        settings = data.load_settings()
        settings.setdefault("etd_changes", {})[self.profile.name] = marks
        data.save_settings(settings)

    def etd_text(self, value):
        column = next(c for c in self.columns if c.key == self.etd_key())
        return data.format_value(value, column.kind)

    def note_etd_change(self, row, old, new):
        code = data.unique_key(row.get("A"))
        if not code:
            return
        marks = self.etd_marks()
        before, after = self.etd_text(old), self.etd_text(new)
        if code in marks:
            if after == marks[code]:
                del marks[code]
        elif before and after and before != after:
            marks[code] = before
        else:
            return
        self.save_etd_marks(marks)

    def etd_styles(self, styles):
        key = self.etd_key()
        if not key:
            return styles
        marks = self.etd_marks()
        if not marks:
            return styles
        index = next(i for i, c in enumerate(self.columns) if c.key == key)
        cells = []
        for number, row in enumerate(self.visible):
            first = marks.get(data.unique_key(row.get("A")))
            if first and self.etd_text(row.get(key)) not in ("", first):
                cells.append((number, index))
        if cells:
            styles = dict(styles or {})
            styles[(None, "#C4B5FD" if is_dark() else "#6D28D9", "etd")] = cells
        return styles

    def show_etd_change(self, row, column):
        key = self.etd_key()
        if not key or row >= len(self.visible) or column >= len(self.columns) or self.columns[column].key != key:
            return
        first = self.etd_marks().get(data.unique_key(self.visible[row].get("A")))
        current = self.etd_text(self.visible[row].get(key))
        if first and current and current != first:
            self.manager.ui.status.configure(text=f"ETD berubah: {first} menjadi {current}.")

    def clear_etd_marks(self):
        chosen = [self.visible[i] for i in self.grid_widget.selected_rows() if i < len(self.visible)]
        if not chosen:
            messagebox.showinfo("Dockit", "Pilih dulu baris yang tanda ETD-nya mau dihapus.")
            return
        marks = self.etd_marks()
        for row in chosen:
            marks.pop(data.unique_key(row.get("A")), None)
        self.save_etd_marks(marks)
        self.refresh()

    def set_sort(self, column, reverse=False):
        """Mengurutkan hanya tampilan. Data di file tetap mengikuti urutan paste."""
        self.sort_by = (column, reverse) if column else None
        self.refresh()

    def sort_matches(self, matches):
        if not self.sort_by:
            return matches
        column, reverse = self.sort_by

        def value_of(item):
            value = data.format_value(item[0].get(column.key), column.kind)
            if column.kind == "num":
                try:
                    return (0, float(value.replace(",", "")), "")
                except ValueError:
                    pass
            return (0, 0.0, value.lower())

        filled = [item for item in matches if not data.is_blank(item[0].get(column.key))]
        blank = [item for item in matches if data.is_blank(item[0].get(column.key))]
        return sorted(filled, key=value_of, reverse=reverse) + blank

    def passes_filters(self, row, skip=None):
        for key, kind, allowed in self.active_filters:
            if key != skip and (data.format_value(row.get(key), kind) or EMPTY_VALUE) not in allowed:
                return False
        return True

    @property
    def active_filters(self):
        kinds = {c.key: c.kind for c in self.columns}
        return [(key, kinds[key], allowed) for key, allowed in self.filters.items() if allowed]

    def column_values(self, column):
        values = {
            data.format_value(row.get(column.key), column.kind) or EMPTY_VALUE
            for row in self.rows() if self.passes_filters(row, skip=column.key)
        }
        return sorted(values)

    def column_counts(self, column):
        counts = {}
        for row in self.rows():
            if self.passes_filters(row, skip=column.key):
                value = data.format_value(row.get(column.key), column.kind) or EMPTY_VALUE
                counts[value] = counts.get(value, 0) + 1
        return counts

    def set_filter(self, key, allowed):
        if allowed:
            self.filters[key] = set(allowed)
        else:
            self.filters.pop(key, None)
        self.refresh()

    def clear_filters(self):
        self.filters = {}
        self.sort_by = None
        self.refresh()

    def on_search(self, *_):
        if self.search_job:
            self.after_cancel(self.search_job)
        self.search_job = self.after(250, self.refresh)

    def refresh(self):
        self.search_job = None
        words = self.search.get().lower().split()
        mode = self.mode.get()
        total_flag = total_dirty = total_special = 0
        matches = []
        rows = self.rows()
        for row in rows:
            flag = self.flag(row)
            total_flag += bool(flag)
            total_dirty += bool(row.dirty)
            total_special += self.special_text in flag if self.special_text else 0
            if not self.passes_filters(row):
                continue
            if mode == FILTER_FLAG and not flag:
                continue
            if mode == FILTER_DIRTY and not row.dirty:
                continue
            if mode in (FILTER_TALLY, FILTER_SPECIAL) and self.special_text not in flag:
                continue
            if words:
                text = " ".join(data.format_value(row.get(c.key), c.kind) for c in self.columns).lower()
                if not all(word in text for word in words):
                    continue
            matches.append((row, flag))
        # Urutan data selalu sama dengan file sumber dan urutan paste. Urut di filter hanya mengubah tampilan.
        matches = self.sort_matches(matches)
        self.visible = [row for row, _ in matches]
        lines = [[data.format_value(row.get(c.key), c.kind) for c in self.columns] + [flag] for row, flag in matches]
        tags = ["dirty" if row.dirty else ("flag" if flag else "") for row, flag in matches]
        self.grid_widget.fill(lines, tags, self.search_styles(self.estimate_styles(self.etd_styles(self.pod_styles())), words))
        self.update_legend()
        self.count.configure(text=f"Menampilkan {len(matches):,} dari {len(rows):,} baris")
        for mode_name, number in ((FILTER_ALL, len(rows)), (FILTER_FLAG, total_flag), (FILTER_DIRTY, total_dirty), (self.special_mode, total_special)):
            self.cards[mode_name].show(number, mode == mode_name and mode != FILTER_ALL)
        self.show_chips()
        self.grid_widget.set_marks({key for key, _, _ in self.active_filters})
        self.clear_filter_button.configure(state="normal" if self.filters or self.sort_by else "disabled")
        self.update_save()

    def estimate_styles(self, styles):
        """NW dan GW perkiraan ditampilkan jingga tebal sampai diganti hasil timbang."""
        if self.table != "items" or not any(c.key == "E" for c in self.columns):
            return styles
        marks = data.estimate_marks()
        if not marks:
            return styles
        store = self.manager.store
        cells = []
        indexes = [i for i, c in enumerate(self.columns) if c.key in ("E", "F")]
        for number, row in enumerate(self.visible):
            if str(row.get("A") or "").strip() in marks and store.estimate_note(row):
                cells += [(number, i) for i in indexes]
        if cells:
            styles = dict(styles or {})
            styles[(None, "#F2A65A" if is_dark() else "#B45309", "estimate")] = cells
        return styles

    def pod_index(self):
        if not self.profile:
            return None
        return next(i for i, c in enumerate(self.columns) if c.key == self.profile.pod_key)

    def pod_styles(self):
        index = self.pod_index()
        if index is None:
            return None
        styles = {}
        for number, row in enumerate(self.visible):
            if not data.is_blank(row.get(self.profile.pod_key)):
                styles.setdefault(pod_color(row.get(self.profile.pod_key)), []).append((number, index))
        # BL yang dipakai beberapa DI mendapat warna yang sama per kelompok.
        groups = {}
        for row in self.rows():
            key = data.unique_key(row.get(self.profile.bl_key))
            if key:
                groups.setdefault(key, []).append(row)
        colors = {}
        for key, members in groups.items():
            if len(members) > 1:
                colors[key] = BL_COLORS[len(colors) % len(BL_COLORS)]
        bl_index = next(i for i, c in enumerate(self.columns) if c.key == self.profile.bl_key)
        for number, row in enumerate(self.visible):
            color = colors.get(data.unique_key(row.get(self.profile.bl_key)))
            if color:
                styles.setdefault(color + ("bl",), []).append((number, bl_index))
        return styles

    def update_legend(self):
        if self.pod_index() is None:
            return
        pod = self.profile.pod_key
        names = sorted({" ".join(str(r.get(pod)).upper().split()) for r in self.rows() if not data.is_blank(r.get(pod))})
        if names == self.legend_names:
            return
        self.legend_names = names
        for child in self.legend.winfo_children():
            child.destroy()
        font = self.manager.ui.font(11, "bold")
        for name in names:
            bg, fg = pod_color(name)
            ctk.CTkButton(
                self.legend, text=name, font=font, height=24, width=20, corner_radius=12, fg_color=bg, hover_color=bg, text_color=fg,
                command=lambda value=name: self.filter_pod(value),
            ).pack(side="left", padx=2)

    def filter_pod(self, name):
        pod = self.profile.pod_key
        values = {data.format_value(r.get(pod)) for r in self.rows() if " ".join(str(r.get(pod)).upper().split()) == name}
        self.set_filter(pod, values)

    def update_save(self):
        store = self.manager.store
        count = store.dirty_count(self.table)
        if self.table in ("items", "prices"):
            count += store.dirty_count("prices" if self.table == "items" else "items")
        self.save_button.configure(text=f"Simpan ({count})" if count else "Simpan", state="normal" if count else "disabled")

    def edit(self, iid, column, text, check=False):
        row = self.visible[int(iid)]
        allowed = column.editable and (not column.new_only or row.row is None)
        if check:
            return allowed
        if not allowed:
            self.refresh()
            return False
        try:
            value = data.parse_value(text, column.kind)
        except ValueError as exc:
            messagebox.showwarning("Dockit", str(exc))
            self.refresh()
            return False
        if value == row.get(column.key):
            self.refresh()
            return True
        if column.key == self.etd_key():
            self.note_etd_change(row, row.get(column.key), value)
        if self.table == "items" and column.key in ("E", "F"):
            data.clear_estimates([row.get("A")])
        row.values[column.key] = value
        row.dirty.add(column.key)
        if self.table == "master":
            data.refresh_master_formulas(row)
        self.manager.store.reindex()
        self.manager.refresh_all(exclude=self)
        self.refresh()
        return True

    def add_row(self):
        self.manager.store.new_row(self.table)
        self.search.set("")
        self.filters = {}
        self.mode.set(FILTER_ALL)
        row = self.rows()[-1]
        row.dirty.add(self.columns[0].key)
        self.refresh()
        if row in self.visible:
            self.grid_widget.see(self.visible.index(row))
        self.manager.ui.status.configure(text="Klik dua kali sel pada baris baru untuk mengisi.")

    def paste(self):
        try:
            text = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Dockit", "Clipboard kosong. Salin data dari Excel dulu.")
            return
        columns = [c for c in self.columns if c.key in self.paste_keys]
        order = sorted(columns, key=lambda c: self.paste_keys.index(c.key))
        names = ", ".join(c.label for c in order)
        lines = [line.split("\t") for line in text.splitlines() if line.strip()]
        if not lines:
            return
        if not messagebox.askyesno("Dockit", f"Tambah {len(lines)} baris baru dengan urutan kolom:\n{names}\n\nLanjut?"):
            return
        try:
            parsed = [{c.key: data.parse_value(v, c.kind) for c, v in zip(order, line)} for line in lines]
        except ValueError as exc:
            messagebox.showwarning("Dockit", f"Tidak ada yang ditambahkan.\n{exc}")
            return
        for values in parsed:
            self.manager.store.new_row(self.table, values)
        self.filters = {}
        self.mode.set(FILTER_DIRTY)
        self.manager.refresh_all()

    def delete_rows(self):
        chosen = [self.visible[i] for i in self.grid_widget.selected_rows() if i < len(self.visible)]
        if not chosen:
            messagebox.showinfo("Dockit", "Pilih dulu baris yang mau dihapus. Klik nomor baris di sebelah kiri.")
            return
        unsaved = [row for row in chosen if row.row is None]
        saved = [row for row in chosen if row.row is not None]
        for row in unsaved:
            self.rows().remove(row)
        self.manager.store.reindex()
        self.manager.refresh_all()
        if not saved:
            return
        if self.manager.dirty():
            messagebox.showwarning("Dockit", "Simpan atau muat ulang perubahan yang lain dulu, lalu hapus lagi.")
            return
        key = self.columns[1].key if self.table == "prices" else self.columns[0].key
        names = [data.format_value(row.get(key)) for row in saved]
        note = "\n\nSemua baris dengan SKU yang sama ikut terhapus dari tab Goods Description dan CBM & Berat." if self.table == "items" else ""
        if messagebox.askyesno(
            "Dockit",
            f"Hapus {len(saved)} baris dari file Excel sumber?\n\n" + "\n".join(names[:12])
            + ("\n..." if len(names) > 12 else "") + note + "\n\nBackup dibuat otomatis sebelum dihapus.",
            icon="warning",
        ):
            self.manager.delete(self.table, saved)

    def save(self):
        for row in self.rows():
            if row.row is None and is_incomplete(self.table, row):
                messagebox.showwarning("Dockit", "Baris baru belum punya kunci (DI, SKU, atau DI/PO + SKU).")
                return
        if self.profile:
            store = self.manager.store
            doubles = [
                f"{r.get('A')}: {', '.join(store.duplicate_labels(r, self.profile))}"
                for r in self.rows() if r.dirty and store.duplicate_labels(r, self.profile)
            ]
            if doubles and not messagebox.askyesno(
                "Dockit", "⚠ Ada nilai double:\n\n" + "\n".join(doubles[:12]) + "\n\nTetap simpan?"
            ):
                return
        self.manager.save(self.table)


def is_incomplete(table, row):
    if table == "prices":
        return not (row.get("G") and row.get("H"))
    return not row.get("A")


class BuildPage(ctk.CTkFrame):
    COLUMNS = [
        data.Column("sku", "Item No.", width=190), data.Column("qty", "Qty", "num", width=60),
        data.Column("asin", "ASIN", width=110), data.Column("po", "PO", width=95),
        data.Column("name", "Goods Description", editable=False, width=300), data.Column("nw", "NW", editable=False, width=60),
        data.Column("gw", "GW", editable=False, width=60), data.Column("cbm", "CBM", "num", width=75),
        data.Column("price", "Harga", "num", width=70),
    ]

    PO_COLUMNS = [
        data.Column("sku", "Item No.", width=210), data.Column("qty", "Qty", "num", width=65),
        data.Column("name", "Goods Description", editable=False, width=400), data.Column("nw", "NW", editable=False, width=65),
        data.Column("gw", "GW", editable=False, width=65), data.Column("cbm", "CBM", "num", width=75),
        data.Column("price", "Harga", "num", width=75), data.Column("service", "Jasa", "num", width=65),
    ]

    def __init__(self, parent, manager, profile=None):
        super().__init__(parent, fg_color="transparent")
        self.manager = manager
        self.profile = profile or data.PROFILES["DI"]
        self.COLUMNS = self.PO_COLUMNS if self.profile.name == "PO" else self.COLUMNS
        ui = manager.ui
        self.sets = {}
        self.di_keys = []
        self.active = ""
        self.syncing = False
        self.filled = []
        self.problems = []
        self.number = ctk.StringVar()
        self.bl = ctk.StringVar()
        settings = data.load_settings()
        self.folder = ctk.StringVar(value=settings.get("output_folder", ""))
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ctk.CTkLabel(top, text="Buat Excel", font=ui.font(26, "bold"), text_color=ui.TEXT).pack(side="left", padx=(0, 22))
        ctk.CTkLabel(top, text="Nomor BL", font=ui.font(14, "bold"), text_color=ui.TEXT).pack(side="left")
        bl_entry = ctk.CTkEntry(top, textvariable=self.bl, width=170, height=42, corner_radius=9, border_color=ui.BORDER, font=ui.font(14))
        bl_entry.pack(side="left", padx=(8, 6))
        bl_entry.bind("<Return>", lambda _: self.open_bl())
        ui.button(top, f"Ambil {self.profile.name}", self.open_bl, primary=True, width=96, height=42).pack(side="left", padx=(0, 16))
        ctk.CTkLabel(top, text=f"atau Nomor {self.profile.name}", font=ui.font(14, "bold"), text_color=ui.TEXT).pack(side="left")
        ctk.CTkEntry(top, textvariable=self.number, width=100, height=42, corner_radius=9, border_color=ui.BORDER, font=ui.font(14)).pack(side="left", padx=8)
        self.number.trace_add("write", lambda *_: self.on_number())
        self.info = ctk.CTkLabel(top, text="", font=ui.font(12), text_color=ui.MUTED, anchor="w")
        self.info.pack(side="left", padx=8, fill="x", expand=True)
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        self.selector = ctk.CTkFrame(self, fg_color="transparent", height=1)
        self.selector.grid(row=1, column=0, sticky="ew")
        self.selector_keys = None
        self.selector_buttons = {}
        self.add_button = ui.button(bar, "+ Item", self.add_line, width=74)
        self.add_button.pack(side="left")
        ui.button(bar, "Ambil Ulang Item", self.fill_from_prices, width=120).pack(side="left", padx=(6, 0))
        ui.button(bar, "Tempel Item & Qty", self.paste, width=140).pack(side="left", padx=6)
        ui.button(bar, "Hapus baris", self.remove, width=100).pack(side="left")
        ui.button(bar, "Kosongkan", self.clear, width=90).pack(side="left", padx=6)
        ui.button(bar, "Input File Loading", lambda: self.manager.choose_file("loading"), width=130).pack(side="right")
        self.source_label = ctk.CTkLabel(bar, text="", font=ui.font(12), text_color=ui.MUTED)
        self.source_label.pack(side="right", padx=8)
        holder = ctk.CTkFrame(self, fg_color=ui.CARD, border_color=ui.BORDER, border_width=1, corner_radius=12)
        holder.grid(row=3, column=0, sticky="nsew")
        holder.grid_columnconfigure(0, weight=1)
        holder.grid_rowconfigure(0, weight=1)
        self.grid_widget = Grid(holder, self.COLUMNS, self.edit, name=f"build_{self.profile.name}")
        self.grid_widget.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.totals = ctk.CTkFrame(self, fg_color="transparent")
        self.totals.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        self.total_labels = {}
        for row_number, name in enumerate(("di", "bl")):
            self.totals.grid_rowconfigure(row_number, weight=0)
            title = ctk.CTkLabel(self.totals, text="", font=ui.font(13, "bold"), text_color=ui.ACCENT, anchor="w", width=120, wraplength=120)
            title.grid(row=row_number, column=0, sticky="w", padx=(2, 8), pady=2)
            cells = [title]
            for column in range(1, 8):
                self.totals.grid_columnconfigure(column, weight=1, uniform="total")
                card = ctk.CTkLabel(
                    self.totals, text="", font=ui.font(13, "bold"), text_color=ui.TEXT, fg_color=ui.CARD, corner_radius=9, height=46, justify="left",
                )
                card.grid(row=row_number, column=column, sticky="ew", padx=3, pady=2)
                cells.append(card)
            self.total_labels[name] = cells
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.grid(row=5, column=0, sticky="ew", pady=(6, 0))
        bottom.grid_columnconfigure(0, weight=1)
        self.check_title = ctk.CTkLabel(bottom, text="Cek data", font=ui.font(13, "bold"), text_color=ui.TEXT, anchor="w")
        self.check_title.grid(row=0, column=0, sticky="ew")
        self.fix_button = ui.button(bottom, "Perbaiki otomatis", self.auto_fix, primary=True, height=30)
        self.fix_button.grid(row=0, column=1, sticky="e")
        self.fix_button.grid_remove()
        self.check_text = ctk.CTkTextbox(bottom, height=70, corner_radius=9, font=ui.font(12), fg_color=ui.NEUTRAL, text_color=ui.TEXT, wrap="word", state="disabled")
        self.check_text.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 8))
        output = ctk.CTkFrame(bottom, fg_color="transparent")
        output.grid(row=2, column=0, columnspan=2, sticky="ew")
        output.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(output, text="Simpan ke", font=ui.font(13, "bold"), text_color=ui.TEXT).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkEntry(output, textvariable=self.folder, height=34, font=ui.font(13)).grid(row=0, column=1, sticky="ew")
        ui.button(output, "Pilih...", self.choose_folder, width=70).grid(row=0, column=2, padx=6)
        self.excel_button = ui.button(output, "Buat Excel", lambda: self.create(False), width=130, height=38)
        self.excel_button.grid(row=0, column=3)
        self.pdf_button = ui.button(output, "Buat Excel + PDF", lambda: self.create(True), primary=True, width=170, height=38)
        self.pdf_button.grid(row=0, column=4, padx=(6, 0))
        self.progress_box = ctk.CTkFrame(bottom, fg_color=ui.CARD, border_color=ui.BORDER, border_width=1, corner_radius=12)
        self.progress_box.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.progress_box.grid_columnconfigure(0, weight=1)
        self.progress_title = ctk.CTkLabel(self.progress_box, text="", font=ui.font(13, "bold"), text_color=ui.TEXT, anchor="w")
        self.progress_title.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 0))
        self.progress_pct = ctk.CTkLabel(self.progress_box, text="0%", font=ui.font(20, "bold"), text_color=ui.ACCENT)
        self.progress_pct.grid(row=0, column=1, rowspan=2, padx=14)
        self.pixel_bar = motion.PixelBar(self.progress_box, height=14)
        self.pixel_bar.grid(row=1, column=0, sticky="ew", padx=14, pady=(6, 4))
        self.progress_chips = ctk.CTkFrame(self.progress_box, fg_color="transparent")
        self.progress_chips.grid(row=2, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 10))
        self.progress_box.grid_remove()

    def di_key(self):
        return self.active

    @property
    def lines(self):
        return self.sets.setdefault(self.active, [])

    @lines.setter
    def lines(self, value):
        self.sets[self.active] = value

    def on_number(self):
        if self.syncing:
            return
        text = self.number.get().strip()
        key = f"{self.profile.prefix}{text.split('柜')[-1]}" if text else ""
        self.di_keys = [key] if key else []
        self.active = key
        store = self.manager.store
        if key in store.loading and not self.sets.get(key):
            self.sets[key] = store.lines_for(key)
        self.refresh()

    def open_bl(self):
        store = self.manager.store
        key = data.unique_key(self.bl.get())
        members = store.bl_of(self.profile).get(key, [])
        if not members:
            messagebox.showinfo("Dockit", f"BL {self.bl.get().strip()} tidak ada di Master {self.profile.name}.")
            return
        self.di_keys = list(members)
        filled = 0
        for member in members:
            if not self.sets.get(member):
                self.sets[member] = store.lines_for(member)
                filled += bool(self.sets[member])
        self.select_di(members[0])
        self.manager.ui.status.configure(text=f"BL {self.bl.get().strip()}: {len(members)} {self.profile.name}, {filled} terisi dari {store.line_source(members[0])}. Cek lalu buat.")

    def select_di(self, value):
        key = next((k for k in self.di_keys if k == value or self.label_of(k) == value), value)
        self.active = key
        self.syncing = True
        self.number.set(key.split("柜")[-1])
        self.syncing = False
        self.refresh()

    def label_of(self, key):
        store = self.manager.store
        _, problems = store.check_items(key, self.sets.get(key, []), self.profile)
        mark = "⚠" if problems or not self.sets.get(key) else "✓"
        return f"{mark} {key.split('柜')[-1]}"

    def update_selector(self):
        ui = self.manager.ui
        store = self.manager.store
        keys = list(self.di_keys) if len(self.di_keys) > 1 else []
        if keys != self.selector_keys:
            self.selector_keys = keys
            for child in self.selector.winfo_children():
                child.destroy()
            self.selector_buttons = {}
            for index, key in enumerate(keys):
                self.selector.grid_columnconfigure(index, weight=1, uniform="di")
                button = ctk.CTkButton(
                    self.selector, text="", font=ui.font(13, "bold"), height=58, corner_radius=12, border_width=1,
                    command=lambda k=key: self.select_di(k),
                )
                button.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 10, 0), pady=(0, 8))
                self.selector_buttons[key] = button
        for key, button in self.selector_buttons.items():
            lines = self.sets.get(key, [])
            _, problems = store.check_items(key, lines, self.profile)
            state = f"⚠ {len(problems)} perlu dicek" if problems else ("Belum ada item" if not lines else "✓ Lengkap")
            active = key == self.active
            button.configure(
                text=f"{key}\n{state} · {len(lines)} item",
                fg_color=ui.ACCENT_SOFT if active else ui.CARD, hover_color=ui.ACCENT_SOFT, text_color=ui.TEXT,
                border_color=ui.ACCENT if active else ui.BORDER, border_width=2 if active else 1,
            )
        count = len(self.di_keys)
        suffix = f" ({count} {self.profile.name})" if count > 1 else ""
        self.excel_button.configure(text=f"Buat Excel{suffix}")
        self.pdf_button.configure(text=f"Buat Excel + PDF{suffix}")

    def fill_from_prices(self):
        if not self.active:
            messagebox.showinfo("Dockit", "Isi nomor BL atau nomor DI dulu.")
            return
        store = self.manager.store
        lines = store.lines_for(self.active)
        if not lines:
            messagebox.showinfo("Dockit", f"{self.active} belum ada di file loading maupun tab Harga.")
            return
        if self.lines and not messagebox.askyesno("Dockit", f"Ganti daftar item DI ini dengan data dari {store.line_source(self.active)}?"):
            return
        self.lines = lines
        self.refresh()

    def refresh(self):
        store = self.manager.store
        key = self.di_key()
        master = store.index_of(self.profile).get(key)
        self.update_selector()
        self.source_label.configure(
            text=f"Loading: {shorten_name(store.load_path)}" if store.load_path else "File loading belum dipilih, item dari tab Harga"
        )
        if not key:
            self.info.configure(text=f"Isi nomor BL untuk semua {self.profile.name}-nya, atau nomor {self.profile.name}.")
        elif master is None:
            self.info.configure(text=f"{key} tidak ada di Master {self.profile.name}.")
        else:
            keys = "EGHJMNO" if self.profile.name == "DI" else "BHIEFG"
            parts = [master.get(k) for k in keys]
            self.info.configure(text="  ·  ".join(data.format_value(p) for p in parts if not data.is_blank(p)))
        if key and self.lines:
            split = store.split_prices(key, self.lines)
            if len(split) != len(self.lines):
                self.lines = split
        self.filled, self.problems = store.check_items(key, self.lines, self.profile) if key else ([], [])
        rows = [[self.display(line, c) for c in self.COLUMNS] + ["; ".join(line["notes"])] for line in self.filled]
        self.grid_widget.fill(rows, ["flag" if line["notes"] else "" for line in self.filled], self.price_styles())
        self.show_totals()
        ui = self.manager.ui
        if not key or not self.lines:
            title, color = "Cek data", ui.TEXT
        elif self.problems:
            title, color = f"⚠ {len(self.problems)} hal perlu diperbaiki", ui.DANGER
        else:
            title, color = "✓ Data lengkap", ui.ACCENT
        self.check_title.configure(text=title, text_color=color)
        fixes, _ = store.auto_fixes(self.problem_skus()) if key and self.lines else ([], [])
        if fixes:
            self.fix_button.configure(text=f"Perbaiki otomatis ({len(fixes)})")
            self.fix_button.grid()
        else:
            self.fix_button.grid_remove()
        self.check_text.configure(state="normal")
        self.check_text.delete("1.0", "end")
        splits = sorted({line["sku"] for line in self.filled if line.get("split")})
        info = [f"ℹ Harga berbeda dipecah per baris (sel biru): {', '.join(splits)}"] if splits else []
        self.check_text.insert("1.0", "\n".join(self.problems + info))
        self.check_text.configure(state="disabled")

    def price_styles(self):
        """SKU dengan harga berbeda: sel Harga dan Qty diberi warna biru supaya kelihatan."""
        index = {c.key: i for i, c in enumerate(self.COLUMNS)}
        cells = []
        for number, line in enumerate(self.filled):
            if line.get("split"):
                cells += [(number, index[k]) for k in ("price", "qty") if k in index]
        if not cells:
            return {}
        bg, fg = ("#1D2A44", "#9DB8F0") if is_dark() else ("#DCE8FA", "#1D3B8A")
        return {(bg, fg, "price2"): cells}

    def problem_skus(self):
        """SKU di semua PO/DI yang sedang dibuka, supaya satu klik memperbaiki semuanya."""
        return {str(line.get("sku") or "").strip() for key in (self.di_keys or [self.active]) for line in self.sets.get(key, self.lines if key == self.active else [])}

    def auto_fix(self):
        self.manager.auto_fix(self.problem_skus())

    def display(self, line, column):
        value = line[column.key]
        if column.key == "cbm" and isinstance(value, (int, float)):
            return f"{data.round_down_2(value):.2f}"
        return data.format_value(value, column.kind)

    def show_totals(self):
        store = self.manager.store

        def write(name, title, totals):
            cells = self.total_labels[name]
            cells[0].configure(text=title)
            parts = []
            if totals:
                parts = [
                    f"Qty\n{totals['qty']:,.0f}", f"CTNS\n{totals['ctns']:,.0f}", f"NW\n{totals['nw']:,.2f}", f"GW\n{totals['gw']:,.2f}",
                    f"CBM\n{totals['cbm']:,.2f}", f"Harga satuan\n{totals['price']:,.2f}", f"Amount\n{totals['amount']:,.2f}",
                ]
            for cell, text in zip(cells, [title] + (parts or [""] * 7)):
                cell.configure(text=text)
                if parts:
                    cell.grid()
                else:
                    cell.grid_remove()

        short = self.active.split("柜")[-1]
        write("di", f"Total {self.profile.name} {short}" if self.active else f"Total {self.profile.name}", data.line_totals(self.filled) if self.active else None)
        if len(self.di_keys) > 1:
            everything = []
            for key in self.di_keys:
                everything.extend(store.check_items(key, self.sets.get(key, []), self.profile)[0])
            write("bl", f"Total BL ({len(self.di_keys)} {self.profile.name})", data.line_totals(everything))
        else:
            write("bl", "", None)

    def edit(self, iid, column, text, check=False):
        if check:
            return column.editable
        line = self.filled[int(iid)] if int(iid) < len(self.filled) else None
        # Nilai yang tampil sudah dibulatkan. Kalau tidak diubah, nilai aslinya tetap dipakai.
        if line is not None and text.strip() == self.display(line, column):
            self.refresh()
            return True
        try:
            value = data.parse_value(text, column.kind)
        except ValueError as exc:
            messagebox.showwarning("Dockit", str(exc))
            return False
        self.lines[int(iid)][column.key] = value
        self.refresh()
        return True

    def add_line(self):
        if not self.active:
            messagebox.showinfo("Dockit", "Isi nomor BL atau nomor DI dulu.")
            return
        self.lines.append({"sku": "", "qty": None})
        self.refresh()

    def paste(self):
        try:
            text = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Dockit", "Clipboard kosong. Salin kolom Item No. dan Qty dari Excel dulu.")
            return
        if not self.active:
            messagebox.showinfo("Dockit", "Isi nomor BL atau nomor DI dulu.")
            return
        for line in text.splitlines():
            parts = line.split("\t")
            if not parts[0].strip():
                continue
            quantity = None
            if len(parts) > 1:
                try:
                    quantity = data.parse_value(parts[1], "num")
                except ValueError:
                    pass
            self.lines.append({"sku": parts[0].strip(), "qty": quantity})
        self.refresh()

    def remove(self):
        for index in reversed(self.grid_widget.selected_rows()):
            if index < len(self.lines):
                del self.lines[index]
        self.refresh()

    def clear(self):
        self.lines = []
        self.refresh()

    def choose_folder(self):
        folder = filedialog.askdirectory(title="Pilih folder penyimpanan")
        if folder:
            self.folder.set(os.path.normpath(folder))

    def create(self, with_pdf):
        store = self.manager.store
        folder = self.folder.get().strip()
        keys = list(self.di_keys)
        if not keys:
            messagebox.showwarning("Dockit", "Isi nomor BL atau nomor DI dulu.")
            return
        if not folder:
            messagebox.showwarning("Dockit", "Pilih folder penyimpanan dulu.")
            return
        jobs, warnings = [], []
        for key in keys:
            master = store.index_of(self.profile).get(key)
            lines = self.sets.get(key, [])
            short = key.split("柜")[-1]
            if master is None:
                messagebox.showwarning("Dockit", f"{key} belum ada di Master {self.profile.name}.")
                return
            if master.row is None or master.dirty:
                messagebox.showwarning("Dockit", f"{key} punya perubahan di Master {self.profile.name} yang belum disimpan. Simpan dulu, lalu buat Excel.")
                return
            if not lines or len(lines) > self.profile.max_items:
                messagebox.showwarning("Dockit", f"{self.profile.name} {short}: isi 1 sampai {self.profile.max_items} item.")
                self.select_di(key)
                return
            filled, problems = store.check_items(key, lines, self.profile)
            if any(not line["sku"] for line in filled):
                messagebox.showwarning("Dockit", f"{self.profile.name} {short}: ada baris tanpa Item No.")
                self.select_di(key)
                return
            warnings.extend(f"{self.profile.name} {short} · {problem}" for problem in problems)
            jobs.append((key, master, filled))
        if warnings and not messagebox.askyesno(
            "Dockit", f"⚠ {len(warnings)} hal masih perlu diperbaiki:\n\n" + "\n".join(warnings[:12])
            + ("\n..." if len(warnings) > 12 else "") + "\n\nTetap buat file?"
        ):
            return
        os.makedirs(folder, exist_ok=True)
        outputs = [os.path.join(folder, f"{key}_{store.file_code(master, self.profile)}.xlsx") for key, master, _ in jobs]
        existing = [os.path.basename(path) for path in outputs if os.path.exists(path)]
        if existing and not messagebox.askyesno("Dockit", "File ini sudah ada dan akan ditimpa:\n\n" + "\n".join(existing) + "\n\nLanjut?"):
            return
        self.run_build(jobs, outputs, folder, with_pdf)

    def run_build(self, jobs, outputs, folder, with_pdf):
        """Buat Excel satu per satu sambil progress terisi piksel demi piksel."""
        ui = self.manager.ui
        for child in self.progress_chips.winfo_children():
            child.destroy()
        chips = []
        for key, _, _ in jobs[:8]:
            chip = ctk.CTkLabel(self.progress_chips, text=key.split("柜")[-1], font=ui.font(12, "bold"), text_color=ui.MUTED,
                                fg_color=ui.NEUTRAL, corner_radius=8, height=28, width=78)
            chip.pack(side="left", padx=(0, 6))
            chips.append(chip)
        self.pixel_bar.set(0)
        self.progress_pct.configure(text="0%")
        self.progress_box.grid()
        self.excel_button.configure(state="disabled")
        self.pdf_button.configure(state="disabled")
        total = len(jobs)

        def finish(error=None):
            self.excel_button.configure(state="normal")
            self.pdf_button.configure(state="normal")
            if error:
                self.progress_box.grid_remove()
                self.manager.notify("Excel gagal dibuat", str(error), "error",
                                    action=lambda: messagebox.showerror("Dockit", f"Excel gagal dibuat:\n\n{error}"), action_text="Detail")
                return
            settings = data.load_settings()
            settings["output_folder"] = folder
            data.save_settings(settings)
            ui.status.configure(text=f"{len(outputs)} Excel dibuat di {folder}")

            def done():
                self.progress_box.grid_remove()
                open_folder = (lambda: os.startfile(folder)) if os.name == "nt" else None
                self.manager.notify(f"{total} Excel selesai dibuat" + (", lanjut PDF" if with_pdf else ""), folder, "ok",
                                    action=open_folder, action_text="Buka folder" if open_folder else "")
                self.manager.open_generated(outputs, with_pdf)

            self.after(700, done)

        def step(index):
            if index >= total:
                finish()
                return
            key, master, filled = jobs[index]
            self.progress_title.configure(text=f"Membuat Excel · {index + 1} dari {total} · {key}")
            self.update_idletasks()
            try:
                data.build_excel(self.manager.template_path(self.profile), outputs[index], key, master, filled, self.profile)
            except Exception as exc:
                finish(exc)
                return
            self.pixel_bar.set((index + 1) / total)
            self.progress_pct.configure(text=f"{round((index + 1) / total * 100)}%")
            if index < len(chips):
                chips[index].configure(text="✓ " + key.split("柜")[-1], text_color=("#14563F", "#5FD3A0"), fg_color=("#E6F2EA", "#1C3A2C"))
            self.after(260, lambda: step(index + 1))

        self.after(120, lambda: step(0))


class DataManager:
    """Mengatur sumber data, muat, simpan, dan halaman DI."""

    def __init__(self, ui, parent):
        self.ui = ui
        self.store = data.Store("", "")
        self.running = False
        self.previous_settings = None
        self.pages = {}
        store = self.store
        self.pages[PAGE_DASHBOARD] = DashboardPage(parent, self, data.PROFILES["DI"])
        self.pages[PAGE_DASHBOARD_PO] = DashboardPage(parent, self, data.PROFILES["PO"])
        from calendar_page import CalendarPage
        self.pages[PAGE_CALENDAR] = CalendarPage(parent, self, data.PROFILES["DI"])
        self.pages[PAGE_CALENDAR_PO] = CalendarPage(parent, self, data.PROFILES["PO"])
        self.pages[PAGE_MASTER] = DataPage(parent, self, "Master DI", "Warna kolom BL = BL yang sama. Klik warna POD atau tombol di judul kolom untuk filter.",
                                           data.MASTER_COLUMNS, "master", lambda r: self.store.master_flag(r, data.PROFILES["DI"]),
                                           [c.key for c in data.MASTER_COLUMNS if c.editable], profile=data.PROFILES["DI"])
        self.pages[PAGE_MASTER_PO] = DataPage(parent, self, "Master PO", "Warna kolom BL = BL yang sama. Klik warna POD atau tombol di judul kolom untuk filter.",
                                              data.MASTER_PO_COLUMNS, "po_master", lambda r: self.store.master_flag(r, data.PROFILES["PO"]),
                                              [c.key for c in data.MASTER_PO_COLUMNS if c.editable], profile=data.PROFILES["PO"])
        self.pages[PAGE_GOODS] = DataPage(parent, self, "Goods Description", "Klik tombol di judul kolom untuk filter.",
                                          data.GOODS_COLUMNS, "items", lambda r: self.store.goods_flag(r), ["A", "B", "C"])
        self.pages[PAGE_WEIGHT] = DataPage(parent, self, "CBM & Berat", "Per 1 PCS. Pilih CBM tidak tally untuk melihat yang beda dengan file loading.", data.WEIGHT_COLUMNS, "items",
                                           lambda r: self.store.weight_flag(r), ["A", "D", "E", "F"], extra_modes=[FILTER_TALLY])
        self.pages[PAGE_PRICE] = DataPage(parent, self, "Harga", "⚠ DI/PO + SKU dengan lebih dari satu harga. Klik tombol di judul kolom untuk filter.",
                                          data.PRICE_COLUMNS, "prices", lambda r: self.store.price_flag(r), ["G", "H", "D", "E", "F"])
        self.pages[PAGE_BUILD] = BuildPage(parent, self)
        self.pages[PAGE_BUILD_PO] = BuildPage(parent, self, data.PROFILES["PO"])
        del store

    def notify(self, title, subtitle="", kind="ok", action=None, action_text=""):
        toast = getattr(self.ui, "toast", None)
        if toast is None:
            return
        try:
            toast.show(title, subtitle, kind, action=action, action_text=action_text)
        except Exception:
            pass

    @staticmethod
    def keys_of(rows):
        return {data.unique_key(row.get("A")) for row in rows if data.unique_key(row.get("A"))}

    def upload_summary(self, table, path, old, store):
        name = os.path.basename(path)
        if table in ("master", "po_master"):
            rows = store.master if table == "master" else store.po_master
            label = "DI" if table == "master" else "PO"
            new = len(self.keys_of(rows) - self.keys_of(old.master if table == "master" else old.po_master)) if old.loaded else 0
            extra = f" · {new:,} {label} baru" if new else ""
            return f"Master {label} berhasil diperbarui", f"{name} · {len(rows):,} {label}{extra}"
        if table == "items":
            return "File data berhasil diperbarui", f"{name} · {len(store.items):,} SKU · {len(store.prices):,} harga"
        return "File loading berhasil dimuat", f"{name} · {len(store.cbm_tally()):,} CBM tidak tally"

    def template_path(self, profile=None):
        name = profile.template if profile else "di_template.xlsx"
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)

    def ensure_sources(self):
        settings = data.load_settings()
        for key, title in (("master_path", "Pilih file Master Data DI"), ("item_path", "Pilih file Goods, CBM, Berat, dan Harga")):
            if not settings.get(key) or not os.path.exists(settings[key]):
                path = filedialog.askopenfilename(title=title, filetypes=[("File Excel", "*.xlsx *.xlsm *.xls")])
                if not path:
                    return None
                settings[key] = os.path.normpath(path)
                dialog = ctk.CTkInputDialog(title="Password", text=f"Password {os.path.basename(path)}\nKosongkan kalau tidak ada.")
                settings[key.replace("_path", "_password")] = dialog.get_input() or ""
        data.save_settings(settings)
        return settings

    def choose_sources(self):
        settings = data.load_settings()
        settings.pop("master_path", None)
        settings.pop("item_path", None)
        data.save_settings(settings)
        self.reload(force=True)

    def dirty(self):
        return sum(self.store.dirty_count(table) for table in ("master", "po_master", "items", "prices"))

    def ensure_loaded(self):
        if not self.store.loaded and not self.running:
            self.reload(force=True)

    def reload(self, force=False, notify=False):
        if self.running:
            return
        upload = getattr(self, "pending_upload", None)
        self.pending_upload = None
        old_store = self.store
        if not force and self.dirty() and not messagebox.askyesno("Dockit", f"{self.dirty()} perubahan belum disimpan dan akan hilang. Lanjut?"):
            return
        settings = self.ensure_sources()
        if not settings:
            self.ui.status.configure(text="Pilih file sumber dulu untuk membuka data DI.")
            return
        store = data.Store(
            settings["master_path"], settings["item_path"], settings.get("master_password", ""),
            settings.get("item_password", ""), settings.get("load_path", ""),
            settings.get("po_master_path", ""), settings.get("po_master_password", ""),
        )
        self.ui.status.configure(text="Memuat data DI...")

        def done(_, error):
            if error:
                detail = f"Data gagal dimuat. Pastikan filenya benar dan password-nya sesuai.\n\n{error}"
                if upload:
                    self.notify("Upload gagal, data tidak diubah", f"{os.path.basename(upload[1])} · cek file atau password", "error",
                                action=lambda: messagebox.showerror("Dockit", detail), action_text="Detail")
                else:
                    messagebox.showerror("Dockit", detail)
                self.ui.status.configure(text="Data DI gagal dimuat.")
                if self.previous_settings is not None:
                    data.save_settings(self.previous_settings)
                    self.previous_settings = None
                return
            self.previous_settings = None
            self.store = store
            self.refresh_all()
            self.ui.status.configure(text=f"Data DI dimuat: {len(store.master)} DI, {len(store.items)} SKU, {len(store.po_master)} PO, {len(store.prices)} harga, {len(store.cbm_tally())} CBM tidak tally.")
            if upload:
                self.notify(*self.upload_summary(upload[0], upload[1], old_store, store))
            elif notify:
                self.notify("Data diperbarui", f"{len(store.po_master):,} PO · {len(store.master):,} DI · {len(store.items):,} SKU · {len(store.prices):,} harga")

        self.run(store.load, done)

    def auto_fix(self, skus=None, kinds=("cbm", "typo")):
        """Samakan CBM dengan file loading dan perbaiki typo deskripsi, lalu simpan ke file Goods/CBM."""
        if self.running:
            messagebox.showinfo("Dockit", "Proses masih berjalan. Coba lagi sebentar.")
            return
        fixes, manual = self.store.auto_fixes(skus, kinds)
        note = ""
        if manual:
            shown = ", ".join(manual[:5]) + (f", dan {len(manual) - 5} lainnya" if len(manual) > 5 else "")
            note = f"\n\nTidak bisa diperbaiki otomatis karena tidak ada kata inch di deskripsi, perbaiki manual: {shown}."
        if not fixes:
            messagebox.showinfo("Dockit", "Tidak ada CBM atau deskripsi yang bisa diperbaiki otomatis." + note)
            return
        lines = []
        for fix in fixes[:12]:
            if fix["key"] == "D":
                old = "kosong" if data.is_blank(fix["old"]) else data.format_value(fix["old"], "num")
                lines.append(f"• {fix['sku']}: CBM {old} → {data.format_value(fix['new'], 'num')}")
            else:
                lines.append(f"• {fix['sku']}: \"{' '.join(str(fix['old']).split())}\" → \"{' '.join(fix['new'].split())}\"")
        if len(fixes) > 12:
            lines.append(f"• dan {len(fixes) - 12} perubahan lainnya")
        cbm = sum(1 for f in fixes if f["key"] == "D")
        typo = len(fixes) - cbm
        summary = ", ".join(part for part in (f"{cbm} CBM disamakan dengan file loading" if cbm else "", f"{typo} deskripsi disamakan dengan SKU" if typo else "") if part)
        if not messagebox.askyesno("Dockit", f"{summary}:\n\n" + "\n".join(lines) + note + "\n\nLangsung simpan ke file Goods, CBM, dan Harga? Backup dibuat otomatis."):
            return
        self.store.apply_fixes(fixes)
        self.save("items")

    def import_weights(self):
        """Baca file hasil timbang dan update NW/GW dengan berat paling kecil per SKU."""
        if self.running:
            messagebox.showinfo("Dockit", "Proses masih berjalan. Coba lagi sebentar.")
            return
        path = filedialog.askopenfilename(title="Pilih file hasil timbang (SKU, NW, GW)", filetypes=[("File Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        self.ui.status.configure(text="Membaca file timbang...")
        self.ui.update_idletasks()
        try:
            weights = data.read_weight_file(path)
        except Exception as error:
            messagebox.showerror("Dockit", f"File timbang tidak bisa dibaca.\n\n{error}")
            self.ui.status.configure(text="File timbang gagal dibaca.")
            return
        if not weights:
            messagebox.showinfo("Dockit", "Tidak ditemukan kolom SKU, NET WEIGHT, dan GROSS WEIGHT di file ini.")
            self.ui.status.configure(text="")
            return
        self.ui.status.configure(text=f"File timbang: {len(weights):,} SKU.")
        WeightDialog(self, os.path.basename(path), weights)

    def save(self, table):
        if self.running:
            return
        job = {"master": self.store.save_master, "po_master": self.store.save_po}.get(table, self.store.save_items)
        self.ui.status.configure(text="Menyimpan ke file sumber...")

        def done(count, error):
            if error:
                messagebox.showerror("Dockit", f"Belum tersimpan.\n\n{error}")
                self.ui.status.configure(text="Perubahan belum tersimpan.")
            else:
                self.ui.status.configure(text=f"{count} baris tersimpan ke file sumber.")
                names = {"master": "Master DI", "po_master": "Master PO"}
                self.notify(f"{names.get(table, 'File data')} berhasil disimpan", f"{count} baris tersimpan ke file sumber · backup dibuat otomatis")
            self.refresh_all()

        self.run(job, done)

    def delete(self, table, rows):
        if self.running:
            return
        self.ui.status.configure(text="Menghapus baris dari file sumber...")

        def done(count, error):
            if error:
                messagebox.showerror("Dockit", f"Baris belum terhapus.\n\n{error}")
                self.ui.status.configure(text="Baris belum terhapus.")
            else:
                self.ui.status.configure(text=f"{count} baris terhapus dari file sumber.")
            self.refresh_all()

        self.run(lambda app: self.store.delete_rows(app, table, rows), done)

    def choose_file(self, table):
        title = {
            "master": "Pilih file Master Data DI", "po_master": "Pilih file Master Data PO", "items": "Pilih file Goods, CBM, Berat, dan Harga",
            "loading": "Pilih file 装柜明细清单 (sheet DI-柜)",
        }[table]
        path = filedialog.askopenfilename(title=title, filetypes=[("File Excel", "*.xlsx *.xlsm *.xls")])
        if path:
            self.set_source(table, os.path.normpath(path))

    def set_source(self, table, path):
        if self.running:
            messagebox.showinfo("Dockit", "Proses masih berjalan. Coba lagi sebentar.")
            return
        detected = data.detect_source(path)
        if detected == "":
            detail = (f"{os.path.basename(path)} bukan file sumber yang dikenali.\n\n"
                      f"Master DI: sheet {data.MASTER_SHEET}\nMaster PO: sheet {data.MASTER_PO_SHEET}\nFile data: sheet {data.ITEM_SHEET}, {data.WEIGHT_SHEET}, {data.PRICE_SHEET}\n"
                      f"File loading: sheet {data.LOADING_SHEET}")
            self.notify("Upload gagal, file tidak dikenali", f"{os.path.basename(path)} · sheet yang dibutuhkan tidak ada", "error",
                        action=lambda: messagebox.showwarning("Dockit", detail), action_text="Detail")
            return
        # Jenis file mengikuti isinya, bukan tab yang sedang terbuka.
        table = detected or table
        if self.dirty() and not messagebox.askyesno("Dockit", f"{self.dirty()} perubahan belum disimpan dan akan hilang. Lanjut?"):
            return
        previous = data.load_settings()
        settings = dict(previous)
        key = {"master": "master", "po_master": "po_master", "items": "item", "loading": "load"}[table]
        if key != "load" and settings.get(f"{key}_path") != path:
            dialog = ctk.CTkInputDialog(title="Password", text=f"Password {os.path.basename(path)}\nKosongkan kalau tidak ada.")
            settings[f"{key}_password"] = dialog.get_input() or ""
        settings[f"{key}_path"] = path
        data.save_settings(settings)
        self.previous_settings = previous
        self.pending_upload = (table, path)
        self.reload(force=True)

    def handle_drop(self, paths):
        """Dipanggil saat file diseret ke aplikasi."""
        page = self.ui.current_page
        tables = {PAGE_MASTER: "master", PAGE_MASTER_PO: "po_master", PAGE_GOODS: "items", PAGE_WEIGHT: "items", PAGE_PRICE: "items", PAGE_BUILD: "loading"}
        if page not in tables:
            return False
        excel = [p for p in paths if p.lower().endswith((".xlsx", ".xlsm", ".xls"))]
        if len(excel) != 1:
            messagebox.showinfo("Dockit", "Seret satu file Excel saja ke tab ini.")
            return True
        self.set_source(tables[page], os.path.normpath(excel[0]))
        return True

    def run(self, job, callback):
        self.running = True

        def worker():
            import pythoncom

            pythoncom.CoInitialize()
            engine = None
            result = error = None
            try:
                engine, _ = self.ui.engine_starter()
                result = job(engine)
            except Exception as exc:
                error = str(exc) or type(exc).__name__
            finally:
                self.ui.engine_stopper(engine)
                pythoncom.CoUninitialize()
                self.ui.post_event("di_job", (self.finish, callback, result, error))

        threading.Thread(target=worker, daemon=True).start()

    def finish(self, callback, result, error):
        self.running = False
        callback(result, error)

    def refresh_all(self, exclude=None):
        for page in self.pages.values():
            if page is not exclude:
                page.refresh()

    def open_generated(self, paths, with_pdf):
        ui = self.ui
        ui.show_page(PAGE_PDF)
        normal = [os.path.normpath(path) for path in paths]
        ui.convert_after_read = with_pdf
        ui.convert_only = set(normal) if with_pdf else None
        known = [path for path in normal if ui.file_by_path(path) is not None]
        new = [path for path in normal if ui.file_by_path(path) is None]
        for path in known:
            ui.read_queue.put(path)
        if new:
            ui.add_paths(new)
        elif known:
            ui.start_reader()
            ui.update_controls()
