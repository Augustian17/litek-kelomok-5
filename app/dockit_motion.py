"""Animasi UI Dockit V1.26: notifikasi Outline Draw, progress Pixel Dissolve, refresh Sync Spin,
menu stagger 65 ms, dan search bar dengan saran Blur Focus.
Tkinter tidak punya blur atau transparansi, jadi blur dan fade dibuat dengan transisi warna dan geser kecil."""
import math
import random
import tkinter as tk

import customtkinter as ctk
from PIL import Image

try:
    import dockit_icons as icons
except ImportError:  # dijalankan sebagai paket
    from . import dockit_icons as icons

FONT_FAMILY = "Segoe UI"
FRAME_MS = 16


# ------------------------------------------------------------------ warna
def dark():
    return ctk.get_appearance_mode() == "Dark"


def pick(color):
    if isinstance(color, (tuple, list)):
        return color[1] if dark() else color[0]
    return color


def rgb(color):
    color = pick(color).lstrip("#")
    return tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))


def mix(a, b, t):
    t = max(0.0, min(1.0, t))
    ra, rb = rgb(a), rgb(b)
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * t) for x, y in zip(ra, rb))


def ease_out(t):
    return 1 - (1 - t) ** 3


def ease_in(t):
    return t ** 3


def scaling(widget):
    try:
        return ctk.ScalingTracker.get_widget_scaling(widget)
    except Exception:
        return 1.0


def tk_font(widget, size, weight="normal"):
    return (FONT_FAMILY, -round(size * scaling(widget)), weight)


BG = ("#F8F0E5", "#0F1420")
CARD = ("#FFFFFF", "#1A2030")
LINE = ("#E0D2BF", "#2E3548")
TEXT = ("#16202E", "#EEF0F5")
MUTED = ("#5E6675", "#A0A8B8")
NAVY = ("#0F2C59", "#3B5F9E")
SAND = "#DAC0A3"
HIT = ("#F6E3B4", "#4A3B12")

TOAST_KINDS = {
    "ok": {"bg": ("#0F2C59", "#16305C"), "line": "#DAC0A3", "icon": "#9CDBB8", "mark": "✓"},
    "info": {"bg": ("#0F2C59", "#16305C"), "line": "#DAC0A3", "icon": "#DAC0A3", "mark": "i"},
    "warn": {"bg": ("#3A2A10", "#3A2A10"), "line": "#F2B544", "icon": "#F2B544", "mark": "!"},
    "error": {"bg": ("#4A1712", "#4A1712"), "line": "#F08A7E", "icon": "#F08A7E", "mark": "!"},
}


def rounded_path(x1, y1, x2, y2, r, steps=8):
    """Titik keliling kotak bulat, mulai dari kiri atas searah jarum jam."""
    points = []
    corners = [(x2 - r, y1 + r, -90), (x2 - r, y2 - r, 0), (x1 + r, y2 - r, 90), (x1 + r, y1 + r, 180)]
    edges = [((x1 + r, y1), (x2 - r, y1)), ((x2, y1 + r), (x2, y2 - r)), ((x2 - r, y2), (x1 + r, y2)), ((x1, y2 - r), (x1, y1 + r))]
    for (start, end), (cx, cy, angle) in zip(edges, corners):
        length = math.hypot(end[0] - start[0], end[1] - start[1])
        count = max(2, int(length / 6))
        for i in range(count):
            t = i / count
            points.append((start[0] + (end[0] - start[0]) * t, start[1] + (end[1] - start[1]) * t))
        for i in range(steps + 1):
            a = math.radians(angle + 90 * i / steps)
            points.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    points.append(points[0])
    return points


def flat(points):
    return [v for p in points for v in p]


# --------------------------------------------------------------- notifikasi
class Toast:
    """Notifikasi pojok kanan bawah. Garis menggambar kotak, isi muncul, lalu teks tersapu masuk."""

    W, H = 430, 70

    def __init__(self, root):
        self.root = root
        self.canvas = None
        self.jobs = []
        self.closing = False
        self.hover = False
        self.close_job = None

    # dipanggil dari mana saja di app
    def show(self, title, subtitle="", kind="ok", action=None, action_text="", seconds=4.5):
        self.stop()
        s = scaling(self.root)
        w, h = round(self.W * s), round(self.H * s)
        style = TOAST_KINDS.get(kind, TOAST_KINDS["ok"])
        self.style, self.action, self.seconds = style, action, seconds
        self.closing = False
        if self.canvas is None or not self.canvas.winfo_exists():
            self.canvas = tk.Canvas(self.root, highlightthickness=0, bd=0, cursor="arrow")
            self.canvas.bind("<Enter>", self.enter)
            self.canvas.bind("<Leave>", self.leave)
        c = self.canvas
        c.delete("all")
        c.configure(width=w + 4, height=h + 4, bg=pick(BG))
        c.place(relx=1.0, rely=1.0, anchor="se", x=-round(26 * s), y=-round(26 * s))
        c.tk.call("raise", c._w)
        self.w, self.h, self.s = w, h, s
        r = round(12 * s)
        self.path = rounded_path(2, 2, w + 2, h + 2, r)
        bg = pick(style["bg"])
        self.fill = c.create_polygon(flat(self.path), fill=pick(BG), outline="", smooth=False)
        self.line = c.create_line(0, 0, 0, 0, fill=style["line"], width=max(2, round(2 * s)), capstyle="round", joinstyle="round")
        cx, cy = 2 + round(30 * s), 2 + h // 2
        rad = round(13 * s)
        items = [
            c.create_oval(cx - rad, cy - rad, cx + rad, cy + rad, outline=style["icon"], width=max(2, round(2 * s))),
        ]
        if style["mark"] == "✓":
            u = rad / 13
            items.append(c.create_line(cx - 6 * u, cy, cx - 1.5 * u, cy + 4.5 * u, cx + 6.5 * u, cy - 4.5 * u, fill=style["icon"],
                                       width=max(2, round(2.4 * s)), capstyle="round", joinstyle="round"))
        else:
            items.append(c.create_text(cx, cy, text=style["mark"], fill=style["icon"], font=tk_font(self.root, 14, "bold")))
        x = 2 + round(56 * s)
        right = w - round(46 * s)
        if action and action_text:
            link = c.create_text(right, cy, text=action_text, anchor="e", fill=style["line"], font=tk_font(self.root, 12, "bold"))
            c.tag_bind(link, "<Button-1>", lambda _e: self.run_action())
            c.tag_bind(link, "<Enter>", lambda _e: c.configure(cursor="hand2"))
            c.tag_bind(link, "<Leave>", lambda _e: c.configure(cursor="arrow"))
            items.append(link)
            right = c.bbox(link)[0] - round(12 * s)
        title_font, sub_font = tk_font(self.root, 14, "bold"), tk_font(self.root, 12)
        room = right - x
        items.append(c.create_text(x, cy - round(10 * s), text=self.fit(title, title_font, room), anchor="w", fill="#FFFFFF", font=title_font))
        if subtitle:
            items.append(c.create_text(x, cy + round(11 * s), text=self.fit(subtitle, sub_font, room), anchor="w", fill="#EADBC8", font=sub_font))
        close = c.create_text(w - round(18 * s), cy, text="✕", fill="#FFFFFF", font=tk_font(self.root, 13, "bold"))
        c.tag_bind(close, "<Button-1>", lambda _e: self.close())
        c.tag_bind(close, "<Enter>", lambda _e: c.configure(cursor="hand2"))
        c.tag_bind(close, "<Leave>", lambda _e: c.configure(cursor="arrow"))
        items.append(close)
        if action and not action_text:
            for item in items[:-1] + [self.fill]:
                c.tag_bind(item, "<Button-1>", lambda _e: self.run_action())
        self.items = items
        for item in items:
            c.itemconfigure(item, state="hidden")
        self.cover = c.create_rectangle(x - 4, 4, w, h, fill=bg, outline="", state="hidden")
        self.text_x = (x - 4, w)
        self.animate(560, self.draw_step, lambda: None)
        self.later(380, lambda: self.animate(320, self.fill_step, lambda: None))
        self.later(540, self.show_text)
        self.schedule_close()

    def fit(self, text, font, room):
        try:
            import tkinter.font as tkfont

            f = tkfont.Font(root=self.root, font=font)
        except Exception:
            return text
        if f.measure(text) <= room:
            return text
        while text and f.measure(text + "…") > room:
            text = text[:-1]
        return text + "…"

    def draw_step(self, t):
        n = max(2, int(len(self.path) * ease_out(t)))
        self.canvas.coords(self.line, flat(self.path[:n]))

    def fill_step(self, t):
        self.canvas.itemconfigure(self.fill, fill=mix(BG, self.style["bg"], ease_out(t)))

    def show_text(self):
        c = self.canvas
        for item in self.items:
            c.itemconfigure(item, state="normal")
        c.itemconfigure(self.cover, state="normal")
        c.tag_raise(self.cover)
        x0, x1 = self.text_x

        def wipe(t):
            c.coords(self.cover, x0 + (x1 - x0) * ease_out(t), 4, x1, self.h)

        self.animate(380, wipe, lambda: c.itemconfigure(self.cover, state="hidden"))

    def schedule_close(self):
        if self.close_job:
            self.root.after_cancel(self.close_job)
        self.close_job = self.root.after(int(self.seconds * 1000), self.auto_close)

    def auto_close(self):
        self.close_job = None
        if self.hover:
            self.close_job = self.root.after(800, self.auto_close)
            return
        self.close()

    def enter(self, _event=None):
        self.hover = True

    def leave(self, _event=None):
        self.hover = False

    def run_action(self):
        action = self.action
        self.close()
        if action:
            self.root.after(10, action)

    def close(self):
        if self.closing or self.canvas is None:
            return
        self.closing = True
        self.cancel_jobs()
        c = self.canvas
        x0, x1 = self.text_x
        c.itemconfigure(self.cover, state="normal", fill=pick(self.style["bg"]))
        c.tag_raise(self.cover)

        def unwipe(t):
            c.coords(self.cover, x0, 4, x0 + (x1 - x0) * ease_out(t), self.h)

        def undraw(t):
            n = int(len(self.path) * ease_in(t))
            c.coords(self.line, flat(self.path[n:]) if len(self.path) - n > 1 else (0, 0, 0, 0))
            c.itemconfigure(self.fill, fill=mix(self.style["bg"], BG, ease_out(min(1, t * 1.6))))
            c.itemconfigure(self.cover, fill=mix(self.style["bg"], BG, ease_out(min(1, t * 1.6))))

        def gone():
            c.place_forget()
            self.closing = False

        self.animate(240, unwipe, lambda: self.animate(420, undraw, gone))

    def stop(self):
        self.cancel_jobs()
        if self.close_job:
            self.root.after_cancel(self.close_job)
            self.close_job = None

    def cancel_jobs(self):
        for job in self.jobs:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
        self.jobs = []

    def later(self, ms, fn):
        self.jobs.append(self.root.after(ms, fn))

    def animate(self, ms, step, done):
        frames = max(1, ms // FRAME_MS)

        def tick(i):
            if self.canvas is None or not self.canvas.winfo_exists():
                return
            step(i / frames)
            if i < frames:
                self.jobs.append(self.root.after(FRAME_MS, lambda: tick(i + 1)))
            else:
                done()

        tick(0)


# ---------------------------------------------------------------- progress
class PixelBar(ctk.CTkFrame):
    """Pengganti CTkProgressBar. Bagian yang baru terisi muncul sebagai piksel kecil yang berkilau."""

    def __init__(self, master, height=12, track=("#EFE6DA", "#283044"), fill=NAVY, sparkle=SAND, **kwargs):
        super().__init__(master, fg_color="transparent", height=height, **kwargs)
        self.track, self.fill_color, self.sparkle = track, fill, sparkle
        self.height_px = height
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, height=round(height * scaling(self)))
        self.canvas.pack(fill="x", expand=True)
        self.canvas.bind("<Configure>", lambda _e: self.redraw())
        self.value = 0.0
        self.shown_cols = 0
        self.cells = {}
        self.jobs = []
        self.seed = random.Random(7)

    def geometry(self):
        s = scaling(self)
        cell = max(3, round(4 * s))
        width = max(1, self.canvas.winfo_width())
        height = round(self.height_px * s)
        rows = max(1, (height - 2) // cell)
        cols = max(1, (width - 2) // cell)
        return cell, rows, cols, width, height

    def redraw(self):
        c = self.canvas
        c.delete("all")
        try:
            parent_bg = self.master.cget("fg_color")
            c.configure(bg=pick(parent_bg) if parent_bg != "transparent" else pick(CARD))
        except Exception:
            c.configure(bg=pick(CARD))
        cell, rows, cols, width, height = self.geometry()
        r = height // 2
        c.create_polygon(flat(rounded_path(0, 0, width - 1, height - 1, min(r, height // 2), 6)), fill=pick(self.track), outline="", smooth=False)
        self.cells = {}
        self.shown_cols = int(cols * self.value)
        for col in range(self.shown_cols):
            for row in range(rows):
                self.draw_cell(col, row, 1.0)

    def draw_cell(self, col, row, t):
        cell, rows, cols, width, height = self.geometry()
        top = (height - rows * cell) // 2
        x, y = 1 + col * cell, top + row * cell
        inset = (cell / 2) * (1 - min(1, t * 1.6))
        color = mix(self.sparkle, self.fill_color, max(0, (t - 0.45) / 0.55)) if t < 1 else pick(self.fill_color)
        if col * 7 + row * 3 in (0,) or (t >= 1 and (col * 13 + row * 7) % 11 == 0):
            color = mix(self.fill_color, "#3B6BB5", 0.6)
        key = (col, row)
        if key in self.cells:
            self.canvas.coords(self.cells[key], x + inset, y + inset, x + cell - 0.5 - inset, y + cell - 0.5 - inset)
            self.canvas.itemconfigure(self.cells[key], fill=color)
        else:
            self.cells[key] = self.canvas.create_rectangle(x + inset, y + inset, x + cell - 0.5 - inset, y + cell - 0.5 - inset, fill=color, outline="")

    def set(self, value):
        value = max(0.0, min(1.0, float(value)))
        if value < self.value - 1e-6:
            self.value = value
            for job in self.jobs:
                try:
                    self.after_cancel(job)
                except Exception:
                    pass
            self.jobs = []
            self.redraw()
            return
        self.value = value
        cell, rows, cols, width, height = self.geometry()
        target = int(cols * value)
        start = self.shown_cols
        if target <= start:
            return
        self.shown_cols = target
        span = max(1, target - start)
        for col in range(start, target):
            for row in range(rows):
                delay = int(((col - start) / span) * 320 + self.seed.random() * 220)
                self.jobs.append(self.after(delay, lambda c=col, r=row: self.pop(c, r, 0)))

    def get(self):
        return self.value

    def pop(self, col, row, frame):
        frames = 7
        self.draw_cell(col, row, frame / frames)
        if frame < frames:
            self.jobs.append(self.after(FRAME_MS * 2, lambda: self.pop(col, row, frame + 1)))


# ------------------------------------------------------------- sync spin
class SyncSpin:
    """Tombol refresh ikon saja: berputar selama proses berjalan, lalu berubah jadi centang."""

    def __init__(self, button, busy, color=NAVY, done_color=("#14795C", "#5FD3A0"), size=18):
        self.button, self.busy, self.size = button, busy, size
        self.color, self.done_color = color, done_color
        self.frames = self.rotations("refresh-cw", color, 24)
        self.check = self.still("circle-check", done_color)
        self.base_fg = button.cget("fg_color")
        self.spinning = False
        button.configure(image=self.frames[0], text="")

    def mask(self, name):
        try:
            return icons.tiles(name)[0]
        except Exception:
            image = Image.new("L", (64, 64), 0)
            return image

    def rotations(self, name, color, count):
        mask = self.mask(name)
        light, dark_ = color if isinstance(color, (tuple, list)) else (color, color)
        frames = []
        for i in range(count):
            m = mask.rotate(-360 * i / count, resample=Image.BICUBIC)
            frames.append(ctk.CTkImage(light_image=icons.tinted(m, light), dark_image=icons.tinted(m, dark_), size=(self.size, self.size)))
        return frames

    def still(self, name, color):
        mask = self.mask(name)
        light, dark_ = color if isinstance(color, (tuple, list)) else (color, color)
        return ctk.CTkImage(light_image=icons.tinted(mask, light), dark_image=icons.tinted(mask, dark_), size=(self.size + 2, self.size + 2))

    def start(self):
        if self.spinning:
            return
        self.spinning = True
        self.started = 0
        self.spin(0)

    def spin(self, i):
        # minimal 2 putaran (0,9 detik), lalu terus berputar sampai proses selesai
        if i >= 2 * len(self.frames) and not self.busy():
            self.finish()
            return
        eased = i if i >= 2 * len(self.frames) else int(2 * len(self.frames) * (1 - (1 - i / (2 * len(self.frames))) ** 2))
        self.button.configure(image=self.frames[eased % len(self.frames)])
        self.button.after(38, lambda: self.spin(i + 1))

    def finish(self):
        self.button.configure(image=self.check, fg_color=("#E6F2EA", "#1C3A2C"))
        self.button.after(1400, self.reset)

    def reset(self):
        self.spinning = False
        self.button.configure(image=self.frames[0], fg_color=self.base_fg)


# --------------------------------------------------------------- menu
class StaggerMenu:
    """Menu tarik-turun. Kartu turun sedikit, tiap baris menyusul 65 ms."""

    STAGGER = 65
    ROW_H = 40

    def __init__(self, anchor, items, width=270):
        """items: list (teks, keterangan, fungsi) atau None untuk garis pemisah. Teks diawali '!' berwarna merah."""
        self.anchor, self.items, self.width = anchor, items, width
        self.top = None

    def toggle(self):
        if self.top is not None and self.top.winfo_exists():
            self.close()
        else:
            self.open()

    def open(self):
        root = self.anchor.winfo_toplevel()
        s = scaling(self.anchor)
        top = tk.Toplevel(root)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        self.top = top
        rows = [i for i in self.items]
        height = sum(9 if i is None else self.ROW_H for i in rows) + 12
        w, h = round(self.width * s), round(height * s)
        x = self.anchor.winfo_rootx() + self.anchor.winfo_width() - w
        y = self.anchor.winfo_rooty() + self.anchor.winfo_height() + round(6 * s)
        top.geometry(f"{w}x{h}+{x}+{y - round(6 * s)}")
        top.configure(bg=pick(LINE))
        card = tk.Frame(top, bg=pick(CARD))
        card.place(x=1, y=1, width=w - 2, height=h - 2)
        self.card = card
        self.widgets = []
        offset = round(6 * s)
        for index, item in enumerate(rows):
            if item is None:
                line = tk.Frame(card, bg=pick(CARD), height=1)
                line.place(x=round(10 * s), y=offset + round(4 * s), width=w - round(22 * s), height=1)
                self.widgets.append((line, offset + round(4 * s), None))
                offset += round(9 * s)
                continue
            text, hint, fn = item
            danger = text.startswith("!")
            text = text.lstrip("!")
            row = tk.Frame(card, bg=pick(CARD), cursor="hand2")
            row.place(x=round(6 * s), y=offset, width=w - round(14 * s), height=round((self.ROW_H - 2) * s))
            label = tk.Label(row, text=text, bg=pick(CARD), fg=pick(CARD), font=tk_font(self.anchor, 13), anchor="w")
            label.pack(side="left", padx=round(10 * s))
            note = tk.Label(row, text=hint or "", bg=pick(CARD), fg=pick(CARD), font=tk_font(self.anchor, 11), anchor="e")
            note.pack(side="right", padx=round(10 * s))
            final = ("#9A3412" if not dark() else "#F2A65A") if danger else pick(TEXT)
            for widget in (row, label, note):
                widget.bind("<Enter>", lambda _e, r=row, l=label, n=note: [x.configure(bg=pick(("#F3EADF", "#283044"))) for x in (r, l, n)])
                widget.bind("<Leave>", lambda _e, r=row, l=label, n=note: [x.configure(bg=pick(CARD)) for x in (r, l, n)])
                widget.bind("<Button-1>", lambda _e, f=fn: self.pick(f))
            self.widgets.append((row, offset, (label, note, final)))
            offset += round(self.ROW_H * s)
        # kartu turun 6 px
        self.slide(top, x, y, s, 0)
        for order, (widget, y0, parts) in enumerate(self.widgets):
            widget.after(40 + order * self.STAGGER, lambda wd=widget, yy=y0, p=parts: self.enter_row(wd, yy, p, 0))
        root.bind("<Button-1>", self.outside, add="+")
        top.bind("<Escape>", lambda _e: self.close())
        self.root_bind = root
        top.focus_set()

    def slide(self, top, x, y, s, i):
        frames = 14
        if not top.winfo_exists():
            return
        t = ease_out(i / frames)
        top.geometry(f"+{x}+{round(y - 6 * s * (1 - t))}")
        if i < frames:
            top.after(FRAME_MS, lambda: self.slide(top, x, y, s, i + 1))

    def enter_row(self, widget, y0, parts, i):
        frames = 16
        if not widget.winfo_exists():
            return
        t = ease_out(i / frames)
        s = scaling(self.anchor)
        info = widget.place_info()
        widget.place_configure(y=round(y0 - 8 * s * (1 - t)))
        if parts:
            label, note, final = parts
            label.configure(fg=mix(CARD, final, t))
            note.configure(fg=mix(CARD, MUTED, t))
        else:
            widget.configure(bg=mix(CARD, LINE, t))
        if i < frames:
            widget.after(FRAME_MS, lambda: self.enter_row(widget, y0, parts, i + 1))

    def outside(self, event):
        if self.top is None or not self.top.winfo_exists():
            return
        widget = event.widget
        try:
            if str(widget).startswith(str(self.anchor)) or str(widget).startswith(str(self.top)):
                return
        except Exception:
            pass
        self.close()

    def pick(self, fn):
        self.close()
        if fn:
            self.anchor.after(10, fn)

    def close(self):
        if self.top is not None and self.top.winfo_exists():
            self.top.destroy()
        self.top = None


# ------------------------------------------------------------- search bar
class SearchBar(ctk.CTkFrame):
    """Kotak cari dengan ikon, tombol hapus, dan saran yang muncul dari samar menjadi jelas."""

    def __init__(self, master, variable, placeholder, suggest=None, width=340, height=42):
        super().__init__(master, fg_color=CARD, border_color=LINE, border_width=1, corner_radius=10, width=width, height=height)
        self.variable, self.suggest = variable, suggest
        self.pack_propagate(False)
        self.icon = ctk.CTkLabel(self, text="", width=20, height=20, fg_color=CARD)
        self.icon.pack(side="left", padx=(12, 4), pady=4)
        try:
            self.icon_anim = icons.Animator(self.icon, "search", MUTED, 16, hover_widgets=(self,))
        except Exception:
            self.icon.configure(text="⌕")
        self.entry = ctk.CTkEntry(self, textvariable=variable, border_width=0, fg_color=CARD, bg_color=CARD, corner_radius=0,
                                  font=ctk.CTkFont(family=FONT_FAMILY, size=14), height=height - 12, text_color=TEXT)
        self.entry.pack(side="left", fill="x", expand=True, pady=4)
        # CTkEntry tidak menampilkan placeholder kalau memakai textvariable, jadi dibuat sendiri
        self.hint = ctk.CTkLabel(self.entry, text=placeholder, font=ctk.CTkFont(family=FONT_FAMILY, size=14), text_color=("#9AA0AB", "#6B7385"),
                                 fg_color=CARD, anchor="w", height=height - 16)
        self.hint.place(x=6, rely=0.5, anchor="w")
        self.hint.bind("<Button-1>", lambda _e: self.entry.focus_set())
        self.kbd = ctk.CTkLabel(self, text="Ctrl+F", font=ctk.CTkFont(family=FONT_FAMILY, size=11), text_color=MUTED,
                                fg_color=("#F3EADF", "#283044"), corner_radius=5, width=46, height=20)
        self.kbd.pack(side="right", padx=(4, 10))
        self.clear_button = ctk.CTkButton(self, text="✕", width=26, height=26, corner_radius=13, fg_color=("#F3EADF", "#283044"),
                                          hover_color=("#E5D8C6", "#323B52"), text_color=MUTED, font=ctk.CTkFont(size=12), command=self.clear)
        self.popup = None
        self.rows = []
        self.active = -1
        self.job = None
        self.picked = False
        variable.trace_add("write", lambda *_: self.changed())
        self.entry.bind("<FocusIn>", lambda _e: self.focus_style(True))
        self.entry.bind("<FocusOut>", lambda _e: (self.focus_style(False), self.after(180, self.hide)))
        self.entry.bind("<Down>", lambda _e: self.move(1))
        self.entry.bind("<Up>", lambda _e: self.move(-1))
        self.entry.bind("<Return>", lambda _e: self.choose(self.active if self.active >= 0 else 0))
        self.entry.bind("<Escape>", lambda _e: self.hide() if self.popup else self.clear())

    def focus(self):
        self.entry.focus_set()
        self.entry.select_range(0, "end")

    def focus_style(self, on):
        self.configure(border_color=NAVY if on or self.variable.get() else LINE)

    def clear(self):
        self.variable.set("")
        self.hide()

    def changed(self):
        text = self.variable.get()
        if text:
            self.hint.place_forget()
        else:
            self.hint.place(x=6, rely=0.5, anchor="w")
        if text:
            self.kbd.pack_forget()
            self.clear_button.pack(side="right", padx=(4, 8))
        else:
            self.clear_button.pack_forget()
            self.kbd.pack(side="right", padx=(4, 10))
        try:
            focused = self.focus_get() == self.entry._entry
        except Exception:
            focused = False
        self.configure(border_color=NAVY if text or focused else LINE)
        if self.picked:
            self.picked = False
            return
        if self.job:
            self.after_cancel(self.job)
        self.job = self.after(140, self.update_popup)

    def update_popup(self):
        self.job = None
        query = self.variable.get().strip()
        if not query or not self.suggest:
            self.hide()
            return
        try:
            rows = self.suggest(query)[:6]
        except Exception:
            rows = []
        if not rows:
            self.hide()
            return
        self.show(query, rows)

    def show(self, query, rows):
        s = scaling(self)
        root = self.winfo_toplevel()
        row_h, head = round(40 * s), round(28 * s)
        w = self.winfo_width()
        h = head + row_h * len(rows) + round(10 * s)
        x, y = self.winfo_rootx(), self.winfo_rooty() + self.winfo_height() + round(6 * s)
        if self.popup is None or not self.popup.winfo_exists():
            self.popup = tk.Toplevel(root)
            self.popup.overrideredirect(True)
            self.popup.attributes("-topmost", True)
            self.canvas = tk.Canvas(self.popup, highlightthickness=1, bd=0)
            self.canvas.pack(fill="both", expand=True)
            self.canvas.bind("<Motion>", self.hover)
            self.canvas.bind("<Button-1>", lambda e: self.choose(self.row_at(e.y)))
        self.popup.geometry(f"{w}x{h}+{x}+{y}")
        c = self.canvas
        c.configure(bg=pick(CARD), highlightbackground=pick(LINE))
        c.delete("all")
        self.rows, self.active, self.row_h, self.head = rows, -1, row_h, head
        self.hover_rect = c.create_rectangle(0, 0, 0, 0, fill=pick(CARD), outline="")
        c.create_text(round(12 * s), round(16 * s), text=f"SARAN · {len(rows)} cocok", anchor="w", fill=pick(MUTED), font=tk_font(self, 10, "bold"))
        bold, normal, small = tk_font(self, 13, "bold"), tk_font(self, 13), tk_font(self, 11)
        import tkinter.font as tkfont

        fb, fn = tkfont.Font(root=self, font=bold), tkfont.Font(root=self, font=normal)
        self.anim = []
        for i, (tag, value, side) in enumerate(rows):
            top = head + i * row_h
            mid = top + row_h // 2
            k = value.lower().find(query.lower())
            pre, hit, post = (value[:k], value[k:k + len(query)], value[k + len(query):]) if k >= 0 else (value, "", "")
            if len(pre) > 22:
                pre = "…" + pre[-14:]
            items = []
            bx = round(12 * s)
            pill = c.create_rectangle(bx, mid - round(11 * s), bx + round(40 * s), mid + round(11 * s), fill=pick(CARD), outline="")
            items.append((pill, "fill", ("#EADBC8", "#2E3D5C")))
            items.append((c.create_text(bx + round(20 * s), mid, text=tag, fill=pick(CARD), font=tk_font(self, 9, "bold")), "fill", NAVY))
            x = bx + round(52 * s)
            side_x = w - round(12 * s)
            limit = side_x - (fn.measure(side) + round(16 * s) if side else 0)
            for part, f, font in ((pre, fn, normal), (hit, fb, bold), (post, fn, normal)):
                if not part or x >= limit:
                    continue
                while part and x + f.measure(part) > limit:
                    part = part[:-2] + "…" if len(part) > 2 else ""
                items.append((c.create_text(x, mid, text=part, anchor="w", fill=pick(CARD), font=font), "fill", TEXT))
                x += f.measure(part)
            if side:
                items.append((c.create_text(side_x, mid, text=side, anchor="e", fill=pick(CARD), font=small), "fill", MUTED))
            self.anim.append((i, items))
        # muncul dari samar (warna kartu) dan sedikit turun, tiap baris menyusul 40 ms
        for i, items in self.anim:
            self.after(30 + i * 40, lambda it=items, idx=i: self.fade(it, idx, 0))

    def fade(self, items, index, frame):
        frames = 14
        if self.popup is None or not self.popup.winfo_exists():
            return
        t = ease_out(frame / frames)
        c = self.canvas
        for item, option, final in items:
            try:
                c.itemconfigure(item, **{option: mix(CARD, final, t)})
                if frame == 0:
                    c.move(item, 0, round(5 * scaling(self)))
                else:
                    c.move(item, 0, -round(5 * scaling(self)) / frames)
            except tk.TclError:
                return
        if frame < frames:
            self.after(FRAME_MS, lambda: self.fade(items, index, frame + 1))

    def row_at(self, y):
        index = (y - self.head) // self.row_h if self.row_h else -1
        return index if 0 <= index < len(self.rows) else -1

    def hover(self, event):
        self.highlight(self.row_at(event.y))

    def highlight(self, index):
        self.active = index
        c = self.canvas
        if index < 0:
            c.coords(self.hover_rect, 0, 0, 0, 0)
            return
        s = scaling(self)
        top = self.head + index * self.row_h
        c.coords(self.hover_rect, round(4 * s), top + 1, self.winfo_width() - round(4 * s), top + self.row_h - 1)
        c.itemconfigure(self.hover_rect, fill=pick(("#F3EADF", "#283044")))
        c.tag_lower(self.hover_rect)

    def move(self, step):
        if not self.rows or self.popup is None:
            return
        self.highlight((self.active + step) % len(self.rows))

    def choose(self, index):
        if index is None or index < 0 or index >= len(self.rows):
            return
        value = self.rows[index][1]
        self.picked = True
        self.variable.set(value)
        self.hide()
        self.entry.icursor("end")

    def hide(self):
        if self.popup is not None and self.popup.winfo_exists():
            self.popup.destroy()
        self.popup = None
