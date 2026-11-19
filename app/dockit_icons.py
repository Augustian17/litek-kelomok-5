"""Ikon Moving Icons (movingicons.dev, lisensi MIT, berbasis Lucide) untuk Dockit.
Animasi aslinya berbasis web, jadi tiap ikon direkam menjadi beberapa frame lalu diputar saat mouse diarahkan."""
import json
import os

import customtkinter as ctk
from PIL import Image, ImageTk

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
_meta = None
_tiles = {}
_ctk = {}
_photo = {}


def meta():
    global _meta
    if _meta is None:
        try:
            with open(os.path.join(BASE, "icons.json"), encoding="utf-8") as file:
                _meta = json.load(file)
        except OSError:
            _meta = {}
    return _meta


def available(name):
    return name in meta()


def tiles(name):
    """Masker alpha: frame diam lalu frame animasi."""
    if name not in _tiles:
        info = meta()[name]
        strip = Image.open(os.path.join(BASE, f"{name}.png")).convert("LA")
        size = info["size"]
        alpha = strip.split()[1]
        _tiles[name] = [alpha.crop((i * size, 0, (i + 1) * size, size)) for i in range(strip.size[0] // size)]
    return _tiles[name]


def tinted(mask, color):
    image = Image.new("RGBA", mask.size, color)
    image.putalpha(mask)
    return image


def ctk_frames(name, color, size):
    """Daftar CTkImage: indeks 0 diam, sisanya animasi. Warna bisa pasangan (terang, gelap)."""
    key = (name, str(color), size)
    if key not in _ctk:
        light, dark = color if isinstance(color, (tuple, list)) else (color, color)
        _ctk[key] = [ctk.CTkImage(light_image=tinted(m, light), dark_image=tinted(m, dark), size=(size, size)) for m in tiles(name)]
    return _ctk[key]


def photo(name, color, pixels):
    """Gambar diam untuk kanvas tabel."""
    key = (name, color, pixels)
    if key not in _photo:
        _photo[key] = ImageTk.PhotoImage(tinted(tiles(name)[0], color).resize((pixels, pixels), Image.LANCZOS))
    return _photo[key]


class Animator:
    """Memasang ikon di tombol atau label, dan memutar animasinya saat mouse masuk."""

    def __init__(self, widget, name, color, size=18, hover_widgets=()):
        self.widget, self.name, self.size = widget, name, size
        self.index = 0
        self.job = None
        self.inside = False
        self.set_color(color)
        for target in (widget,) + tuple(hover_widgets):
            target.bind("<Enter>", self.enter, add="+")
            target.bind("<Leave>", self.leave, add="+")

    def set_color(self, color):
        self.color = color
        self.images = ctk_frames(self.name, color, self.size)
        self.show(self.index if self.index < len(self.images) else 0)

    def show(self, index):
        self.index = index
        try:
            self.widget.configure(image=self.images[index])
        except Exception:
            pass

    def enter(self, _event=None):
        if self.inside:
            return
        self.inside = True
        if len(self.images) < 2:
            return
        self.cancel()
        step = max(16, meta()[self.name]["duration"] // max(1, len(self.images) - 2))
        self.play(1, step)

    def play(self, index, step):
        if not self.inside or index >= len(self.images):
            self.job = None
            return
        self.show(index)
        self.job = self.widget.after(step, lambda: self.play(index + 1, step))

    def leave(self, _event=None):
        self.inside = False
        self.cancel()
        self.show(0)

    def cancel(self):
        if self.job is not None:
            try:
                self.widget.after_cancel(self.job)
            except Exception:
                pass
            self.job = None


BUTTON_ICONS = [
    ("buat excel + pdf", "download"), ("buat excel", "file-plus"), ("simpan", "file-check"), ("muat ulang", "refresh-cw"),
    ("+ baris", "square-plus"), ("+ item", "square-plus"), ("tempel", "clipboard"), ("hapus baris", "trash-2"),
    ("hapus tanda", "rotate-ccw"), ("input file", "file-up"), ("kosongkan", "x"), ("pilih", "folder-plus"),
    ("ambil ulang", "refresh-cw"), ("ambil", "search"), ("reset", "rotate-ccw"), ("lihat di master", "ship"),
    ("batal", "x"), ("ok", "circle-check"), ("convert", "file-down"), ("hapus", "trash-2"),
    ("tambah", "square-plus"), ("buka", "folder-plus"), ("zip", "download"),
]


def icon_for(text):
    words = " ".join(str(text).lower().split())
    for prefix, name in BUTTON_ICONS:
        if words.startswith(prefix) and available(name):
            return name
    return None


CARD_ICONS = [("total etd", "chart-column-increasing"), ("etd terpadat", "chart-column-increasing"), ("etd", "calendar-days"), ("bulan", "calendar-days"),
              ("total", "clipboard-list"), ("perlu dicek", "triangle-alert"), ("belum disimpan", "circle-alert"), ("double", "file-stack"),
              ("cbm tidak tally", "boxes"), ("harga ganda", "hand-coins")]


def card_icon(text):
    words = " ".join(str(text).lower().split())
    for prefix, name in CARD_ICONS:
        if words.startswith(prefix) and available(name):
            return name
    return None
