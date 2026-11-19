"""Kalender ETD: blok per BL selama 2 hari (berangkat H-1 lalu ETD), dikelompokkan per POD."""
import datetime
import re
import tkinter as tk
from tkinter import font as tkfont

import customtkinter as ctk

import di_data as data
import dockit_motion as motion

DATE_VALUE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})")
DAYS = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
ETD_KEY = {"DI": "L", "PO": "D"}
TABLE = {"DI": "master", "PO": "po_master"}
CONTAINER_KEY = {"DI": "G", "PO": "H"}
SEAL_KEY = {"DI": "H", "PO": "I"}

# warna: (isi, garis, teks) untuk terang dan gelap
STATUS = {
    "today": ("Berangkat hari ini", (("#0F2C59", "#0F2C59", "#FFFFFF"), ("#3B5F9E", "#3B5F9E", "#FFFFFF"))),
    "ready": ("Siap: kontainer + seal", (("#3B6BB5", "#3B6BB5", "#FFFFFF"), ("#2E4F86", "#2E4F86", "#FFFFFF"))),
    "moved": ("ETD berubah", (("#FDF0D9", "#E8A33D", "#8A4B00"), ("#3A2A10", "#E8A33D", "#F2C879"))),
    "wait": ("Menunggu kontainer", (("#FFFFFF", "#C9B79E", "#16202E"), ("#1A2030", "#4A5268", "#EEF0F5"))),
    "gone": ("Sudah lewat", (("#F1E6D7", "#E8DCCB", "#5E6675"), ("#232A3A", "#2E3548", "#A0A8B8"))),
}
ORDER = ["today", "ready", "moved", "wait", "gone"]


def short_date(day):
    return f"{day.day} {MONTHS[day.month - 1]}"


def parse_day(text):
    match = DATE_VALUE.match(text or "")
    if not match:
        return None
    try:
        return datetime.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except ValueError:
        return None


class CalendarPage(ctk.CTkFrame):
    DAYS_SHOWN = 14
    LABEL_W = 230
    LANE_H = 52
    POD_H = 30
    HEAD_H = 40

    def __init__(self, parent, manager, profile):
        super().__init__(parent, fg_color="transparent")
        self.manager, self.profile = manager, profile
        ui = manager.ui
        self.ui = ui
        self.shift = 0
        self.off = set()
        self.selected = None
        self.lanes = []
        self.hits = []
        self.hover = None
        self.search = ctk.StringVar()
        self.search.trace_add("write", lambda *_: self.after_idle(self.draw))
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        top = ctk.CTkFrame(self, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        titles = ctk.CTkFrame(top, fg_color="transparent")
        titles.pack(side="left")
        ctk.CTkLabel(titles, text="Kalender ETD", font=ui.font(26, "bold"), text_color=ui.TEXT, height=32).pack(anchor="w")
        self.summary = ctk.CTkLabel(titles, text="", font=ui.font(12), text_color=ui.MUTED, height=18)
        self.summary.pack(anchor="w")
        self.build_button = ui.button(top, "Buat Excel", self.open_build, primary=True, width=130, height=42)
        self.build_button.pack(side="right")
        nav = ctk.CTkFrame(top, fg_color=ui.CARD, border_color=ui.BORDER, border_width=1, corner_radius=10, height=42)
        nav.pack(side="right", padx=8)
        ctk.CTkButton(nav, text="‹", width=34, height=36, fg_color="transparent", hover_color=ui.NEUTRAL, text_color=ui.ACCENT,
                      font=ui.font(18, "bold"), command=lambda: self.move(-1)).pack(side="left", padx=(3, 0), pady=3)
        self.range_label = ctk.CTkLabel(nav, text="", font=ui.font(13, "bold"), text_color=ui.TEXT, width=150)
        self.range_label.pack(side="left")
        ctk.CTkButton(nav, text="›", width=34, height=36, fg_color="transparent", hover_color=ui.NEUTRAL, text_color=ui.ACCENT,
                      font=ui.font(18, "bold"), command=lambda: self.move(1)).pack(side="left", pady=3)
        ctk.CTkButton(nav, text="Hari ini", width=70, height=30, corner_radius=8, fg_color=ui.NEUTRAL, hover_color=ui.ACCENT_SOFT,
                      text_color=ui.TEXT, font=ui.font(12), command=self.go_today).pack(side="left", padx=(4, 6))
        self.search_bar = motion.SearchBar(top, self.search, "Cari PO, DI, BL, kontainer", self.suggest, width=300, height=42)
        self.search_bar.pack(side="right", padx=(0, 0))

        self.legend = ctk.CTkFrame(self, fg_color="transparent")
        self.legend.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=2, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)
        card = ctk.CTkFrame(body, fg_color=ui.CARD, border_color=ui.BORDER, border_width=1, corner_radius=14)
        card.grid(row=0, column=0, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(card, text="ETD per BL", font=ui.font(15, "bold"), text_color=ui.TEXT, anchor="w").grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 0))
        ctk.CTkLabel(card, text="Tiap blok 2 hari: berangkat (H-1) lalu ETD. Klik blok untuk detail, klik status untuk menyaring.",
                     font=ui.font(12), text_color=ui.MUTED, anchor="w").grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 6))
        holder = tk.Frame(card, bd=0, highlightthickness=0)
        holder.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        holder.grid_columnconfigure(0, weight=1)
        holder.grid_rowconfigure(1, weight=1)
        self.head = tk.Canvas(holder, highlightthickness=0, bd=0, height=self.px(self.HEAD_H))
        self.head.grid(row=0, column=0, sticky="ew")
        self.canvas = tk.Canvas(holder, highlightthickness=0, bd=0)
        self.canvas.grid(row=1, column=0, sticky="nsew")
        bar = ctk.CTkScrollbar(holder, command=self.canvas.yview)
        bar.grid(row=1, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=bar.set)
        self.canvas.bind("<Configure>", lambda _e: self.draw())
        self.canvas.bind("<Motion>", self.on_motion)
        self.canvas.bind("<Leave>", lambda _e: self.set_hover(None))
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))
        self.canvas.bind("<Button-4>", lambda _e: self.canvas.yview_scroll(-2, "units"))
        self.canvas.bind("<Button-5>", lambda _e: self.canvas.yview_scroll(2, "units"))

        side = ctk.CTkFrame(body, fg_color="transparent", width=300)
        side.grid(row=0, column=1, sticky="ns", padx=(12, 0))
        side.grid_propagate(False)
        side.grid_columnconfigure(0, weight=1)
        side.grid_rowconfigure(2, weight=1)
        self.detail = ctk.CTkFrame(side, fg_color=ui.CARD, border_color=ui.BORDER, border_width=1, corner_radius=14)
        self.detail.grid(row=0, column=0, sticky="ew")
        self.detail_title = ctk.CTkLabel(self.detail, text="Detail BL", font=ui.font(15, "bold"), text_color=ui.TEXT, anchor="w")
        self.detail_title.pack(fill="x", padx=16, pady=(12, 0))
        self.detail_text = ctk.CTkLabel(self.detail, text="Klik blok di kalender.", font=ui.font(12), text_color=ui.MUTED, anchor="w",
                                        justify="left", wraplength=262)
        self.detail_text.pack(fill="x", padx=16, pady=(2, 12))
        pods = ctk.CTkFrame(side, fg_color=ui.CARD, border_color=ui.BORDER, border_width=1, corner_radius=14)
        pods.grid(row=1, column=0, sticky="ew", pady=12)
        ctk.CTkLabel(pods, text="Kontainer per POD", font=ui.font(15, "bold"), text_color=ui.TEXT, anchor="w").pack(fill="x", padx=16, pady=(12, 0))
        ctk.CTkLabel(pods, text="Dalam 2 minggu yang tampil", font=ui.font(12), text_color=ui.MUTED, anchor="w").pack(fill="x", padx=16)
        self.pod_canvas = tk.Canvas(pods, highlightthickness=0, bd=0, height=self.px(150))
        self.pod_canvas.pack(fill="x", padx=12, pady=(6, 12))
        changes = ctk.CTkFrame(side, fg_color=ui.CARD, border_color=ui.BORDER, border_width=1, corner_radius=14)
        changes.grid(row=2, column=0, sticky="nsew")
        ctk.CTkLabel(changes, text="Perubahan ETD", font=ui.font(15, "bold"), text_color=ui.TEXT, anchor="w").pack(fill="x", padx=16, pady=(12, 4))
        self.change_list = ctk.CTkScrollableFrame(changes, fg_color="transparent")
        self.change_list.pack(fill="both", expand=True, padx=6, pady=(0, 8))

    # ------------------------------------------------------------- data
    def px(self, value):
        return round(value * motion.scaling(self))

    def today(self):
        return datetime.date.today()

    def start(self):
        today = self.today()
        return today - datetime.timedelta(days=today.weekday() + 7) + datetime.timedelta(days=self.shift * self.DAYS_SHOWN)

    def load(self):
        store = self.manager.store
        profile = self.profile
        name = profile.name
        columns = data.MASTER_PO_COLUMNS if name == "PO" else data.MASTER_COLUMNS
        kinds = {c.key: c.kind for c in columns}
        etd_key = ETD_KEY[name]
        marks = data.load_settings().get("etd_changes", {}).get(name, {})
        groups = {}
        seen = set()
        for row in store.rows_of(TABLE[name]):
            code = data.unique_key(row.get("A"))
            if not code or code in seen:
                continue
            seen.add(code)
            etd_text = data.format_value(row.get(etd_key), kinds.get(etd_key, "date"))
            etd = parse_day(etd_text)
            if etd is None:
                continue
            bl = data.format_value(row.get(profile.bl_key), "text").strip() or f"(tanpa BL) {code}"
            group = groups.setdefault(bl, {"bl": bl, "codes": [], "etds": [], "pods": [], "containers": [], "seals": [], "qty": 0.0, "old": []})
            group["codes"].append(code)
            group["etds"].append(etd)
            group["pods"].append(data.format_value(row.get(profile.pod_key), "text").strip().upper())
            group["containers"].append(data.format_value(row.get(CONTAINER_KEY[name]), "text").strip())
            group["seals"].append(data.format_value(row.get(SEAL_KEY[name]), "text").strip())
            value = row.get(profile.qty_key)
            try:
                group["qty"] += float(value) if isinstance(value, (int, float)) else float(data.parse_value(str(value or "0").replace(",", ""), "num") or 0)
            except (TypeError, ValueError):
                pass
            first = marks.get(code)
            if first and first != etd_text:
                old = parse_day(first)
                if old and old != etd:
                    group["old"].append(old)
        today = self.today()
        lanes = []
        for group in groups.values():
            etd = max(set(group["etds"]), key=group["etds"].count)
            pod = max(set(group["pods"]), key=group["pods"].count) if any(group["pods"]) else "(tanpa POD)"
            pod = pod or "(tanpa POD)"
            filled = sum(1 for c, s in zip(group["containers"], group["seals"]) if c and s and not c.startswith("#") and not s.startswith("#"))
            count = len(group["codes"])
            old = min(group["old"]) if group["old"] else None
            if etd < today:
                status = "gone"
            elif etd - datetime.timedelta(days=1) == today or etd == today:
                status = "today"
            elif old:
                status = "moved"
            elif filled == count:
                status = "ready"
            else:
                status = "wait"
            group.update(etd=etd, pod=pod, status=status, count=count, filled=filled, old_etd=old)
            lanes.append(group)
        self.all_lanes = lanes

    def refresh(self):
        if not self.manager.store.loaded:
            return
        self.load()
        self.draw()

    def move(self, step):
        self.shift += step
        self.draw()

    def go_today(self):
        self.shift = 0
        self.draw()

    def code_text(self, codes):
        if not codes:
            return ""
        prefix = codes[0].split("柜")[0] + "柜" if "柜" in codes[0] else ""
        tail = [c.split("柜")[-1] for c in codes]
        text = prefix + ", ".join(tail[:4])
        return text + (f" +{len(tail) - 4}" if len(tail) > 4 else "")

    def matches(self, lane, words):
        if not words:
            return True
        text = " ".join([lane["bl"], lane["pod"]] + lane["codes"] + lane["containers"]).lower()
        return all(w in text for w in words)

    def suggest(self, query):
        if not hasattr(self, "all_lanes"):
            return []
        q = query.lower()
        found, seen = [], set()
        for lane in self.all_lanes:
            for tag, values in (("BL", [lane["bl"]]), (self.profile.name, lane["codes"]), ("CN", lane["containers"])):
                for value in values:
                    if value and q in value.lower() and value not in seen:
                        seen.add(value)
                        found.append((tag, value, "ETD " + short_date(lane["etd"])))
                        if len(found) >= 6:
                            return found
        return found

    # ------------------------------------------------------------- gambar
    def colors(self, status):
        light, dark = STATUS[status][1]
        return dark if motion.dark() else light

    def draw(self):
        if not hasattr(self, "all_lanes"):
            return
        c, h = self.canvas, self.head
        c.delete("all")
        h.delete("all")
        bg = motion.pick(("#FFFFFF", "#1A2030"))
        for canvas in (c, h):
            canvas.configure(bg=bg)
        line = motion.pick(("#EEE4D6", "#2E3548"))
        muted, text = motion.pick(motion.MUTED), motion.pick(motion.TEXT)
        weekend = motion.pick(("#FBF6EF", "#1E2536"))
        start = self.start()
        end = start + datetime.timedelta(days=self.DAYS_SHOWN - 1)
        today = self.today()
        width = max(c.winfo_width(), self.px(600))
        label_w = self.px(self.LABEL_W)
        col_w = (width - label_w) / self.DAYS_SHOWN
        words = self.search.get().lower().split()
        visible = []
        for lane in self.all_lanes:
            first = lane["etd"] - datetime.timedelta(days=1)
            in_range = first <= end and lane["etd"] >= start
            ghost = lane["old_etd"] and lane["old_etd"] - datetime.timedelta(days=1) <= end and lane["old_etd"] >= start
            if (in_range or ghost) and self.matches(lane, words):
                visible.append(lane)
        counts = {s: 0 for s in ORDER}
        for lane in visible:
            counts[lane["status"]] += lane["count"]
        shown = [lane for lane in visible if lane["status"] not in self.off]
        shown.sort(key=lambda l: (l["pod"], l["etd"], l["bl"]))
        self.range_label.configure(text=f"{start.day} {MONTHS[start.month - 1]} – {short_date(end)} {end.year}")
        pods = sorted({l["pod"] for l in visible})
        moved = sum(1 for l in visible if l["old_etd"])
        self.summary.configure(text=f"{len({l['pod'] for l in shown})} POD · {len(shown)} BL · {sum(l['count'] for l in shown)} {self.profile.name} · {moved} ETD berubah")
        self.draw_legend(counts)

        f_day, f_day_b = motion.tk_font(self, 11), motion.tk_font(self, 11, "bold")
        f_pod, f_bl, f_sub = motion.tk_font(self, 11, "bold"), motion.tk_font(self, 12, "bold"), motion.tk_font(self, 11)
        f_block, f_block_sub = motion.tk_font(self, 11, "bold"), motion.tk_font(self, 10)
        head_h = self.px(self.HEAD_H)
        h.configure(height=head_h)
        for d in range(self.DAYS_SHOWN):
            day = start + datetime.timedelta(days=d)
            x = label_w + d * col_w
            is_today = day == today
            h.create_text(x + col_w / 2, head_h * 0.34, text=DAYS[day.weekday()], fill=text if is_today else muted, font=f_day_b if is_today else f_day)
            h.create_text(x + col_w / 2, head_h * 0.72, text=short_date(day), fill=text if is_today else muted, font=f_day_b if is_today else f_day)
        h.create_line(0, head_h - 1, width, head_h - 1, fill=line)
        if start <= today <= end:
            x = label_w + ((today - start).days + 0.5) * col_w
            h.create_rectangle(x - self.px(28), 2, x + self.px(28), self.px(18), fill=motion.pick(motion.NAVY), outline="")
            h.create_text(x, self.px(10), text="Hari ini", fill="#FFFFFF", font=motion.tk_font(self, 10, "bold"))

        lane_h, pod_h = self.px(self.LANE_H), self.px(self.POD_H)
        rows = []
        y = 0
        last_pod = None
        for lane in shown:
            if lane["pod"] != last_pod:
                rows.append(("pod", lane["pod"], y))
                y += pod_h
                last_pod = lane["pod"]
            rows.append(("lane", lane, y))
            y += lane_h
        total_h = max(y, c.winfo_height())
        # latar akhir pekan dan garis hari
        for d in range(self.DAYS_SHOWN):
            day = start + datetime.timedelta(days=d)
            x = label_w + d * col_w
            if day.weekday() >= 5:
                c.create_rectangle(x, 0, x + col_w, total_h, fill=weekend, outline="")
            c.create_line(x, 0, x, total_h, fill=line)
        if start <= today <= end:
            x = label_w + ((today - start).days + 0.5) * col_w
            c.create_line(x, 0, x, total_h, fill=motion.pick(motion.NAVY), width=max(1, self.px(1.5)), dash=(4, 3))
        self.hits = []
        for kind, item, y in rows:
            if kind == "pod":
                c.create_rectangle(0, y, width, y + pod_h, fill=motion.pick(("#F8F0E5", "#141A28")), outline="")
                c.create_text(self.px(10), y + pod_h / 2, text=item, anchor="w", fill=motion.pick(motion.NAVY) if not motion.dark() else "#DAC0A3", font=f_pod)
                continue
            lane = item
            c.create_line(0, y + lane_h, width, y + lane_h, fill=line)
            c.create_text(self.px(14), y + lane_h * 0.36, text=lane["bl"], anchor="w", fill=text, font=f_bl)
            c.create_text(self.px(14), y + lane_h * 0.7, text=f"{lane['count']} kontainer · {lane['filled']}/{lane['count']} seal", anchor="w", fill=muted, font=f_sub)
            if lane["old_etd"]:
                self.block(c, lane, lane["old_etd"], start, label_w, col_w, y, lane_h, ghost=True, fonts=(f_block, f_block_sub))
            self.block(c, lane, lane["etd"], start, label_w, col_w, y, lane_h, fonts=(f_block, f_block_sub))
        if not shown:
            c.create_text(width / 2, self.px(60), text="Tidak ada BL di rentang ini" + (f" yang cocok dengan \"{self.search.get()}\"" if words else ""), fill=muted, font=f_sub)
        c.configure(scrollregion=(0, 0, width, total_h))
        self.draw_pods(shown)
        self.draw_changes(visible)
        self.show_detail()

    def block(self, c, lane, etd, start, label_w, col_w, y, lane_h, ghost=False, fonts=None):
        first = (etd - datetime.timedelta(days=1) - start).days
        x1 = label_w + first * col_w + self.px(3)
        x2 = x1 + 2 * col_w - self.px(6)
        y1, y2 = y + self.px(6), y + lane_h - self.px(6)
        fill, outline, fg = self.colors(lane["status"])
        selected = self.selected == lane["bl"]
        hovered = self.hover == lane["bl"]
        r = self.px(8)
        points = motion.flat(motion.rounded_path(x1, y1, x2, y2, r, 4))
        if ghost:
            c.create_polygon(points, fill="", outline=motion.pick(("#C9B79E", "#4A5268")), dash=(4, 3), width=1)
            c.create_text((x1 + x2) / 2, (y1 + y2) / 2, text=f"ETD lama {short_date(etd)}", fill=motion.pick(motion.MUTED), font=fonts[1])
            return
        width = 2 if (selected or hovered) else 1
        if selected:
            outline = motion.pick(motion.NAVY) if lane["status"] not in ("today", "ready") else "#DAC0A3"
        c.create_polygon(points, fill=fill, outline=outline, width=self.px(width))
        room = x2 - x1 - self.px(16)
        title = self.fit(self.code_text(lane["codes"]), fonts[0], room)
        sub = self.fit(f"Berangkat {short_date(etd - datetime.timedelta(days=1))} · ETD {short_date(etd)}", fonts[1], room)
        c.create_text(x1 + self.px(8), y1 + (y2 - y1) * 0.34, text=title, anchor="w", fill=fg, font=fonts[0])
        c.create_text(x1 + self.px(8), y1 + (y2 - y1) * 0.72, text=sub, anchor="w", fill=fg, font=fonts[1])
        self.hits.append((x1, y1, x2, y2, lane["bl"]))

    def fit(self, text, font, room):
        f = tkfont.Font(root=self, font=font)
        if f.measure(text) <= room:
            return text
        while text and f.measure(text + "…") > room:
            text = text[:-1]
        return text + "…" if text else ""

    def draw_legend(self, counts):
        for child in self.legend.winfo_children():
            child.destroy()
        ui = self.ui
        for status in ORDER:
            name = STATUS[status][0]
            fill, outline, _ = self.colors(status)
            off = status in self.off
            chip = ctk.CTkButton(
                self.legend, text=f"     {name}  {counts.get(status, 0)}", font=ui.font(12, "bold" if not off else "normal"), height=30, corner_radius=15,
                fg_color=ui.CARD if off else ("#FBF6EF", "#232A3A"), border_width=1, border_color=ui.BORDER,
                hover_color=ui.NEUTRAL, text_color=ui.MUTED if off else ui.TEXT, width=40, command=lambda s=status: self.toggle(s),
            )
            chip.pack(side="left", padx=(0, 6))
            # titik warna status
            dot = outline if status in ("wait",) else fill
            swatch = tk.Canvas(chip, width=self.px(10), height=self.px(10), highlightthickness=0, bd=0, bg=motion.pick(("#FBF6EF", "#232A3A")) if not off else motion.pick(motion.CARD))
            swatch.create_rectangle(0, 0, self.px(10), self.px(10), fill=dot if not off else motion.pick(("#E0D2BF", "#2E3548")), outline=outline)
            swatch.place(x=self.px(12), rely=0.5, anchor="w")
            swatch.bind("<Button-1>", lambda _e, s=status: self.toggle(s))

    def toggle(self, status):
        if status in self.off:
            self.off.discard(status)
        else:
            self.off.add(status)
        self.draw()

    def draw_pods(self, lanes):
        c = self.pod_canvas
        c.delete("all")
        c.configure(bg=motion.pick(motion.CARD))
        totals = {}
        for lane in lanes:
            totals[lane["pod"]] = totals.get(lane["pod"], 0) + lane["count"]
        items = sorted(totals.items(), key=lambda x: -x[1])[:6]
        width = max(c.winfo_width(), self.px(260))
        row = self.px(24)
        c.configure(height=max(self.px(30), row * len(items) + self.px(4)))
        if not items:
            c.create_text(self.px(4), self.px(14), text="Belum ada BL di rentang ini.", anchor="w", fill=motion.pick(motion.MUTED), font=motion.tk_font(self, 12))
            return
        top = items[0][1]
        label_w = self.px(104)
        for i, (pod, count) in enumerate(items):
            y = i * row + self.px(2)
            c.create_text(0, y + row / 2, text=self.fit(pod.title(), motion.tk_font(self, 11), label_w - self.px(6)), anchor="w",
                          fill=motion.pick(motion.TEXT), font=motion.tk_font(self, 11))
            bar_w = (width - label_w - self.px(30)) * count / top
            c.create_rectangle(label_w, y + self.px(5), width - self.px(30), y + row - self.px(5), fill=motion.pick(("#F3EADF", "#232A3A")), outline="")
            c.create_rectangle(label_w, y + self.px(5), label_w + max(self.px(4), bar_w), y + row - self.px(5), fill=motion.pick(motion.NAVY), outline="")
            c.create_text(width - self.px(2), y + row / 2, text=str(count), anchor="e", fill=motion.pick(motion.TEXT), font=motion.tk_font(self, 11, "bold"))

    def draw_changes(self, lanes):
        for child in self.change_list.winfo_children():
            child.destroy()
        ui = self.ui
        moved = [l for l in lanes if l["old_etd"]]
        if not moved:
            ctk.CTkLabel(self.change_list, text="Tidak ada ETD yang berubah di rentang ini.", font=ui.font(12), text_color=ui.MUTED,
                         anchor="w", wraplength=250, justify="left").pack(fill="x", padx=8, pady=4)
            return
        for lane in moved:
            days = (lane["etd"] - lane["old_etd"]).days
            row = ctk.CTkFrame(self.change_list, fg_color=("#FBF6EF", "#232A3A"), corner_radius=10)
            row.pack(fill="x", padx=4, pady=3)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(row, text=f"{self.code_text(lane['codes'])} · {lane['bl']}", font=ui.font(12, "bold"), text_color=ui.TEXT, anchor="w",
                         wraplength=190, justify="left").grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 0))
            ctk.CTkLabel(row, text=f"{lane['pod'].title()} · ETD {short_date(lane['old_etd'])} → {short_date(lane['etd'])}", font=ui.font(11),
                         text_color=ui.MUTED, anchor="w").grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
            late = days > 0
            ctk.CTkLabel(row, text=f"{'+' if late else '−'}{abs(days)} hari", font=ui.font(11, "bold"), corner_radius=8, height=24,
                         fg_color=("#FDF0D9", "#3A2A10") if late else ("#E3ECF8", "#1D2A44"),
                         text_color=("#8A4B00", "#F2C879") if late else ("#1D3B8A", "#9DB8F0")).grid(row=0, column=1, rowspan=2, padx=8)
            for widget in [row] + row.winfo_children():
                widget.bind("<Button-1>", lambda _e, bl=lane["bl"]: self.select(bl))

    def show_detail(self):
        lane = next((l for l in getattr(self, "all_lanes", []) if l["bl"] == self.selected), None)
        if lane is None:
            self.detail_title.configure(text="Detail BL")
            self.detail_text.configure(text="Klik blok di kalender untuk melihat detail, lalu Buat Excel untuk BL itu.")
            return
        first = lane["etd"] - datetime.timedelta(days=1)
        qty = f"{lane['qty']:,.0f}" if lane["qty"] else "-"
        containers = ", ".join(c for c in lane["containers"] if c) or "belum ada"
        lines = [
            f"{self.profile.name}: {', '.join(lane['codes'])}",
            f"POD: {lane['pod'].title()} · {STATUS[lane['status']][0]}",
            f"Berangkat {short_date(first)} · ETD {short_date(lane['etd'])} {lane['etd'].year}",
            f"Kontainer: {containers}",
            f"Seal lengkap: {lane['filled']}/{lane['count']} · Qty {qty}",
        ]
        if lane["old_etd"]:
            lines.append(f"ETD lama: {short_date(lane['old_etd'])}")
        self.detail_title.configure(text=lane["bl"])
        self.detail_text.configure(text="\n".join(lines))

    # ------------------------------------------------------------ interaksi
    def hit(self, event):
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        for x1, y1, x2, y2, bl in self.hits:
            if x1 <= x <= x2 and y1 <= y <= y2:
                return bl
        return None

    def on_motion(self, event):
        self.set_hover(self.hit(event))

    def set_hover(self, bl):
        self.canvas.configure(cursor="hand2" if bl else "")
        if bl != self.hover:
            self.hover = bl
            self.draw()

    def on_click(self, event):
        bl = self.hit(event)
        if bl:
            self.select(bl)

    def select(self, bl):
        self.selected = None if self.selected == bl else bl
        self.draw()

    def open_build(self):
        import di_ui

        ui = self.manager.ui
        ui.show_page(di_ui.PAGE_BUILD)
        if self.selected and not self.selected.startswith("(tanpa BL)"):
            page = self.manager.pages[di_ui.page_key(ui.tab, di_ui.PAGE_BUILD)]
            page.bl.set(self.selected)
            page.open_bl()
