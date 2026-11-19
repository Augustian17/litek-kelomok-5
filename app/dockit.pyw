"""Dockit (sebelumnya PDF Gen) - convert sheet pilihan dari banyak file Excel sekaligus.

Butuh Windows, pywin32, customtkinter, pypdfium2, pillow, dan Microsoft Excel atau WPS Spreadsheets.
"""
import fnmatch
import gc
import os
import queue
import re
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from tkinter import filedialog, messagebox

import customtkinter as ctk

import di_data
import di_ui
import dockit_icons as icons
import dockit_motion as motion

APP_VERSION = "1.28"
APP_TITLE = f"Dockit V{APP_VERSION}"
APP_ID = f"Dockit.App.{APP_VERSION}"
def asset_path(name):
    """File logo ada di samping app; saat dijalankan dari folder kerja, cari juga di samping modul data."""
    here = os.path.join(os.path.dirname(os.path.abspath(__file__)), name)
    if os.path.exists(here):
        return here
    return os.path.join(os.path.dirname(os.path.abspath(di_data.__file__)), name)


ICON_PATH = asset_path("dockit.ico")
LOGO_PATH = asset_path("dockit.png")
EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xlsb", ".xls")
EXCEL_TYPES = [("File Excel", "*.xlsx *.xlsm *.xlsb *.xls"), ("Semua file", "*.*")]
ENGINES = [
    ("Microsoft Excel", "Excel.Application"),
    ("WPS Spreadsheets", "KET.Application"),
    ("WPS Spreadsheets", "ET.Application"),
]
XL_TYPE_PDF = 0
XL_SHEET_VISIBLE = -1
XL_PORTRAIT = 1
XL_LANDSCAPE = 2
FORCE_SECURITY_DISABLE_MACROS = 3
NO_PASSWORD = "__tanpa_password__"
# Sheet yang otomatis dipilih, beserta awalan nama PDF-nya.
AUTO_SHEETS = {"INV": "CI", "PL": "PL"}
TAB_PO = "PO"
TAB_DI = "DI"
SI_SHEET = "FINAL SI"
SI_QUANTITY_CELL = "L21"
SI_POD_CELL = "A34"
SI_FILE_NAME = re.compile(
    r"^(.+?)\s+(\d+\s*X\s*\d+\s*[A-Z]+)\s+(.+?)\s+SI(\s|$)", re.IGNORECASE
)
MAX_PREVIEW_PAGES = 20
MAX_PREVIEW_CACHE = 10
# Kolom tabel CI/PL yang wajib terisi.
VALUE_COLUMNS = [
    ("Quantity", r"\b(quantity|qty)\b"),
    ("Net Weight", r"net\s*w(eigh)?t|\bn\.?\s*w\.?(\s|$)"),
    ("Gross Weight", r"gross\s*w(eigh)?t|\bg\.?\s*w\.?(\s|$)"),
    ("Measurement", r"measurement|\bcbm\b|\bvolume\b"),
    ("Unit Price", r"\bprice\b"),
    ("Amount", r"\bamount\b"),
]
ID_COLUMNS = [("Item No.", r"\bitem\b|prototype|\bsku\b|\bstyle\b|\bmodel\b")]
DESCRIPTION_COLUMNS = [("Description", r"description")]
END_OF_TABLE = re.compile(
    r"^\W*(sub\s*-?\s*total|grand\s*total|total|container|remark|say\b|signed|authori[sz]ed)",
    re.IGNORECASE,
)
# Nilai Value2 untuk error Excel seperti #N/A, #REF!, dan #VALUE!.
EXCEL_ERRORS = {
    -2146826288, -2146826281, -2146826273, -2146826265,
    -2146826259, -2146826252, -2146826246,
}

ACCENT = ("#0F2C59", "#3B5F9E")
ACCENT_HOVER = ("#0A1F40", "#466DB0")
ACCENT_SOFT = ("#F1E6D7", "#22304A")
HEADER_BAND = ("#0F2C59", "#0A1A33")
TAB_IDLE = ("#1B3E73", "#1A2C4A")
TAB_IDLE_HOVER = ("#24508F", "#24385C")
TAB_ACTIVE = ("#3B6BB5", "#3B5F9E")
TAB_ACTIVE_HOVER = ("#4A7AC4", "#466DB0")
DANGER_SOFT = ("#FDE7E4", "#4A2320")
DANGER_SOFT_HOVER = ("#FAD3CD", "#5C2B27")
BACKGROUND = ("#F8F0E5", "#0F1420")
CARD = ("#FFFFFF", "#1A2030")
BORDER = ("#E8DCCB", "#2E3548")
TEXT = ("#16202E", "#EEF0F5")
MUTED = ("#5E6675", "#A0A8B8")
WARNING = ("#B45309", "#F2A65A")
DANGER = ("#B42318", "#F97066")
NEUTRAL = ("#EFE6DA", "#283044")
NEUTRAL_HOVER = ("#E5D8C6", "#323B52")
FONT_FAMILY = "Segoe UI" if os.name == "nt" else "DejaVu Sans"


def font(size=13, weight="normal"):
    return ctk.CTkFont(family=FONT_FAMILY, size=size, weight=weight)


def clean_filename(name):
    """Hapus karakter yang tidak boleh ada di nama file Windows."""
    name = name.strip()
    if name.lower().endswith(".pdf"):
        name = name[:-4]
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    return name.strip(" .")


def shorten(text, limit):
    if len(text) <= limit:
        return text
    keep = limit - 1
    return text[: keep - keep // 3] + "…" + text[-(keep // 3):]


def parse_rules(text):
    """Contoh 'Invoice*=A; Packing*=B' menjadi [('Invoice*', 'A'), ('Packing*', 'B')]."""
    rules = []
    for part in re.split(r"[;\n]", text):
        part = part.strip()
        if not part:
            continue
        pattern, _, pdf_name = part.partition("=")
        rules.append((pattern.strip().lower(), clean_filename(pdf_name)))
    return rules


def document_code(stem):
    """'PO-柜9892_INV-PO2026-2341' menjadi 'INV-PO2026-2341'."""
    matches = re.findall(r"INV-[^_\s]+", stem, flags=re.IGNORECASE)
    return matches[-1] if matches else stem


def si_parts(stem):
    """'280564681 3X40HC LOS ANGELES SI BREAKDOWN' menjadi BL, jumlah kontainer, POD."""
    match = SI_FILE_NAME.match(stem)
    if not match:
        return None
    return match.group(1), re.sub(r"\s+", "", match.group(2)).upper(), match.group(3)


def si_pdf_name(stem):
    parts = si_parts(stem)
    if parts:
        return f"{parts[0]} {parts[1]} {parts[2]} SI COMBINED"
    return f"{stem} SI COMBINED"


def is_si_sheet(sheet_name):
    return " ".join(sheet_name.upper().split()) == SI_SHEET


def zip_member_name(info):
    """Nama Mandarin di ZIP Windows sering tersimpan sebagai GBK."""
    if info.flag_bits & 0x800:
        return info.filename
    try:
        return info.filename.encode("cp437").decode("gbk")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return info.filename


def extract_zip(zip_path):
    """Ekstrak ke folder di samping ZIP. File yang sudah ada tidak ditimpa."""
    target = os.path.splitext(zip_path)[0]
    root = os.path.abspath(target)
    with zipfile.ZipFile(zip_path) as archive:
        for info in archive.infolist():
            path = os.path.abspath(os.path.join(root, zip_member_name(info)))
            if not path.startswith(root + os.sep):
                continue
            if info.is_dir():
                os.makedirs(path, exist_ok=True)
            elif not os.path.exists(path):
                os.makedirs(os.path.dirname(path), exist_ok=True)
                with archive.open(info) as source, open(path, "wb") as output:
                    shutil.copyfileobj(source, output)
    return target


def excel_files_in(folder, recursive):
    paths = []
    for current, folders, names in os.walk(folder):
        for name in sorted(names):
            if name.lower().endswith(EXCEL_EXTENSIONS) and not name.startswith("~$"):
                paths.append(os.path.join(current, name))
        if not recursive:
            folders.clear()
    return paths


def is_empty(value):
    if value is None:
        return True
    if isinstance(value, str):
        text = value.strip()
        return text in ("", "-") or text.startswith("#N/A")
    if isinstance(value, (int, float)):
        return value == 0 or value in EXCEL_ERRORS
    return False


def has_digit(value):
    if is_empty(value):
        return False
    return not isinstance(value, str) or any(char.isdigit() for char in value)


def match_header(text, columns):
    if not isinstance(text, str) or len(text) > 40:
        return None
    text = " ".join(text.lower().split())
    return next((label for label, pattern in columns if re.search(pattern, text)), None)


def check_table(values, first_row=1, first_col=1):
    """Cari baris item yang kosong. None berarti tabelnya tidak dikenali."""
    if not isinstance(values, (tuple, list)):
        return None
    rows = [list(row) for row in values]
    header_index, best = None, 1
    for index, row in enumerate(rows):
        found = {match_header(cell, VALUE_COLUMNS) for cell in row} - {None}
        if len(found) > best:
            header_index, best = index, len(found)
    if header_index is None:
        return None

    header = rows[header_index]
    width = len(header)
    positions = [col for col in range(width) if not is_empty(header[col])]
    spans = []
    for number, col in enumerate(positions):
        end = positions[number + 1] if number + 1 < len(positions) else width
        spans.append((header[col], range(col, end)))

    def spans_of(columns):
        result = []
        for text, span in spans:
            label = match_header(text, columns)
            if label:
                result.append((label, span))
        return result

    value_spans = spans_of(VALUE_COLUMNS)
    id_spans = spans_of(ID_COLUMNS)
    description_spans = spans_of(DESCRIPTION_COLUMNS)

    def filled(row, span_list):
        return any(not is_empty(row[col]) for _, span in span_list for col in span)

    items = []
    for index in range(header_index + 1, len(rows)):
        row = rows[index]
        if any(isinstance(cell, str) and END_OF_TABLE.match(cell) for cell in row):
            break
        has_id = filled(row, id_spans)
        has_description = filled(row, description_spans)
        has_values = filled(row, value_spans)
        continues_item = (
            bool(items) and items[-1][0] == index - 1
            and any(has_digit(row[col]) for _, span in value_spans for col in span)
        )
        if (has_id and has_description) or (has_values and (has_id or has_description)) or continues_item:
            items.append((index, row))
    if not items:
        return None

    required = []
    value_labels = {label for label, _ in value_spans}
    for label, span in id_spans + description_spans + value_spans:
        span_is_value = label in value_labels
        columns = [
            col for col in span
            if sum(not is_empty(row[col]) for _, row in items) * 2 >= len(items)
            or (span_is_value and any(has_digit(row[col]) for _, row in items))
        ]
        required.append((label, columns or [span[0]]))

    text_labels = {label for label, _ in id_spans + description_spans}
    item_indexes = {index for index, _ in items}
    problems = []
    for index, row in items:
        missing = []
        for label, columns in required:
            # Item No. dan Description bisa digabung dari baris item di atasnya.
            if label in text_labels and index - 1 in item_indexes:
                continue
            if any(is_empty(row[col]) for col in columns) and label not in missing:
                missing.append(label)
        for label, span in value_spans:
            if not any(has_digit(row[col]) for col in span) and label not in missing:
                missing.append(label)
        if not missing:
            continue
        name = next(
            (str(row[col]).strip() for _, span in id_spans + description_spans
             for col in span if not is_empty(row[col])),
            "",
        )
        problems.append({"row": first_row + index, "item": name, "missing": missing})
    return problems


def find_duplicates(paths):
    seen, duplicates = set(), []
    for path in paths:
        key = path.lower()
        if key in seen:
            duplicates.append(path)
        seen.add(key)
    return duplicates


def auto_prefix(sheet_name):
    return AUTO_SHEETS.get(sheet_name.strip().upper().rstrip("-"))


class Sheet:
    def __init__(self, name, hidden):
        self.name = name
        self.hidden = hidden
        self.selected = False
        self.pdf_name = clean_filename(name)


class ExcelFile:
    def __init__(self, path, tab=TAB_PO):
        self.path = path
        self.name = os.path.basename(path)
        self.stem = clean_filename(os.path.splitext(self.name)[0])
        self.mtime = 0
        self.issues = {}
        self.unchecked = []
        self.preview_sheet = None
        self.tab = tab
        upper_stem = self.stem.upper()
        if tab == TAB_PO:
            self.role = "PO"
            self.code = document_code(self.stem)
        elif upper_stem.startswith("VGM"):
            self.role = "IGNORED"
            self.code = self.stem
        elif upper_stem.startswith("DI"):
            self.role = "DI"
            self.code = self.stem.split("_", 1)[1] if "_" in self.stem else self.stem
        else:
            self.role = "SI"
            self.code = self.stem
        self.notes = []
        self.sheets = []
        self.state = "ignored" if self.role == "IGNORED" else "loading"
        self.error = ""

    def set_sheets(self, sheets):
        previous = {sheet.name: sheet for sheet in self.sheets}
        self.sheets = [Sheet(name, hidden) for name, hidden in sheets]
        for sheet in self.sheets:
            if sheet.name in previous:
                sheet.selected = previous[sheet.name].selected
                sheet.pdf_name = previous[sheet.name].pdf_name
                continue
            if self.role == "SI":
                if is_si_sheet(sheet.name):
                    sheet.selected = True
                    sheet.pdf_name = si_pdf_name(self.stem)
                continue
            prefix = auto_prefix(sheet.name)
            if prefix:
                sheet.selected = True
                sheet.pdf_name = f"{prefix}_{self.code}"

    def update_notes(self, si_values):
        self.notes = []
        if self.role != "SI":
            return
        if not any(is_si_sheet(sheet.name) for sheet in self.sheets):
            self.notes.append(f"Sheet {SI_SHEET} tidak ditemukan.")
            return
        parts = si_parts(self.stem)
        if not parts:
            self.notes.append("Nama file tidak mengikuti pola: BL 3X40HC POD SI ...")
            return
        quantity, pod = si_values or (None, None)
        quantity = re.sub(r"\s+", "", str(quantity or "")).upper()
        pod = " ".join(str(pod or "").upper().split())
        if quantity != parts[1]:
            self.notes.append(f"Nama file {parts[1]}, tapi {SI_QUANTITY_CELL} berisi {quantity or 'kosong'}.")
        if pod != " ".join(parts[2].upper().split()):
            self.notes.append(f"Nama file POD {parts[2]}, tapi {SI_POD_CELL} berisi {pod or 'kosong'}.")

    def selected_sheets(self):
        return [sheet for sheet in self.sheets if sheet.selected]

    def issue_count(self):
        return sum(len(rows) for rows in self.issues.values())


# ---------------------------------------------------------------- Excel / WPS
def set_quietly(target, attribute, value):
    try:
        setattr(target, attribute, value)
    except Exception:
        pass


def start_engine():
    """Buka instance baru agar Excel yang sedang kamu pakai tidak terganggu."""
    import win32com.client

    errors = []
    for label, prog_id in ENGINES:
        try:
            app = win32com.client.DispatchEx(prog_id)
            app.Visible = False
            app.DisplayAlerts = False
            set_quietly(app, "AskToUpdateLinks", False)
            set_quietly(app, "EnableEvents", False)
            set_quietly(app, "AutomationSecurity", FORCE_SECURITY_DISABLE_MACROS)
            return app, label
        except Exception as exc:
            errors.append(f"{prog_id}: {exc}")
    raise RuntimeError(
        "Microsoft Excel atau WPS Spreadsheets tidak ditemukan.\n" + "\n".join(errors)
    )


def open_workbook(app, path):
    # Mode baca saja. Password palsu membuat file terkunci langsung gagal,
    # bukan memunculkan jendela password yang menghentikan seluruh proses.
    return app.Workbooks.Open(os.path.abspath(path), 0, True, 5, NO_PASSWORD)


def close_workbook(workbook):
    if workbook is not None:
        try:
            workbook.Close(False)
        except Exception:
            pass


def quit_engine(app):
    if app is not None:
        try:
            app.Quit()
        except Exception:
            pass


def engine_alive(app):
    try:
        _ = app.Version
        return True
    except Exception:
        return False


def is_wide(sheet, setup):
    """Tentukan orientasi dari area yang dicetak."""
    try:
        area = setup.PrintArea
        target = sheet.Range(area) if area else sheet.UsedRange
        return float(target.Width) > float(target.Height)
    except Exception:
        return False


def fit_one_page(sheet):
    setup = sheet.PageSetup
    orientation = XL_LANDSCAPE if is_wide(sheet, setup) else XL_PORTRAIT
    try:
        sheet.ResetAllPageBreaks()
    except Exception:
        pass
    set_quietly(setup, "Orientation", orientation)
    setup.Zoom = False
    setup.FitToPagesWide = 1
    setup.FitToPagesTall = 1


def export_sheet(workbook, sheet_name, pdf_path, one_page):
    sheet = workbook.Sheets.Item(sheet_name)
    if sheet.Visible != XL_SHEET_VISIBLE:
        sheet.Visible = XL_SHEET_VISIBLE
    if one_page:
        fit_one_page(sheet)
    try:
        workbook.Application.CalculateFull()
    except Exception:
        pass
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
    sheet.ExportAsFixedFormat(XL_TYPE_PDF, os.path.abspath(pdf_path))


def check_sheet(sheet):
    """Periksa satu sheet. None berarti tabelnya tidak dikenali."""
    used = sheet.UsedRange
    problems = check_table(used.Value2, used.Row, used.Column)
    if problems is None:
        return None
    visible = []
    for problem in problems:
        try:
            if sheet.Rows.Item(problem["row"]).Hidden:
                continue
        except Exception:
            pass
        visible.append(problem)
    return visible


def read_worker(paths, cancel, report):
    import pythoncom

    pythoncom.CoInitialize()
    app = None
    try:
        while not cancel.is_set():
            try:
                path = paths.get_nowait()
            except queue.Empty:
                break
            workbook = None
            try:
                if app is None or not engine_alive(app):
                    app, _ = start_engine()
                mtime = os.path.getmtime(path)
                workbook = open_workbook(app, path)
                try:
                    app.CalculateFull()
                except Exception:
                    pass
                sheets = []
                issues = {}
                unchecked = []
                si_values = None
                for index in range(1, workbook.Sheets.Count + 1):
                    sheet = workbook.Sheets.Item(index)
                    sheets.append((sheet.Name, sheet.Visible != XL_SHEET_VISIBLE))
                    if is_si_sheet(sheet.Name):
                        try:
                            si_values = (
                                sheet.Range(SI_QUANTITY_CELL).Value2,
                                sheet.Range(SI_POD_CELL).Value2,
                            )
                        except Exception:
                            si_values = None
                    stem = os.path.basename(path).upper()
                    is_si_file = is_si_sheet(sheet.Name) and not stem.startswith(("DI", "PO"))
                    if auto_prefix(sheet.Name):
                        try:
                            problems = check_sheet(sheet)
                        except Exception:
                            problems = None
                        if problems is None:
                            unchecked.append(sheet.Name)
                        elif problems:
                            issues[sheet.Name] = problems
                    if auto_prefix(sheet.Name) or is_si_file:
                        try:
                            used = sheet.UsedRange
                            typos = di_data.scan_sheet_typos(used.Value2, used.Row)
                        except Exception:
                            typos = []
                        if typos:
                            issues.setdefault(sheet.Name, []).extend(typos)
                report("file_read", (path, sheets, issues, unchecked, mtime, si_values))
            except Exception as exc:
                report("file_error", (path, str(exc)))
            finally:
                close_workbook(workbook)
    finally:
        quit_engine(app)
        pythoncom.CoUninitialize()
        report("read_finished", None)


def convert_worker(plan, one_page, cancel, report):
    """plan berisi (path Excel, [(nama sheet, path PDF), ...])."""
    import pythoncom

    pythoncom.CoInitialize()
    app = None
    done = 0
    succeeded = 0
    failed = []
    total = sum(len(jobs) for _, jobs in plan)
    started = time.time()
    try:
        for file_number, (path, jobs) in enumerate(plan, start=1):
            if cancel.is_set():
                break
            name = os.path.basename(path)
            report("status", f"File {file_number}/{len(plan)}: {name}")
            workbook = None
            try:
                if app is None or not engine_alive(app):
                    app, _ = start_engine()
                workbook = open_workbook(app, path)
            except Exception as exc:
                failed.extend((name, sheet, str(exc)) for sheet, _ in jobs)
                done += len(jobs)
                report("progress", (done, total))
                continue
            for sheet_name, pdf_path in jobs:
                if cancel.is_set():
                    break
                try:
                    export_sheet(workbook, sheet_name, pdf_path, one_page)
                    succeeded += 1
                except Exception as exc:
                    failed.append((name, sheet_name, str(exc)))
                done += 1
                report("progress", (done, total))
            # Tutup tanpa menyimpan, jadi file Excel asli tidak berubah.
            close_workbook(workbook)
    finally:
        quit_engine(app)
        pythoncom.CoUninitialize()
        report(
            "convert_done",
            {
                "succeeded": succeeded,
                "failed": failed,
                "cancelled": cancel.is_set(),
                "seconds": time.time() - started,
            },
        )


def preview_worker(requests, report):
    import pythoncom

    pythoncom.CoInitialize()
    app = None
    try:
        while True:
            try:
                request = requests.get(timeout=60)
            except queue.Empty:
                break
            while request is not None and not requests.empty():
                request = requests.get_nowait()
            if request is None:
                break
            key, path, sheet_name, one_page, pdf_path = request
            workbook = None
            try:
                if app is None or not engine_alive(app):
                    app, _ = start_engine()
                workbook = open_workbook(app, path)
                export_sheet(workbook, sheet_name, pdf_path, one_page)
                close_workbook(workbook)
                workbook = None
                import pypdfium2

                document = pypdfium2.PdfDocument(pdf_path)
                total = len(document)
                pages = [
                    document[number].render(scale=2).to_pil().convert("RGB")
                    for number in range(min(total, MAX_PREVIEW_PAGES))
                ]
                document.close()
                report("preview_ready", (key, pages, total, pdf_path))
            except Exception as exc:
                report("preview_error", (key, str(exc)))
            finally:
                close_workbook(workbook)
    finally:
        quit_engine(app)
        pythoncom.CoUninitialize()
        report("preview_stopped", None)


# ------------------------------------------------------------------- Tampilan
NAV_ICONS = {
    "Dasbor": "layout-dashboard", "Kalender ETD": "calendar-days", "Master DI": "ship", "Master PO": "ship", "Goods Description": "file-text",
    "CBM & Berat": "boxes", "Harga": "hand-coins", "Buat Excel": "file-plus", "PDF": "file-down",
}


class SideNav:
    """Menu samping dengan nama lengkap, dibagi Data dan Proses."""

    def __init__(self, parent, command):
        self.command = command
        self.buttons = {}
        self.values = []
        self.frame = ctk.CTkFrame(parent, fg_color="transparent")
        self.frame.pack(fill="both", expand=True, padx=16)
        self.configure(values=di_ui.PO_LABELS)

    def configure(self, values=None):
        self.values = list(values)
        for child in self.frame.winfo_children():
            child.destroy()
        self.buttons = {}
        self.animators = {}
        sections = (("RINGKASAN", self.values[:2]), ("DATA", self.values[2:6]), ("PROSES", self.values[6:]))
        collapsed = getattr(self, "collapsed", False)
        self.section_labels = []
        for title, names in sections:
            label = ctk.CTkLabel(
                self.frame, text="" if collapsed else title, font=font(12, "bold"), text_color="#EADBC8", anchor="w", height=10 if collapsed else 26,
            )
            label.pack(fill="x", padx=10, pady=(10, 2))
            self.section_labels.append((label, title))
            for name in names:
                button = ctk.CTkButton(
                    self.frame, text="" if collapsed else name, font=font(15), height=42, corner_radius=9, anchor="center" if collapsed else "w",
                    fg_color="transparent", hover_color=TAB_IDLE_HOVER, text_color="#FFFFFF", command=lambda n=name: self.command(n),
                )
                button.pack(fill="x", pady=2)
                self.buttons[name] = button
                icon = NAV_ICONS.get(name)
                if icon and icons.available(icon):
                    self.animators[name] = icons.Animator(button, icon, "#FFFFFF", 20)

    def set_collapsed(self, collapsed):
        self.collapsed = collapsed
        self.frame.pack_configure(padx=8 if collapsed else 16)
        for label, title in getattr(self, "section_labels", []):
            label.configure(text="" if collapsed else title, height=10 if collapsed else 26)
        for name, button in self.buttons.items():
            button.configure(text="" if collapsed else name, anchor="center" if collapsed else "w")

    def set(self, name):
        self.current = name
        for label, button in self.buttons.items():
            active = label == name
            button.configure(
                fg_color="#FFFFFF" if active else "transparent", text_color=HEADER_BAND[0] if active else "#FFFFFF",
                hover_color="#FFFFFF" if active else TAB_IDLE_HOVER, font=font(15, "bold" if active else "normal"),
            )
            if label in self.animators:
                self.animators[label].set_color(HEADER_BAND[0] if active else "#FFFFFF")

    def cget(self, _name):
        return self.values


class App(ctk.CTk):
    def setup_drop(self):
        """Seret file ke aplikasi. Kalau library-nya gagal, tombol tetap bisa dipakai."""
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD

            TkinterDnD._require(self)
            for widget in [self] + list(self.data_manager.pages.values()):
                widget.tk.call("tkdnd::drop_target", "register", widget._w, (DND_FILES,))
            for widget in [self] + list(self.data_manager.pages.values()):
                self.bind_drop(widget)
        except Exception:
            pass

    def bind_drop(self, widget):
        def on_drop(data):
            self.after_idle(lambda: self.handle_drop(list(self.tk.splitlist(data))))
            return "copy"

        command = widget.register(on_drop)
        widget.tk.call("bind", widget._w, "<<Drop:DND_Files>>", f"{command} %D")

    def handle_drop(self, paths):
        if self.data_manager.handle_drop(paths):
            return
        if self.current_page == di_ui.PAGE_BUILD:
            return
        if self.converting:
            return
        excel = []
        for path in paths:
            if path.lower().endswith(".zip") and self.tab == TAB_DI:
                self.add_zip(os.path.normpath(path))
            elif os.path.isdir(path):
                excel.extend(excel_files_in(path, recursive=self.tab == TAB_DI))
            elif path.lower().endswith(EXCEL_EXTENSIONS):
                excel.append(path)
        self.add_paths(excel)

    def __init__(self):
        super().__init__(fg_color=BACKGROUND)
        ctk.set_appearance_mode("system")
        self.title(APP_TITLE)
        width = min(1600, self.winfo_screenwidth() - 40)
        height = min(860, self.winfo_screenheight() - 90)
        self.geometry(f"{width}x{height}")
        self.minsize(1240, 680)
        if os.name == "nt" and os.path.exists(ICON_PATH):
            self.iconbitmap(ICON_PATH)

        self.tab = TAB_PO
        self.tabs = {
            name: {"files": [], "current": None, "rows": {}, "list": None, "empty": None}
            for name in (TAB_PO, TAB_DI)
        }
        self.events = queue.Queue()
        self.read_queue = queue.Queue()
        self.cancel = threading.Event()
        self.reading = False
        self.converting = False
        self.result_folder = ""
        self.convert_after_read = False
        self.convert_only = None
        self.temp_dir = tempfile.mkdtemp(prefix="pdf_gen_")
        self.preview_requests = queue.Queue()
        self.preview_thread = None
        self.preview_cache = {}
        self.preview_key = None
        self.preview_pages = []
        self.preview_total = 0
        self.preview_path = None
        self.preview_number = 0
        self.preview_image = None
        self.preview_resize_job = None
        self.rules = ctk.StringVar()
        self.one_page = ctk.BooleanVar(value=True)

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=5, uniform="panel")
        self.content.grid_columnconfigure(1, weight=6, uniform="panel")
        self.content.grid_columnconfigure(2, weight=7, uniform="panel")
        self.content.grid_rowconfigure(1, weight=1)
        self.build_header()
        self.build_page_title()
        self.build_file_panel()
        self.build_sheet_panel()
        self.build_preview_panel()
        self.build_footer()
        self.build_di_pages()
        self.toast = motion.Toast(self)
        self.bind_all("<Control-f>", self.focus_search, add="+")
        self.bind_all("<Control-F>", self.focus_search, add="+")
        self.setup_drop()
        self.show_current_file()
        self.update_summary()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(100, self.poll_events)
        self.after(5000, self.collect_garbage)

    def focus_search(self, _event=None):
        """Ctrl+F membuka kotak cari di halaman data yang sedang terbuka."""
        page = self.data_manager.pages.get(di_ui.page_key(self.tab, self.current_page)) if hasattr(self, "data_manager") else None
        bar = getattr(page, "search_bar", None)
        if bar is not None:
            bar.focus()
            return "break"
        return None

    @property
    def files(self):
        return self.tabs[self.tab]["files"]

    @property
    def current(self):
        return self.tabs[self.tab]["current"]

    @current.setter
    def current(self, item):
        self.tabs[self.tab]["current"] = item

    @property
    def file_rows(self):
        return self.tabs[self.tab]["rows"]

    @property
    def file_list(self):
        return self.tabs[self.tab]["list"]

    @property
    def file_empty(self):
        return self.tabs[self.tab]["empty"]

    def build_di_pages(self):
        for name in ("ACCENT", "ACCENT_HOVER", "ACCENT_SOFT", "BORDER", "CARD", "DANGER", "MUTED", "NEUTRAL", "TEXT"):
            setattr(self, name, globals()[name])
        self.font = font
        self.engine_starter = start_engine
        self.engine_stopper = quit_engine
        self.di_area = ctk.CTkFrame(self.content, fg_color=BACKGROUND, corner_radius=0)
        self.di_area.grid_columnconfigure(0, weight=1)
        self.di_area.grid_rowconfigure(0, weight=1)
        self.data_manager = di_ui.DataManager(self, self.di_area)
        for page in self.data_manager.pages.values():
            page.grid(row=0, column=0, sticky="nsew", padx=24, pady=(20, 0))
            page.grid_remove()
        self.current_page = di_ui.PAGE_PDF

    def show_page(self, name):
        self.current_page = name
        self.page_selector.set(name)
        target = di_ui.page_key(self.tab, name)
        for page_name, page in self.data_manager.pages.items():
            if page_name == target:
                page.grid()
            else:
                page.grid_remove()
        if name == di_ui.PAGE_PDF:
            self.di_area.grid_remove()
            self.title_row.grid()
            self.update_summary()
            self.update_controls()
            return
        self.title_row.grid_remove()
        self.di_area.grid(row=0, rowspan=2, column=0, columnspan=3, sticky="nsew")
        self.di_area.lift()
        self.update_controls()
        self.data_manager.ensure_loaded()
        self.data_manager.pages[target].refresh()

    def switch_tab(self, tab):
        if tab == self.tab:
            return
        self.file_list.grid_remove()
        self.tab = tab
        self.file_list.grid()
        if tab == TAB_DI:
            self.add_zip_button.pack(side="left", after=self.add_folder_button)
        else:
            self.add_zip_button.pack_forget()
        mapping = {di_ui.PAGE_MASTER: di_ui.PAGE_MASTER_PO, di_ui.PAGE_MASTER_PO: di_ui.PAGE_MASTER}
        self.tab_selector.set(tab)
        if hasattr(self, "tab_mini"):
            self.tab_mini.configure(text=tab)
        self.page_selector.configure(values=di_ui.PAGES if tab == TAB_DI else di_ui.PO_LABELS)
        self.show_page(mapping.get(self.current_page, self.current_page))
        self.show_current_file()
        self.request_preview()
        self.update_summary()
        self.update_controls()
        self.progress.set(0)
        self.status.configure(text=f"Tab {tab}: {len(self.files)} file.")

    def card(self, parent):
        return ctk.CTkFrame(
            parent, fg_color=CARD, border_color=BORDER, border_width=1, corner_radius=12
        )

    def button(self, parent, text, command, primary=False, danger=False, **options):
        colors = (
            {"fg_color": ACCENT, "hover_color": ACCENT_HOVER, "text_color": "#FFFFFF"}
            if primary
            else {"fg_color": NEUTRAL, "hover_color": NEUTRAL_HOVER, "text_color": TEXT}
        )
        if danger:
            colors = {"fg_color": DANGER_SOFT, "hover_color": DANGER_SOFT_HOVER, "text_color": DANGER}
        options.setdefault("height", 34)
        icon = None if options.pop("plain", False) or "image" in options else icons.icon_for(text)
        if icon and "width" in options and len(text) * 8 + 44 > options["width"]:
            icon = None
        if icon and text.startswith("+ "):
            text = text[2:]
        button = ctk.CTkButton(
            parent, text=text, command=command, font=font(13, "bold" if primary else "normal"),
            corner_radius=8, compound="left", **colors, **options,
        )
        if icon:
            button.icon_animator = icons.Animator(button, icon, colors["text_color"], 16)
        return button

    def build_header(self):
        side = ctk.CTkFrame(self, fg_color=HEADER_BAND, corner_radius=0, width=236)
        side.grid(row=0, column=0, sticky="nsew")
        side.grid_propagate(False)
        side.pack_propagate(False)
        brand = ctk.CTkFrame(side, fg_color="transparent")
        brand.pack(fill="x", padx=20, pady=(22, 16))
        try:
            from PIL import Image

            self.logo = ctk.CTkImage(Image.open(LOGO_PATH if os.path.exists(LOGO_PATH) else ICON_PATH), size=(40, 40))
            ctk.CTkLabel(brand, image=self.logo, text="").pack(side="left", padx=(0, 10))
        except Exception:
            pass
        names = ctk.CTkFrame(brand, fg_color="transparent")
        names.pack(side="left")
        ctk.CTkLabel(names, text="Dockit", font=font(21, "bold"), text_color="#FFFFFF", height=26).pack(anchor="w")
        ctk.CTkLabel(names, text=f"Versi {APP_VERSION}", font=font(13), text_color="#EADBC8", height=18).pack(anchor="w")
        self.tab_selector = ctk.CTkSegmentedButton(
            side, values=[TAB_PO, TAB_DI], command=self.switch_tab, font=font(15, "bold"), height=42,
            selected_color=TAB_ACTIVE, selected_hover_color=TAB_ACTIVE_HOVER, unselected_color=TAB_IDLE,
            unselected_hover_color=TAB_IDLE_HOVER, fg_color=TAB_IDLE, text_color="#FFFFFF", text_color_disabled="#EADBC8",
        )
        self.tab_selector.set(TAB_PO)
        self.tab_selector.pack(fill="x", padx=16, pady=(0, 18))
        self.side, self.brand_names = side, names
        self.collapse_button = ctk.CTkButton(
            side, text="«  Lipat sidebar", font=font(13), height=36, corner_radius=9, fg_color="transparent", border_width=1,
            border_color="#2B4A7A", hover_color=TAB_IDLE_HOVER, text_color="#EADBC8", command=self.toggle_sidebar,
        )
        self.collapse_button.pack(side="bottom", fill="x", padx=16, pady=16)
        self.page_selector = SideNav(side, self.show_page)
        if di_ui.data.load_settings().get("sidebar_collapsed"):
            self.after(50, lambda: self.toggle_sidebar(True))

    def toggle_sidebar(self, collapsed=None):
        """Sidebar bisa dilipat jadi ikon saja supaya tabel lebih lebar."""
        collapsed = not getattr(self, "sidebar_collapsed", False) if collapsed is None else collapsed
        self.sidebar_collapsed = collapsed
        self.side.configure(width=76 if collapsed else 236)
        if not hasattr(self, "tab_mini"):
            self.tab_mini = ctk.CTkButton(
                self.side, text=self.tab, font=font(14, "bold"), height=42, corner_radius=9, fg_color=TAB_ACTIVE, hover_color=TAB_ACTIVE_HOVER,
                text_color="#FFFFFF", command=lambda: self.switch_tab(TAB_DI if self.tab == TAB_PO else TAB_PO),
            )
        if collapsed:
            self.brand_names.pack_forget()
            self.tab_selector.pack_forget()
            self.tab_mini.configure(text=self.tab)
            self.tab_mini.pack(fill="x", padx=8, pady=(0, 18), before=self.page_selector.frame)
        else:
            self.brand_names.pack(side="left")
            self.tab_mini.pack_forget()
            self.tab_selector.pack(fill="x", padx=16, pady=(0, 18), before=self.page_selector.frame)
        self.collapse_button.configure(text="»" if collapsed else "«  Lipat sidebar")
        self.collapse_button.pack_configure(padx=8 if collapsed else 16)
        self.page_selector.set_collapsed(collapsed)
        settings = di_ui.data.load_settings()
        if settings.get("sidebar_collapsed") != collapsed:
            settings["sidebar_collapsed"] = collapsed
            di_ui.data.save_settings(settings)

    def build_page_title(self):
        self.title_row = ctk.CTkFrame(self.content, fg_color="transparent")
        self.title_row.grid(row=0, column=0, columnspan=3, sticky="ew", padx=24, pady=(20, 12))
        ctk.CTkLabel(self.title_row, text="Buat PDF", font=font(26, "bold"), text_color=TEXT).pack(side="left")
        self.summary = ctk.CTkLabel(
            self.title_row, text="", font=font(14, "bold"), text_color=ACCENT, fg_color=ACCENT_SOFT, corner_radius=9, height=40,
        )
        self.summary.pack(side="right", ipadx=14)

    def build_file_panel(self):
        panel = self.card(self.content)
        panel.grid(row=1, column=0, sticky="nsew", padx=(24, 8))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            panel, text="File Excel", font=font(16, "bold"), text_color=TEXT
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(14, 8))
        actions = ctk.CTkFrame(panel, fg_color="transparent")
        actions.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))
        self.add_files_button = self.button(
            actions, "+ File", self.choose_files, primary=True, width=72
        )
        self.add_files_button.pack(side="left")
        self.add_folder_button = self.button(actions, "+ Folder", self.choose_folder, width=84)
        self.add_folder_button.pack(side="left", padx=6)
        self.add_zip_button = self.button(actions, "+ ZIP", self.choose_zip, width=66)
        self.clear_button = self.button(actions, "Kosongkan", self.clear_files, width=90)
        self.clear_button.pack(side="right")

        for name, hint in (
            (TAB_PO, "Belum ada file PO.\nTambahkan file atau folder."),
            (TAB_DI, "Belum ada file DI.\nTambahkan ZIP, folder, atau file."),
        ):
            file_list = ctk.CTkScrollableFrame(panel, fg_color="transparent")
            file_list.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 8))
            file_list.grid_columnconfigure(0, weight=1)
            self.tabs[name]["list"] = file_list
            self.tabs[name]["empty"] = ctk.CTkLabel(
                file_list, text=hint, font=font(13), text_color=MUTED
            )
        self.tabs[TAB_DI]["list"].grid_remove()

    def build_sheet_panel(self):
        panel = self.card(self.content)
        panel.grid(row=1, column=1, sticky="nsew", padx=8)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(2, weight=1)

        self.sheet_title = ctk.CTkLabel(
            panel, text="", font=font(16, "bold"), text_color=TEXT, anchor="w"
        )
        self.sheet_title.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))
        actions = ctk.CTkFrame(panel, fg_color="transparent")
        actions.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))
        self.button(actions, "Pilih semua", lambda: self.mark_current(True)).pack(side="left")
        self.button(actions, "Hapus pilihan", lambda: self.mark_current(False)).pack(
            side="left", padx=6
        )
        self.button(actions, "Terapkan ke semua", self.apply_to_all, primary=True).pack(
            side="right"
        )

        self.sheet_list = ctk.CTkScrollableFrame(panel, fg_color="transparent")
        self.sheet_list.grid(row=2, column=0, sticky="nsew", padx=8)
        self.sheet_list.grid_columnconfigure(0, weight=1)
        self.sheet_list.grid_columnconfigure(1, weight=1)
        self.sheet_list.grid_columnconfigure(2, weight=0)

        rule_box = ctk.CTkFrame(panel, fg_color="transparent")
        rule_box.grid(row=3, column=0, sticky="ew", padx=16, pady=(8, 14))
        rule_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            rule_box, text="Aturan sheet lain, contoh: INV-2=CI2_{kode}",
            font=font(12), text_color=MUTED,
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        ctk.CTkEntry(
            rule_box, textvariable=self.rules, font=font(13), height=34, border_color=BORDER,
        ).grid(row=1, column=0, sticky="ew", padx=(0, 6), pady=(4, 0))
        self.button(rule_box, "Pakai aturan", self.apply_rules).grid(row=1, column=1, pady=(4, 0))

    def build_preview_panel(self):
        panel = self.card(self.content)
        panel.grid(row=1, column=2, sticky="nsew", padx=(8, 24))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(panel, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))
        top.grid_columnconfigure(0, weight=1)
        self.preview_title = ctk.CTkLabel(
            top, text="Preview PDF", font=font(16, "bold"), text_color=TEXT, anchor="w"
        )
        self.preview_title.grid(row=0, column=0, sticky="ew")
        self.previous_page_button = self.button(top, "‹", lambda: self.turn_page(-1), width=32)
        self.previous_page_button.grid(row=0, column=1, padx=(6, 0))
        self.page_label = ctk.CTkLabel(top, text="", font=font(12), text_color=MUTED, width=56)
        self.page_label.grid(row=0, column=2)
        self.next_page_button = self.button(top, "›", lambda: self.turn_page(1), width=32)
        self.next_page_button.grid(row=0, column=3)
        self.open_pdf_button = self.button(top, "Buka PDF", self.open_preview_pdf, width=80)
        self.open_pdf_button.grid(row=0, column=4, padx=(6, 0))

        self.preview_list = ctk.CTkScrollableFrame(panel, fg_color=NEUTRAL)
        self.preview_list.grid(row=1, column=0, sticky="nsew", padx=12)
        self.preview_list.grid_columnconfigure(0, weight=1)
        self.preview_list.bind("<Configure>", self.on_preview_resize)
        self.preview_content = None

        notice = ctk.CTkFrame(panel, fg_color="transparent")
        notice.grid(row=2, column=0, sticky="ew", padx=16, pady=(8, 14))
        notice.grid_columnconfigure(0, weight=1)
        self.check_title = ctk.CTkLabel(
            notice, text="Cek data CI/PL", font=font(13, "bold"), text_color=TEXT, anchor="w"
        )
        self.check_title.grid(row=0, column=0, sticky="ew")
        self.button(notice, "Buka Excel", self.open_excel, width=90, height=28).grid(
            row=0, column=1, padx=6
        )
        self.button(notice, "Cek ulang", self.recheck_current, width=90, height=28).grid(
            row=0, column=2
        )
        self.check_text = ctk.CTkTextbox(
            notice, height=96, font=font(12), fg_color=NEUTRAL, text_color=TEXT,
            wrap="word", state="disabled",
        )
        self.check_text.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 0))

    def build_footer(self):
        footer = self.card(self.content)
        footer.grid(row=2, column=0, columnspan=3, sticky="ew", padx=24, pady=(12, 18))
        footer.grid_columnconfigure(0, weight=1)

        self.progress = motion.PixelBar(footer, height=12)
        self.progress.set(0)
        self.progress.grid(row=0, column=0, sticky="ew", padx=16, pady=(18, 0))
        self.status = ctk.CTkLabel(
            footer, text="Tambahkan file Excel untuk mulai.", font=font(12),
            text_color=MUTED, anchor="w",
        )
        self.status.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))
        ctk.CTkSwitch(
            footer, text="Muat 1 halaman", variable=self.one_page, font=font(13),
            text_color=TEXT, progress_color=ACCENT, command=self.request_preview,
        ).grid(row=0, column=1, rowspan=2, padx=(4, 16))
        self.cancel_button = self.button(footer, "Batal", self.request_cancel, height=40, width=200)
        self.convert_button = self.button(
            footer, "Convert ke PDF", self.start_convert, primary=True, height=40, width=200
        )
        self.action_grid = {"row": 0, "column": 2, "rowspan": 2, "padx": (0, 16), "pady": 14}
        self.convert_button.grid(**self.action_grid)

    # ---------------------------------------------------------- daftar file
    def choose_files(self):
        paths = filedialog.askopenfilenames(title="Pilih file Excel", filetypes=EXCEL_TYPES)
        self.add_paths(paths)

    def choose_folder(self):
        folder = filedialog.askdirectory(title="Pilih folder berisi file Excel")
        if not folder:
            return
        # Tab DI ikut membaca subfolder karena satu folder bisa berisi beberapa pengiriman.
        paths = excel_files_in(folder, recursive=self.tab == TAB_DI)
        if not paths:
            messagebox.showinfo(APP_TITLE, "Tidak ada file Excel di folder itu.")
        self.add_paths(paths)

    def choose_zip(self):
        zip_path = filedialog.askopenfilename(
            title="Pilih file ZIP", filetypes=[("File ZIP", "*.zip")]
        )
        if zip_path:
            self.add_zip(os.path.normpath(zip_path))

    def add_zip(self, zip_path):
        self.status.configure(text="Mengekstrak ZIP...")
        self.update_idletasks()
        try:
            folder = extract_zip(zip_path)
        except (OSError, zipfile.BadZipFile) as exc:
            self.status.configure(text="ZIP gagal dibuka.")
            messagebox.showerror(APP_TITLE, f"ZIP gagal dibuka:\n\n{exc}")
            return
        paths = excel_files_in(folder, recursive=True)
        if not paths:
            messagebox.showinfo(APP_TITLE, "Tidak ada file Excel di dalam ZIP itu.")
        self.add_paths(paths)

    def add_paths(self, paths):
        known = {item.path.lower() for item in self.files}
        added = []
        for path in paths:
            path = os.path.normpath(path)
            if path.lower() in known:
                continue
            known.add(path.lower())
            item = ExcelFile(path, self.tab)
            self.files.append(item)
            self.add_file_row(item)
            if item.state == "loading":
                self.read_queue.put(path)
            added.append(item)
        if not added:
            return
        if self.current is None:
            self.select_file(next((item for item in added if item.state == "loading"), added[0]))
        self.start_reader()
        self.update_summary()
        self.update_controls()

    def start_reader(self):
        if not self.reading:
            self.reading = True
            self.cancel.clear()
            threading.Thread(
                target=read_worker,
                args=(self.read_queue, self.cancel, self.post_event),
                daemon=True,
            ).start()

    def post_event(self, kind, data):
        self.events.put((kind, data))

    def add_file_row(self, item):
        self.file_empty.grid_forget()
        row = ctk.CTkFrame(self.file_list, fg_color="transparent", corner_radius=8)
        row.grid(row=len(self.file_rows), column=0, sticky="ew", pady=1)
        row.grid_columnconfigure(0, weight=1)
        name = ctk.CTkLabel(
            row, text=shorten(item.name, 34), font=font(13), text_color=TEXT,
            anchor="w", height=20,
        )
        name.grid(row=0, column=0, sticky="ew", padx=(10, 4), pady=(6, 0))
        detail = ctk.CTkLabel(
            row, text="", font=font(11), text_color=MUTED, anchor="w", height=16
        )
        detail.grid(row=1, column=0, sticky="ew", padx=(10, 4), pady=(0, 6))
        remove = ctk.CTkButton(
            row, text="×", width=26, height=26, font=font(15), fg_color="transparent",
            hover_color=NEUTRAL_HOVER, text_color=MUTED, command=lambda: self.remove_file(item),
        )
        remove.grid(row=0, column=1, rowspan=2, padx=6)
        for widget in (row, name, detail):
            widget.bind("<Button-1>", lambda _, target=item: self.select_file(target))
        self.file_rows[item.path] = (row, detail, remove)
        self.refresh_file_row(item)

    def refresh_file_row(self, item):
        row_widgets = self.file_rows.get(item.path)
        if not row_widgets:
            return
        row, detail, _ = row_widgets
        row.configure(fg_color=ACCENT_SOFT if item is self.current else "transparent")
        count = len(item.selected_sheets())
        problems = item.issue_count()
        if item.state == "ignored":
            detail.configure(text="Dilewati, bukan SI/DI", text_color=MUTED)
        elif item.state == "ready" and problems:
            detail.configure(text=f"⚠ {problems} baris perlu dicek", text_color=DANGER)
        elif item.state == "ready" and item.notes:
            detail.configure(text=f"⚠ {count} PDF, perlu dicek", text_color=WARNING)
        elif item.state == "ready" and item.unchecked and count:
            detail.configure(text=f"{count} PDF, perlu cek manual", text_color=WARNING)
        elif item.state == "loading":
            detail.configure(text="Membaca sheet...", text_color=MUTED)
        elif item.state == "error":
            detail.configure(text="Tidak bisa dibuka", text_color=DANGER)
        elif count == 0:
            detail.configure(text=f"{len(item.sheets)} sheet, belum ada yang dipilih",
                             text_color=WARNING)
        else:
            detail.configure(text=f"{count} PDF dari {len(item.sheets)} sheet", text_color=MUTED)

    def select_file(self, item):
        previous, self.current = self.current, item
        if previous is not None:
            self.refresh_file_row(previous)
        self.refresh_file_row(item)
        self.show_current_file()
        self.request_preview()

    def remove_file(self, item):
        if self.converting:
            return
        self.files.remove(item)
        self.file_rows.pop(item.path)[0].destroy()
        if self.current is item:
            self.current = None
            if self.files:
                self.select_file(self.files[0])
            else:
                self.show_current_file()
                self.request_preview()
        self.after_file_change()

    def clear_files(self):
        if self.converting:
            return
        self.files.clear()
        for row_widgets in self.file_rows.values():
            row_widgets[0].destroy()
        self.file_rows.clear()
        self.current = None
        self.show_current_file()
        self.request_preview()
        self.after_file_change()

    def after_file_change(self):
        if not self.files:
            self.file_empty.grid(row=0, column=0, pady=60)
        self.update_summary()
        self.update_controls()

    # --------------------------------------------------------- daftar sheet
    def show_current_file(self):
        self.show_check_result()
        for widget in self.sheet_list.winfo_children():
            widget.destroy()
        if not self.files:
            self.file_empty.grid(row=0, column=0, pady=60)
        item = self.current
        if item is None:
            self.sheet_title.configure(text="Sheet")
            self.sheet_message("Pilih file di sebelah kiri untuk melihat sheet-nya.")
            return
        self.sheet_title.configure(text=shorten(item.name, 60))
        if item.state == "ignored":
            self.sheet_message("File ini dilewati karena bukan file SI atau DI.")
            return
        if item.state == "loading":
            self.sheet_message("Membaca daftar sheet...")
            return
        if item.state == "error":
            self.sheet_message(f"File tidak bisa dibuka.\n{item.error}", DANGER)
            return
        ctk.CTkLabel(
            self.sheet_list, text="Sheet", font=font(12, "bold"), text_color=MUTED, anchor="w"
        ).grid(row=0, column=0, sticky="ew", padx=8, pady=(0, 4))
        ctk.CTkLabel(
            self.sheet_list, text="Nama PDF", font=font(12, "bold"), text_color=MUTED, anchor="w"
        ).grid(row=0, column=1, sticky="ew", padx=8, pady=(0, 4))
        for number, sheet in enumerate(item.sheets, start=1):
            self.add_sheet_row(item, sheet, number)

    def sheet_message(self, text, color=MUTED):
        ctk.CTkLabel(
            self.sheet_list, text=text, font=font(13), text_color=color, wraplength=520
        ).grid(row=0, column=0, columnspan=2, pady=60)

    def add_sheet_row(self, item, sheet, number):
        label = f"{shorten(sheet.name, 18)} (tersembunyi)" if sheet.hidden else shorten(sheet.name, 26)
        if sheet.name in item.issues:
            label = "⚠ " + label
        selected = ctk.BooleanVar(value=sheet.selected)

        def on_toggle():
            sheet.selected = selected.get()
            self.refresh_file_row(item)
            self.update_summary()

        ctk.CTkCheckBox(
            self.sheet_list, text=label, variable=selected, command=on_toggle, font=font(13),
            text_color=TEXT, fg_color=ACCENT, hover_color=ACCENT_HOVER, border_color=MUTED,
            checkbox_width=20, checkbox_height=20,
        ).grid(row=number, column=0, sticky="w", padx=8, pady=4)
        name = ctk.StringVar(value=sheet.pdf_name)
        name.trace_add("write", lambda *_: setattr(sheet, "pdf_name", name.get()))
        ctk.CTkEntry(
            self.sheet_list, textvariable=name, font=font(13), height=32, border_color=BORDER
        ).grid(row=number, column=1, sticky="ew", padx=(8, 4), pady=4)
        active = sheet.name == self.active_preview_sheet(item)
        self.button(
            self.sheet_list, "Lihat", lambda: self.preview_sheet(item, sheet),
            primary=active, width=52, height=32,
        ).grid(row=number, column=2, padx=(0, 8), pady=4)

    def preview_sheet(self, item, sheet):
        item.preview_sheet = sheet.name
        self.show_current_file()
        self.request_preview()

    def mark_current(self, selected):
        if self.current is None:
            return
        for sheet in self.current.sheets:
            sheet.selected = selected
        self.show_current_file()
        self.refresh_file_row(self.current)
        self.update_summary()

    def apply_to_all(self):
        source = self.current
        if source is None or not source.selected_sheets():
            messagebox.showinfo(APP_TITLE, "Centang dulu sheet yang dibutuhkan di file ini.")
            return
        chosen = {sheet.name.lower(): sheet.pdf_name for sheet in source.selected_sheets()}
        missing = 0
        for item in self.files:
            if item is source or item.state != "ready" or item.role != source.role:
                continue
            for sheet in item.sheets:
                sheet.selected = sheet.name.lower() in chosen
                if sheet.selected:
                    sheet.pdf_name = chosen[sheet.name.lower()].replace(source.code, item.code)
            if len(item.selected_sheets()) < len(chosen):
                missing += 1
        self.refresh_all_rows()
        note = f" {missing} file tidak punya semua sheet itu." if missing else ""
        self.status.configure(text=f"Pilihan diterapkan ke semua file.{note}")

    def apply_rules(self):
        rules = parse_rules(self.rules.get())
        if not rules:
            messagebox.showinfo(APP_TITLE, "Tulis aturannya dulu, contoh: INV-2=CI2_{kode}")
            return
        missing = 0
        for item in self.files:
            if item.state != "ready" or (self.current and item.role != self.current.role):
                continue
            for sheet in item.sheets:
                sheet.selected = False
            for pattern, pdf_name in rules:
                matches = [s for s in item.sheets if fnmatch.fnmatchcase(s.name.lower(), pattern)]
                if not matches:
                    missing += 1
                for match_number, sheet in enumerate(matches, start=1):
                    sheet.selected = True
                    if pdf_name:
                        suffix = f" {match_number}" if len(matches) > 1 else ""
                        sheet.pdf_name = pdf_name.replace("{kode}", item.code) + suffix
        self.refresh_all_rows()
        note = f" {missing} aturan tidak menemukan sheet." if missing else ""
        self.status.configure(text=f"Aturan dipakai pada semua file.{note}")

    def show_check_result(self):
        item = self.current
        lines = []
        title = "Cek data CI/PL"
        color = TEXT
        if item is not None and item.state == "ready" and item.role == "SI":
            title = "Cek nama SI"
            if item.notes or item.issues:
                title = "⚠ Perlu dicek"
                color = WARNING
                lines.extend(item.notes)
                for sheet_name, problems in item.issues.items():
                    for problem in problems:
                        lines.append(f"{sheet_name} baris {problem['row']} · {problem['item']}: " + ", ".join(problem["missing"]))
            else:
                title = "✓ Nama SI sesuai isi sheet"
                color = ACCENT
                lines.append(f"Jumlah kontainer dan POD cocok dengan sheet {SI_SHEET}.")
                lines.append("Nomor BL diambil dari nama file.")
        elif item is not None and item.state == "ready":
            count = item.issue_count()
            if count:
                title = f"⚠ {count} baris perlu dicek"
                color = DANGER
                for sheet_name, problems in item.issues.items():
                    for problem in problems:
                        name = f" · {problem['item']}" if problem["item"] else ""
                        lines.append(
                            f"{sheet_name} baris {problem['row']}{name}: "
                            + ", ".join(problem["missing"])
                        )
                lines.append("Perbaiki di Excel, simpan, lalu klik Cek ulang.")
            elif not item.unchecked and any(auto_prefix(s.name) for s in item.sheets):
                title = "✓ Data CI/PL lengkap"
                color = ACCENT
                lines.append("Tidak ada data kosong atau typo di tabel item.")
            elif not item.unchecked:
                lines.append("Tidak ada sheet CI/PL di file ini.")
            for sheet_name in item.unchecked:
                lines.append(f"{sheet_name}: tabelnya tidak dikenali, cek manual lewat preview.")
            if item.unchecked and not count:
                title = "Perlu cek manual"
                color = WARNING
        self.check_title.configure(text=title, text_color=color)
        self.check_text.configure(state="normal")
        self.check_text.delete("1.0", "end")
        self.check_text.insert("1.0", "\n".join(lines))
        self.check_text.configure(state="disabled")

    def open_excel(self):
        if self.current is not None and os.name == "nt":
            os.startfile(self.current.path)

    def recheck_current(self):
        if self.current is None or self.converting or self.current.state == "ignored":
            return
        self.read_queue.put(self.current.path)
        self.status.configure(text=f"Mengecek ulang {self.current.name}...")
        self.start_reader()
        self.update_controls()

    # -------------------------------------------------------------- preview
    def active_preview_sheet(self, item):
        names = [sheet.name for sheet in item.sheets]
        if item.preview_sheet in names:
            return item.preview_sheet
        selected = item.selected_sheets()
        if selected:
            return selected[0].name
        return names[0] if names else None

    def request_preview(self):
        item = self.current
        sheet_name = self.active_preview_sheet(item) if item and item.state == "ready" else None
        if sheet_name is None:
            self.preview_key = None
            self.set_preview_message("Pilih file untuk melihat preview.")
            return
        sheet = next(entry for entry in item.sheets if entry.name == sheet_name)
        one_page = self.one_page.get()
        key = (item.path, sheet_name, one_page, item.mtime)
        self.preview_key = key
        self.preview_title.configure(text=shorten(clean_filename(sheet.pdf_name) + ".pdf", 34))
        if key in self.preview_cache:
            self.show_preview(*self.preview_cache[key])
            return
        self.set_preview_message("Membuat preview...")
        pdf_path = os.path.join(self.temp_dir, uuid.uuid4().hex + ".pdf")
        self.preview_requests.put((key, item.path, sheet_name, one_page, pdf_path))
        self.start_preview_worker()

    def start_preview_worker(self):
        if self.preview_thread is None:
            self.preview_thread = threading.Thread(
                target=preview_worker,
                args=(self.preview_requests, self.post_event),
                daemon=True,
            )
            self.preview_thread.start()

    def clear_preview_content(self):
        if self.preview_content is not None:
            self.preview_content.destroy()
            self.preview_content = None

    def set_preview_message(self, text, color=MUTED):
        self.preview_pages = []
        self.preview_total = 0
        self.preview_path = None
        self.clear_preview_content()
        self.preview_content = ctk.CTkLabel(
            self.preview_list, text=text, font=font(13), text_color=color, wraplength=420
        )
        self.preview_content.grid(row=0, column=0, pady=80)
        self.update_page_controls()

    def show_preview(self, pages, total, pdf_path):
        self.preview_pages = pages
        self.preview_total = total
        self.preview_path = pdf_path
        self.preview_number = 0
        self.draw_preview_page()

    def draw_preview_page(self):
        if not self.preview_pages:
            return
        page = self.preview_pages[self.preview_number]
        scaling = self._get_window_scaling()
        width = max(240, int(self.preview_list.winfo_width() / scaling) - 34)
        height = int(width * page.height / page.width)
        self.preview_image = ctk.CTkImage(light_image=page, size=(width, height))
        self.clear_preview_content()
        self.preview_content = ctk.CTkLabel(self.preview_list, image=self.preview_image, text="")
        self.preview_content.grid(row=0, column=0, pady=4)
        self.update_page_controls()

    def update_page_controls(self):
        count = len(self.preview_pages)
        self.page_label.configure(
            text=f"{self.preview_number + 1} / {self.preview_total}" if count else ""
        )
        previous_state = "normal" if count and self.preview_number > 0 else "disabled"
        next_state = "normal" if self.preview_number + 1 < count else "disabled"
        self.previous_page_button.configure(state=previous_state)
        self.next_page_button.configure(state=next_state)
        self.open_pdf_button.configure(state="normal" if self.preview_path else "disabled")

    def turn_page(self, step):
        self.preview_number += step
        self.draw_preview_page()

    def open_preview_pdf(self):
        if self.preview_path and os.name == "nt":
            os.startfile(self.preview_path)

    def on_preview_resize(self, _event):
        if self.preview_resize_job is not None:
            self.after_cancel(self.preview_resize_job)
        self.preview_resize_job = self.after(200, self.draw_preview_page)

    def handle_preview_ready(self, key, pages, total, pdf_path):
        while len(self.preview_cache) >= MAX_PREVIEW_CACHE:
            self.preview_cache.pop(next(iter(self.preview_cache)))
        self.preview_cache[key] = (pages, total, pdf_path)
        if key == self.preview_key:
            self.show_preview(pages, total, pdf_path)

    def refresh_all_rows(self):
        for item in self.files:
            self.refresh_file_row(item)
        self.show_current_file()
        self.update_summary()

    def update_summary(self):
        pdf_count = sum(len(item.selected_sheets()) for item in self.files)
        self.summary.configure(text=f"{len(self.files)} file  ·  {pdf_count} PDF")

    def update_controls(self):
        busy = self.converting
        state = "disabled" if busy else "normal"
        self.tab_selector.configure(state="disabled" if busy or self.reading else "normal")
        self.add_zip_button.configure(state=state)
        for widget in (self.add_files_button, self.add_folder_button, self.clear_button):
            widget.configure(state=state)
        for row_widgets in self.file_rows.values():
            row_widgets[2].configure(state=state)
        can_convert = not busy and not self.reading and bool(self.files)
        self.convert_button.configure(state="normal" if can_convert else "disabled")
        if busy or self.reading:
            self.convert_button.grid_forget()
            self.cancel_button.grid(**self.action_grid)
        else:
            self.cancel_button.grid_forget()
            if self.current_page == di_ui.PAGE_PDF:
                self.convert_button.grid(**self.action_grid)
            else:
                self.convert_button.grid_forget()

    # -------------------------------------------------------------- convert
    def build_plan(self):
        plan = []
        for item in self.files:
            if self.convert_only and item.path not in self.convert_only:
                continue
            jobs = []
            for sheet in item.selected_sheets():
                pdf_name = clean_filename(sheet.pdf_name)
                if not pdf_name:
                    raise ValueError(f"Nama PDF kosong:\n{item.name}, sheet {sheet.name}")
                pdf_path = os.path.join(os.path.dirname(item.path), pdf_name + ".pdf")
                jobs.append((sheet.name, pdf_path))
            if jobs:
                plan.append((item.path, jobs))
        return plan

    def changed_files(self):
        changed = []
        for item in self.files:
            try:
                if item.state == "ready" and os.path.getmtime(item.path) != item.mtime:
                    changed.append(item)
            except OSError:
                pass
        return changed

    def start_convert(self):
        self.focus_set()
        changed = self.changed_files()
        if changed:
            # File sudah diubah sejak dibaca, jadi cek ulang dulu sebelum convert.
            self.convert_after_read = True
            for item in changed:
                self.read_queue.put(item.path)
            self.status.configure(text=f"Mengecek ulang {len(changed)} file yang berubah...")
            self.start_reader()
            self.update_controls()
            return
        try:
            plan = self.build_plan()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return
        if not plan:
            messagebox.showwarning(APP_TITLE, "Centang minimal satu sheet.")
            return
        pdf_paths = [pdf for _, jobs in plan for _, pdf in jobs]
        duplicates = find_duplicates(pdf_paths)
        if duplicates:
            messagebox.showwarning(
                APP_TITLE,
                "Beberapa PDF di folder yang sama akan memakai nama yang sama. "
                "Ubah namanya dulu, atau pakai {kode} di aturan:\n\n"
                + "\n".join(os.path.basename(path) for path in duplicates[:10]),
            )
            return
        scope = [item for item in self.files if not self.convert_only or item.path in self.convert_only]
        skipped = [item.name for item in scope if item.state == "ready" and not item.selected_sheets()]
        failed_open = [item.name for item in scope if item.state == "error"]
        existing = [path for path in pdf_paths if os.path.exists(path)]
        notes = [f"Akan membuat {len(pdf_paths)} PDF dari {len(plan)} file."]
        if skipped:
            notes.append(f"{len(skipped)} file dilewati karena tidak ada sheet yang dipilih.")
        if failed_open:
            notes.append(f"{len(failed_open)} file dilewati karena tidak bisa dibuka.")
        if existing:
            notes.append(f"{len(existing)} PDF yang sudah ada akan ditimpa.")
        incomplete = []
        for item in scope:
            for sheet in item.selected_sheets():
                if sheet.name in item.issues:
                    incomplete.append(f"{item.name} ({sheet.name})")
            if item.notes:
                incomplete.append(f"{item.name} (nama SI)")
        if incomplete:
            notes.append(f"\n⚠ {len(incomplete)} sheet masih perlu dicek:")
            notes.extend(incomplete[:10])
            if len(incomplete) > 10:
                notes.append(f"dan {len(incomplete) - 10} lainnya")
        if (skipped or failed_open or existing or incomplete) and not messagebox.askyesno(
            APP_TITLE, "\n".join(notes) + "\n\nLanjut?"
        ):
            return

        self.convert_only = None
        self.result_folder = os.path.dirname(plan[0][0])
        self.cancel.clear()
        self.converting = True
        self.progress.set(0)
        self.update_controls()
        threading.Thread(
            target=convert_worker,
            args=(plan, self.one_page.get(), self.cancel, self.post_event),
            daemon=True,
        ).start()

    def request_cancel(self):
        self.cancel.set()
        self.status.configure(text="Membatalkan setelah proses yang berjalan selesai...")

    def finish_convert(self, result):
        self.converting = False
        self.update_controls()
        succeeded, failed = result["succeeded"], result["failed"]
        seconds = int(result["seconds"])
        headline = "Dibatalkan" if result["cancelled"] else "Selesai"
        self.status.configure(
            text=f"{headline}: {succeeded} PDF berhasil, {len(failed)} gagal, {seconds} detik."
        )
        message = f"{headline}: {succeeded} PDF berhasil, {len(failed)} gagal."
        if failed:
            report_path = os.path.join(self.result_folder, "laporan_gagal.txt")
            try:
                with open(report_path, "w", encoding="utf-8") as report:
                    for name, sheet, error in failed:
                        report.write(f"{name} | {sheet} | {error}\n")
                message += f"\n\nDetailnya ada di:\n{report_path}"
            except OSError as exc:
                message += f"\n\nLaporan gagal disimpan: {exc}"
        folder = self.result_folder
        open_folder = (lambda: os.startfile(folder)) if os.name == "nt" else None
        if succeeded and not failed:
            self.toast.show(f"{succeeded} PDF selesai dibuat" if not result["cancelled"] else f"Dibatalkan · {succeeded} PDF sudah dibuat",
                            folder, "ok" if not result["cancelled"] else "warn", action=open_folder, action_text="Buka folder" if open_folder else "")
        elif succeeded:
            self.toast.show(f"{succeeded} PDF selesai, {len(failed)} gagal", "Detail ada di laporan_gagal.txt", "warn",
                            action=open_folder, action_text="Buka folder" if open_folder else "")
            messagebox.showwarning(APP_TITLE, message)
        else:
            messagebox.showwarning(APP_TITLE, message)

    # --------------------------------------------------------------- events
    def poll_events(self):
        try:
            while True:
                kind, data = self.events.get_nowait()
                self.handle_event(kind, data)
        except queue.Empty:
            pass
        self.after(100, self.poll_events)

    def collect_garbage(self):
        # Objek Tk hanya boleh dibersihkan dari thread utama. Kalau GC berjalan
        # di thread Excel, aplikasi bisa crash.
        gc.collect()
        self.after(5000, self.collect_garbage)

    def file_by_path(self, path):
        return next((item for item in self.files if item.path == path), None)

    def handle_di_job(self, finish, callback, result, error):
        finish(callback, result, error)

    def handle_event(self, kind, data):
        if kind in ("file_read", "file_error"):
            item = self.file_by_path(data[0])
            if item is None:
                return
            if kind == "file_read":
                item.set_sheets(data[1])
                item.issues = data[2]
                item.unchecked = data[3]
                item.mtime = data[4]
                item.update_notes(data[5])
                item.state = "ready"
            else:
                item.state = "error"
                item.error = data[1]
            self.refresh_file_row(item)
            self.update_summary()
            if item is self.current:
                self.show_current_file()
                self.request_preview()
            loaded = sum(1 for entry in self.files if entry.state != "loading")
            self.status.configure(text=f"Membaca sheet: {loaded}/{len(self.files)} file")
            self.progress.set(loaded / len(self.files))
        elif kind == "read_finished":
            self.finish_reading()
        elif kind == "status":
            self.status.configure(text=data)
        elif kind == "progress":
            self.progress.set(data[0] / data[1])
        elif kind == "convert_done":
            self.finish_convert(data)
        elif kind == "di_job":
            self.handle_di_job(*data)
        elif kind == "preview_ready":
            self.handle_preview_ready(*data)
        elif kind == "preview_error":
            if data[0] == self.preview_key:
                self.set_preview_message(f"Preview gagal dibuat.\n{data[1]}", DANGER)
        elif kind == "preview_stopped":
            self.preview_thread = None
            if not self.preview_requests.empty():
                self.start_preview_worker()

    def finish_reading(self):
        self.reading = False
        if self.cancel.is_set():
            while not self.read_queue.empty():
                self.read_queue.get_nowait()
            for item in [entry for entry in self.files if entry.state == "loading"]:
                self.remove_file(item)
            self.status.configure(text="Pembacaan dibatalkan.")
            self.convert_after_read = False
        elif not self.read_queue.empty():
            self.start_reader()
            return
        else:
            problems = sum(1 for item in self.files if item.state == "error")
            note = f" {problems} tidak bisa dibuka." if problems else ""
            self.status.configure(text=f"{len(self.files)} file siap.{note}")
        self.update_controls()
        if self.convert_after_read:
            self.convert_after_read = False
            self.start_convert()

    def on_close(self):
        if self.converting or self.reading or self.data_manager.running:
            messagebox.showinfo(APP_TITLE, "Proses masih berjalan. Tunggu atau klik Batal dulu.")
            return
        unsaved = self.data_manager.dirty()
        if unsaved and not messagebox.askyesno(APP_TITLE, f"{unsaved} perubahan data belum disimpan dan akan hilang. Tutup aplikasi?"):
            return
        if self.preview_thread is not None:
            self.preview_requests.put(None)
            self.preview_thread.join(timeout=5)
        shutil.rmtree(self.temp_dir, ignore_errors=True)
        self.destroy()


if __name__ == "__main__":
    gc.disable()
    if os.name == "nt":
        # Agar taskbar memakai ikon Dockit, bukan ikon Python.
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    App().mainloop()
