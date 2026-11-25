"""Data DI PDF Gen: baca dan simpan file sumber, pengecekan, dan pembuatan Excel DI."""
try:
    from xlsx_rows import add_item_rows
except ImportError:
    from .xlsx_rows import add_item_rows
import datetime
import json
import math
import os
import re
import shutil
import time
import zipfile

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
Q = "{%s}" % NS
MASTER_SHEET = "MASTER DATA DI-柜"
ITEM_SHEET = "Sheet1"
WEIGHT_SHEET = "MODEL -1 (2)"
PRICE_SHEET = "HARGA"
LOADING_SHEET = "DI-柜"
LOADING_COLUMNS = {"di": "B", "sku": "C", "qty": "I", "planned": "J", "po": "N", "asin": "O", "cbm": "AF"}
CBM_TOLERANCE = 1e-6
MAX_ITEMS = 19
MASTER_LINK_PART = "xl/externalLinks/externalLink1.xml"
MASTER_FORMULA_CELLS = (
    ("G3", "E"), ("L3", "I"), ("P5", "G"), ("P6", "H"), ("P7", "J"), ("A32", "O"), ("E32", "L"), ("E34", "M"), ("C37", "N"),
)
PO_FORMULA_CELLS = (
    ("G3", "B"), ("L3", "C"), ("P5", "H"), ("P6", "I"), ("A10", "Y"), ("A16", "Z"), ("E34", "E"), ("C37", "F"),
)
MASTER_PO_SHEET = "MASTER DATA PO-柜"
LOADING_PO_SHEET = "PO-柜"
LOADING_PO_COLUMNS = {"di": "B", "sku": "C", "qty": "I", "planned": "J", "cbm": "AE"}
PO_LAST_COLUMN = 32
FIRST_YEAR = "2026"
UNIQUE_MASTER_COLUMNS = [("D", "Rolling"), ("G", "Kontainer"), ("H", "Seal")]
FIRST_ITEM_ROW = 3
EXCEL_ERRORS = {
    -2146826288: "#NULL!", -2146826281: "#DIV/0!", -2146826273: "#VALUE!", -2146826265: "#REF!",
    -2146826259: "#NAME?", -2146826252: "#NUM!", -2146826246: "#N/A",
}
NO_PASSWORD = "__tanpa_password__"
XL_MANUAL, XL_AUTOMATIC = -4135, -4105


class StaleFileError(Exception):
    pass


class Column:
    def __init__(self, key, label, kind="text", editable=True, width=110, new_only=False):
        self.key, self.label, self.kind = key, label, kind
        self.editable, self.width, self.new_only = editable, width, new_only


MASTER_COLUMNS = [
    Column("A", "DI", new_only=True, width=95), Column("B", "Amazon No.", width=135), Column("C", "PO", width=90),
    Column("D", "Rolling", width=85), Column("E", "Invoice", editable=False, width=105),
    Column("F", "Kode nama file", editable=False, width=145), Column("G", "Kontainer", width=105),
    Column("H", "Seal", width=105), Column("I", "Tgl invoice", "date", width=90),
    Column("J", "Size", width=60), Column("K", "ETD Batam", "date", width=90), Column("L", "ETD MV", "date", width=90),
    Column("M", "POD", width=100), Column("N", "BL", width=115), Column("O", "Kapal", width=170),
    Column("P", "Notes", width=85), Column("Q", "Qty", "num", width=60), Column("S", "Net Weight", "num", width=85),
    Column("T", "Gross Weight", "num", width=85), Column("U", "CTNS", "num", width=60), Column("V", "CBM", "num", width=65),
    Column("W", "Total harga satuan", "num", width=95), Column("X", "Invoice Amount", "num", width=95),
    Column("Z", "Kurs", "num", width=70), Column("AA", "AJU", width=70), Column("AD", "PEB No.", width=80),
    Column("AE", "PPJK", width=150), Column("AF", "Shipper", width=100),
]
GOODS_COLUMNS = [
    Column("A", "SKU", new_only=True, width=210), Column("B", "Goods Description", width=480), Column("C", "HS Code", width=110),
]
WEIGHT_COLUMNS = [
    Column("A", "SKU", new_only=True, width=210), Column("B", "Goods Description", editable=False, width=400),
    Column("D", "CBM", "num", width=90), Column("LC", "CBM Loading", "num", editable=False, width=100), Column("E", "Net Weight", "num", width=90), Column("F", "Gross Weight", "num", width=90),
]
PRICE_COLUMNS = [
    Column("G", "DI / PO", new_only=True, width=120), Column("H", "SKU", new_only=True, width=210), Column("D", "Qty", "num", width=70),
    Column("E", "Harga PEB", "num", width=95), Column("F", "Jasa", "num", width=80), Column("B", "Harga total", "num", editable=False, width=95),
]


MASTER_PO_COLUMNS = [
    Column("A", "PO", new_only=True, width=105), Column("B", "SI Number", width=130), Column("C", "SI Date", "date", width=90),
    Column("D", "ETD Batam", "date", width=90), Column("E", "POD", width=105), Column("F", "BL", width=125),
    Column("G", "Forwarder", width=90), Column("H", "Kontainer", width=110), Column("I", "Seal", width=105),
    Column("J", "Qty", "num", width=60), Column("L", "Net Weight", "num", width=85), Column("M", "Gross Weight", "num", width=85),
    Column("N", "CTNS", "num", width=60), Column("O", "CBM", "num", width=65), Column("P", "Total harga satuan", "num", width=95),
    Column("Q", "Invoice Amount", "num", width=95), Column("R", "AJU PEB", width=75), Column("T", "PEB No.", width=75),
    Column("V", "PPJK", width=150), Column("W", "Kurs", "num", width=70), Column("Y", "Consignee", width=190),
    Column("Z", "Notify", width=190), Column("AB", "Type", width=70), Column("AF", "Shipper", width=95),
]


class Profile:
    def __init__(self, **values):
        self.__dict__.update(values)


class Row:
    __slots__ = ("row", "values", "dirty")

    def __init__(self, row, values):
        self.row, self.values, self.dirty = row, values, set()

    def get(self, key):
        return self.values.get(key)


def _profiles():
    return {
        "DI": Profile(
            name="DI", prefix="DI-柜", table="master", sheet=MASTER_SHEET, first_row=3, max_column=None, template="di_template.xlsx",
            formula_cells=MASTER_FORMULA_CELLS, max_items=99, code_key="F", qty_key="Q", bl_key="N", pod_key="M", year_key="D",
            unique=[("D", "Rolling"), ("G", "Kontainer"), ("H", "Seal")],
            required={"C": "PO", "G": "Kontainer", "H": "Seal", "I": "Tgl", "J": "Size", "L": "ETD", "M": "POD", "N": "BL", "O": "Kapal"},
            columns=MASTER_COLUMNS, line_required=(("qty", "Qty"), ("asin", "ASIN"), ("po", "PO"), ("nw", "NW"), ("gw", "GW"), ("cbm", "CBM"), ("hs", "HS Code")),
        ),
        "PO": Profile(
            name="PO", prefix="PO-柜", table="po_master", sheet=MASTER_PO_SHEET, first_row=3, max_column=PO_LAST_COLUMN, template="po_template.xlsx",
            formula_cells=PO_FORMULA_CELLS, max_items=99, code_key="B", qty_key="J", bl_key="F", pod_key="E", year_key="B",
            unique=[("B", "SI Number"), ("H", "Kontainer"), ("I", "Seal")],
            required={"B": "SI Number", "C": "Tgl", "E": "POD", "F": "BL", "H": "Kontainer", "I": "Seal"},
            columns=MASTER_PO_COLUMNS, line_required=(("qty", "Qty"), ("nw", "NW"), ("gw", "GW"), ("cbm", "CBM"), ("hs", "HS Code")),
        ),
    }


def column_number(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def column_letters(number):
    out = ""
    while number:
        number, rem = divmod(number - 1, 26)
        out = chr(65 + rem) + out
    return out


def clean_cell(value):
    if isinstance(value, (int, float)) and value in EXCEL_ERRORS:
        return EXCEL_ERRORS[value]
    if isinstance(value, datetime.datetime):
        return to_serial(value)
    return value


def to_serial(value):
    return (value.replace(tzinfo=None) - datetime.datetime(1899, 12, 30)).total_seconds() / 86400


def clean_hs(value):
    """HS Code angka tidak boleh tampil dengan .0 di belakang."""
    if is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, str) and re.fullmatch(r"\s*\d+\.0\s*", value):
        return int(float(value))
    return value


def format_value(value, kind="text"):
    if value is None:
        return ""
    if kind == "date" and isinstance(value, (int, float)):
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=value)).strftime("%Y-%m-%d")
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(round(value, 6))
    return str(value)


def parse_value(text, kind="text"):
    text = text.strip()
    if text == "":
        return None
    if kind == "num":
        try:
            return float(text.replace(",", ""))
        except ValueError:
            raise ValueError(f"'{text}' bukan angka.")
    if kind == "date":
        for pattern in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
            try:
                return to_serial(datetime.datetime.strptime(text, pattern))
            except ValueError:
                pass
        raise ValueError(f"'{text}' bukan tanggal. Pakai format 2026-11-05.")
    return text


def is_blank(value):
    return value in (None, "", 0) or (isinstance(value, str) and (not value.strip() or value.startswith("#")))


def check_description(sku, name):
    """None berarti cocok. Selain itu berisi keterangan typo inch atau type."""
    sku = str(sku or "").strip()
    parts = sku.split("-")
    if len(parts) < 3 or sku.upper().endswith("-ID"):
        return None
    if parts[-2].isdigit():
        inch, kind = parts[-2], parts[-1]
    elif len(parts) > 3 and parts[-3].isdigit():
        inch, kind = parts[-3], parts[-2]
    else:
        return None
    text = str(name or "").upper().replace("\xa0", " ").replace("INCHES", "INCH")
    for mark in "()（）":
        text = text.replace(mark, " ")
    text = " ".join(text.split()).replace(" INCH", "INCH")
    token = f"{inch}INCH"
    if f" {token} " not in f" {text} ":
        return f"Inch tidak cocok, SKU {inch} inch"
    if f" {kind.upper()} " not in f" {text[text.find(token) + len(token):]} ":
        return f"Type tidak cocok, SKU {kind}"
    return None


def sku_size(sku):
    """Inch dan type dari SKU, misal A-2503M02822-10-T menjadi ("10", "T")."""
    parts = str(sku or "").strip().split("-")
    if len(parts) < 3 or str(sku).strip().upper().endswith("-ID"):
        return None
    if parts[-2].isdigit():
        return parts[-2], parts[-1]
    if len(parts) > 3 and parts[-3].isdigit():
        return parts[-3], parts[-2]
    return None


INCH_WORD = re.compile(r"(\d+(?:\.\d+)?)(\s*)(inch(?:es)?)", re.I)


def fix_description(sku, name):
    """Perbaiki inch dan type di deskripsi supaya sama dengan SKU. None kalau tidak perlu atau tidak bisa."""
    if not check_description(sku, name):
        return None
    size = sku_size(sku)
    text = str(name or "")
    match = INCH_WORD.search(text)
    if not size or not match:
        return None
    inch, kind = size
    fixed = text[:match.start(1)] + inch + text[match.end(1):]
    match = INCH_WORD.search(fixed)
    tail = fixed[match.end():]
    # Hanya kode ukuran huruf besar (T, TXL, Q, K, ...) yang diganti; kata biasa seperti Twin dibiarkan.
    word = re.match(r"(\s*[()（）]?\s*)([A-Z]{1,4})(?![A-Za-z])", tail)
    if word and word.group(2) != kind.upper():
        start = match.end() + word.start(2)
        fixed = fixed[:start] + kind + fixed[start + len(word.group(2)):]
    elif not word and not re.match(r"\s*[A-Za-z]", tail):
        fixed = fixed[:match.end()] + f" {kind}" + tail
    return fixed if check_description(sku, fixed) is None and fixed != text else None


def read_weight_file(path):
    """File hasil timbang (Sheet1: SKU, 3 baris NW/GW per SKU). Per SKU dipakai kelompok timbang terbaru,
    lalu baris dengan NW paling kecil beserta GW di baris itu."""
    import openpyxl
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        best = {}
        for sheet in workbook.worksheets:
            rows = sheet.iter_rows(values_only=True)
            header, columns = 0, None
            for header, row in enumerate(rows, start=1):
                texts = [str(c or "").upper() for c in row]
                sku = next((i for i, t in enumerate(texts) if t.strip() == "SKU"), None)
                nw = next((i for i, t in enumerate(texts) if "NET WEIGHT" in t or "净重" in t), None)
                gw = next((i for i, t in enumerate(texts) if "GROSS WEIGHT" in t or "毛重" in t), None)
                date = next((i for i, t in enumerate(texts) if "DATE" in t or "日期" in t), None)
                if None not in (sku, nw, gw):
                    columns = (sku, nw, gw, date)
                    break
                if header >= 10:
                    break
            if not columns:
                continue
            sku_col, nw_col, gw_col, date_col = columns
            current = None
            group = 0
            for number, row in enumerate(rows, start=header + 1):
                cell = row[sku_col] if sku_col < len(row) else None
                if cell not in (None, "") and str(cell).strip():
                    group += 1
                    stamp = row[date_col] if date_col is not None and date_col < len(row) else None
                    current = (str(cell).strip().upper(), group, stamp)
                if current is None:
                    continue
                nw, gw = to_weight(row[nw_col] if nw_col < len(row) else None), to_weight(row[gw_col] if gw_col < len(row) else None)
                if nw is None or gw is None:
                    continue
                sku, group_id, stamp = current
                order = (stamp.timestamp() if hasattr(stamp, "timestamp") else 0, sheet.title == "Sheet1", group_id)
                entry = best.get(sku)
                if entry is None or order > entry["order"]:
                    day = stamp.date() if hasattr(stamp, "date") and callable(stamp.date) else None
                    best[sku] = entry = {"order": order, "nw": nw, "gw": gw, "row": number, "sheet": sheet.title, "count": 0, "date": day}
                if order == entry["order"]:
                    entry["count"] += 1
                    if (nw, gw) < (entry["nw"], entry["gw"]):
                        entry.update(nw=nw, gw=gw, row=number)
            break  # hanya sheet pertama yang punya kolom SKU, NW, dan GW
        return best
    finally:
        workbook.close()


def to_weight(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value) if value > 0 else None
    try:
        number = float(str(value).replace(",", ".").strip())
    except (TypeError, ValueError):
        return None
    return number if number > 0 else None


# Ukuran kasur per type (cm), untuk memperkirakan berat dari ukuran dan tebal.
MATTRESS_SIZES = {"T": (99, 190), "TXL": (99, 203), "F": (137, 190), "FXL": (137, 203), "Q": (152, 203), "K": (193, 203),
                  "CK": (183, 213), "EK": (193, 203), "CAK": (183, 213), "S": (90, 190), "D": (137, 190)}


def parse_mattress(sku):
    """(inch, type, model) dari SKU, misal A-2604M01852-12-K menjadi (12, "K", "A-2604M01852")."""
    parts = str(sku or "").strip().upper().split("-")
    for i in range(len(parts) - 1, 0, -1):
        if parts[i] in MATTRESS_SIZES and parts[i - 1].replace(".", "", 1).isdigit():
            return float(parts[i - 1]), parts[i], "-".join(parts[:i - 1])
    return None


def mattress_kind(text):
    text = str(text or "").upper()
    if "HYBRID" in text or "混合" in text:
        return "hybrid"
    if "SPRING" in text or "弹簧" in text:
        return "spring"
    if "FOAM" in text or "海绵" in text:
        return "foam"
    return "lain"


def solve_least_squares(rows, targets):
    """Regresi linear kecil tanpa library tambahan."""
    size = len(rows[0])
    a = [[sum(r[i] * r[j] for r in rows) for j in range(size)] for i in range(size)]
    b = [sum(r[i] * t for r, t in zip(rows, targets)) for i in range(size)]
    for i in range(size):
        pivot = max(range(i, size), key=lambda k: abs(a[k][i]))
        a[i], a[pivot], b[i], b[pivot] = a[pivot], a[i], b[pivot], b[i]
        if abs(a[i][i]) < 1e-12:
            return None
        for k in range(i + 1, size):
            factor = a[k][i] / a[i][i]
            a[k] = [x - factor * y for x, y in zip(a[k], a[i])]
            b[k] -= factor * b[i]
    result = [0.0] * size
    for i in reversed(range(size)):
        result[i] = (b[i] - sum(a[i][j] * result[j] for j in range(i + 1, size))) / a[i][i]
    return result


_estimate_cache = None


def estimate_marks():
    """Tanda NW/GW perkiraan per SKU. Disimpan di pengaturan, dibaca sekali lalu disimpan di memori."""
    global _estimate_cache
    if _estimate_cache is None:
        _estimate_cache = dict(load_settings().get("weight_estimates", {}))
    return _estimate_cache


def save_estimate_marks(marks):
    global _estimate_cache
    _estimate_cache = dict(marks)
    settings = load_settings()
    settings["weight_estimates"] = _estimate_cache
    save_settings(settings)


def clear_estimates(skus):
    marks = dict(estimate_marks())
    changed = False
    for sku in skus:
        changed |= marks.pop(str(sku).strip(), None) is not None
    if changed:
        save_estimate_marks(marks)


def scan_sheet_typos(values, first_row=1):
    """Cari baris berisi SKU dan deskripsi yang tidak cocok pada INV, PL, atau FINAL SI."""
    problems = []
    if not isinstance(values, (tuple, list)):
        return problems
    sku_pattern = re.compile(r"^[A-Z0-9]+(-[A-Z0-9]+){2,}$")
    for index, row in enumerate(values):
        if not isinstance(row, (tuple, list)):
            continue
        sku = next((c.strip() for c in row if isinstance(c, str) and sku_pattern.match(c.strip())), None)
        name = next((c for c in row if isinstance(c, str) and "INCH" in c.upper()), None)
        if sku and name:
            result = check_description(sku, name)
            if result:
                problems.append({"row": first_row + index, "item": sku, "missing": [f"Typo deskripsi: {' '.join(name.split())} ({result})"]})
    return problems


def workbook_path(workbook):
    try:
        return workbook.FullName
    except Exception:
        return "ini"


def read_sheet(workbook, name, columns, first_row, max_column=None):
    try:
        sheet = workbook.Sheets.Item(name)
    except Exception:
        raise ValueError(f"Sheet '{name}' tidak ada di file {os.path.basename(str(workbook_path(workbook)))}.")
    used = sheet.UsedRange
    if max_column:
        # Master PO punya ribuan kolom kosong, jadi hanya kolom data yang dibaca.
        last = used.Row + used.Rows.Count - 1
        values = sheet.Range(sheet.Cells(1, 1), sheet.Cells(last, max_column)).Value2
        origin_row, origin_col = 1, 1
    else:
        values = used.Value2
        origin_row, origin_col = used.Row, used.Column
    rows = []
    indexes = [(key, column_number(key) - origin_col) for key in columns]
    last_row = origin_row - 1
    if isinstance(values, (tuple, list)):
        last_row = origin_row + len(values) - 1
        for offset, line in enumerate(values):
            number = origin_row + offset
            if number < first_row:
                continue
            data = {key: clean_cell(line[i]) if 0 <= i < len(line) else None for key, i in indexes}
            rows.append((number, data))
    return rows, last_row


def open_source(app, path, password, write=False):
    secret = password or NO_PASSWORD
    return app.Workbooks.Open(os.path.abspath(path), 0, not write, 5, secret, password or "")


def detect_source(path):
    """Kenali jenis file dari nama sheet. None berarti file terkunci, "" berarti bukan file sumber."""
    try:
        with zipfile.ZipFile(path) as archive:
            book = archive.read("xl/workbook.xml").decode("utf-8", "ignore")
    except (zipfile.BadZipFile, KeyError, OSError):
        return None
    names = set(re.findall(r'<sheet\b[^>]*?\bname="([^"]*)"', book))
    if MASTER_SHEET in names:
        return "master"
    if MASTER_PO_SHEET in names:
        return "po_master"
    if LOADING_SHEET in names:
        return "loading"
    if {ITEM_SHEET, WEIGHT_SHEET, PRICE_SHEET} <= names:
        return "items"
    return ""


def read_loading_rows(path, app=None):
    """Baca sheet DI-柜dan PO-柜 apa adanya. Setiap baris tetap terpisah dan urutannya tidak berubah."""
    layouts = ((LOADING_SHEET, LOADING_COLUMNS), (LOADING_PO_SHEET, LOADING_PO_COLUMNS))
    collected = []
    try:
        from openpyxl import load_workbook

        book = load_workbook(path, data_only=True, read_only=True)
        try:
            if LOADING_SHEET not in book.sheetnames:
                raise ValueError(f"Sheet '{LOADING_SHEET}' tidak ada di file {os.path.basename(path)}.")
            for name, columns in layouts:
                if name in book.sheetnames:
                    collected.append((columns, list(book[name].iter_rows(min_row=2, values_only=True))))
        finally:
            book.close()
    except (zipfile.BadZipFile, ImportError, KeyError, OSError):
        workbook = open_source(app, path, "")
        try:
            for name, columns in layouts:
                collected.append((columns, list(workbook.Sheets.Item(name).UsedRange.Value2)[1:]))
        finally:
            workbook.Close(False)
    rows = []
    for columns, lines in collected:
        numbers = {key: column_number(letter) - 1 for key, letter in columns.items()}
        for line in lines:
            item = {key: clean_cell(line[i]) if i < len(line) else None for key, i in numbers.items()}
            if item["di"] and item["sku"]:
                item["di"] = str(item["di"]).strip()
                item["sku"] = str(item["sku"]).strip()
                rows.append(item)
    return rows


class Store:
    def __init__(self, master_path, item_path, master_password="", item_password="", load_path="", po_path="", po_password=""):
        self.master_path, self.item_path = master_path, item_path
        self.master_password, self.item_password = master_password, item_password
        self.load_path = load_path
        self.po_path, self.po_password = po_path, po_password
        self.po_master = []
        self.po_last = 2
        self.po_mtime = 0
        self.loading = {}
        self.loading_cbm = {}
        self.master, self.items, self.prices = [], [], []
        self.master_last = self.item_last = self.price_last = self.weight_last = 2
        self.weight_rows = {}
        self.master_mtime = self.item_mtime = 0
        self.loaded = False
        self.reindex()

    def load(self, app):
        master_keys = [c.key for c in MASTER_COLUMNS] + ["Y"]
        workbook = open_source(app, self.master_path, self.master_password)
        try:
            self.master_mtime = os.path.getmtime(self.master_path)
            rows, self.master_last = read_sheet(workbook, MASTER_SHEET, master_keys, 3)
            # Hanya menyaring tampilan. Baris lama di file sumber tidak diubah.
            all_master = [Row(n, d) for n, d in rows if d.get("A")]
            self.master = [r for r in all_master if not is_old(r)]
            if not all_master:
                raise ValueError("File master tidak berisi data DI.")
            old_di = {str(r.get("A")).strip() for r in all_master if is_old(r)}
        finally:
            workbook.Close(False)
        if self.po_path:
            profile = PROFILES["PO"]
            keys = [c.key for c in profile.columns]
            workbook = open_source(app, self.po_path, self.po_password)
            try:
                self.po_mtime = os.path.getmtime(self.po_path)
                rows, self.po_last = read_sheet(workbook, profile.sheet, keys, profile.first_row, profile.max_column)
                all_po = [Row(n, d) for n, d in rows if d.get("A")]
                if not all_po:
                    raise ValueError("File Master PO tidak berisi data PO.")
                self.po_master = [r for r in all_po if not is_old_po(r)]
                old_di |= {str(r.get("A")).strip() for r in all_po if is_old_po(r)}
            finally:
                workbook.Close(False)
        workbook = open_source(app, self.item_path, self.item_password)
        try:
            self.item_mtime = os.path.getmtime(self.item_path)
            rows, self.item_last = read_sheet(workbook, ITEM_SHEET, list("ABCDEFG"), 3)
            self.items = [Row(n, d) for n, d in rows if d.get("A")]
            rows, self.weight_last = read_sheet(workbook, WEIGHT_SHEET, list("CFG"), 2)
            self.weight_rows = {}
            for number, data in rows:
                if data.get("C"):
                    self.weight_rows.setdefault(str(data["C"]).strip(), []).append(number)
            rows, self.price_last = read_sheet(workbook, PRICE_SHEET, list("BDEFGH"), 2)
            self.prices = [
                Row(n, d) for n, d in rows
                if (d.get("G") or d.get("H")) and str(d.get("G") or "").strip() not in old_di
            ]
        finally:
            workbook.Close(False)
        if self.load_path:
            self.load_loading(app)
        self.apply_loading_cbm()
        self.reindex()
        self.loaded = True

    def load_loading(self, app):
        self.loading = {}
        self.loading_cbm = {}
        for row in read_loading_rows(self.load_path, app):
            self.loading.setdefault(row["di"], []).append(row)
            if isinstance(row["cbm"], (int, float)) and row["cbm"]:
                self.loading_cbm[row["sku"]] = row["cbm"]

    def apply_loading_cbm(self):
        for row in self.items:
            row.values["LC"] = self.loading_cbm.get(str(row.get("A")).strip())

    def lines_for(self, di_key):
        """Baris per DI: dari file loading kalau ada, kalau tidak dari tab Harga."""
        if di_key in self.loading:
            return [
                {"sku": r["sku"], "qty": r["qty"], "po": text_or_blank(r.get("po")), "asin": text_or_blank(r.get("asin")),
                 "cbm": number_or_none(r["cbm"]), "planned": r["planned"]}
                for r in self.loading[di_key]
            ]
        return self.lines_from_prices(di_key)

    def line_source(self, di_key):
        return "file loading" if di_key in self.loading else "tab Harga"

    def auto_fixes(self, skus=None, kinds=("cbm", "typo")):
        """Perubahan yang bisa dilakukan otomatis: CBM disamakan dengan file loading, deskripsi disamakan dengan SKU.
        Mengembalikan (perubahan, SKU typo yang tidak bisa diperbaiki otomatis)."""
        fixes, manual = [], []
        seen = set()
        for row in self.items:
            sku = str(row.get("A") or "").strip()
            if not sku or (skus is not None and sku not in skus) or sku in seen:
                continue
            seen.add(sku)
            if "cbm" in kinds and cbm_mismatch(row.get("D"), row.get("LC")):
                fixes.append({"row": row, "key": "D", "sku": sku, "old": row.get("D"), "new": row.get("LC"), "label": "CBM"})
            if "typo" in kinds and not is_blank(row.get("B")) and check_description(sku, row.get("B")):
                fixed = fix_description(sku, row.get("B"))
                if fixed:
                    fixes.append({"row": row, "key": "B", "sku": sku, "old": row.get("B"), "new": fixed, "label": "Deskripsi"})
                else:
                    manual.append(sku)
        return fixes, manual

    def estimate_weights(self):
        """Perkiraan NW/GW untuk SKU yang belum punya berat.
        Tinggi: dari SKU satu model dengan ukuran lain (rata-rata 3 terdekat, disesuaikan volume kasur).
        Sedang: dari regresi per kategori (foam, spring, hybrid) memakai volume dan luas kasur.
        Hasil uji pada data: rata-rata meleset 3.5% (tinggi) dan 6.5% (sedang)."""
        known, todo, skipped, seen = [], [], [], set()
        for row in self.items:
            sku = str(row.get("A") or "").strip()
            if not sku or sku in seen:
                continue
            seen.add(sku)
            nw, gw = row.get("E"), row.get("F")
            size = parse_mattress(sku)
            has = isinstance(nw, (int, float)) and nw > 0 and isinstance(gw, (int, float)) and gw > nw
            if has and size:
                inch, kind, model = size
                width, length = MATTRESS_SIZES[kind]
                known.append({"sku": sku, "model": model, "inch": inch, "area": width * length / 1e4, "vol": width * length * inch * 2.54 / 1e6,
                              "kind": mattress_kind(row.get("B")), "nw": nw, "gw": gw})
            elif not has:
                (todo if size else skipped).append(row)
        models = {}
        for item in known:
            models.setdefault(item["model"], []).append(item)
        fits = {}
        for kind in {k["kind"] for k in known} | {"semua"}:
            pool = [k for k in known if k["kind"] == kind or kind == "semua"]
            if len(pool) >= 20:
                nw = solve_least_squares([[k["vol"], k["area"], 1] for k in pool], [k["nw"] for k in pool])
                pack = solve_least_squares([[k["area"], 1] for k in pool], [k["gw"] - k["nw"] for k in pool])
                if nw and pack:
                    fits[kind] = (nw, pack)
        fixes, marks = [], {}
        for row in todo:
            sku = str(row.get("A")).strip()
            inch, kind, model = parse_mattress(sku)
            width, length = MATTRESS_SIZES[kind]
            area, vol = width * length / 1e4, width * length * inch * 2.54 / 1e6
            siblings = models.get(model)
            if siblings:
                near = sorted(siblings, key=lambda s: (abs(math.log(s["vol"] / vol)), s["inch"] != inch))[:3]
                weights = [1 / (1e-3 + abs(math.log(s["vol"] / vol))) for s in near]
                nw = sum(w * s["nw"] * vol / s["vol"] for w, s in zip(weights, near)) / sum(weights)
                pack = sum(w * (s["gw"] - s["nw"]) * (area / s["area"]) ** 0.5 for w, s in zip(weights, near)) / sum(weights)
                level, source = "Tinggi", "dari " + ", ".join(s["sku"] for s in near)
            else:
                category = mattress_kind(row.get("B"))
                fit = fits.get(category) or fits.get("semua")
                if not fit:
                    skipped.append(row)
                    continue
                nw = fit[0][0] * vol + fit[0][1] * area + fit[0][2]
                pack = fit[1][0] * area + fit[1][1]
                level, source = "Sedang", f"dari rata-rata kategori {category if category in fits else 'semua'}"
            if nw <= 0:
                skipped.append(row)
                continue
            nw, gw = round(nw, 2), round(nw + max(pack, 0.3), 2)
            for key, value, label in (("E", nw, "NW"), ("F", gw, "GW")):
                fixes.append({"row": row, "key": key, "sku": sku, "old": row.get(key), "new": value, "label": label, "level": level, "source": source})
            marks[sku] = {"nw": nw, "gw": gw, "level": level, "source": source}
        return fixes, marks, [str(r.get("A")).strip() for r in skipped]

    def estimate_note(self, row):
        mark = estimate_marks().get(str(row.get("A") or "").strip())
        if not mark:
            return ""
        same = all(isinstance(row.get(k), (int, float)) and abs(row.get(k) - mark[v]) < 1e-6 for k, v in (("E", "nw"), ("F", "gw")))
        return f"NW/GW perkiraan ({mark['level'].lower()})" if same else ""

    def weight_fixes(self, weights, since=None):
        """NW dan GW dari file timbang yang berbeda dengan data. since: hanya SKU yang ditimbang sejak tanggal itu."""
        fixes, seen = [], set()
        for row in self.items:
            sku = str(row.get("A") or "").strip()
            entry = weights.get(sku.upper())
            if not entry or sku in seen:
                continue
            seen.add(sku)
            if since is not None and (entry["date"] is None or entry["date"] < since):
                continue
            for key, value, label in (("E", entry["nw"], "NW"), ("F", entry["gw"], "GW")):
                old = row.get(key)
                if not isinstance(old, (int, float)) or abs(old - value) > 1e-6:
                    fixes.append({"row": row, "key": key, "sku": sku, "old": old, "new": value, "label": label, "date": entry["date"]})
        return fixes

    def item_skus(self):
        return {str(row.get("A") or "").strip().upper() for row in self.items}

    def apply_fixes(self, fixes):
        for fix in fixes:
            # SKU yang muncul di beberapa baris ikut diubah semuanya.
            for row in self.item_groups.get(fix["sku"], [fix["row"]]):
                row.values[fix["key"]] = fix["new"]
                row.dirty.add(fix["key"])
        self.reindex()

    def cbm_tally(self):
        """SKU yang CBM-nya di data berbeda dengan file loading."""
        result = []
        for row in self.items:
            flag = cbm_mismatch(row.get("D"), row.get("LC"))
            if flag:
                result.append((row, flag))
        return result

    def duplicate_labels(self, row, profile=None):
        profile = profile or PROFILES["DI"]
        labels = []
        for key, label in profile.unique:
            value = unique_key(row.get(key))
            if value and self.counts[profile.name][key].get(value, 0) > 1:
                labels.append(f"{label} {format_value(row.get(key))}")
        return labels

    def rows_of(self, table):
        return {"master": self.master, "po_master": self.po_master, "items": self.items, "prices": self.prices}[table]

    def index_of(self, profile):
        return self.di_index if profile.name == "DI" else self.po_index

    def bl_of(self, profile):
        return self.bl_index if profile.name == "DI" else self.po_bl_index

    def reindex(self):
        self.counts = {}
        for profile, rows in ((PROFILES["DI"], self.master), (PROFILES["PO"], self.po_master)):
            counts = self.counts[profile.name] = {key: {} for key, _ in profile.unique}
            for row in rows:
                for key, _ in profile.unique:
                    value = unique_key(row.get(key))
                    if value:
                        counts[key][value] = counts[key].get(value, 0) + 1
        self.po_index = {}
        self.po_bl_index = {}
        for row in self.po_master:
            self.po_index.setdefault(str(row.get("A")).strip(), row)
            bl = unique_key(row.get("F"))
            if bl:
                self.po_bl_index.setdefault(bl, []).append(str(row.get("A")).strip())
        self.di_index = {}
        self.bl_index = {}
        for row in self.master:
            self.di_index.setdefault(str(row.get("A")).strip(), row)
            bl = unique_key(row.get("N"))
            if bl:
                self.bl_index.setdefault(bl, []).append(str(row.get("A")).strip())
        self.item_index = {}
        self.item_groups = {}
        for row in self.items:
            key = str(row.get("A")).strip()
            self.item_index.setdefault(key, row)
            self.item_groups.setdefault(key, []).append(row)
        self.price_index = {}
        for row in self.prices:
            key = f"{row.get('G') or ''}{row.get('H') or ''}"
            self.price_index.setdefault(key, []).append(row)

    def lines_from_prices(self, di_key):
        """Item dan Qty per DI dari tab Harga, urut sesuai kemunculan pertama."""
        lines = {}
        for row in self.prices:
            if str(row.get("G") or "").strip() != di_key or not row.get("H"):
                continue
            sku = str(row.get("H")).strip()
            line = lines.setdefault(sku, {"sku": sku, "qty": 0})
            if isinstance(row.get("D"), (int, float)):
                line["qty"] += row.get("D")
        return list(lines.values())

    def price_groups(self, key):
        """Harga berbeda untuk satu DI/PO + SKU, urut kemunculan, dengan total Qty (kolom D) per harga."""
        groups = {}
        for row in self.price_index.get(key, []):
            price = row.get("E")
            if not isinstance(price, (int, float)):
                continue
            group = groups.setdefault(round(price, 2), {"price": price, "service": None, "qty": 0, "has_qty": True})
            if group["service"] is None and isinstance(row.get("F"), (int, float)):
                group["service"] = row.get("F")
            if isinstance(row.get("D"), (int, float)):
                group["qty"] += row.get("D")
            else:
                group["has_qty"] = False
        return list(groups.values())

    def split_prices(self, di_key, lines):
        """SKU dengan 2 harga atau lebih: satu baris per harga, Qty mengikuti tab Harga.
        Kalau file loading sudah punya SKU itu di beberapa baris, tiap baris dipasangkan dengan harganya (Qty sama didahulukan).
        Baris yang harganya sudah diisi manual tidak diubah."""
        open_lines = {}
        for line in lines:
            sku = str(line.get("sku") or "").strip()
            if sku and line.get("price") is None and not line.get("split"):
                open_lines.setdefault(sku, []).append(line)
        plans = {}
        for sku, group_lines in open_lines.items():
            groups = self.price_groups(f"{di_key}{sku}")
            if len(groups) < 2 or not all(g["has_qty"] and g["qty"] > 0 for g in groups):
                continue
            plans[sku] = self.price_plan(groups, group_lines)
        if not plans:
            return lines
        result, done = [], set()
        for line in lines:
            sku = str(line.get("sku") or "").strip()
            if sku not in plans or line.get("price") is not None or line.get("split"):
                result.append(line)
                continue
            if sku in done:
                continue
            done.add(sku)
            result.extend(plans[sku])
        return result

    def price_plan(self, groups, group_lines):
        count = len(groups)
        planned_total = sum(l.get("planned") for l in group_lines if isinstance(l.get("planned"), (int, float))) if any(
            isinstance(l.get("planned"), (int, float)) for l in group_lines) else None
        if len(group_lines) == count:
            # loading sudah dipisah: pasangkan baris dengan harga yang Qty-nya sama, sisanya berurutan
            unused = list(range(count))
            pairs = [None] * count
            for i, line in enumerate(group_lines):
                match = next((g for g in unused if isinstance(line.get("qty"), (int, float)) and abs(groups[g]["qty"] - line["qty"]) < 1e-6), None)
                if match is not None:
                    pairs[i] = match
                    unused.remove(match)
            for i in range(count):
                if pairs[i] is None:
                    pairs[i] = unused.pop(0)
            out = []
            for line, g in zip(group_lines, pairs):
                new = dict(line)
                new.update(price=groups[g]["price"], service=groups[g]["service"], split=count, split_note="")
                if isinstance(line.get("qty"), (int, float)) and abs(groups[g]["qty"] - line["qty"]) > 1e-6:
                    new["split_note"] = (f"Qty {format_value(float(line['qty']))} beda dengan Qty di tab Harga "
                                         f"{format_value(float(groups[g]['qty']))} untuk harga {format_value(float(groups[g]['price']))}")
                out.append(new)
            return out
        base = dict(group_lines[0])
        quantities = [l.get("qty") for l in group_lines]
        qty = sum(q for q in quantities if isinstance(q, (int, float))) if all(isinstance(q, (int, float)) for q in quantities) else None
        qtys = [g["qty"] for g in groups]
        total, note = sum(qtys), ""
        if qty is not None and abs(total - qty) > 1e-6:
            rest = qty - sum(qtys[:-1])
            if rest > 0:
                note = (f"Qty di tab Harga {format_value(float(total))} beda dengan Qty Order {format_value(float(qty))}; "
                        f"Qty harga {format_value(float(groups[-1]['price']))} disesuaikan jadi {format_value(float(rest))}")
                qtys[-1] = rest
            else:
                note = f"Qty di tab Harga {format_value(float(total))} beda dengan Qty Order {format_value(float(qty))}"
        out = []
        for index, (group, part) in enumerate(zip(groups, qtys)):
            new = dict(base)
            new.update(qty=part, price=group["price"], service=group["service"], split=count, planned=None,
                       split_note=note if index == count - 1 else "")
            if index == 0:
                new.update(planned_total=planned_total, qty_total=qty)
            out.append(new)
        return out

    def distinct_prices(self, key):
        return sorted({round(r.get("E"), 2) for r in self.price_index.get(key, []) if isinstance(r.get("E"), (int, float))})

    def master_flag(self, row, profile=None):
        profile = profile or PROFILES["DI"]
        missing = [label for key, label in profile.required.items() if is_blank(row.get(key))]
        parts = []
        duplicates = self.duplicate_labels(row, profile)
        if duplicates:
            parts.append("Double: " + ", ".join(duplicates))
        if missing:
            parts.append("Kosong: " + ", ".join(missing))
        return "; ".join(parts)

    def goods_flag(self, row):
        if is_blank(row.get("B")):
            return "Deskripsi kosong"
        return check_description(row.get("A"), row.get("B")) or ""

    def weight_flag(self, row):
        missing = [label for key, label in (("D", "CBM"), ("E", "NW"), ("F", "GW")) if is_blank(row.get(key))]
        parts = []
        estimate = self.estimate_note(row)
        if estimate:
            parts.append(estimate)
        tally = cbm_mismatch(row.get("D"), row.get("LC"))
        if tally:
            parts.append(tally)
        if missing:
            parts.append("Kosong: " + ", ".join(missing))
        return "; ".join(parts)

    def price_flag(self, row):
        if not isinstance(row.get("E"), (int, float)) or row.get("E") == 0:
            return "Harga kosong"
        count = len(self.distinct_prices(f"{row.get('G') or ''}{row.get('H') or ''}"))
        return f"{count} harga berbeda" if count > 1 else ""

    def new_row(self, table, values=None):
        row = Row(None, dict(values or {}))
        row.dirty = {key for key, value in row.values.items() if value not in (None, "")}
        self.rows_of(table).append(row)
        self.reindex()
        return row

    def dirty_count(self, table):
        return sum(1 for r in self.rows_of(table) if r.dirty)

    def backup(self, path):
        base = os.path.join(settings_dir(), "backup")
        os.makedirs(base, exist_ok=True)
        stem, ext = os.path.splitext(os.path.basename(path))
        shutil.copy2(path, os.path.join(base, f"{stem}_{time.strftime('%Y%m%d_%H%M%S')}{ext}"))
        old = sorted(n for n in os.listdir(base) if n.startswith(stem + "_"))
        for name in old[:-10]:
            os.remove(os.path.join(base, name))

    def open_for_write(self, app, path, password, mtime):
        if os.path.getmtime(path) != mtime:
            raise StaleFileError(f"{os.path.basename(path)} sudah berubah sejak dimuat. Klik Muat ulang dulu.")
        self.backup(path)
        workbook = open_source(app, path, password, write=True)
        if workbook.ReadOnly:
            workbook.Close(False)
            raise StaleFileError(f"{os.path.basename(path)} sedang dibuka orang lain atau hanya bisa dibaca.")
        return workbook

    def delete_rows(self, app, table, rows):
        """Hapus baris dari file sumber, dari bawah ke atas, lalu muat ulang semua data."""
        if table == "master":
            path, password, mtime, sheet_name, key = self.master_path, self.master_password, self.master_mtime, MASTER_SHEET, "A"
        elif table == "po_master":
            path, password, mtime, sheet_name, key = self.po_path, self.po_password, self.po_mtime, MASTER_PO_SHEET, "A"
        elif table == "items":
            path, password, mtime, sheet_name, key = self.item_path, self.item_password, self.item_mtime, ITEM_SHEET, "A"
        else:
            path, password, mtime, sheet_name, key = self.item_path, self.item_password, self.item_mtime, PRICE_SHEET, "H"
        targets = {}
        for row in rows:
            group = self.item_groups.get(str(row.get("A")).strip(), [row]) if table == "items" else [row]
            for target in group:
                targets[target.row] = target
        workbook = self.open_for_write(app, path, password, mtime)
        try:
            sheet = workbook.Sheets.Item(sheet_name)
            calculation(app, XL_MANUAL)
            for number in sorted(targets):
                verify(sheet, number, key, targets[number].get(key))
            for number in sorted(targets, reverse=True):
                sheet.Rows.Item(number).Delete()
            finish_write(app, workbook)
        except Exception:
            calculation(app, XL_AUTOMATIC)
            workbook.Close(False)
            raise
        workbook.Close(False)
        self.load(app)
        return len(targets)

    def save_master(self, app):
        rows = [r for r in self.master if r.dirty]
        if not rows:
            return 0
        workbook = self.open_for_write(app, self.master_path, self.master_password, self.master_mtime)
        try:
            sheet = workbook.Sheets.Item(MASTER_SHEET)
            calculation(app, XL_MANUAL)
            for row in rows:
                if row.row is None:
                    self.master_last += 1
                    row.row = self.master_last
                    n = row.row
                    sheet.Cells(n, column_number("E")).Formula = f'="INV"&"-"&D{n}'
                    sheet.Cells(n, column_number("F")).Formula = f'=D{n}&"-"&C{n}'
                    sheet.Cells(n, column_number("Y")).Formula = f"=SUMIF($N:$N,N{n},$X:$X)"
                    row.dirty |= {"A"}
                else:
                    verify(sheet, row.row, "A", row.get("A"))
                for key in sorted(row.dirty):
                    sheet.Cells(row.row, column_number(key)).Value2 = row.get(key)
            finish_write(app, workbook)
        except Exception:
            calculation(app, XL_AUTOMATIC)
            workbook.Close(False)
            raise
        workbook.Close(False)
        for row in rows:
            row.dirty.clear()
            refresh_master_formulas(row)
        self.master_mtime = os.path.getmtime(self.master_path)
        self.reindex()
        return len(rows)

    def save_po(self, app):
        rows = [r for r in self.po_master if r.dirty]
        if not rows:
            return 0
        workbook = self.open_for_write(app, self.po_path, self.po_password, self.po_mtime)
        try:
            sheet = workbook.Sheets.Item(MASTER_PO_SHEET)
            calculation(app, XL_MANUAL)
            for row in rows:
                if row.row is None:
                    self.po_last += 1
                    row.row = self.po_last
                    row.dirty |= {"A"}
                else:
                    verify(sheet, row.row, "A", row.get("A"))
                for key in sorted(row.dirty):
                    sheet.Cells(row.row, column_number(key)).Value2 = row.get(key)
            finish_write(app, workbook)
        except Exception:
            calculation(app, XL_AUTOMATIC)
            workbook.Close(False)
            raise
        workbook.Close(False)
        for row in rows:
            row.dirty.clear()
        self.po_mtime = os.path.getmtime(self.po_path)
        self.reindex()
        return len(rows)

    def save_items(self, app):
        items = [r for r in self.items if r.dirty]
        prices = [r for r in self.prices if r.dirty]
        if not items and not prices:
            return 0
        workbook = self.open_for_write(app, self.item_path, self.item_password, self.item_mtime)
        try:
            calculation(app, XL_MANUAL)
            if items:
                self._write_items(workbook, items)
            if prices:
                self._write_prices(workbook, prices)
            finish_write(app, workbook)
        except Exception:
            calculation(app, XL_AUTOMATIC)
            workbook.Close(False)
            raise
        workbook.Close(False)
        for row in items + prices:
            row.dirty.clear()
        self.item_mtime = os.path.getmtime(self.item_path)
        self.reindex()
        return len(items) + len(prices)

    def _write_items(self, workbook, rows):
        sheet = workbook.Sheets.Item(ITEM_SHEET)
        weights = workbook.Sheets.Item(WEIGHT_SHEET)
        for row in rows:
            sku = str(row.get("A") or "").strip()
            if row.row is None:
                self.item_last += 1
                row.row = n = self.item_last
                sheet.Cells(n, 1).Value2 = sku
                lookup = f"'{WEIGHT_SHEET}'!$C$1:$G$9987"
                sheet.Cells(n, 5).Formula = f"=VLOOKUP(A{n},{lookup},4,0)"
                sheet.Cells(n, 6).Formula = f"=VLOOKUP(A{n},{lookup},5,0)"
                sheet.Cells(n, 8).Formula = f"=F{n}-E{n}"
                targets = [row]
            else:
                targets = self.item_groups.get(sku, [row])
            for target in targets:
                if target.row != row.row or row.row <= self.item_last:
                    verify(sheet, target.row, "A", sku)
                for key in sorted(row.dirty - {"A", "E", "F"}):
                    sheet.Cells(target.row, column_number(key)).Value2 = row.get(key)
                for key, source in (("E", "F"), ("F", "G")):
                    if key not in row.dirty:
                        continue
                    cell = sheet.Cells(target.row, column_number(key))
                    if cell.HasFormula:
                        self._write_weight(weights, sku, source, row.get(key))
                    else:
                        cell.Value2 = row.get(key)

    def _write_weight(self, sheet, sku, column, value):
        numbers = self.weight_rows.get(sku)
        if not numbers:
            self.weight_last += 1
            numbers = self.weight_rows[sku] = [self.weight_last]
            sheet.Cells(self.weight_last, 3).Value2 = sku
        for number in numbers:
            sheet.Cells(number, column_number(column)).Value2 = value

    def _write_prices(self, workbook, rows):
        sheet = workbook.Sheets.Item(PRICE_SHEET)
        for row in rows:
            if row.row is None:
                self.price_last += 1
                row.row = n = self.price_last
                sheet.Cells(n, 1).Formula = f"=G{n}&H{n}"
                sheet.Cells(n, 2).Formula = f"=E{n}+F{n}"
                sheet.Cells(n, 3).Formula = f"=D{n}*E{n}"
            else:
                verify(sheet, row.row, "H", row.get("H"))
            for key in sorted(row.dirty - {"B"}):
                sheet.Cells(row.row, column_number(key)).Value2 = row.get(key)
            if isinstance(row.get("E"), (int, float)) and isinstance(row.get("F"), (int, float)):
                row.values["B"] = round(row.get("E") + row.get("F"), 2)

    def check_items(self, di_key, lines, profile=None):
        """Isi baris item dari data app dan kumpulkan semua hal yang perlu diperbaiki."""
        profile = profile or PROFILES["DI"]
        problems = []
        master = self.index_of(profile).get(di_key)
        if master is None:
            problems.append(f"{di_key} tidak ada di Master {profile.name}.")
        else:
            flag = self.master_flag(master, profile)
            if flag:
                problems.append(f"Master {di_key}: {flag}")
        if len(lines) > profile.max_items:
            problems.append(f"Maksimal {profile.max_items} item per {profile.name}.")
        result = []
        lines = self.split_prices(di_key, lines)
        for line in lines:
            sku = str(line.get("sku") or "").strip()
            item = self.item_index.get(sku)
            prices = self.distinct_prices(f"{di_key}{sku}")
            price_rows = self.price_index.get(f"{di_key}{sku}", [])
            first = next((r.get("E") for r in price_rows if isinstance(r.get("E"), (int, float))), None)
            service = next((r.get("F") for r in price_rows if isinstance(r.get("F"), (int, float))), None)
            filled = {
                "sku": sku, "qty": line.get("qty"),
                "asin": line.get("asin") or text_or_blank(item and item.get("G")),
                "po": line.get("po") or (text_or_blank(master and master.get("C")) if profile.name == "DI" else ""),
                "name": text_or_blank(item and item.get("B")), "hs": clean_hs(item and item.get("C")),
                "nw": number_or_none(item and item.get("E")), "gw": number_or_none(item and item.get("F")),
                "cbm": line.get("cbm") if line.get("cbm") is not None else number_or_none(item and item.get("D")),
                "price": line.get("price") if line.get("price") is not None else first,
                "service": line.get("service") if line.get("service") is not None else service,
            }
            notes = []
            if not sku:
                notes.append("Item No. kosong")
            elif item is None:
                notes.append("SKU tidak ada di Goods Description")
            else:
                typo = self.goods_flag(item)
                if typo:
                    notes.append(f"Typo: {typo}")
                estimate = self.estimate_note(item)
                if estimate:
                    notes.append(estimate)
            if item is not None and line.get("cbm") is not None:
                tally = cbm_mismatch(item.get("D"), line.get("cbm"))
                if tally:
                    notes.append(tally)
            filled["split"] = line.get("split") or 0
            planned, ordered = line.get("planned"), filled["qty"]
            if line.get("planned_total") is not None:
                planned, ordered = line.get("planned_total"), line.get("qty_total")
            if isinstance(planned, (int, float)) and isinstance(ordered, (int, float)) and planned != ordered:
                notes.append(f"Qty Order {format_value(float(ordered))} beda dengan Qty Diatur {format_value(float(planned))}")
            if line.get("split_note"):
                notes.append(line["split_note"])
            if line.get("price") is None:
                if len(prices) > 1:
                    notes.append(f"{len(prices)} harga: {', '.join(map(str, prices))}")
                elif not prices:
                    notes.append("Harga tidak ditemukan")
            for key, label in profile.line_required:
                if is_blank(filled[key]):
                    notes.append(f"{label} kosong")
            filled["notes"] = notes
            result.append(filled)
            if notes:
                problems.append(f"{sku or '(baris kosong)'}: " + "; ".join(notes))
        expected = master.get(profile.qty_key) if master else None
        total = sum(line["qty"] for line in result if isinstance(line["qty"], (int, float)))
        if result and isinstance(expected, (int, float)) and expected and abs(total - expected) > 1e-6:
            problems.append(
                f"Total Qty item {format_value(float(total))} berbeda dengan Qty di Master {profile.name} {format_value(float(expected))}"
                f" (kolom {profile.qty_key}). Cek angka Qty di Master atau item yang belum masuk."
            )
        return result, problems

    def file_code(self, master, profile=None):
        profile = profile or PROFILES["DI"]
        code = master.get(profile.code_key)
        if is_blank(code):
            code = f"{master.get('D') or ''}-{master.get('C') or ''}" if profile.name == "DI" else str(master.get("A"))
        return re.sub(r'[<>:"/\\|?*]', "_", str(code))


def unique_key(value):
    return "" if is_blank(value) else format_value(value).strip().upper()


def round_down_2(value):
    """CBM selalu dibulatkan ke bawah. round() menghilangkan sisa kecil dari hitungan float."""
    return math.floor(round(value * 100, 6)) / 100


def line_totals(lines):
    """Total per DI dengan rumus yang sama seperti sheet INPUT2."""
    number = lambda value: value if isinstance(value, (int, float)) else 0
    totals = {"qty": 0, "ctns": 0, "nw": 0, "gw": 0, "cbm": 0, "price": 0, "amount": 0}
    for line in lines:
        qty = number(line.get("qty"))
        totals["qty"] += qty
        totals["ctns"] += qty
        totals["nw"] += round(number(line.get("nw")) * qty, 2)
        totals["gw"] += round(number(line.get("gw")) * qty, 2)
        totals["cbm"] += round_down_2(number(line.get("cbm")) * qty)
        totals["price"] += number(line.get("price"))
        totals["amount"] += round(number(line.get("price")) * qty, 2)
    return {key: round(value, 2) for key, value in totals.items()}


def cbm_mismatch(master_value, loading_value):
    if not isinstance(loading_value, (int, float)) or not loading_value:
        return ""
    if not isinstance(master_value, (int, float)) or not master_value:
        return f"CBM tidak tally: kosong, loading {format_value(float(loading_value))}"
    if abs(master_value - loading_value) > CBM_TOLERANCE:
        return f"CBM tidak tally: {format_value(float(master_value))}, loading {format_value(float(loading_value))}"
    return ""


def is_old_po(row):
    match = re.search(r"(20\d\d)", str(row.get("B") or ""))
    return bool(match) and match.group(1) < FIRST_YEAR


def is_old(row):
    rolling = str(row.get("D") or "")
    return rolling[:4].isdigit() and rolling[:4] < FIRST_YEAR


def text_or_blank(value):
    return "" if is_blank(value) else format_value(value)


def number_or_none(value):
    return value if isinstance(value, (int, float)) and value not in EXCEL_ERRORS else None


def refresh_master_formulas(row):
    rolling, po = row.get("D"), row.get("C")
    if rolling:
        row.values["E"] = f"INV-{rolling}"
        row.values["F"] = f"{rolling}-{po or ''}"


def verify(sheet, number, key, expected):
    actual = sheet.Cells(number, column_number(key)).Value2
    if str(actual or "").strip() != str(expected or "").strip():
        raise StaleFileError(f"Baris {number} berubah di file sumber. Klik Muat ulang dulu.")


def calculation(app, mode):
    try:
        app.Calculation = mode
    except Exception:
        pass


def finish_write(app, workbook):
    calculation(app, XL_AUTOMATIC)
    try:
        app.Calculate()
    except Exception:
        pass
    workbook.Save()


def settings_dir():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "PDF Gen")
    os.makedirs(path, exist_ok=True)
    return path


def load_settings():
    try:
        with open(os.path.join(settings_dir(), "settings.json"), encoding="utf-8") as handle:
            return json.load(handle)
    except (OSError, ValueError):
        return {}


def save_settings(values):
    with open(os.path.join(settings_dir(), "settings.json"), "w", encoding="utf-8") as handle:
        json.dump(values, handle, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------- Excel DI
def _column_of(ref):
    return column_number(re.match(r"[A-Z]+", ref).group())


def get_cell(sheet, ref):
    number = int(re.search(r"\d+", ref).group())
    data = sheet.find(Q + "sheetData")
    row = next((r for r in data if int(r.get("r")) == number), None)
    if row is None:
        row = etree.Element(Q + "row", r=str(number))
        data.insert(len([r for r in data if int(r.get("r")) < number]), row)
    for index, cell in enumerate(row):
        if cell.get("r") == ref:
            return cell
        if _column_of(cell.get("r")) > _column_of(ref):
            new = etree.Element(Q + "c", r=ref)
            row.insert(index, new)
            return new
    return etree.SubElement(row, Q + "c", r=ref)


def set_cell(sheet, ref, value=None, formula=None):
    cell = get_cell(sheet, ref)
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)
    cell.attrib.pop("cm", None)
    if formula is not None:
        etree.SubElement(cell, Q + "f").text = formula
    elif isinstance(value, str) and value != "":
        cell.set("t", "inlineStr")
        text = etree.SubElement(etree.SubElement(cell, Q + "is"), Q + "t")
        text.text = value
    elif isinstance(value, (int, float)):
        etree.SubElement(cell, Q + "v").text = repr(float(value)) if isinstance(value, float) else str(value)


def set_cached(cell, value):
    for child in cell.findall(Q + "v"):
        cell.remove(child)
    cell.attrib.pop("t", None)
    if value in (None, ""):
        value = 0
    if isinstance(value, str):
        cell.set("t", "e" if value.startswith("#") else "str")
    etree.SubElement(cell, Q + "v").text = repr(float(value)) if isinstance(value, float) else str(value)


def master_link_cache(master, profile=None):
    """Salinan satu baris master di dalam link, dipakai sampai link diperbarui dari server."""
    relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    root = etree.Element(Q + "externalLink", nsmap={None: NS})
    book = etree.SubElement(root, Q + "externalBook", nsmap={"r": relationships})
    book.set("{%s}id" % relationships, "rId1")
    names = etree.SubElement(book, Q + "sheetNames")
    profile = profile or PROFILES["DI"]
    etree.SubElement(names, Q + "sheetName", val=profile.sheet)
    dataset = etree.SubElement(book, Q + "sheetDataSet")
    sheet_data = etree.SubElement(dataset, Q + "sheetData", sheetId="0")
    row = etree.SubElement(sheet_data, Q + "row", r=str(master.row))
    keys = sorted({"A"} | {key for _, key in profile.formula_cells}, key=column_number)
    for key in keys:
        value = master.get(key)
        if value in (None, ""):
            continue
        cell = etree.SubElement(row, Q + "cell", r=f"{key}{master.row}")
        if isinstance(value, str):
            cell.set("t", "e" if value.startswith("#") else "str")
        etree.SubElement(cell, Q + "v").text = repr(float(value)) if isinstance(value, float) else str(value)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def fill_po_input2(sheet, lines):
    for index in range(max(len(lines), 21)):
        r = FIRST_ITEM_ROW + index
        line = lines[index] if index < len(lines) else None
        cells = {}
        if line:
            cells = {"A": line["sku"], "B": line["name"], "C": line["qty"], "D": "PCS", "I": line["price"], "J": line.get("service"),
                     "R": line["nw"], "S": line["gw"], "T": line["cbm"], "U": line["hs"]}
        for column in "ABCDIJRSTU":
            set_cell(sheet, f"{column}{r}", cells.get(column))
        guard = lambda body: f'IF($A{r}="","",{body})'
        for column, body in (("E", f"ROUND(N($R{r})*N($C{r}),2)"), ("F", f"ROUND(N($S{r})*N($C{r}),2)"), ("G", f"$C{r}"),
                             ("H", f"ROUNDDOWN(N($T{r})*N($C{r}),2)"), ("K", f"ROUND(N($C{r})*N($I{r}),2)"),
                             ("V", f'LEFT($A{r},FIND("~",SUBSTITUTE($A{r},"-","~",LEN($A{r})-LEN(SUBSTITUTE($A{r},"-",""))))-1)')):
            set_cell(sheet, f"{column}{r}", formula=guard(body))


def build_excel(template, output, di_key, master, lines, profile=None):
    profile = profile or PROFILES["DI"]
    source = zipfile.ZipFile(template)
    files = {i.filename: source.read(i.filename) for i in source.infolist() if not i.is_dir()}
    source.close()
    # template punya baris item tetap (PO 20, DI 19); kalau item lebih banyak, baris CI/PL/SI/INPUT2 ditambah
    add_item_rows(files, profile.name, len(lines))
    sheet = etree.fromstring(files["xl/worksheets/sheet1.xml"])
    number = di_key.split("柜")[-1]
    set_cell(sheet, "O2", int(number) if number.isdigit() else number)
    # Rumus XLOOKUP ke master server tetap ada. Hasil terakhir disimpan agar langsung tampil.
    for ref, key in profile.formula_cells:
        set_cached(get_cell(sheet, ref), master.get(key))
    # C36 tetap rumus TEXTJOIN dari INPUT2, jadi foam dan spring otomatis tampil dua kode.
    hs_codes = " / ".join(dict.fromkeys(str(line["hs"]) for line in lines if line["hs"]))
    hs_cell = get_cell(sheet, "C36")
    if hs_codes:
        set_cached(hs_cell, hs_codes)
    else:
        for child in hs_cell.findall(Q + "v"):
            hs_cell.remove(child)
        hs_cell.attrib.pop("t", None)
    files["xl/worksheets/sheet1.xml"] = etree.tostring(sheet, xml_declaration=True, encoding="UTF-8", standalone=True)
    sheet = etree.fromstring(files["xl/worksheets/sheet2.xml"])
    if profile.name == "PO":
        fill_po_input2(sheet, lines)
    for index in range(max(len(lines), MAX_ITEMS) if profile.name == "DI" else 0):
        r = FIRST_ITEM_ROW + index
        line = lines[index] if index < len(lines) else None
        cells = {}
        if line:
            cells = {"A": line["sku"], "B": line["asin"], "C": line["po"], "D": line["name"], "E": line["qty"], "F": "PCS",
                     "K": line["price"], "N": line["nw"], "O": line["gw"], "P": line["cbm"], "Q": line["hs"]}
        for column in "ABCDEFKNOPQ":
            set_cell(sheet, f"{column}{r}", cells.get(column))
        guard = lambda body: f'IF($A{r}="","",{body})'
        for column, body in (("G", f"ROUND(N($N{r})*N($E{r}),2)"), ("H", f"ROUND(N($O{r})*N($E{r}),2)"), ("I", f"$E{r}"),
                             ("J", f"ROUNDDOWN(N($P{r})*N($E{r}),2)"), ("L", f"ROUND(N($K{r})*N($E{r}),2)"),
                             ("R", f'LEFT($A{r},FIND("~",SUBSTITUTE($A{r},"-","~",LEN($A{r})-LEN(SUBSTITUTE($A{r},"-",""))))-1)')):
            set_cell(sheet, f"{column}{r}", formula=guard(body))
    files["xl/worksheets/sheet2.xml"] = etree.tostring(sheet, xml_declaration=True, encoding="UTF-8", standalone=True)
    files[MASTER_LINK_PART] = master_link_cache(master, profile)
    temporary = output + ".tmp"
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in files.items():
            out.writestr(name, data)
    os.replace(temporary, output)


def make_template(source, output, profile=None):
    profile = profile or PROFILES["DI"]
    """Dipakai sekali saat membuat installer: link master tetap ada, link lain dan data DI dibuang."""
    archive = zipfile.ZipFile(source)
    files = {i.filename: archive.read(i.filename) for i in archive.infolist() if not i.is_dir()}
    for name in ("xl/externalLinks/externalLink2.xml", "xl/externalLinks/_rels/externalLink2.xml.rels"):
        files.pop(name, None)
    files[MASTER_LINK_PART] = (
        f'<externalLink xmlns="{NS}"><externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">'
        f'<sheetNames><sheetName val="{profile.sheet}"/></sheetNames></externalBook></externalLink>'
    ).encode("utf-8")
    for name in [n for n in files if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)]:
        sheet = etree.fromstring(files[name])
        for cell in sheet.iter(Q + "c"):
            if cell.find(Q + "f") is not None:
                value = cell.find(Q + "v")
                if value is not None:
                    cell.remove(value)
                cell.attrib.pop("t", None)
        files[name] = etree.tostring(sheet, xml_declaration=True, encoding="UTF-8", standalone=True)
    types = files["[Content_Types].xml"].decode()
    files["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/externalLinks/externalLink2.xml"[^>]*/>', "", types).encode()
    book = files["xl/workbook.xml"].decode()
    rels = files["xl/_rels/workbook.xml.rels"].decode()
    target = re.search(r'<Relationship Id="(\w+)"[^>]*Target="externalLinks/externalLink2.xml"[^>]*/>', rels)
    rels = rels.replace(target.group(0), "")
    book = book.replace(f'<externalReference r:id="{target.group(1)}"/>', "")
    book = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="191029" fullCalcOnLoad="1"/>', book)
    files["xl/_rels/workbook.xml.rels"] = rels.encode()
    files["xl/workbook.xml"] = book.encode()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as out:
        for name, content in files.items():
            out.writestr(name, content)


PROFILES = _profiles()
