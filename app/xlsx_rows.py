"""Menambah baris item di template Excel (CI, PL, SI, INPUT2) tanpa merusak rumus, merge, gambar, dan area cetak.
Cara kerjanya sama seperti 'Insert row' di Excel: baris di bawahnya turun, rumus yang menunjuk ke sana ikut disesuaikan,
dan rentang SUM yang berakhir di baris item terakhir ikut diperpanjang."""
import copy
import re

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
Q = "{%s}" % NS
XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

CELL = re.compile(r"^([A-Z]{1,3})(\d+)$")
# rentang atau sel tunggal, boleh diawali nama sheet
REF = re.compile(
    r"(?<![A-Za-z0-9_\]!.$])((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![A-Za-z0-9_(])"
)


def col_number(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def sheet_of(prefix):
    if not prefix:
        return None
    name = prefix[:-1]
    if name.startswith("'"):
        name = name[1:-1].replace("''", "'")
    return name


def map_refs(formula, fn):
    """Jalankan fn pada setiap referensi di luar teks berkutip."""
    parts = re.split(r'("(?:[^"]|"")*")', formula)
    for i in range(0, len(parts), 2):
        parts[i] = REF.sub(fn, parts[i])
    return "".join(parts)


def shift_relative(formula, drow, dcol=0):
    """Geser referensi relatif (seperti saat rumus disalin ke baris lain)."""

    def one(dollar_c, col, dollar_r, row):
        c = col if dollar_c else col_letters(col_number(col) + dcol)
        r = row if dollar_r else str(int(row) + drow)
        return f"{dollar_c}{c}{dollar_r}{r}"

    def fn(m):
        prefix = m.group(1) or ""
        text = prefix + one(m.group(2), m.group(3), m.group(4), m.group(5))
        if m.group(7):
            text += ":" + one(m.group(6), m.group(7), m.group(8), m.group(9))
        return text

    return map_refs(formula, fn)


def shift_inserted(formula, home, target, after, count, first):
    """Sesuaikan referensi ke sheet `target` setelah `count` baris disisipkan di bawah baris `after`.
    `home` = sheet tempat rumus berada (None untuk defined name)."""

    def fn(m):
        prefix = m.group(1) or ""
        sheet = sheet_of(prefix) or home
        if sheet != target:
            return m.group(0)
        r1 = int(m.group(5))
        nr1 = r1 + count if r1 > after else r1
        text = f"{prefix}{m.group(2)}{m.group(3)}{m.group(4)}{nr1}"
        if m.group(7):
            r2 = int(m.group(9))
            if r2 > after or (r2 == after and first <= r1 <= after and r1 < r2):
                r2 += count
            text += f":{m.group(6)}{m.group(7)}{m.group(8)}{r2}"
        return text

    return map_refs(formula, fn)


def shift_sqref(value, after, count):
    out = []
    for part in value.split():
        pieces = []
        for ref in part.split(":"):
            m = re.match(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$", ref)
            if m:
                row = int(m.group(4))
                pieces.append(f"{m.group(1)}{m.group(2)}{m.group(3)}{row + count if row > after else row}")
            else:
                pieces.append(ref)
        out.append(":".join(pieces))
    return " ".join(out)


class Book:
    def __init__(self, files):
        self.files = files
        book = etree.fromstring(files["xl/workbook.xml"])
        rels = etree.fromstring(files["xl/_rels/workbook.xml.rels"])
        targets = {r.get("Id"): r.get("Target") for r in rels}
        self.book = book
        self.paths = {}
        for sheet in book.iter(Q + "sheet"):
            target = targets[sheet.get(REL + "id")].lstrip("/")
            self.paths[sheet.get("name")] = target if target.startswith("xl/") else "xl/" + target
        self.trees = {}
        self.expanded = set()

    def tree(self, name):
        if name not in self.trees:
            self.trees[name] = etree.fromstring(self.files[self.paths[name]])
        return self.trees[name]

    def save(self):
        for name, tree in self.trees.items():
            self.files[self.paths[name]] = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.files["xl/workbook.xml"] = etree.tostring(self.book, xml_declaration=True, encoding="UTF-8", standalone=True)
        # urutan hitung lama tidak berlaku lagi; Excel membuat ulang
        if "xl/calcChain.xml" in self.files:
            del self.files["xl/calcChain.xml"]
            types = self.files["[Content_Types].xml"].decode("utf-8")
            self.files["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", types).encode("utf-8")
            rels = self.files["xl/_rels/workbook.xml.rels"].decode("utf-8")
            self.files["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship [^>]*Target="calcChain.xml"[^>]*/>', "", rels).encode("utf-8")

    def expand_shared(self, name):
        """Rumus 'shared' diubah jadi rumus biasa supaya aman saat baris digeser."""
        if name in self.expanded:
            return
        self.expanded.add(name)
        tree = self.tree(name)
        masters = {}
        for cell in tree.iter(Q + "c"):
            f = cell.find(Q + "f")
            if f is not None and f.get("t") == "shared" and f.get("ref") and f.text:
                masters[f.get("si")] = (cell.get("r"), f.text)
        for cell in tree.iter(Q + "c"):
            f = cell.find(Q + "f")
            if f is None or f.get("t") != "shared":
                continue
            master = masters.get(f.get("si"))
            if master is None:
                continue
            mc, mr = CELL.match(master[0]).groups()
            cc, cr = CELL.match(cell.get("r")).groups()
            f.text = shift_relative(master[1], int(cr) - int(mr), col_number(cc) - col_number(mc))
            for key in ("t", "ref", "si"):
                f.attrib.pop(key, None)

    def insert_rows(self, name, after, count, first):
        """Sisipkan `count` baris di bawah baris `after` pada sheet `name`, salin format dan rumus baris `after`."""
        if count <= 0:
            return
        for other in self.paths:
            self.expand_shared(other)
        tree = self.tree(name)
        data = tree.find(Q + "sheetData")
        template = None
        for row in list(data):
            r = int(row.get("r"))
            if r == after:
                template = row
            if r > after:
                row.set("r", str(r + count))
                for cell in row:
                    col, _ = CELL.match(cell.get("r")).groups()
                    cell.set("r", f"{col}{r + count}")
        # rumus di seluruh workbook yang menunjuk ke sheet ini
        for other in self.paths:
            for f in self.tree(other).iter(Q + "f"):
                if f.text:
                    f.text = shift_inserted(f.text, other, name, after, count, first)
        for dn in self.book.iter(Q + "definedName"):
            if dn.text and "#REF!" not in dn.text:
                dn.text = shift_inserted(dn.text, None, name, after, count, first)
        # baris baru: salinan baris `after`, rumus relatif turun satu per baris
        if template is not None:
            position = list(data).index(template)
            for i in range(1, count + 1):
                new = copy.deepcopy(template)
                new.set("r", str(after + i))
                for cell in new:
                    col, _ = CELL.match(cell.get("r")).groups()
                    cell.set("r", f"{col}{after + i}")
                    f = cell.find(Q + "f")
                    if f is not None and f.text:
                        f.text = shift_relative(f.text, i)
                    v = cell.find(Q + "v")
                    if v is not None and f is not None:
                        cell.remove(v)
                    if f is None and v is not None:
                        cell.remove(v)
                        cell.attrib.pop("t", None)
                    for child in cell.findall(Q + "is"):
                        cell.remove(child)
                        cell.attrib.pop("t", None)
                data.insert(position + i, new)
        # merge, format bersyarat, validasi, hyperlink, filter
        merges = tree.find(Q + "mergeCells")
        if merges is not None:
            extra = []
            for m in merges:
                ref = m.get("ref")
                a, b = ref.split(":")
                ra, rb = int(CELL.match(a).group(2)), int(CELL.match(b).group(2))
                if ra == rb == after:
                    ca, cb = CELL.match(a).group(1), CELL.match(b).group(1)
                    extra += [f"{ca}{after + i}:{cb}{after + i}" for i in range(1, count + 1)]
                m.set("ref", shift_sqref(ref, after, count))
            for ref in extra:
                etree.SubElement(merges, Q + "mergeCell", ref=ref)
            merges.set("count", str(len(merges)))
        for tag, attr in (("conditionalFormatting", "sqref"), ("dataValidation", "sqref"), ("hyperlink", "ref"), ("autoFilter", "ref")):
            for node in tree.iter(Q + tag):
                if node.get(attr):
                    node.set(attr, shift_sqref(node.get(attr), after, count))
        for selection in tree.iter(Q + "selection"):
            selection.attrib.pop("sqref", None)
            selection.attrib.pop("activeCell", None)
        dim = tree.find(Q + "dimension")
        if dim is not None and ":" in dim.get("ref", ""):
            a, b = dim.get("ref").split(":")
            m = CELL.match(b)
            if m:
                dim.set("ref", f"{a}:{m.group(1)}{int(m.group(2)) + count}")
        breaks = tree.find(Q + "rowBreaks")
        for brk in (breaks if breaks is not None else []):
            if int(brk.get("id", 0)) > after:
                brk.set("id", str(int(brk.get("id")) + count))
        self.shift_drawings(name, after, count)

    def shift_drawings(self, name, after, count):
        path = self.paths[name]
        rels_path = path.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if rels_path not in self.files:
            return
        rels = etree.fromstring(self.files[rels_path])
        for rel in rels:
            if not rel.get("Type", "").endswith("/drawing"):
                continue
            target = "xl/drawings/" + rel.get("Target").split("/")[-1]
            if target not in self.files:
                continue
            drawing = etree.fromstring(self.files[target])
            for anchor in list(drawing):
                for part in ("from", "to"):
                    node = anchor.find(XDR + part)
                    if node is None:
                        continue
                    row = node.find(XDR + "row")
                    # nomor baris di gambar mulai dari 0
                    if row is not None and int(row.text) + 1 > after:
                        row.text = str(int(row.text) + count)
            self.files[target] = etree.tostring(drawing, xml_declaration=True, encoding="UTF-8", standalone=True)


# Blok item tiap template: sheet, baris item pertama, baris item terakhir. Urutan penting: INPUT2 dulu, lalu sheet yang merujuknya.
ITEM_BLOCKS = {
    "PO": [("INPUT2", 3, 22), ("INV-", 13, 32), ("PL-", 36, 55), ("FINAL SI", 42, 61)],
    "DI": [("INPUT2", 3, 21), ("INV-", 13, 31), ("PL", 39, 57), ("FINAL SI", 41, 59)],
}


def capacity(profile_name):
    blocks = ITEM_BLOCKS.get(profile_name, [])
    return min(last - first + 1 for _, first, last in blocks) if blocks else 0


def add_item_rows(files, profile_name, items):
    """Tambah baris di semua blok item kalau jumlah item lebih banyak dari baris template."""
    extra = items - capacity(profile_name)
    if extra <= 0:
        return 0
    book = Book(files)
    for name, first, last in ITEM_BLOCKS[profile_name]:
        if name in book.paths:
            book.insert_rows(name, last, extra, first)
    book.save()
    return extra
