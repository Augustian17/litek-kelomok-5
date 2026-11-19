/* Dockit Web: versi browser dari Dockit (cek data DI dan dokumen CI/PL).
   Semua file diproses di browser. Tidak ada data yang dikirim ke server,
   kecuali pertanyaan ke Asisten AI bila pengguna memasukkan API key Claude sendiri. */
(() => {
"use strict";

/* ---------- Konfigurasi ---------- */
const VERSION = "1.0";
// MODE DEMO: login dicek di browser. Ini BUKAN pengamanan sungguhan (kode bisa dibaca siapa saja).
// Untuk produksi, ganti dengan layanan login seperti Firebase Authentication atau Supabase Auth.
const DEMO_USERS = [{ email: "demo@dockit.app", name: "Admin Shipping", hash: "0b453894bb8bc970de416f19c32700ddae45d0ffe687c83fd18bc111ab673e11" }];
const FIRST_YEAR = "2026";
const MAX_ITEMS = 19;
const SHEETS = { master: "MASTER DATA DI-柜", loading: "DI-柜", items: "Sheet1", weights: "MODEL -1 (2)", prices: "HARGA" };
const REQUIRED = { po: "PO", cont: "Kontainer", seal: "Seal", invDate: "Tgl invoice", size: "Size", etdmv: "ETD MV", pod: "POD", bl: "BL", ship: "Kapal" };
const AI_MODEL = "claude-sonnet-5-5";
const POD_COLORS = [["#DCE6F7", "#1E3E73"], ["#FCEBD6", "#7A400B"], ["#E3F1E7", "#1F5E37"], ["#F3E3F4", "#6B2A70"], ["#FDE4E2", "#8E1B13"], ["#E8E6FB", "#3A3391"]];
const DEFAULT_PROFILE = {
  shipper: "PT CONTOH EKSPOR INDONESIA\nJL. CONTOH INDUSTRI NO. 1, KAWASAN INDUSTRI CONTOH, KOTA BATAM",
  buyer: "SAMPLE BUYER (HONG KONG) LIMITED\nUNIT 1, SAMPLE TOWER, 1 EXAMPLE ROAD, KOWLOON",
  consignee: "SAMPLE CONSIGNEE LLC.\n100 EXAMPLE AVE, SAMPLE CITY, WA 00000",
  terms: "FOB SINGAPORE", loadPort: "SINGAPORE", payment: "60 days NET", origin: "MADE IN INDONESIA",
  bank: "[NAMA BANK]\nAccount: 000-000000-000",
};

/* ---------- Utilitas ---------- */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];
const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const blank = (v) => v === null || v === undefined || String(v).trim() === "";
const num = (v) => (typeof v === "number" && isFinite(v) ? v : (typeof v === "string" && v.trim() !== "" && isFinite(+v.replace(",", ".")) ? +v.replace(",", ".") : null));
const r2 = (v) => Math.round(v * 100) / 100;
const down2 = (v) => Math.floor(Math.round(v * 100 * 1e6) / 1e6) / 100; // CBM selalu dibulatkan ke bawah
const money = (v) => (v === null || v === undefined ? "" : Number(v).toLocaleString("id-ID", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
const fmt = (v, d = 2) => (v === null || v === undefined ? "–" : Number(v).toLocaleString("id-ID", { minimumFractionDigits: 0, maximumFractionDigits: d }));
const MONTHS = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
const MON = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
const toDate = (v) => {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === "number" && v > 20000 && v < 80000) { const d = new Date(Math.round((v - 25569) * 864e5)); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); }
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
    if (m) { const y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return new Date(y, +m[2] - 1, +m[1]); }
    const iso = v.trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return new Date(+iso[1], +iso[2] - 1, +iso[3]);
  }
  return null;
};
const dkey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const dShow = (d) => (d ? `${d.getDate()} ${MON[d.getMonth()]} ${d.getFullYear()}` : "–");
const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
const store = {
  get(k, s = localStorage) { try { return JSON.parse(s.getItem(k)); } catch { return null; } },
  set(k, v, s = localStorage) { try { s.setItem(k, JSON.stringify(v)); } catch { /* storage penuh atau diblokir */ } },
  del(k, s = localStorage) { try { s.removeItem(k); } catch { /* abaikan */ } },
};
async function sha256(text) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}
function toast(text) {
  $(".toast")?.remove();
  const el = document.createElement("div");
  el.className = "toast"; el.setAttribute("role", "status"); el.textContent = text;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 3600);
}

/* ---------- Ikon (Lucide, MIT) ---------- */
const I = (p, s = 19) => `<svg viewBox="0 0 24 24" width="${s}" height="${s}" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${p}</svg>`;
const ICON = {
  upload: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>',
  check: '<path d="m3 17 2 2 4-4"/><path d="m3 7 2 2 4-4"/><path d="M13 6h8"/><path d="M13 12h8"/><path d="M13 18h8"/>',
  file: '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M16 13H8"/><path d="M16 17H8"/>',
  cal: '<rect width="18" height="18" x="3" y="4" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/><path d="M3 10h18"/>',
  ai: '<path d="M9.94 15.5A2 2 0 0 0 8.5 14.06l-6.14-1.58a.5.5 0 0 1 0-.96L8.5 9.94A2 2 0 0 0 9.94 8.5l1.58-6.14a.5.5 0 0 1 .96 0l1.58 6.14a2 2 0 0 0 1.44 1.44l6.14 1.58a.5.5 0 0 1 0 .96l-6.14 1.58a2 2 0 0 0-1.44 1.44l-1.58 6.14a.5.5 0 0 1-.96 0z"/>',
  shield: '<path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z"/><path d="m9 12 2 2 4-4"/>',
  x: '<path d="M18 6 6 18"/><path d="m6 6 12 12"/>',
  arrow: '<path d="M5 12h14"/><path d="m12 5 7 7-7 7"/>',
  down: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>',
  send: '<path d="m22 2-7 20-4-9-9-4Z"/><path d="M22 2 11 13"/>',
  eye: '<path d="M2.06 12.35a1 1 0 0 1 0-.7 10.75 10.75 0 0 1 19.88 0 1 1 0 0 1 0 .7 10.75 10.75 0 0 1-19.88 0"/><circle cx="12" cy="12" r="3"/>',
  eyeOff: '<path d="M10.73 5.08A10.43 10.43 0 0 1 12 5c7 0 10 7 10 7a13.16 13.16 0 0 1-1.67 2.68"/><path d="M6.61 6.61A13.53 13.53 0 0 0 2 12s3 7 10 7a9.74 9.74 0 0 0 5.39-1.61"/><path d="m2 2 20 20"/><path d="M14.12 14.12a3 3 0 1 1-4.24-4.24"/>',
  alert: '<circle cx="12" cy="12" r="10"/><path d="M12 8v4"/><path d="M12 16h.01"/>',
  left: '<path d="m15 18-6-6 6-6"/>', right: '<path d="m9 18 6-6-6-6"/>',
  key: '<path d="m15.5 7.5 2.3 2.3a1 1 0 0 0 1.4 0l2.1-2.1a1 1 0 0 0 0-1.4L19 4"/><path d="m21 2-9.6 9.6"/><circle cx="7.5" cy="15.5" r="5.5"/>',
  printer: '<path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><path d="M6 9V3a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v6"/><rect x="6" y="14" width="12" height="8" rx="1"/>',
  sample: '<path d="M14 2v6a2 2 0 0 0 .245.96l5.51 10.08A2 2 0 0 1 18 22H6a2 2 0 0 1-1.755-2.96l5.51-10.08A2 2 0 0 0 10 8V2"/><path d="M6.453 15h11.094"/><path d="M8.5 2h7"/>',
};

/* ---------- State ---------- */
const S = {
  user: store.get("dockit.session", sessionStorage) || store.get("dockit.session"),
  files: [],            // {id, name, status, note, kind[], data}
  data: null,           // hasil gabungan semua file
  filter: "all",
  selected: new Set(),
  docTab: "ci",
  preview: null,
  cal: null,
  calPick: null,
  profile: Object.assign({}, DEFAULT_PROFILE, store.get("dockit.profile") || {}),
  chat: [],
  apiKey: store.get("dockit.key", sessionStorage) || "",
  busy: false,
  introPlayed: false,
};

/* ---------- Membaca file Excel / ZIP ---------- */
const sheetRows = (ws, firstRow) => XLSX.utils.sheet_to_json(ws, { header: "A", range: firstRow - 1, defval: null, raw: true, blankrows: false });
const cleanStr = (v) => (v instanceof Date ? v : blank(v) ? null : typeof v === "string" ? v.trim() : v);

function parseWorkbook(buf, name) {
  let wb;
  try { wb = XLSX.read(buf, { type: "array", cellDates: false }); }
  catch (e) {
    const locked = /password|encrypt/i.test(String(e && e.message));
    return { status: "err", note: locked ? "File terkunci password. Buka kunci di Excel dulu." : "File tidak bisa dibaca sebagai Excel." };
  }
  const out = { master: [], loading: [], items: [], weights: [], prices: [] };
  const kinds = [];
  const ws = (n) => wb.Sheets[n];
  if (ws(SHEETS.master)) {
    kinds.push("master");
    sheetRows(ws(SHEETS.master), 3).forEach((r, i) => {
      if (blank(r.A)) return;
      out.master.push({
        row: i + 3, file: name, key: String(r.A).trim(), amazon: cleanStr(r.B), po: cleanStr(r.C), rolling: cleanStr(r.D), invoice: cleanStr(r.E),
        code: cleanStr(r.F), cont: cleanStr(r.G), seal: cleanStr(r.H), invDate: toDate(r.I), size: cleanStr(r.J), etd: toDate(r.K), etdmv: toDate(r.L),
        pod: cleanStr(r.M), bl: cleanStr(r.N), ship: cleanStr(r.O), notes: cleanStr(r.P), qty: num(r.Q), nw: num(r.S), gw: num(r.T), ctns: num(r.U),
        cbm: num(r.V), amount: num(r.X),
      });
    });
  }
  if (ws(SHEETS.loading)) {
    kinds.push("loading");
    sheetRows(ws(SHEETS.loading), 2).forEach((r, i) => {
      if (blank(r.B) || blank(r.C)) return;
      out.loading.push({ row: i + 2, di: String(r.B).trim(), sku: String(r.C).trim(), qty: num(r.I), po: cleanStr(r.N), asin: cleanStr(r.O), cbm: num(r.AF) });
    });
  }
  if (ws(SHEETS.items) && ws(SHEETS.weights)) {
    kinds.push("items");
    sheetRows(ws(SHEETS.items), 2).forEach((r) => { if (!blank(r.A)) out.items.push({ sku: String(r.A).trim(), name: cleanStr(r.B), hs: cleanStr(r.C) }); });
    sheetRows(ws(SHEETS.weights), 2).forEach((r) => { if (!blank(r.A)) out.weights.push({ sku: String(r.A).trim(), cbm: num(r.D), nw: num(r.E), gw: num(r.F) }); });
    if (ws(SHEETS.prices)) sheetRows(ws(SHEETS.prices), 2).forEach((r) => { if (!blank(r.G) && !blank(r.H)) out.prices.push({ di: String(r.G).trim(), sku: String(r.H).trim(), qty: num(r.D), price: num(r.E), service: num(r.F) }); });
  }
  if (!kinds.length) return { status: "skip", note: "Dilewati: bukan file master, loading, atau data barang." };
  return { status: "ok", kinds, data: out };
}

async function addFiles(list) {
  if (!window.XLSX || !window.JSZip) { toast("Library Excel belum termuat. Muat ulang halaman."); return; }
  const queue = [];
  for (const f of list) {
    if (/\.zip$/i.test(f.name)) {
      try {
        const zip = await JSZip.loadAsync(await f.arrayBuffer());
        const entries = Object.values(zip.files).filter((e) => !e.dir && /\.xls[xm]?$/i.test(e.name) && !/(^|\/)(~\$|__MACOSX)/.test(e.name));
        if (!entries.length) { S.files.push({ id: crypto.randomUUID(), name: f.name, status: "skip", note: "ZIP tidak berisi file Excel." }); continue; }
        for (const e of entries) queue.push({ name: e.name.split("/").pop(), from: f.name, buf: await e.async("arraybuffer") });
      } catch { S.files.push({ id: crypto.randomUUID(), name: f.name, status: "err", note: "ZIP gagal dibuka." }); }
    } else if (/\.xls[xm]?$/i.test(f.name)) {
      queue.push({ name: f.name, buf: await f.arrayBuffer() });
    } else {
      S.files.push({ id: crypto.randomUUID(), name: f.name, status: "skip", note: "Format tidak didukung. Pakai .xlsx, .xls, atau .zip." });
    }
  }
  for (const q of queue) {
    const res = parseWorkbook(q.buf, q.name);
    S.files.push({ id: crypto.randomUUID(), name: q.name, from: q.from, ...res });
  }
  rebuild();
}

/* ---------- Gabungkan data & cek ---------- */
const ukey = (v) => (blank(v) ? "" : String(v).trim().toUpperCase());
const isOld = (m) => { const y = String(m.rolling || "").slice(0, 4); return /^\d{4}$/.test(y) && y < FIRST_YEAR; };

function rebuild() {
  const all = { master: [], loading: [], items: [], weights: [], prices: [] };
  S.files.filter((f) => f.status !== "err" && f.data).forEach((f) => Object.keys(all).forEach((k) => all[k].push(...f.data[k])));
  const items = new Map(all.items.map((x) => [x.sku, x]));
  const weights = new Map(all.weights.map((x) => [x.sku, x]));
  const prices = new Map(all.prices.map((x) => [x.di + "|" + x.sku, x]));
  const loading = new Map();
  all.loading.forEach((x) => { if (!loading.has(x.di)) loading.set(x.di, []); loading.get(x.di).push(x); });
  const seen = new Map();
  const rows = []; let old = 0;
  all.master.forEach((m) => { if (isOld(m)) { old++; return; } if (!seen.has(m.key)) { seen.set(m.key, m); rows.push(m); } });

  const linesFor = (key) => {
    const src = loading.get(key) || all.prices.filter((p) => p.di === key).map((p) => ({ sku: p.sku, qty: p.qty }));
    return src.map((l) => {
      const it = items.get(l.sku) || {}, w = weights.get(l.sku) || {}, p = prices.get(key + "|" + l.sku) || {};
      return { sku: l.sku, qty: l.qty ?? p.qty ?? null, po: l.po ?? null, asin: l.asin ?? null, name: it.name ?? null, hs: it.hs ?? null,
        nw: w.nw ?? null, gw: w.gw ?? null, cbm: w.cbm ?? l.cbm ?? null, price: p.price ?? null };
    });
  };
  const totals = (lines) => {
    const t = { qty: 0, ctns: 0, nw: 0, gw: 0, cbm: 0, amount: 0 };
    lines.forEach((l) => { const q = l.qty || 0; t.qty += q; t.ctns += q; t.nw += r2((l.nw || 0) * q); t.gw += r2((l.gw || 0) * q); t.cbm += down2((l.cbm || 0) * q); t.amount += r2((l.price || 0) * q); });
    Object.keys(t).forEach((k) => (t[k] = r2(t[k])));
    return t;
  };

  rows.forEach((m) => { m.lines = linesFor(m.key); m.totals = totals(m.lines); m.issues = []; });
  const issue = (m, level, text, hint) => m.issues.push({ level, text, hint });

  // Duplikat: hanya Rolling, Kontainer, dan Seal yang tidak boleh sama
  [["rolling", "Rolling"], ["cont", "Nomor kontainer"], ["seal", "Nomor seal"]].forEach(([k, label]) => {
    const groups = new Map();
    rows.forEach((m) => { const v = ukey(m[k]); if (v) { if (!groups.has(v)) groups.set(v, []); groups.get(v).push(m); } });
    groups.forEach((g) => { if (g.length > 1) g.forEach((m) => issue(m, "err", `${label} ${m[k]} juga dipakai di ${g.filter((o) => o !== m).map((o) => o.bl || o.key).join(", ")}`, `Cek ${label.toLowerCase()} di file sumber`)); });
  });
  const hasLoading = all.loading.length > 0, hasItems = all.items.length > 0;
  rows.forEach((m) => {
    const miss = Object.entries(REQUIRED).filter(([k]) => blank(m[k]) || (m[k] instanceof Date && isNaN(m[k]))).map(([, l]) => l);
    if (miss.length) issue(m, "err", `${miss.join(", ")} masih kosong`, "Lengkapi di master data");
    if (!m.lines.length) issue(m, "err", hasLoading || all.prices.length ? "Tidak ada baris barang untuk DI ini" : "Belum ada file loading atau data barang", hasLoading ? "Cek sheet DI-柜 / HARGA" : "Tambahkan file loading dan file data barang");
    if (m.lines.length > MAX_ITEMS) issue(m, "err", `${m.lines.length} baris barang, template hanya muat ${MAX_ITEMS}`, "Pecah menjadi dua DI");
    if (m.lines.length && hasItems) {
      const noW = m.lines.filter((l) => l.nw === null || l.gw === null).map((l) => l.sku);
      if (noW.length) issue(m, "warn", `NW/GW belum ada untuk ${noW.join(", ")}`, "Update dari file timbang");
      const noHs = m.lines.filter((l) => blank(l.hs)).map((l) => l.sku);
      if (noHs.length) issue(m, "warn", `HS Code kosong untuk ${noHs.join(", ")}`, "Lengkapi di Sheet1");
      const noPrice = m.lines.filter((l) => l.price === null).map((l) => l.sku);
      if (noPrice.length && all.prices.length) issue(m, "warn", `Harga belum ada untuk ${noPrice.join(", ")}`, "Lengkapi di tab HARGA");
      if (m.cbm !== null && m.totals.cbm && Math.abs(m.cbm - m.totals.cbm) > 0.005) issue(m, "warn", `CBM tidak tally: master ${fmt(m.cbm)}, hitungan ${fmt(m.totals.cbm)}`, "Samakan CBM di master atau data barang");
      if (m.qty !== null && m.totals.qty && m.qty !== m.totals.qty) issue(m, "warn", `Qty tidak tally: master ${fmt(m.qty, 0)}, loading ${fmt(m.totals.qty, 0)}`, "Cek qty di file loading");
    }
    if (m.lines.length && m.lines.some((l) => blank(l.asin) || blank(l.po)) && hasLoading) issue(m, "warn", "ASIN atau PO barang ada yang kosong", "Lengkapi di file loading");
  });
  // Deteksi anomali: kepadatan berat (GW/CBM) jauh dari kontainer lain dengan ukuran sama
  const bySize = new Map();
  rows.forEach((m) => { const gw = m.gw ?? m.totals.gw, cbm = m.cbm ?? m.totals.cbm; if (gw && cbm) { const s = ukey(m.size) || "?"; if (!bySize.has(s)) bySize.set(s, []); bySize.get(s).push({ m, d: gw / cbm, gw }); } });
  bySize.forEach((list) => {
    if (list.length < 3) return;
    const sorted = list.map((x) => x.d).sort((a, b) => a - b), med = sorted[Math.floor(sorted.length / 2)];
    const gws = list.map((x) => x.gw).sort((a, b) => a - b), gmed = gws[Math.floor(gws.length / 2)];
    list.forEach(({ m, d, gw }) => { if (Math.abs(d - med) / med > 0.35) issue(m, "ai", `GW ${fmt(gw)} kg tidak wajar untuk ${m.size || "kontainer"} ini (median kontainer sejenis ±${fmt(gmed, 0)} kg)`, "Pastikan qty atau berat per karton"); });
  });
  // Berangkat selalu H-1 dari ETD
  rows.forEach((m) => { m.depart = m.etd ? addDays(m.etd, -1) : null; m.level = m.issues.some((i) => i.level === "err") ? "err" : m.issues.some((i) => i.level === "warn") ? "warn" : m.issues.length ? "ai" : "ok"; });
  rows.sort((a, b) => (a.etd || 0) - (b.etd || 0) || String(a.bl).localeCompare(String(b.bl)));

  S.data = { rows, old, counts: { items: all.items.length, loading: all.loading.length } };
  [...S.selected].forEach((k) => { const m = rows.find((r) => r.key === k); if (!m || m.level === "err") S.selected.delete(k); });
  if (!S.selected.size) rows.filter((m) => m.level !== "err").slice(0, 4).forEach((m) => S.selected.add(m.key));
  render();
}

const fileStats = (f) => {
  if (!f.data) return null;
  const keys = f.data.master.filter((m) => !isOld(m));
  return { di: keys.length, cont: new Set(keys.map((m) => ukey(m.cont)).filter(Boolean)).size, loading: f.data.loading.length, items: f.data.items.length };
};

/* ---------- Login ---------- */
function loginScene() {
  const wave = (y, a, p = 160, w = 3200) => `M0 ${y} q${p / 4} ${-a} ${p / 2} 0 ` + Array.from({ length: Math.ceil(w / (p / 2)) }, () => `t${p / 2} 0`).join(" ");
  const fill = (y, a, p) => wave(y, a, p) + " V900 H0 Z";
  const boxes = [[66, 256, "#F2C879", "#0A1F40", ".18"], [176, 256, "#EADBC8", "#0A1F40", ".18"], [286, 256, "#3B5F9E", "#fff", ".22"], [396, 256, "#FBF6EF", "#0A1F40", ".14"],
    [121, 206, "#3B5F9E", "#fff", ".22"], [231, 206, "#F2C879", "#0A1F40", ".18"], [341, 206, "#EADBC8", "#0A1F40", ".18"], [176, 156, "#FBF6EF", "#0A1F40", ".14"], [286, 156, "#F2C879", "#0A1F40", ".18"]]
    .map(([x, y, f, s, o], i) => `<g class="box" style="animation-delay:${600 + i * 80}ms"><rect x="${x}" y="${y}" width="104" height="44" rx="3" fill="${f}"/><path d="${[1, 2, 3, 4, 5].map((k) => `M${x + 18 * k} ${y + 6}v32`).join("")}" stroke="${s}" stroke-opacity="${o}" stroke-width="2"/></g>`).join("");
  return `<svg class="scene" viewBox="0 0 1440 900" preserveAspectRatio="xMidYMax slice" aria-hidden="true">
    <circle cx="1060" cy="650" r="130" fill="#F2C879" opacity=".92"/><circle cx="1060" cy="650" r="210" fill="#F2C879" opacity=".10"/>
    <rect x="0" y="640" width="1440" height="260" fill="#0A1F40"/>
    <g class="shimmer"><path d="M960 668h200M985 684h150M1005 700h110M1025 716h70" stroke="#F2C879" stroke-opacity=".55" stroke-width="3" stroke-linecap="round"/></g>
    <g class="waveA"><path d="${fill(648, 5, 160)}" fill="#11284D"/></g>
    <g transform="translate(0 292)"><g class="sail"><g class="bob">
      <path d="M40 300 H600 L560 372 H92 Z" fill="#FBF6EF"/><rect x="40" y="300" width="560" height="12" fill="#F2C879"/>
      <rect x="514" y="214" width="58" height="86" rx="4" fill="#EADBC8"/><rect x="524" y="226" width="38" height="12" rx="2" fill="#0F2C59"/><rect x="536" y="188" width="12" height="26" fill="#EADBC8"/>
      ${boxes}</g></g></g>
    <g class="waveB"><path d="${fill(656, 6, 200)}" fill="#0A1F40"/></g>
    <g class="waveC"><path d="${wave(690, 4, 240)}" fill="none" stroke="#3B5F9E" stroke-opacity=".5" stroke-width="2"/></g></svg>`;
}

function viewLogin(error = "", email = "") {
  const intro = !S.introPlayed; S.introPlayed = true;
  return `<main class="login ${intro ? "intro" : ""}">
    ${loginScene()}
    <div class="brand reveal" style="animation-delay:1500ms">
      <div class="brand-row"><img src="assets/logo.png" alt="">Dockit</div>
      <div class="tag"><p>Dokumen ekspor beres sebelum kapal berangkat.</p><p>CI/PL, pengecekan data, dan jadwal ETD dalam satu tempat untuk tim shipping.</p></div>
    </div>
    <section class="glass reveal" style="animation-delay:1700ms" aria-labelledby="lg-t">
      <h1 id="lg-t">Masuk ke Dockit</h1>
      <p class="lead">Khusus karyawan dan admin shipping. Akun dibuat oleh admin.</p>
      <form id="login" novalidate>
        ${error ? `<div class="alert" role="alert">${I(ICON.alert, 18)}<span>${esc(error)}</span></div>` : ""}
        <div class="f"><label class="lbl" for="email">Email kantor</label>
          <input class="field" id="email" name="email" type="email" autocomplete="username" placeholder="nama@perusahaan.co.id" value="${esc(email)}" ${error ? 'aria-invalid="true"' : ""} required></div>
        <div class="f"><div style="display:flex;justify-content:space-between;align-items:baseline"><label class="lbl" for="pw">Password</label><a href="#" data-act="forgot" style="font-size:13.5px;font-weight:600">Lupa password?</a></div>
          <div class="pw"><input class="field" id="pw" name="pw" type="password" autocomplete="current-password" ${error ? 'aria-invalid="true"' : ""} required>
          <button type="button" class="eye" data-act="eye" aria-label="Tampilkan password">${I(ICON.eye, 20)}</button></div></div>
        <label class="check"><input type="checkbox" name="remember">Ingat saya di perangkat ini</label>
        <button class="go" type="submit">Masuk</button>
      </form>
      <p class="demo">Mode demo: masuk dengan <code>demo@dockit.app</code> / <code>Dockit2026!</code></p>
      <p class="foot">Belum punya akun? Hubungi admin shipping untuk dibuatkan akses.</p>
    </section></main>`;
}

async function doLogin(form) {
  const email = form.email.value.trim().toLowerCase(), pw = form.pw.value;
  const btn = $(".go", form); btn.disabled = true; btn.textContent = "Memeriksa…";
  const hash = await sha256(pw);
  const user = DEMO_USERS.find((u) => u.email === email && u.hash === hash);
  if (!user) {
    $("#app").innerHTML = viewLogin(!email || !pw ? "Isi email kantor dan password dulu." : "Email atau password salah. Coba lagi, atau klik \"Lupa password?\" untuk minta reset ke admin.", email);
    $(".glass").classList.add("shake"); $("#pw").focus(); return;
  }
  S.user = { email: user.email, name: user.name };
  store.set("dockit.session", S.user, sessionStorage);
  if (form.remember.checked) store.set("dockit.session", S.user); else store.del("dockit.session");
  location.hash = "#/file"; render();
}

/* ---------- Shell ---------- */
const ROUTES = { file: "Tambah File", cek: "Cek Data", dokumen: "Dokumen CI/PL", kalender: "Kalender ETD" };
function shell(route, body) {
  const errs = S.data ? S.data.rows.filter((m) => m.level === "err").length : 0;
  const link = (r, icon, label, extra = "") => `<a href="#/${r}" ${route === r ? 'aria-current="page"' : ""}>${I(icon)}${label}${extra}</a>`;
  return `<div class="shell">
    <aside class="side">
      <div class="logo"><img src="assets/logo.png" alt=""><div><b>Dockit</b><small>Versi Web ${VERSION}</small></div></div>
      <nav class="nav" aria-label="Menu utama">
        ${link("file", ICON.upload, "Tambah File")}
        ${link("cek", ICON.check, "Cek Data", errs ? `<span class="badge num" aria-label="${errs} error">${errs}</span>` : "")}
        ${link("dokumen", ICON.file, "Dokumen CI/PL")}
        ${link("kalender", ICON.cal, "Kalender ETD")}
        <a href="#/cek?ai=1">${I(ICON.ai)}Asisten AI</a>
      </nav>
      <div class="note">${I(`<g stroke="#F2C879">${ICON.shield}</g>`, 18)}<span>File diproses di browser kamu. Tidak ada data yang dikirim ke server.</span></div>
      <div class="user"><span title="${esc(S.user.email)}">${esc(S.user.name)}</span><button class="out" data-act="logout">Keluar</button></div>
    </aside>
    <main class="main" id="main">${body}</main></div>`;
}

/* ---------- Halaman: Tambah File ---------- */
function viewFiles() {
  const rows = S.files.map((f) => {
    const st = fileStats(f);
    const pill = f.status === "err" ? `<span class="pill err">Error</span>` : f.status === "skip" ? `<span class="pill skip">Dilewati</span>` : `<span class="pill ok">Siap</span>`;
    const kind = f.kinds ? f.kinds.map((k) => ({ master: "Master DI", loading: "Loading", items: "Data barang" }[k])).join(" + ") : "–";
    return `<tr><td><strong>${esc(f.name)}</strong>${f.from ? `<br><small>dari ${esc(f.from)}</small>` : ""}${f.note ? `<br><small>${esc(f.note)}</small>` : ""}</td>
      <td>${kind}</td><td class="num">${st && st.di ? st.di : "–"}</td><td class="num">${st && st.cont ? st.cont : "–"}</td><td>${pill}</td>
      <td style="text-align:right"><button class="icon-btn" data-act="rm-file" data-id="${f.id}" aria-label="Hapus ${esc(f.name)}">${I(ICON.x, 18)}</button></td></tr>`;
  }).join("");
  const d = S.data, active = d ? d.rows.length : 0;
  const conts = d ? new Set(d.rows.map((m) => ukey(m.cont)).filter(Boolean)).size : 0;
  return `<header class="head"><div><h1>Tambah file DI</h1><p>Masukkan file master DI, file loading, dan file data barang (Excel atau ZIP).</p></div>
      <a class="btn" href="#/cek" ${active ? "" : 'aria-disabled="true" tabindex="-1"'}>Cek data ${I(ICON.arrow, 18)}</a></header>
    <label class="drop" id="drop">
      <span class="ic">${I(`<g stroke="#0F2C59">${ICON.upload}</g>`, 26)}</span>
      <b>Tarik file ke sini, atau klik untuk memilih</b><span>.xlsx, .xls atau .zip · bisa banyak file sekaligus</span>
      <input id="file-in" type="file" multiple accept=".xlsx,.xls,.xlsm,.zip">
    </label>
    <section class="card">
      <div class="card-h"><h2>${S.files.length ? `${S.files.length} file dimuat` : "Belum ada file"}</h2>
        <p class="num">${active ? `${active} DI · ${conts} kontainer · ${FIRST_YEAR}+` : ""}${d && d.old ? ` · ${d.old} DI lama dilewati` : ""}</p></div>
      ${S.files.length ? `<div class="scroll"><table class="tbl"><thead><tr><th>File</th><th>Isi</th><th>DI</th><th>Kontainer</th><th>Status</th><th><span class="sr">Aksi</span></th></tr></thead><tbody>${rows}</tbody></table></div>`
        : `<div class="empty"><b>Tambahkan file Excel untuk mulai.</b>Belum punya file? <button class="btn soft" data-act="sample" style="margin-left:8px;height:38px">${I(ICON.sample, 17)}Coba dengan data contoh</button></div>`}
    </section>
    <p class="hint">Dockit mengenali file dari nama sheet: <b>${esc(SHEETS.master)}</b> (master), <b>${esc(SHEETS.loading)}</b> (loading), serta <b>Sheet1</b> + <b>${esc(SHEETS.weights)}</b> + <b>HARGA</b> (data barang). Data sebelum ${FIRST_YEAR} dilewati otomatis.</p>`;
}

/* ---------- Halaman: Cek Data ---------- */
function viewCheck(focusAi) {
  const d = S.data;
  if (!d || !d.rows.length) return `<header class="head"><div><h1>Cek data</h1><p>Belum ada data DI untuk dicek.</p></div></header>
    <section class="card"><div class="empty"><b>Tambahkan file master DI dulu.</b><a class="btn" href="#/file" style="margin-top:12px">${I(ICON.upload, 18)}Tambah file</a></div></section>${aiPanel(focusAi, true)}`;
  const list = d.rows.flatMap((m) => m.issues.map((i) => ({ m, ...i })));
  const c = { err: 0, warn: 0, ai: 0 }; list.forEach((x) => c[x.level]++);
  const shown = S.filter === "all" ? list : list.filter((x) => x.level === S.filter);
  const label = { err: "Error", warn: "Peringatan", ai: "Anomali" };
  const chip = (k, cls, text, n) => `<button class="chip ${cls}" data-act="filter" data-f="${k}" aria-pressed="${S.filter === k}">${text} <span class="num">${n}</span></button>`;
  const okCount = d.rows.filter((m) => m.level === "ok").length;
  return `<div class="row">
    <div style="flex:1 1 560px;min-width:0;display:flex;flex-direction:column;gap:20px">
      <header class="head"><div><h1>Cek data</h1><p>Hasil pengecekan ${d.rows.length} DI sebelum dokumen dibuat. ${okCount} DI tanpa catatan.</p></div></header>
      <div class="chips" role="group" aria-label="Filter hasil">${chip("all", "all", "Semua", list.length)}${chip("err", "err", "Error", c.err)}${chip("warn", "warn", "Peringatan", c.warn)}${chip("ai", "ai", "Anomali", c.ai)}</div>
      <section class="card">${shown.length ? `<div class="scroll"><table class="tbl"><thead><tr><th>Jenis</th><th>BL / DI</th><th>Masalah</th><th>Saran</th></tr></thead><tbody>
        ${shown.map((x) => `<tr><td><span class="pill ${x.level}">${label[x.level]}</span></td><td class="num"><strong>${esc(x.m.bl || "–")}</strong><br><small>${esc(x.m.key)} · ${esc(x.m.file)} baris ${x.m.row}</small></td><td>${esc(x.text)}</td><td style="color:var(--ink-2)">${esc(x.hint || "")}</td></tr>`).join("")}
        </tbody></table></div>` : `<div class="empty"><b>${list.length ? "Tidak ada catatan di filter ini." : "Semua data bersih."}</b>${list.length ? "" : "Lanjut ke Dokumen CI/PL."}</div>`}</section>
      <p class="hint">Aturan: Rolling, Kontainer, dan Seal tidak boleh dobel. CBM dibulatkan ke bawah. Berangkat dihitung H-1 dari ETD. Anomali dihitung dari kepadatan berat (GW/CBM) dibanding kontainer seukuran.</p>
    </div>${aiPanel(focusAi)}</div>`;
}

/* ---------- Asisten AI ---------- */
function aiPanel(focus, standalone) {
  const log = S.chat.length ? S.chat.map((m) => `<div class="msg ${m.role === "user" ? "me" : m.role === "sys" ? "sys" : "bot"}">${esc(m.text)}</div>`).join("")
    : `<div class="msg bot">Halo! Aku bisa merangkum data DI yang sudah dimuat. Coba tombol di bawah, atau masukkan API key Claude untuk bertanya bebas.</div>`;
  return `<aside class="aip" aria-label="Asisten AI" ${standalone ? 'style="position:static;margin-top:24px"' : ""}>
    <div class="ai-h"><span class="mark">${I(`<g stroke="#F2C879">${ICON.ai}</g>`, 18)}</span><div><h2>Asisten AI</h2><p>${S.apiKey ? "Terhubung ke Claude" : "Ringkasan lokal · Claude opsional"}</p></div>
      <button class="icon-btn" data-act="key" aria-label="Atur API key Claude" aria-expanded="false">${I(ICON.key, 18)}</button></div>
    <div class="keybox" id="keybox" hidden>
      <label for="key-in" style="font-weight:700">API key Claude</label>
      <div class="rowk"><input class="input" id="key-in" type="password" placeholder="sk-ant-…" value="${esc(S.apiKey)}" autocomplete="off"><button class="btn" data-act="save-key" style="height:40px;padding:0 14px">Simpan</button></div>
      <p>Key disimpan hanya di tab ini dan hilang saat tab ditutup. Pertanyaan beserta ringkasan data DI akan dikirim ke Anthropic.</p>
    </div>
    <div class="ai-log" id="ai-log" aria-live="polite">${log}</div>
    <div class="ai-chips">
      <button data-act="ask" data-q="revisi">BL yang perlu direvisi</button>
      <button data-act="ask" data-q="cbm">Total CBM per bulan</button>
      <button data-act="ask" data-q="pod">Kontainer per POD</button>
    </div>
    <form id="ai-form"><label for="ask" class="sr">Tanya asisten</label>
      <input class="input" id="ask" style="flex:1" placeholder="Tanya soal data pengiriman…" autocomplete="off" ${focus ? "autofocus" : ""}>
      <button class="send" type="submit" aria-label="Kirim pertanyaan" ${S.busy ? "disabled" : ""}>${I(ICON.send, 18)}</button></form>
  </aside>`;
}

function localAnswer(q) {
  const d = S.data;
  if (!d || !d.rows.length) return "Belum ada data DI. Tambahkan file master DI dulu di menu Tambah File.";
  if (q === "revisi") {
    const bad = d.rows.filter((m) => m.level === "err" || m.level === "warn");
    if (!bad.length) return "Tidak ada BL yang perlu direvisi. Semua data lolos pengecekan.";
    return `Ada ${bad.length} BL yang perlu dicek:\n` + bad.map((m, i) => `${i + 1}. ${m.bl || m.key} (ETD ${dShow(m.etd)}): ${m.issues.filter((x) => x.level !== "ai").map((x) => x.text).join("; ")}`).join("\n");
  }
  if (q === "cbm") {
    const by = new Map();
    d.rows.forEach((m) => { if (!m.etd) return; const k = `${MONTHS[m.etd.getMonth()]} ${m.etd.getFullYear()}`; by.set(k, r2((by.get(k) || 0) + (m.cbm ?? m.totals.cbm ?? 0))); });
    return "Total CBM berdasarkan bulan ETD:\n" + [...by].map(([k, v]) => `• ${k}: ${fmt(v)} CBM`).join("\n");
  }
  if (q === "pod") {
    const by = new Map();
    d.rows.forEach((m) => { const k = m.pod || "(POD kosong)"; if (!by.has(k)) by.set(k, new Set()); if (m.cont) by.get(k).add(ukey(m.cont)); });
    return "Jumlah kontainer per POD:\n" + [...by].sort((a, b) => b[1].size - a[1].size).map(([k, v]) => `• ${k}: ${v.size} kontainer`).join("\n");
  }
  return "";
}

function dataContext() {
  const d = S.data;
  if (!d) return "Belum ada data.";
  return JSON.stringify(d.rows.map((m) => ({ di: m.key, bl: m.bl, po: m.po, kontainer: m.cont, seal: m.seal, size: m.size, berangkat: m.depart && dkey(m.depart), etd: m.etd && dkey(m.etd), pod: m.pod, kapal: m.ship,
    qty: m.qty ?? m.totals.qty, gw: m.gw ?? m.totals.gw, cbm: m.cbm ?? m.totals.cbm, catatan: m.issues.map((i) => `${i.level}: ${i.text}`) })));
}

async function askAI(text) {
  S.chat.push({ role: "user", text });
  if (!S.apiKey) {
    S.chat.push({ role: "bot", text: "Untuk pertanyaan bebas, masukkan API key Claude lewat ikon kunci di atas. Tanpa key, pakai tombol ringkasan di bawah." });
    return refreshChat();
  }
  S.busy = true; S.chat.push({ role: "sys", text: "Claude sedang mengetik…" }); refreshChat();
  const history = S.chat.filter((m) => m.role === "user" || m.role === "bot").slice(-10);
  const messages = history.map((m) => ({ role: m.role === "user" ? "user" : "assistant", content: m.text }));
  while (messages.length && messages[0].role !== "user") messages.shift();
  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "content-type": "application/json", "x-api-key": S.apiKey, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
      body: JSON.stringify({ model: AI_MODEL, max_tokens: 1024,
        system: `Kamu Asisten AI di Dockit, aplikasi tim shipping untuk dokumen ekspor CI/PL. Jawab dalam Bahasa Indonesia, singkat dan jelas, pakai daftar bila membantu. Jawab hanya dari data berikut; kalau datanya tidak ada, katakan terus terang. Aturan: kontainer berangkat H-1 dari ETD, CBM dibulatkan ke bawah, Rolling/Kontainer/Seal tidak boleh dobel. Hari ini ${dkey(new Date())}.\n\nDATA DI (JSON):\n${dataContext()}`,
        messages }),
    });
    const body = await res.json().catch(() => ({}));
    S.chat.pop();
    if (!res.ok) {
      const msg = res.status === 401 ? "API key tidak valid. Cek lagi key-nya." : res.status === 429 ? "Batas pemakaian API tercapai. Coba lagi sebentar." : `Claude tidak bisa dihubungi (${res.status}). ${body?.error?.message || ""}`;
      S.chat.push({ role: "sys", text: msg });
    } else {
      S.chat.push({ role: "bot", text: (body.content || []).filter((c) => c.type === "text").map((c) => c.text).join("\n").trim() || "(Tidak ada jawaban.)" });
    }
  } catch {
    S.chat.pop(); S.chat.push({ role: "sys", text: "Gagal terhubung ke Claude. Cek koneksi internet." });
  }
  S.busy = false; refreshChat();
}
function refreshChat() {
  const panel = $(".aip"); if (!panel) return;
  const html = aiPanel(false, panel.style.position === "static");
  const tmp = document.createElement("div"); tmp.innerHTML = html;
  $("#ai-log").innerHTML = $("#ai-log", tmp).innerHTML;
  $(".send").disabled = S.busy;
  $(".ai-h p").textContent = S.apiKey ? "Terhubung ke Claude" : "Ringkasan lokal · Claude opsional";
  const log = $("#ai-log"); log.scrollTop = log.scrollHeight;
}

/* ---------- Halaman: Dokumen ---------- */
const lines2 = (t) => esc(t).replace(/\n/g, "<br>");
function docCI(m) {
  const p = S.profile;
  const body = m.lines.map((l) => `<tr><td>${esc(l.sku)}</td><td>${esc(l.asin || "")}</td><td>${esc(l.po || m.po || "")}</td><td>${esc(l.name || "")}</td><td class="r">${fmt(l.qty, 0)}</td><td class="r">${l.price !== null ? money(l.price) : ""}</td><td class="r">${l.price !== null ? money(r2(l.price * (l.qty || 0))) : ""}</td></tr>`).join("");
  return `<article class="paper"><div class="ttl"><b>${lines2(p.shipper.split("\n")[0])}</b><span>${lines2(p.shipper.split("\n").slice(1).join("\n"))}</span><i>COMMERCIAL INVOICE</i></div>
    <div class="two"><div><strong>To:</strong><br>${lines2(p.buyer)}</div><div class="num"><strong>Invoice No.:</strong> ${esc(m.invoice || m.key)}<br><strong>Date:</strong> ${dShow(m.invDate)}<br><strong>PO:</strong> ${esc(m.po || "")}</div>
      <div><strong>Ship to:</strong><br>${lines2(p.consignee)}</div><div class="num"><strong>Container / Seal:</strong> ${esc(m.cont || "–")} / ${esc(m.seal || "–")}<br><strong>Vessel:</strong> ${esc(m.ship || "–")}<br><strong>BL:</strong> ${esc(m.bl || "–")}</div></div>
    <table class="num"><thead><tr><th>Item No.</th><th>ASIN</th><th>PO</th><th>Goods Description</th><th class="r">Qty (PCS)</th><th class="r">Unit Price (USD)</th><th class="r">Amount (USD)</th></tr></thead>
      <tbody>${body}<tr class="tot"><td colspan="4">TOTAL</td><td class="r">${fmt(m.totals.qty, 0)}</td><td></td><td class="r">${money(m.totals.amount)}</td></tr></tbody></table>
    <div class="two" style="margin-top:18px"><div><strong>Port of Loading:</strong> ${esc(p.loadPort)}<br><strong>Port of Discharge:</strong> ${esc(m.pod || "–")}<br><strong>Terms of Delivery:</strong> ${esc(p.terms)}</div>
      <div><strong>Payment Term:</strong> ${esc(p.payment)}<br><strong>Country of origin:</strong> ${esc(p.origin)}<br><strong>HS Code:</strong> ${esc([...new Set(m.lines.map((l) => l.hs).filter(Boolean))].join(" / ") || "–")}</div>
      <div><strong>Payment Instruction:</strong><br>${lines2(p.bank)}</div></div>
    <div class="sign"><div>Authorized Signature</div></div></article>`;
}
function docPL(m) {
  const p = S.profile;
  const body = m.lines.map((l) => { const q = l.qty || 0; return `<tr><td>${esc(l.sku)}</td><td>${esc(l.name || "")}</td><td class="r">${fmt(q, 0)}</td><td class="r">${fmt(q, 0)}</td><td class="r">${l.nw !== null ? fmt(r2(l.nw * q)) : ""}</td><td class="r">${l.gw !== null ? fmt(r2(l.gw * q)) : ""}</td><td class="r">${l.cbm !== null ? fmt(down2(l.cbm * q)) : ""}</td></tr>`; }).join("");
  return `<article class="paper"><div class="ttl"><b>${lines2(p.shipper.split("\n")[0])}</b><span>${lines2(p.shipper.split("\n").slice(1).join("\n"))}</span><i>PACKING LIST</i></div>
    <div class="two"><div><strong>Consignee / Ship to:</strong><br>${lines2(p.consignee)}</div><div class="num"><strong>Invoice No.:</strong> ${esc(m.invoice || m.key)}<br><strong>Date:</strong> ${dShow(m.invDate)}<br><strong>Container / Seal:</strong> ${esc(m.cont || "–")} / ${esc(m.seal || "–")}<br><strong>Size:</strong> ${esc(m.size || "–")}</div></div>
    <table class="num"><thead><tr><th>Item No.</th><th>Goods Description</th><th class="r">Qty</th><th class="r">CTNS</th><th class="r">N.W (kg)</th><th class="r">G.W (kg)</th><th class="r">CBM</th></tr></thead>
      <tbody>${body}<tr class="tot"><td colspan="2">TOTAL</td><td class="r">${fmt(m.totals.qty, 0)}</td><td class="r">${fmt(m.totals.ctns, 0)}</td><td class="r">${fmt(m.totals.nw)}</td><td class="r">${fmt(m.totals.gw)}</td><td class="r">${fmt(m.totals.cbm)}</td></tr></tbody></table>
    <div class="two" style="margin-top:18px"><div><strong>Country of origin:</strong> ${esc(p.origin)}<br><strong>Term of delivery:</strong> ${esc(p.terms)}<br><strong>ETD:</strong> ${dShow(m.etd)} · <strong>POD:</strong> ${esc(m.pod || "–")}</div>
      <div><strong>REMARK:</strong><br>WE DECLARE THAT THE SHIPMENT CONTAINS NO SOLID WOOD PACKING MATERIALS.</div></div>
    <div class="sign"><div>Authorized Signature</div></div></article>`;
}

function viewDocs() {
  const d = S.data;
  if (!d || !d.rows.length) return `<header class="head"><div><h1>Dokumen CI/PL</h1><p>Belum ada data DI.</p></div></header><section class="card"><div class="empty"><b>Tambahkan file master DI dulu.</b><a class="btn" href="#/file" style="margin-top:12px">${I(ICON.upload, 18)}Tambah file</a></div></section>`;
  const sel = d.rows.filter((m) => S.selected.has(m.key));
  const cur = d.rows.find((m) => m.key === S.preview && m.level !== "err") || sel[0] || d.rows.find((m) => m.level !== "err");
  const okRows = d.rows.filter((m) => m.level !== "err");
  const allOn = okRows.length && okRows.every((m) => S.selected.has(m.key));
  const p = S.profile;
  return `<header class="head"><div><h1>Dokumen CI/PL</h1><p>Pilih BL, cek preview, lalu unduh dokumennya.</p></div>
      <div style="display:flex;gap:10px;flex-wrap:wrap"><button class="btn ghost" data-act="xlsx" ${sel.length ? "" : "disabled"}>${I(ICON.down, 18)}Excel</button>
      <button class="btn" data-act="pdf" ${sel.length ? "" : "disabled"}>${I(ICON.printer, 18)}PDF · ${sel.length} BL dipilih</button></div></header>
    <div class="row">
      <section class="card bl-list" aria-label="Daftar BL">
        <div class="card-h" style="padding:12px 18px"><label style="display:flex;align-items:center;gap:10px;font-size:14px;font-weight:700"><input type="checkbox" data-act="all" ${allOn ? "checked" : ""} style="width:18px;height:18px;accent-color:var(--navy)">Pilih semua</label><span class="num" style="font-size:13px;color:var(--muted)">${sel.length} / ${d.rows.length}</span></div>
        <ul>${d.rows.map((m) => { const bad = m.level === "err"; const on = S.selected.has(m.key);
          return `<li><label class="${on ? "on" : ""} ${bad ? "bad" : ""}"><input type="checkbox" data-act="pick" data-k="${esc(m.key)}" ${on ? "checked" : ""} ${bad ? "disabled" : ""}>
          <span><strong class="num">${esc(m.bl || m.key)}</strong><small class="num">${bad ? "Perbaiki error dulu" : `ETD ${dShow(m.etd)} · ${esc(m.size || "")}`}</small></span></label></li>`; }).join("")}</ul>
      </section>
      <section class="preview" aria-label="Preview dokumen">
        <div class="tabs" role="tablist" aria-label="Jenis dokumen"><button role="tab" data-act="tab" data-t="ci" aria-selected="${S.docTab === "ci"}">Commercial Invoice</button><button role="tab" data-act="tab" data-t="pl" aria-selected="${S.docTab === "pl"}">Packing List</button></div>
        ${cur ? (S.docTab === "ci" ? docCI(cur) : docPL(cur)) : `<div class="card"><div class="empty"><b>Belum ada BL yang bisa dibuat dokumennya.</b>Perbaiki error di Cek Data.</div></div>`}
        ${cur && sel.length > 1 ? `<p class="hint">Preview menampilkan ${esc(cur.bl || cur.key)} (BL terakhir yang dicentang). Unduhan berisi semua ${sel.length} BL yang dipilih.</p>` : ""}
        <details class="card"><summary>Profil dokumen (shipper, buyer, bank) <span class="hint">disimpan di browser ini</span></summary>
          <form class="settings" id="profile">
            ${[["shipper", "Shipper (baris 1 = nama)"], ["buyer", "Buyer / To"], ["consignee", "Consignee / Ship to"], ["bank", "Payment instruction"], ["terms", "Terms of delivery"], ["loadPort", "Port of loading"], ["payment", "Payment term"], ["origin", "Country of origin"]]
              .map(([k, l]) => `<label>${l}<textarea name="${k}" rows="${p[k].includes("\n") ? 3 : 1}">${esc(p[k])}</textarea></label>`).join("")}
            <div style="display:flex;gap:10px;align-items:end"><button class="btn" type="submit">Simpan profil</button><button class="btn soft" type="button" data-act="reset-profile">Kembalikan contoh</button></div>
          </form></details>
      </section></div>`;
}

function exportXlsx() {
  const sel = S.data.rows.filter((m) => S.selected.has(m.key));
  const wb = XLSX.utils.book_new();
  const used = new Set();
  const sheetName = (base) => { let n = base.replace(/[\\/?*[\]:]/g, "-").slice(0, 31), i = 2; while (used.has(n)) n = `${base.slice(0, 28)}-${i++}`; used.add(n); return n; };
  sel.forEach((m) => {
    const id = String(m.bl || m.key).replace(/^.*柜/, "DI");
    const ci = [["COMMERCIAL INVOICE"], ["Invoice No.", m.invoice || m.key, "Date", dShow(m.invDate)], ["BL", m.bl, "Container / Seal", `${m.cont || ""} / ${m.seal || ""}`], ["POD", m.pod, "Vessel", m.ship], [],
      ["Item No.", "ASIN", "PO", "Goods Description", "Qty (PCS)", "Unit Price", "Amount"],
      ...m.lines.map((l) => [l.sku, l.asin, l.po || m.po, l.name, l.qty, l.price, l.price !== null ? r2(l.price * (l.qty || 0)) : null]), ["TOTAL", "", "", "", m.totals.qty, "", m.totals.amount]];
    const pl = [["PACKING LIST"], ["Invoice No.", m.invoice || m.key, "Container / Seal", `${m.cont || ""} / ${m.seal || ""}`], [],
      ["Item No.", "Goods Description", "Qty", "CTNS", "N.W (kg)", "G.W (kg)", "CBM"],
      ...m.lines.map((l) => { const q = l.qty || 0; return [l.sku, l.name, q, q, l.nw !== null ? r2(l.nw * q) : null, l.gw !== null ? r2(l.gw * q) : null, l.cbm !== null ? down2(l.cbm * q) : null]; }),
      ["TOTAL", "", m.totals.qty, m.totals.ctns, m.totals.nw, m.totals.gw, m.totals.cbm]];
    const wci = XLSX.utils.aoa_to_sheet(ci), wpl = XLSX.utils.aoa_to_sheet(pl);
    wci["!cols"] = [{ wch: 18 }, { wch: 14 }, { wch: 12 }, { wch: 36 }, { wch: 10 }, { wch: 11 }, { wch: 12 }];
    wpl["!cols"] = [{ wch: 18 }, { wch: 36 }, { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 10 }, { wch: 8 }];
    XLSX.utils.book_append_sheet(wb, wci, sheetName(`INV ${id}`));
    XLSX.utils.book_append_sheet(wb, wpl, sheetName(`PL ${id}`));
  });
  XLSX.writeFile(wb, `Dockit CI-PL ${dkey(new Date())}.xlsx`);
  toast(`Excel berisi ${sel.length} BL sudah diunduh.`);
}
function printDocs() {
  const sel = S.data.rows.filter((m) => S.selected.has(m.key));
  $("#print").innerHTML = sel.map((m) => docCI(m) + docPL(m)).join("");
  window.print();
}

/* ---------- Halaman: Kalender ETD ---------- */
function viewCalendar() {
  const d = S.data;
  const rows = d ? d.rows.filter((m) => m.etd) : [];
  if (!S.cal) { const today = new Date(); const next = rows.find((m) => m.etd >= addDays(today, -1)) || rows[0]; const base = next ? next.etd : today; S.cal = new Date(base.getFullYear(), base.getMonth(), 1); }
  const y = S.cal.getFullYear(), mo = S.cal.getMonth();
  const start = addDays(S.cal, -((S.cal.getDay() + 6) % 7));
  const pods = [...new Set(rows.map((m) => m.pod || "(POD kosong)"))];
  const color = (pod) => POD_COLORS[pods.indexOf(pod || "(POD kosong)") % POD_COLORS.length];
  const evs = new Map();
  rows.forEach((m) => {
    [[m.depart, "go"], [m.etd, "etd"]].forEach(([dt, kind]) => { if (!dt) return; const k = dkey(dt); if (!evs.has(k)) evs.set(k, []); evs.get(k).push({ m, kind }); });
  });
  const todayK = dkey(new Date());
  let cells = "";
  for (let i = 0; i < 42; i++) {
    const day = addDays(start, i), k = dkey(day), list = evs.get(k) || [];
    cells += `<div class="day ${day.getMonth() !== mo ? "out" : ""} ${k === todayK ? "today" : ""}"><span class="d num">${day.getDate()}</span>
      ${list.slice(0, 4).map(({ m, kind }) => { const [bg, fg] = color(m.pod); return `<button class="ev ${kind}" data-act="ev" data-k="${esc(m.key)}" style="background:${bg};color:${fg}" title="${kind === "go" ? "Berangkat" : "ETD"} ${esc(m.bl || m.key)}">${kind === "go" ? "Berangkat · " : "ETD · "}${esc(m.bl || m.key)}</button>`; }).join("")}
      ${list.length > 4 ? `<small style="color:var(--muted)">+${list.length - 4} lagi</small>` : ""}</div>`;
  }
  const inMonth = rows.filter((m) => m.etd.getMonth() === mo && m.etd.getFullYear() === y);
  const pick = S.calPick && rows.find((m) => m.key === S.calPick);
  return `<header class="head"><div><h1>Kalender ETD</h1><p>Setiap BL tampil 2 hari: berangkat (H-1) lalu ETD. Warna menunjukkan POD.</p></div></header>
    <section class="card">
      <div class="card-h"><div class="cal-bar"><h2>${MONTHS[mo]} ${y}</h2>
        <button class="icon-btn" data-act="cal" data-d="-1" aria-label="Bulan sebelumnya">${I(ICON.left, 18)}</button>
        <button class="icon-btn" data-act="cal" data-d="1" aria-label="Bulan berikutnya">${I(ICON.right, 18)}</button>
        <button class="btn soft" data-act="cal" data-d="0" style="height:36px;padding:0 14px;font-size:14px">Hari ini</button></div>
        <p class="num">${inMonth.length} BL · ${new Set(inMonth.map((m) => ukey(m.cont)).filter(Boolean)).size} kontainer bulan ini</p></div>
      ${pods.length ? `<div class="legend" style="padding:12px 20px;border-bottom:1px solid var(--line-2)">${pods.map((p) => { const [bg, fg] = color(p); return `<span><i style="background:${bg};box-shadow:inset 0 0 0 1px ${fg}"></i>${esc(p)}</span>`; }).join("")}</div>` : ""}
      <div class="scroll"><div class="cal" style="min-width:760px">${["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"].map((x) => `<div class="dow">${x}</div>`).join("")}${cells}</div></div>
    </section>
    ${pick ? `<section class="card" aria-live="polite"><div class="card-h"><h2>${esc(pick.bl || pick.key)}</h2><button class="icon-btn" data-act="ev-close" aria-label="Tutup detail">${I(ICON.x, 18)}</button></div>
      <div class="detail num">${[["DI", pick.key], ["Kontainer", pick.cont], ["Seal", pick.seal], ["Size", pick.size], ["Berangkat", dShow(pick.depart)], ["ETD Batam", dShow(pick.etd)], ["ETD MV", dShow(pick.etdmv)], ["POD", pick.pod], ["Kapal", pick.ship], ["CBM", fmt(pick.cbm ?? pick.totals.cbm)], ["Status", { ok: "Siap", warn: "Ada peringatan", err: "Ada error", ai: "Ada anomali" }[pick.level]]]
        .map(([k, v]) => `<div><small>${k}</small>${esc(v ?? "–")}</div>`).join("")}</div></section>`
      : rows.length ? `<p class="hint">Klik blok di kalender untuk melihat detail BL.</p>` : `<section class="card"><div class="empty"><b>Belum ada ETD.</b>Tambahkan file master DI untuk mengisi kalender.</div></section>`}`;
}

/* ---------- Render & router ---------- */
function route() { const h = location.hash.replace(/^#\/?/, ""); const [r, q] = h.split("?"); return { r: ROUTES[r] ? r : "file", q: new URLSearchParams(q || "") }; }
function render() {
  const app = $("#app");
  if (!S.user) { app.innerHTML = viewLogin(); document.title = "Masuk · Dockit Web"; return; }
  const { r, q } = route();
  const body = r === "cek" ? viewCheck(q.get("ai") === "1") : r === "dokumen" ? viewDocs() : r === "kalender" ? viewCalendar() : viewFiles();
  app.innerHTML = shell(r, body);
  document.title = `${ROUTES[r]} · Dockit Web`;
  if (r === "cek" && q.get("ai") === "1") $("#ask")?.focus();
  const log = $("#ai-log"); if (log) log.scrollTop = log.scrollHeight;
}

/* ---------- Event ---------- */
document.addEventListener("click", async (e) => {
  const t = e.target.closest("[data-act]"); if (!t) return;
  const a = t.dataset.act;
  if (a === "eye") { const pw = $("#pw"); const show = pw.type === "password"; pw.type = show ? "text" : "password"; t.innerHTML = I(show ? ICON.eyeOff : ICON.eye, 20); t.setAttribute("aria-label", show ? "Sembunyikan password" : "Tampilkan password"); }
  else if (a === "forgot") { e.preventDefault(); toast("Hubungi admin shipping untuk reset password."); }
  else if (a === "logout") { S.user = null; store.del("dockit.session", sessionStorage); store.del("dockit.session"); location.hash = ""; render(); }
  else if (a === "rm-file") { S.files = S.files.filter((f) => f.id !== t.dataset.id); rebuild(); }
  else if (a === "sample") {
    t.disabled = true;
    try { const res = await fetch("sample/contoh-data-dockit.xlsx"); if (!res.ok) throw 0; const blob = await res.blob(); await addFiles([new File([blob], "contoh-data-dockit.xlsx")]); toast("Data contoh fiktif dimuat."); }
    catch { toast("Data contoh gagal dimuat."); t.disabled = false; }
  }
  else if (a === "filter") { S.filter = t.dataset.f; render(); }
  else if (a === "key") { const box = $("#keybox"); box.hidden = !box.hidden; t.setAttribute("aria-expanded", String(!box.hidden)); if (!box.hidden) $("#key-in").focus(); }
  else if (a === "save-key") { S.apiKey = $("#key-in").value.trim(); if (S.apiKey) store.set("dockit.key", S.apiKey, sessionStorage); else store.del("dockit.key", sessionStorage); $("#keybox").hidden = true; refreshChat(); toast(S.apiKey ? "API key tersimpan untuk tab ini." : "API key dihapus."); }
  else if (a === "ask") { const q = t.dataset.q; S.chat.push({ role: "user", text: t.textContent }); S.chat.push({ role: "bot", text: localAnswer(q) }); refreshChat(); }
  else if (a === "tab") { S.docTab = t.dataset.t; render(); }
  else if (a === "xlsx") exportXlsx();
  else if (a === "pdf") printDocs();
  else if (a === "reset-profile") { S.profile = { ...DEFAULT_PROFILE }; store.del("dockit.profile"); render(); toast("Profil dikembalikan ke contoh."); }
  else if (a === "cal") { const n = +t.dataset.d; if (n === 0) { const d = new Date(); S.cal = new Date(d.getFullYear(), d.getMonth(), 1); } else S.cal = new Date(S.cal.getFullYear(), S.cal.getMonth() + n, 1); render(); }
  else if (a === "ev") { S.calPick = t.dataset.k; render(); }
  else if (a === "ev-close") { S.calPick = null; render(); }
});
document.addEventListener("change", (e) => {
  const t = e.target;
  if (t.id === "file-in" && t.files.length) { addFiles([...t.files]); t.value = ""; }
  else if (t.dataset.act === "pick") { if (t.checked) { S.selected.add(t.dataset.k); S.preview = t.dataset.k; } else { S.selected.delete(t.dataset.k); if (S.preview === t.dataset.k) S.preview = null; } render(); }
  else if (t.dataset.act === "all") { S.data.rows.filter((m) => m.level !== "err").forEach((m) => (t.checked ? S.selected.add(m.key) : S.selected.delete(m.key))); render(); }
});
document.addEventListener("submit", (e) => {
  e.preventDefault();
  const f = e.target;
  if (f.id === "login") doLogin(f);
  else if (f.id === "ai-form") { const v = $("#ask").value.trim(); if (v && !S.busy) { $("#ask").value = ""; askAI(v); } }
  else if (f.id === "profile") { const p = {}; new FormData(f).forEach((v, k) => (p[k] = String(v))); S.profile = { ...DEFAULT_PROFILE, ...p }; store.set("dockit.profile", S.profile); render(); toast("Profil dokumen disimpan."); }
});
["dragenter", "dragover"].forEach((ev) => document.addEventListener(ev, (e) => { const d = e.target.closest?.("#drop"); if (d) { e.preventDefault(); d.classList.add("over"); } }));
["dragleave", "drop"].forEach((ev) => document.addEventListener(ev, (e) => { const d = e.target.closest?.("#drop"); if (d) { d.classList.remove("over"); } }));
document.addEventListener("drop", (e) => { if (e.target.closest?.("#drop")) { e.preventDefault(); addFiles([...e.dataTransfer.files]); } });
window.addEventListener("hashchange", render);
window.addEventListener("beforeunload", (e) => { if (S.files.length) { e.preventDefault(); e.returnValue = ""; } });

render();
})();
