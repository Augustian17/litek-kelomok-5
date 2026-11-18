# Dockit: Otomasi Dokumen Ekspor Berbantuan AI

> Tugas Kelompok 5: **Design Thinking dan Solusi Digital Berbasis AI (Ideasi & Prototyping)**

Dockit adalah aplikasi desktop Windows yang menyusun dokumen ekspor **Commercial Invoice (CI)** dan **Packing List (PL)** secara otomatis dari banyak file Excel, memvalidasi datanya dengan bantuan AI, lalu mengekspornya ke PDF sekaligus.

Proyek ini berangkat dari masalah nyata di departemen shipping. Kami menyelesaikannya dengan metode **Design Thinking**, dan hasilnya dituangkan dalam **Design Thinking Canvas**.

---

## Daftar Isi

- [Anggota Kelompok](#anggota-kelompok)
- [Latar Belakang](#latar-belakang)
- [Design Thinking Canvas](#design-thinking-canvas)
- [Proses Design Thinking](#proses-design-thinking)
- [Fitur Solusi](#fitur-solusi)
- [Peran AI](#peran-ai)
- [Struktur Repository](#struktur-repository)
- [Rencana Pengembangan](#rencana-pengembangan)

---

## Anggota Kelompok

| No | Nama | NIM | Peran |
|----|------|-----|-------|
| 1 | Augustian Gautama | 2023131006 | Project Manager |
| 2 | [NAMA] | [NIM] | [PERAN] |
| 3 | [NAMA] | [NIM] | [PERAN] |
| 4 | [NAMA] | [NIM] | [PERAN] |
| 5 | [NAMA] | [NIM] | [PERAN] |

**Mata Kuliah:** [NAMA MATA KULIAH]
**Dosen Pengampu:** Yodi
**Universitas:** Universitas Universal, Batam

---

## Latar Belakang

Setiap kali ada pengiriman, staf shipping harus menyalin data dari banyak file Excel (PO, DI, jadwal muat) ke template CI/PL, lalu mengubah tiap sheet menjadi PDF satu per satu. Prosesnya berulang dan memakan waktu.

Satu salah ketik pada nomor kontainer, nomor seal, atau CBM membuat dokumen harus direvisi. Padahal kontainer berangkat **H-1 sebelum ETD**, sehingga setiap revisi berisiko menunda pengiriman.

---

## Design Thinking Canvas

![Design Thinking Canvas — Dockit](docs/design-thinking-canvas.png)

| Bagian | Isi |
|--------|-----|
| **01 People** | Staf shipping/export (pembuat CI/PL) · Supervisor logistik (pemeriksa & penyetuju) · Forwarder & bea cukai (penerima dokumen) · Buyer/customer (penerima CI/PL) |
| **02 Challenges** | Format Excel dari tiap sumber tidak seragam · Data sebelum 2026 tersimpan terpisah · Staf terbiasa proses manual · Tidak ada tim IT khusus |
| **03 Storytelling** | Jadwal pengiriman masuk → staf membuka banyak file Excel → data disalin manual ke template → tiap sheet di-export ke PDF satu per satu → salah ketik nomor kontainer/seal/CBM → dokumen direvisi, keberangkatan H-1 ETD terancam |
| **04 Problem** | CI/PL dibuat manual dari banyak file · Duplikat container/seal/rolling number · Salah pembulatan CBM & qty tidak cocok · Tidak ada validasi otomatis |
| **05 Solution** | Menarik data PO/DI & menyusun CI/PL otomatis · AI mendeteksi kesalahan & anomali · Export banyak sheet ke PDF sekaligus · AI membaca dokumen forwarder (pengembangan) |
| **06 Management** | Standarisasi template CI/PL · Rapikan struktur data PO & DI · Uji coba paralel dengan proses manual · Pelatihan staf · Versioning tiap update · Feedback pengguna tiap rilis |
| **07 Vision** | Dokumen ekspor selesai dalam hitungan menit · Nol kesalahan data kontainer & seal · Tim fokus pada koordinasi, bukan administrasi · Data shipment siap untuk AI lanjutan |
| **08 Impact** | Waktu pembuatan dokumen berkurang · Kesalahan & revisi berkurang · Pengiriman lebih tepat waktu · Data siap dianalisis |

---

## Proses Design Thinking

| Tahap | Yang Dilakukan |
|-------|----------------|
| **1. Empathize** | Mengamati alur kerja staf shipping saat menyiapkan dokumen CI/PL untuk setiap pengiriman |
| **2. Define** | Merumuskan masalah inti: pembuatan dokumen yang manual dan tanpa validasi menyebabkan kesalahan dan keterlambatan |
| **3. Ideate** | Brainstorming solusi: template baku, makro Excel, sampai aplikasi otomasi dengan validasi berbasis AI |
| **4. Prototype** | Membangun Dockit sebagai aplikasi desktop yang membaca file Excel dan menghasilkan CI/PL dalam PDF |
| **5. Test** | Uji coba Dockit secara paralel dengan proses manual pada pengiriman nyata, lalu iterasi berdasarkan feedback |

---

## Fitur Solusi

- **Penyusunan dokumen otomatis.** Menarik data dari banyak file Excel (PO & DI) dan mengisi template CI/PL.
- **Validasi data.** Mendeteksi duplikat container number, seal number, dan rolling number.
- **Aturan bisnis bawaan.** CBM selalu dibulatkan ke bawah, dan tanggal keberangkatan dihitung H-1 dari ETD.
- **Export PDF massal.** Mengubah banyak sheet terpilih dari banyak file menjadi PDF sekaligus.
- **Versioning.** Setiap update aplikasi diberi nomor versi baru.

---

## Peran AI

| Fitur AI | Fungsi | Status |
|----------|--------|--------|
| Deteksi anomali data | Menandai nilai yang tidak wajar (qty, berat, CBM) dibandingkan pola pengiriman sebelumnya | Direncanakan |
| Pembacaan dokumen (OCR + LLM) | Membaca dokumen masuk dari forwarder (PDF/scan) dan mencocokkannya dengan data internal | Direncanakan |
| Pemetaan kolom otomatis | Mengenali kolom dari file Excel yang formatnya berbeda-beda | Direncanakan |

---

## Struktur Repository

```
litek-kelomok-5/
├── README.md
└── docs/
    └── design-thinking-canvas.png   # Hasil Design Thinking Canvas
```

---

## Rencana Pengembangan

- [x] Design Thinking Canvas
- [ ] Wireframe / mockup antarmuka Dockit
- [ ] Prototype fitur deteksi anomali berbasis AI
- [ ] Pengujian dengan pengguna (staf shipping)
- [ ] Presentasi hasil

---

<p align="center">Kelompok 5 · Universitas Universal · 2026</p>
