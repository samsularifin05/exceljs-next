# exceljs-next (Dokumentasi Bahasa Indonesia)

**Bahasa / Languages:** [English](README.md) | **Bahasa Indonesia** | [中文](README_zh.md)

> **Tentang Repositori Ini:**
> [samsularifin05/exceljs-next](https://github.com/samsularifin05/exceljs-next) merupakan *maintained fork* yang dimodernisasi dan diperkeras keamanannya dari repositori asli [exceljs/exceljs](https://github.com/exceljs/exceljs) (yang tidak lagi diperbarui sejak `v4.4.0` pada Januari 2024). Proyek ini menjaga 100% kompatibilitas API lama sekaligus memperbaiki celah keamanan (**0 vulnerabilities** di `npm audit`), bug streaming pada Node.js 16–22+, masalah korupsi file Excel, dukungan Native ESM & TypeScript penuh, serta menambahkan berbagai fitur baru.
> Lihat rincian perubahan lengkap di [CHANGELOG.md](CHANGELOG.md) atau baca tutorial langkah demi langkah di **[GitHub Wiki](https://github.com/samsularifin05/exceljs-next/wiki)**.

[![Build Status](https://github.com/samsularifin05/exceljs-next/actions/workflows/tests.yml/badge.svg?branch=master&event=push)](https://github.com/samsularifin05/exceljs-next/actions/workflows/tests.yml)

Membaca, memanipulasi, dan menulis data serta styling spreadsheet ke format **XLSX, CSV, dan JSON** — mendukung **Node.js `>=16.0.0`**, **Native ESM**, **TypeScript**, dan **Browser**.

---

## Daftar Isi

- [Keunggulan Dibanding `exceljs@4.4.0`](#-keunggulan-dibanding-exceljs440)
- [Instalasi](#-instalasi)
- [Demo Interaktif Browser](#-demo-interaktif-browser)
- [Fitur Baru di `exceljs-next`](#-fitur-baru-di-exceljs-next)
  - [1. File Encryption — Password Buka File (MS-OFFCRYPTO)](#1-file-encryption--password-buka-file-ms-offcrypto)
  - [2. Native Excel Charts (`ws.addChart()`)](#2-native-excel-charts-wsaddchart)
  - [3. Template Engine (`renderTemplate()`)](#3-template-engine-rendertemplate)
  - [4. Helper JSON (`ws.fromJSON()` & `ws.toJSON()`)](#4-helper-json-wsfromjson--wstojson)
  - [5. Auto-Fit Lebar Kolom & Tinggi Baris](#5-auto-fit-lebar-kolom--tinggi-baris)
  - [6. Enhanced Multi-Pivot Tables](#6-enhanced-multi-pivot-tables)
  - [7. Workbook Structure Protection (`workbook.protect()`)](#7-workbook-structure-protection-workbookprotect)
  - [8. Column Page Breaks (`colBreaks`)](#8-column-page-breaks-colbreaks)
  - [9. Streaming Writer & Reader Parity](#9-streaming-writer--reader-parity)
- [Panduan Dasar Penggunaan](#-panduan-dasar-penggunaan)
  - [Membuat & Mengelola Workbook](#membuat--mengelola-workbook)
  - [Mengelola Worksheet, Kolom, Baris, & Sel](#mengelola-worksheet-kolom-baris--sel)
  - [Styling (Font, Alignment, Border, Fill, Number Format)](#styling-font-alignment-border-fill-number-format)
  - [Formula, Merge Cell, Gambar, & Komentar](#formula-merge-cell-gambar--komentar)
  - [Proteksi Worksheet (`ws.protect()`)](#proteksi-worksheet-wsprotect)
  - [Membaca & Menulis XLSX / CSV / Streaming](#membaca--menulis-xlsx--csv--streaming)
  - [Penggunaan di TypeScript & ESM](#penggunaan-di-typescript--esm)
- [Lisensi](#-lisensi)

---

## Keunggulan Dibanding `exceljs@4.4.0`

* **Keamanan Terjamin (0 Vulnerabilities):** Menutup celah *Prototype Pollution* pada `deepMerge`, `note`, `definedNames`, `cellMatrix`, `dataValidations`, dan `cfRule`; menambah validasi batas anti-OOM pada `<mergeCell>` dan koordinat sel (`1..1048576`); mengamankan lookup properti CSV; serta menghapus dependensi rentan (`tmp` & `uuid` diganti implementasi bawaan Node.js).
* **File Encryption (Password Buka File):** Enkripsi file `.xlsx` dengan standar **MS-OFFCRYPTO Agile Encryption (AES-256-CBC + SHA-512 + HMAC-SHA512)** agar Excel / WPS / LibreOffice langsung meminta password saat file dibuka, serta mendukung dekripsi file terenkripsi Agile (`4.4`) & Standard (`3.2`/`4.2`).
* **Native Excel Charts:** Membuat grafik `column`, `bar`, `line`, `area`, `pie`, dan `doughnut` lengkap dengan *data labels* (nilai & persentase).
* **Template Engine:** Mengisi placeholder `{{variabel}}` dan mengekspansi baris array `{{items.nama}}` secara otomatis tanpa merusak format template.
* **JSON Import/Export & Auto-Fit:** `ws.fromJSON()`, `ws.toJSON()`, `ws.autoFitColumns()`, dan `ws.autoFitRows()`.
* **Kompatibilitas Node.js 16–22+:** Memperbaiki bug `WorkbookReader` yang kehilangan `sharedStrings` pada Node.js modern serta korupsi karakter multi-byte UTF-8 (CJK, Arab, Emoji) pada streaming chunk.
* **Native ESM & TypeScript Akurat:** Menyediakan `excel.mjs`, `index.d.mts`, dan pembaruan lengkap `index.d.ts`.

---

## Instalasi

```shell
npm install exceljs-next
# atau dengan pnpm
pnpm add exceljs-next
```

Atau pasang sebagai *drop-in replacement* tanpa perlu mengubah `require('exceljs')` / `import ... from 'exceljs'` di proyek Anda yang sudah ada:

```shell
npm install exceljs@npm:exceljs-next
```

---

## Demo Interaktif Browser

Repositori ini menyertakan halaman playground interaktif di [demo.html](demo.html) untuk menguji hampir seluruh fitur langsung di browser (Auto-Fit, Pivot Table, Chart, Template Engine, JSON Import/Export, Proteksi Workbook, dan File Encryption):

```shell
npx serve .
# Lalu buka http://localhost:3000/demo.html
```

---

## Fitur Baru di `exceljs-next`

### 1. File Encryption — Password Buka File (`MS-OFFCRYPTO`)

Mengenkripsi seluruh isi file `.xlsx` sehingga pengguna harus memasukkan password sebelum file dapat dibuka di Microsoft Excel, WPS Office, maupun LibreOffice.

```js
// Menulis file .xlsx terenkripsi password
await workbook.xlsx.writeFile('rahasia.xlsx', { password: 'PasswordKuat123' });
const encryptedBuffer = await workbook.xlsx.writeBuffer({ password: 'PasswordKuat123' });

// Membaca / mendekripsi file .xlsx ber-password
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.readFile('rahasia.xlsx', { password: 'PasswordKuat123' });
await wb2.xlsx.load(encryptedBuffer, { password: 'PasswordKuat123' });
```

---

### 2. Native Excel Charts (`ws.addChart()`)

Menambahkan grafik native Excel (`column`, `bar`, `line`, `area`, `pie`, `doughnut`) ke dalam worksheet:

```js
ws.addChart({
  type: 'doughnut', // 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut'
  title: 'Pendapatan per Produk',
  range: 'E2:M18',
  legend: { position: 'b' }, // 'r' | 'b' | 't' | 'l' | 'none'
  valAxis: { numFmt: '#,##0,," Jt"' },
  dataLabels: {
    showVal: true,
    showPercent: true,
    separator: '\n',
  },
  series: [
    {
      name: 'Revenue',
      categories: '$A$2:$A$6',
      values: '$C$2:$C$6',
      color: '#10B981',
    },
  ],
});
```

---

### 3. Template Engine (`renderTemplate()`)

Mengisi file template `.xlsx` (atau sheet yang berisi placeholder `{{...}}`) sambil menjaga gaya teks, warna, border, dan tipe data asli (`number`, `Date`, `boolean`). Baris yang berisi placeholder array akan otomatis digandakan sesuai jumlah data.

```js
ws.getCell('A1').value = 'Invoice #{{invoiceNo}} - {{customer.name}}';
ws.getCell('A4').value = '{{items.$number}}. {{items.name}}';
ws.getCell('B4').value = '{{items.qty}}';
ws.getCell('C4').value = '{{items.price}}';

workbook.renderTemplate({
  invoiceNo: 'INV-2026-001',
  customer: { name: 'PT Nusantara Digital' },
  items: [
    { name: 'Cloud Server', qty: 2, price: 4500000 },
    { name: 'Managed DB', qty: 1, price: 3200000 },
  ],
});
```

---

### 4. Helper JSON (`ws.fromJSON()` & `ws.toJSON()`)

Konversi cepat antara array of objects JavaScript dan baris Excel:

```js
// Mengisi worksheet dari array of objects
ws.fromJSON(
  [
    { nama: 'Alice', role: 'Admin', skor: 98 },
    { nama: 'Bob', role: 'Finance', skor: 91 },
  ],
  {
    headerLabels: { nama: 'Nama Lengkap', role: 'Jabatan', skor: 'Skor' },
    headerStyle: { font: { bold: true } },
    autoFit: true,
  }
);

// Mengekspor worksheet kembali ke array of objects
const records = ws.toJSON({ headerRow: 1, includeEmpty: false });
```

---

### 5. Auto-Fit Lebar Kolom & Tinggi Baris

```js
// Otomatis menyesuaikan lebar semua kolom
ws.autoFitColumns({
  minWidth: 10,
  maxWidth: 60,
  padding: 2,
  onlyEmptyWidth: false,
});

// Otomatis menyesuaikan tinggi semua baris (mendukung '\n', wrapText, dan ukuran font)
ws.autoFitRows({
  minHeight: 16,
  maxHeight: 409,
  padding: 4,
  onlyEmptyHeight: false,
});
```

---

### 6. Enhanced Multi-Pivot Tables

Mendukung banyak Pivot Table dalam satu workbook, 11 metrik agregasi (`sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevP`, `var`, `varP`), multi-values, dan `sourceTable`:

```js
const pivotSheet = workbook.addWorksheet('Ringkasan Pivot');
pivotSheet.addPivotTable({
  sourceTable: salesTable, // atau sourceSheet: ws
  rows: ['Region', 'Product'],
  columns: [],
  values: ['Revenue', 'Units'],
  metric: 'average',
});
```

---

### 7. Workbook Structure Protection (`workbook.protect()`)

Mengunci struktur tab sheet (mencegah penambahan, penghapusan, pengubahan nama, atau pemindahan sheet):

```js
await workbook.protect('rahasia123', {
  lockStructure: true,
  lockWindows: false,
});

// Membuka kunci di model
workbook.unprotect();
```

---

### 8. Column Page Breaks (`colBreaks`)

```js
// Memisahkan halaman cetak setelah kolom C (index 3)
ws.getColumn(3).addPageBreak(0, 1048575);
```

---

### 9. Streaming Writer & Reader Parity

`WorksheetWriter` kini mendukung `addTable()`, `addImage()`, extended conditional formatting (`dataBar`/`iconSet` `<extLst>`), dan `colBreaks`. `WorksheetReader` mendukung pembacaan `mergeCells` (`sheetReader.mergeCells` dan event `'mergeCell'`).

---

## Panduan Dasar Penggunaan

### Membuat & Mengelola Workbook

```js
const ExcelJS = require('exceljs-next');

const workbook = new ExcelJS.Workbook();
workbook.creator = 'Tim Engineering';
workbook.created = new Date();
workbook.modified = new Date();
```

### Mengelola Worksheet, Kolom, Baris, & Sel

```js
const sheet = workbook.addWorksheet('Laporan', {
  views: [{ state: 'frozen', ySplit: 1 }],
});

sheet.columns = [
  { header: 'ID', key: 'id', width: 10 },
  { header: 'Nama', key: 'name', width: 32 },
  { header: 'Tanggal Lahir', key: 'dob', width: 15 },
];

sheet.addRow({ id: 1, name: 'Budi Santoso', dob: new Date(1995, 4, 12) });
sheet.addRow([2, 'Siti Aminah', new Date(1998, 8, 20)]);

// Mengakses atau mengubah sel tertentu
sheet.getCell('B2').value = 'Budi Santoso, S.Kom';
```

### Styling (Font, Alignment, Border, Fill, Number Format)

```js
const cell = sheet.getCell('A1');
cell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
cell.border = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' },
};

sheet.getCell('C2').numFmt = 'yyyy-mm-dd';
```

### Formula, Merge Cell, Gambar, & Komentar

```js
// Formula (dengan nilai hasil cached)
sheet.getCell('D5').value = { formula: 'SUM(D2:D4)', result: 150000 };

// Merge Cell
sheet.mergeCells('A10:C10');

// Komentar / Catatan Sel
sheet.getCell('A2').note = 'Data sudah diverifikasi';

// Menambahkan Gambar
const imageId = workbook.addImage({
  filename: 'logo.png',
  extension: 'png',
});
sheet.addImage(imageId, 'E1:F4');
```

### Proteksi Worksheet (`ws.protect()`)

Untuk mengunci isi sel agar tidak dapat diedit pengguna:

```js
await sheet.protect('password-sheet', {
  selectLockedCells: true,
  selectUnlockedCells: true,
  formatCells: false,
  insertRows: false,
  deleteRows: false,
});

// Membuka kunci sel tertentu agar tetap bisa diisi user
sheet.getCell('B2').protection = { locked: false };
```

### Membaca & Menulis XLSX / CSV / Streaming

```js
// XLSX File & Buffer
await workbook.xlsx.writeFile('output.xlsx');
const buffer = await workbook.xlsx.writeBuffer();

const wbRead = new ExcelJS.Workbook();
await wbRead.xlsx.readFile('output.xlsx');

// CSV File
await workbook.csv.writeFile('output.csv');
await wbRead.csv.readFile('output.csv');

// Streaming Writer (hemat memori untuk ratusan ribu baris)
const streamWb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: 'besar.xlsx' });
const streamWs = streamWb.addWorksheet('Data');
for (let i = 1; i <= 100000; i++) {
  streamWs.addRow([i, `Baris ke-${i}`]).commit();
}
await streamWb.commit();
```

### Penggunaan di TypeScript & ESM

```ts
import ExcelJS, { Workbook, ValueType } from 'exceljs-next';

const wb: Workbook = new ExcelJS.Workbook();
const ws = wb.addWorksheet('TypeScript');
ws.fromJSON([{ item: 'Buku', qty: 3 }], { autoFit: true });
ws.autoFitRows();
await wb.xlsx.writeFile('ts-output.xlsx', { password: '123' });
```

---

## Pengembang Asli & Kredit

* **Proyek & Pengembang Asli:** Dibangun di atas repositori asli [exceljs/exceljs](https://github.com/exceljs/exceljs) yang dibuat oleh **Guyon Roche** (`guyon@live.com`) beserta seluruh kontributornya.
* **Maintained Fork (`exceljs-next`):** Pembaruan keamanan, kompatibilitas Node.js 16–22+, dukungan TypeScript/ESM, serta fitur-fitur baru dikelola oleh **[samsularifin05](https://github.com/samsularifin05)**.

---

## Lisensi

[MIT](LICENSE)
