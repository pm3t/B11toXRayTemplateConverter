# X-Ray Cargo Breakdown Report Template Generator

Program untuk mengonversi data manifes BC 1.1 dari file Excel berformat modern (`.xlsx`) maupun file Excel lama (`.xls`) menjadi file Excel template laporan *Breakdown Report Cargo International* (`.xlsx`).

---

## 📋 Pemetaan Kolom (Mapping Rules)

| No | Kolom di `OUT.xlsx` | Sumber di `IN.xlsx` / `*.xls` | Keterangan |
|---|---|---|---|
| 1 | **No** | Nomor Urut (1, 2, 3, ...) | Nomor urut integer |
| 2 | **MAWB** | Sheet `Detil` -> `NO MASTER BLAWB` | Master Air Waybill |
| 3 | **HAWB** | Sheet `Detil` -> `NO HOST BLAWB` | House Air Waybill |
| 4 | **Act Pcs** | Sheet `Detil` / `Detail` -> `JUMLAH KEMASAN` | Jumlah kemasan aktual (integer) |
| 5 | **of** | *(Kosong / None)* | Sesuai template |
| 6 | **Tot Pcs** | Sheet `Detil` -> `TOTAL KEMASAN` | Total kemasan (fallback ke `JUMLAH KEMASAN` jika kosong) |
| 7 | **Act Weight** | *(Kosong / None)* | Sesuai template |
| 8 | **of** | *(Kosong / None)* | Sesuai template |
| 9 | **Tot Weight** | Sheet `Detil` -> `BRUTO` | Total berat kotor |
| 10 | **Nature of Goods** | Sheet `Barang` -> `URAIAN BARANG` | Dipilih berdasarkan kecocokan `ID DETIL`. Nilai `'CONSOL'` dan kode HS `'0000'` difilter secara cerdas. Dibatasi maksimal 50 karakter. |
| 11 | **Actual Times** | *(Kosong / None)* | Sesuai template |
| 12 | **Remarks** | *(Kosong / None)* | Sesuai template |
| 13 | **Status** | *(Kosong / None)* | Sesuai template |
| 14 | **BC11** | Sheet `Header` -> `NO BC 11` | Nomor pendaftaran BC 1.1 |
| 15 | **TanggalBC11** | Sheet `Header` -> `TGL BC 11` | Tanggal BC 1.1 (format tanggal `mm-dd-yy`) |
| 16 | **FlightNumber** | Sheet `Header` -> `NO VOYAGE / ARRIVAL` | Nomor penerbangan / flight (e.g. VJ0929) |
| 17 | **FlightDate** | Sheet `Header` -> `TGL TIBA` | Tanggal penerbangan / tiba (format `mm-dd-yy`) |
| 18 | **Origin** | Sheet `Detil` / `Header` -> `PELABUHAN ASAL` | Kode bandara / pelabuhan asal |
| 19 | **Destination** | Sheet `Detil` / `Header` -> `PELABUHAN BONGKAR` | Kode bandara / pelabuhan tujuan/bongkar |
| 20 | **Shipper ** | Sheet `Detil` -> `NAMA SHIPPER` | Nama pengirim barang |
| 21 | **AlamatShipper ** | Sheet `Detil` -> `ALMT SHIPPER` | Alamat pengirim barang |
| 22 | **Consignee** | Sheet `Detil` -> `NAMA CONSIGNEE` | Nama penerima barang |
| 23 | **AlamatConsignee** | Sheet `Detil` -> `ALMT CONSIGNEE` | Alamat penerima barang |

---

## 🚫 Filter Kargo Transshipment

Kargo transshipment (hanya transit / lanjutan ke luar negeri dan **tidak di-screening**) secara otomatis **dilewati (exclude)** dan tidak masuk ke file `OUT.xlsx`.

Kriteria deteksi transshipment:
1. Kolom `KD KELOMPOK POS` bernilai `'02'` (kargo pos lanjutan/transshipment).
2. Nilai `PELABUHAN BONGKAR` atau `PELABUHAN AKHIR` berbeda dengan pelabuhan kedatangan impor pada Header (misalnya tujuan ke bandara luar negeri seperti `USJFK`, `AEDXB`, `ITMXP`).

---

## 🌐 Cara Menjalankan Versi Web (Siap Deploy Vercel)

Aplikasi ini kini dilengkapi versi **Web UI modern** yang berjalan langsung di browser (client-side) tanpa memerlukan server backend.

### Keunggulan Versi Web:
- **100% Privat & Aman**: Dokumen manifes diproses langsung di memori browser pengguna, tidak pernah diunggah ke server pihak ketiga.
- **Mendukung `.xlsx` & `.xls`**: Menggunakan SheetJS & ExcelJS untuk membaca dan membuat format template secara instan.
- **Preview Interaktif**: Menampilkan ringkasan nomor penerbangan, tanggal, BC 1.1, serta tabel preview data sebelum diunduh.
- **Pencarian Real-Time**: Fitur filter pencarian instan berdasarkan MAWB, HAWB, Shipper, atau jenis barang.

### 1. Menjalankan di Komputer Lokal (Dev Mode)
```bash
npm install
npm run dev
```
Buka browser pada alamat yang muncul (biasanya `http://localhost:5173`).

### 2. Cara Deploy ke Vercel
Proyek ini sudah dikonfigurasi dengan `vercel.json` dan siap dideploy langsung:

**Metode A: Menggunakan GitHub / GitLab (Rekomendasi)**
1. Push repository ini ke akun GitHub/GitLab Anda.
2. Masuk ke [vercel.com](https://vercel.com) &rarr; **Add New Project**.
3. Pilih repository ini.
4. Vercel akan otomatis mendeteksi framework **Vite** dengan Build Command `npm run build` dan Output Directory `dist`.
5. Klik **Deploy**!

**Metode B: Menggunakan Vercel CLI**
```bash
npx vercel
```
Ikuti petunjuk di terminal untuk langsung mempublikasikan.

---

```bash
pip install -r requirements.txt
```
*(Membutuhkan `openpyxl` untuk output/input .xlsx dan `xlrd` untuk input .xls)*

---

## 🚀 Cara Menjalankan Program

### 1. Menggunakan Script Shell (Paling Praktis)
```bash
./run.sh
```
*Secara default membaca `IN.xlsx` dan menghasilkan `OUT.xlsx`.*

Bisa menggunakan file input `.xlsx` maupun `.xls`:
```bash
./run.sh "20261013_VJ0929(1).xls" OUT2.xlsx
```

### 2. Menggunakan Python Langsung
```bash
python main.py "20261013_VJ0929(1).xls" OUT2.xlsx
```
*(Typo ekstensi seperti `.xlxs` akan otomatis dikoreksi menjadi `.xlsx`)*
# JASTemplateConverter
