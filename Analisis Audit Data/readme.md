# 🔬 Laporan Analisis Prosedur Audit Piutang

> **Empat metode audit analitik dalam satu laporan Excel — dari data AR Accurate ke temuan risiko siap presentasi**

Skrip Python satu file yang membaca data AR bulanan dari file `PTM*.xlsx` (file yang sama digunakan bersama `Monitoring AR Outstanding`), lalu secara otomatis menyusun **`Laporan_Analisis_Prosedur_Audit.xlsx`** berisi lima sheet prosedur audit: analisis tren horizontal & vertikal, deteksi anomali Z-Score, uji Hukum Benford forensik, pemetaan konsentrasi risiko Pareto, hingga deteksi duplikasi faktur.

---

## 📋 Daftar Isi

- [Gambaran Umum & Konteks](#-gambaran-umum--konteks)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Detail](#-alur-kerja-detail)
- [Penjelasan 5 Sheet Audit](#-penjelasan-5-sheet-audit)
  - [Sheet 1 — Analisis Horisontal & Vertikal](#sheet-1--analisis-horisontal--vertikal)
  - [Sheet 2 — Anomali Z-Score Makro](#sheet-2--anomali-z-score-makro)
  - [Sheet 3 — Forensik Hukum Benford](#sheet-3--forensik-hukum-benford)
  - [Sheet 4 — Risiko Pareto Customer](#sheet-4--risiko-pareto-customer)
  - [Sheet 5 — Uji Duplikasi Faktur](#sheet-5--uji-duplikasi-faktur)
- [Konfigurasi `config.conf`](#-konfigurasi-configconf)
- [Catatan Teknis & Perbedaan dari Monitoring AR](#-catatan-teknis--perbedaan-dari-monitoring-ar)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum & Konteks

Skrip ini dirancang sebagai **lapisan analitik lanjutan** di atas data yang sama dengan proyek `Monitoring AR Outstanding`. Jika `Monitoring AR` berfokus pada *monitoring operasional* (berapa saldo piutang per bulan, berapa yang 60 hari, dst.), maka skrip ini berfokus pada *audit analitik*: menemukan anomali, pola tidak wajar, dan konsentrasi risiko yang perlu direviu lebih mendalam oleh auditor.

Empat metode yang diimplementasikan adalah bagian dari prosedur analitik standar audit keuangan:

| Metode | Tujuan Audit |
|---|---|
| Analisis Horizontal & Vertikal | Menilai tren dan komposisi portofolio piutang |
| Z-Score | Mendeteksi lonjakan saldo bulanan yang tidak wajar secara statistik |
| Hukum Benford | Menguji distribusi digit angka faktur untuk indikasi fraud/manipulasi |
| Pareto 80/20 | Memetakan konsentrasi risiko gagal bayar pada segelintir pelanggan |
| Uji Duplikasi | Mendeteksi faktur yang diinput ganda dalam satu bulan |

---

## ✨ Fitur Utama

- **Input berbagi dengan Monitoring AR** — Membaca file `PTM*.xlsx` yang sama, dengan logika `clean_key()` dan `auto-detect header` yang identik, sehingga tidak perlu menyiapkan file terpisah.
- **Kalkulasi MoM otomatis** — Perubahan Month-over-Month dihitung dengan `pct_change()` pandas, tanpa formula manual.
- **Deteksi anomali statistik Z-Score** — Threshold `|Z-Score| > 1.2` digunakan untuk menandai bulan dengan fluktuasi saldo tidak wajar; baris anomali di-highlight merah secara otomatis.
- **Analisis forensik Hukum Benford** — Distribusi digit pertama seluruh `Nilai Faktur` dibandingkan dengan probabilitas Benford teoritis; deviasi > 5% ditandai untuk reviu.
- **Pareto 80/20 berbasis data bulan terakhir** — Mengurutkan pelanggan berdasarkan total saldo piutang, menghitung kumulatif kontribusi, dan mengklasifikasikan mereka yang masuk dalam 80% pertama sebagai `TOP 80% CORE RISK (Kritis)`.
- **Deteksi duplikasi lintas semua bulan** — Menggabungkan seluruh transaksi dari semua sheet MMYY, lalu menguji apakah ada nomor faktur yang muncul lebih dari sekali dalam bulan yang sama.
- **Styling audit profesional** — Header biru gelap (`#366092`), baris temuan merah-oranye (`#F2DCDB`), font bold untuk anomali, format angka `#,##0` dan `0.00%`, serta auto-fit lebar kolom di semua sheet.
- **Auto-generate `config.conf`** — Jika file konfigurasi tidak ditemukan, skrip membuat versi default secara otomatis tanpa menghentikan proses.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl numpy
```

| Library | Kegunaan |
|---|---|
| `pandas` | Baca Excel, transformasi data, groupby, pct_change |
| `openpyxl` | Buat workbook, styling sel (font, fill, border, alignment) |
| `numpy` | Z-Score (mean, std), `np.where` untuk klasifikasi kondisional |
| `configparser` | Baca `config.conf` (standard library) |
| `re`, `glob`, `datetime`, `os` | Utilitas (standard library) |

---

## 📁 Struktur Folder & File

```
📦 Buat_Analisis_Audit/
│
├── 📄 Buat_Analisis_Audit.py                ← Skrip utama. Jalankan ini
├── 📄 config.conf                           ← Konfigurasi filter produk
│                                               (dibuat otomatis jika tidak ada)
│
├── 📄 PTM*.xlsx                             ← [INPUT] File data AR bulanan
│                                               (nama bebas asal diawali "PTM")
│
└── 📄 Laporan_Analisis_Prosedur_Audit.xlsx  ← [OUTPUT] Dihasilkan otomatis
```

> Semua file berada dalam **satu folder**. Tidak ada subfolder.

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input PTM

Letakkan file `PTM*.xlsx` di folder yang sama dengan `Buat_Analisis_Audit.py`. File ini adalah ekspor AR bulanan dari Accurate dengan sheet berformat `MMYY` (mis. `0125` untuk Januari 2025).

> Jika sudah memiliki file `PTM*.xlsx` dari proyek `Monitoring AR Outstanding`, **file yang sama** dapat langsung digunakan tanpa perubahan apapun.

Syarat struktur file PTM: setiap sheet `MMYY` harus memiliki kolom-kolom berikut:

| Kolom wajib | Kolom opsional |
|---|---|
| `Sisa Piutang` | `Negara Pelanggan` |
| `Umur AR base on Tgl Faktur` | `Kontak Pelanggan` |
| `Nama Penjual` | `Nilai Faktur` |
| `Nama Pelanggan` | |
| `No. Faktur` | |

### Langkah 2 — (Opsional) Sesuaikan `config.conf`

`config.conf` dibaca skrip ini untuk membangun `filter_produk`, namun pada versi saat ini konfigurasi tersebut **tidak digunakan aktif** dalam kalkulasi output. File tetap dibaca dan dibuat otomatis agar kompatibel dengan skrip lain dalam ekosistem yang sama.

### Langkah 3 — Jalankan

```bash
python Buat_Analisis_Audit.py
```

Program akan menampilkan progress per sheet di terminal:

```
--> Memulai pemrosesan data audit analytics ...
--> Menyusun Sheet 1: Analisis Horisontal & Vertikal...
--> Menyusun Sheet 3: Deteksi Statistik Z-Score...
--> Menyusun Sheet 3: Forensik Digit Angka Hukum Benford...
--> Menyusun Sheet 4: Kosentrasi Risiko Piutang Hukum Pareto...
--> Menyusun Sheet 5: Uji Duplikasi Nomor Faktur...
--> Menyelaraskan ukuran lebar kolom seluruh berkas audit otomatis...
--> SUKSES! Berkas laporan analitis audit 'Laporan_Analisis_Prosedur_Audit.xlsx' telah diterbitkan.
```

### Langkah 4 — Buka laporan

Buka **`Laporan_Analisis_Prosedur_Audit.xlsx`** yang muncul di folder yang sama.

---

## 🔄 Alur Kerja Detail

```
[Mulai]
   │
   ├─── Inisialisasi
   │       Baca config.conf → ambil filter_produk
   │       Buat config.conf default jika tidak ada
   │       Cari PTM*.xlsx dengan glob → ambil yang pertama ditemukan
   │       Filter sheet MMYY valid → urutkan kronologis
   │
   ├─── Bangun dua dataset master (loop semua sheet MMYY)
   │       │
   │       ├─ df_macro_master  (satu baris per bulan, 7 metrik agregat)
   │       │     Sama dengan kalkulasi Monitoring AR:
   │       │     Total AR, AR 60 Hari Up, Bad Debt, jumlah customer distinct
   │       │     (dengan ekslusi FRAUD & filter days)
   │       │
   │       └─ df_micro_master  (semua baris transaksi, lintas semua bulan)
   │             + kolom tambahan: Bulan_Sistem, Cleaned_Sisa_Piutang,
   │               Cleaned_Nilai_Faktur, Cleaned_Days, Cleaned_Cust_ID
   │
   ├─── [Sheet 1] Analisis Horisontal & Vertikal
   │       Hitung MoM % change → Rasio AR 60 / Total → Rasio Bad Debt / Total
   │       Tulis tabel 7 kolom dengan format angka + persen
   │
   ├─── [Sheet 2] Anomali Z-Score
   │       Hitung mean & std Total AR seluruh periode
   │       Z-Score per bulan = (nilai - mean) / std
   │       Threshold |Z| > 1.2 → "ANOMALI / REVIU MENDALAM" + highlight merah
   │
   ├─── [Sheet 3] Forensik Hukum Benford
   │       Ekstrak digit pertama non-nol dari semua Cleaned_Nilai_Faktur
   │       Hitung frekuensi aktual (proporsi) per digit 1–9
   │       Bandingkan dengan distribusi Benford teoritis
   │       Threshold |deviasi| > 5% → "REVIU (Deviasi Tinggi)" + highlight merah
   │
   ├─── [Sheet 4] Pareto 80/20
   │       Filter hanya data bulan terakhir (latest_month)
   │       Groupby Cleaned_Cust_ID → sum Sisa Piutang → urutkan menurun
   │       Hitung Kontribusi% dan Kumulatif%
   │       Kumulatif ≤ 80% → "TOP 80% CORE RISK (Kritis)" + highlight merah
   │       Tampilkan top 25 pelanggan
   │
   ├─── [Sheet 5] Uji Duplikasi Faktur
   │       Normalisasi kolom df_micro_master dengan clean_key()
   │       Groupby (Bulan_Sistem + No. Faktur) → count
   │       Filter count > 1 → merge kembali ke detail transaksi
   │       Tampilkan semua baris duplikat dengan rekomendasi tindakan
   │
   ├─── Auto-fit lebar kolom semua sheet
   │
   └─── Simpan sebagai Laporan_Analisis_Prosedur_Audit.xlsx ✅
```

---

## 📊 Penjelasan 5 Sheet Audit

### Sheet 1 — Analisis Horisontal & Vertikal

**Tujuan audit:** Menilai volatilitas perubahan saldo bulanan dan komposisi risiko portofolio piutang.

**Analisis Horisontal (kolom MoM):**
Mengukur persentase perubahan Total AR Outstanding dari satu bulan ke bulan berikutnya menggunakan `pct_change()`. Peningkatan MoM yang konsisten tanpa diimbangi pembayaran adalah sinyal risiko penumpukan piutang.

**Analisis Vertikal (kolom Rasio):**
Mengukur proporsi AR 60 Hari Up dan Bad Debt terhadap Total AR Outstanding. Rasio yang meningkat dari bulan ke bulan mengindikasikan penurunan kualitas portofolio.

**Kolom yang dihasilkan:**

| Kolom | Format | Keterangan |
|---|---|---|
| Bulan | Teks | Label bulan (mis. `Jan-25`) |
| Total AR Outstanding | `#,##0` | Seluruh sisa piutang bulan ini |
| Perubahan MoM (%) | `0.00%` | Perubahan terhadap bulan sebelumnya. Baris pertama = 0% |
| AR 60 Hari Up | `#,##0` | Piutang umur 60–364 hari (non-FRAUD) |
| Rasio Kontribusi 60+ (%) | `0.00%` | AR 60 Hari Up ÷ Total AR |
| AR Bad Debt (365+) | `#,##0` | Piutang umur ≥ 365 hari (non-FRAUD) |
| Rasio Kontribusi Bad Debt (%) | `0.00%` | Bad Debt ÷ Total AR |

---

### Sheet 2 — Anomali Z-Score Makro

**Tujuan audit:** Menemukan bulan dengan lonjakan atau penurunan saldo ekstrim yang berada di luar batas fluktuasi normal secara statistik.

**Cara kerja Z-Score:**

```
Z-Score (bulan X) = (Total AR bulan X - Rata-rata semua bulan) / Std Dev semua bulan
```

Nilai Z-Score mengukur seberapa jauh saldo bulan tersebut menyimpang dari rata-rata historis, dalam satuan standar deviasi.

| Rentang Z-Score | Interpretasi |
|---|---|
| `-1.2` s.d. `+1.2` | **WAJAR** — fluktuasi dalam batas normal |
| `< -1.2` atau `> +1.2` | **ANOMALI / REVIU MENDALAM** — perlu investigasi penyebab |

Baris dengan status `ANOMALI` diberi **highlight merah-oranye** dan **teks tebal** secara otomatis.

**Kolom yang dihasilkan:**

| Kolom | Format | Keterangan |
|---|---|---|
| Bulan | Teks | Label bulan |
| Total AR Outstanding | `#,##0` | Saldo bulan tersebut |
| Deviasi Nilai dari Rata-Rata (Z-Score) | `0.00` | Nilai Z-Score |
| Status Kesimpulan Audit | Teks | `WAJAR` atau `ANOMALI / REVIU MENDALAM` |

---

### Sheet 3 — Forensik Hukum Benford

**Tujuan audit:** Mendeteksi indikasi manipulasi data atau faktur fiktif melalui pengujian distribusi digit pertama seluruh nilai faktur.

**Dasar teori:** Hukum Benford menyatakan bahwa dalam kumpulan data keuangan organik yang besar, digit pertama angka (1–9) mengikuti distribusi logaritmik yang dapat diprediksi. Data yang telah dimanipulasi cenderung menyimpang dari distribusi ini karena pola buatan manusia tidak mengikuti alam data alami.

**Distribusi Benford teoritis yang digunakan:**

| Digit | Ekspektasi Benford |
|---|---|
| 1 | 30,1% |
| 2 | 17,6% |
| 3 | 12,5% |
| 4 | 9,7% |
| 5 | 7,9% |
| 6 | 6,7% |
| 7 | 5,8% |
| 8 | 5,1% |
| 9 | 4,6% |

**Logika ekstraksi digit pertama:**
Dari kolom `Cleaned_Nilai_Faktur`, semua karakter non-digit dihapus, lalu karakter pertama yang bukan `0` diambil sebagai digit pertama. Baris dengan nilai `0` atau yang tidak bisa diekstrak diabaikan dari analisis.

**Threshold:** `|Proporsi Riil - Ekspektasi Benford| > 5%` → `"REVIU (Deviasi Tinggi)"` + highlight merah.

**Kolom yang dihasilkan:**

| Kolom | Format | Keterangan |
|---|---|---|
| Komponen Digit Pertama | Teks | `Digit 1` s.d. `Digit 9` |
| Proporsi Riil Faktur (%) | `0.00%` | Frekuensi aktual dari data |
| Ekspektasi Hukum Benford (%) | `0.00%` | Nilai teoritis Benford |
| Nilai Deviasi Selisih | `0.00%` | Selisih (Riil - Ekspektasi). Negatif = lebih sedikit dari ekspektasi |
| Status Evaluasi Data | Teks | `NORMAL` atau `REVIU (Deviasi Tinggi)` |

---

### Sheet 4 — Risiko Pareto Customer

**Tujuan audit:** Mengidentifikasi pelanggan-pelanggan inti yang secara kolektif menguasai 80% saldo piutang, untuk menilai risiko konsentrasi dan gagal bayar massal.

**Prinsip Pareto 80/20:** Dalam banyak portofolio piutang, sebagian kecil pelanggan menyumbang sebagian besar saldo. Pelanggan yang masuk dalam 80% kumulatif pertama adalah titik risiko paling kritis yang harus diprioritaskan dalam penagihan dan pemantauan.

**Sumber data:** Hanya menggunakan data bulan **terakhir** yang tersedia di file PTM (bulan paling baru secara kronologis), agar mencerminkan kondisi risiko terkini.

**Urutan proses:**
1. Groupby `Cleaned_Cust_ID` → sum `Cleaned_Sisa_Piutang`
2. Urutkan dari saldo terbesar ke terkecil
3. Hitung `Kontribusi%` = saldo pelanggan ÷ total AR bulan tersebut
4. Hitung `Kumulatif%` = akumulasi kumulatif Kontribusi%
5. Tandai pelanggan dengan `Kumulatif% ≤ 80%` sebagai `TOP 80% CORE RISK (Kritis)`

> Sheet hanya menampilkan **25 pelanggan teratas** (top 25 by saldo).

**Kolom yang dihasilkan:**

| Kolom | Format | Keterangan |
|---|---|---|
| ID/Nama Pelanggan | Teks | Cleaned_Cust_ID (kode atau nama) |
| Total Saldo Piutang (Bulan) | `#,##0` | Total sisa piutang pelanggan ini |
| Proporsi Beban Risiko (%) | `0.00%` | Kontribusi pelanggan terhadap total AR |
| Akumulasi Distribusi Kumulatif (%) | `0.00%` | Kumulatif dari pelanggan 1 s.d. baris ini |
| Grup Prioritas Pengawasan | Teks | `TOP 80% CORE RISK (Kritis)` atau `Regular` |

Baris dengan klasifikasi `Kritis` diberi **highlight merah-oranye** dan **teks tebal**.

---

### Sheet 5 — Uji Duplikasi Faktur

**Tujuan audit:** Mendeteksi nomor faktur yang diinput lebih dari sekali dalam bulan yang sama — indikasi `double input` atau `double journal` yang akan menggelembungkan saldo piutang.

**Logika deteksi:**
1. Gabungkan seluruh transaksi dari semua bulan (`df_micro_master`)
2. Normalisasi nama kolom dengan `clean_key()`
3. Groupby `(Bulan_Sistem + No. Faktur)` → hitung jumlah kemunculan
4. Filter hanya yang count > 1 (duplikat intra-bulan)
5. Merge kembali ke baris detail untuk menampilkan informasi lengkap

**Catatan:** Pengujian dilakukan secara **intra-bulan** — nomor faktur yang sama muncul di bulan berbeda (mis. faktur belum lunas yang carry-over) tidak dianggap duplikat. Hanya kemunculan ganda dalam satu bulan yang sama yang ditandai.

**Kolom yang dihasilkan:**

| Kolom | Format | Keterangan |
|---|---|---|
| Nomor Faktur | Teks | Nomor faktur yang terdeteksi duplikat |
| Jumlah Kemunculan Di Bulan Ini | `#,##0` | Berapa kali nomor ini muncul di bulan tersebut |
| ID/Nama Pelanggan | Teks | Nama pelanggan dari baris data |
| Bulan Transaksi | Teks | Bulan di mana duplikasi terjadi |
| Nilai Faktur | `#,##0` | Nilai faktur asli |
| Sisa Piutang | `#,##0` | Sisa piutang baris tersebut |
| Rekomendasi Tindakan | Teks | `REVIU JURNAL (Duplikasi Intra-Bulan)` |

Jika kolom `No. Faktur` tidak ditemukan di data, sheet menampilkan satu baris pesan: `"Kolom Nomor Faktur tidak ditemukan di dalam berkas data sumber."`

---

## ⚙️ Konfigurasi `config.conf`

```ini
[FILTER]
SHELL
IRC
ZN
GT
FILTER
JIMCO
LAIN
TOP 1
OLI
```

File dibaca saat startup untuk membangun daftar `filter_produk`. Dalam versi skrip ini, konfigurasi tersebut **tidak menghasilkan tabel kontribusi produk** — berbeda dengan `Buat_Data.py` (Monitoring AR) yang menggunakannya untuk mengelompokkan AR-60 dan Bad Debt per produk.

Jika `config.conf` tidak ada saat skrip dijalankan, file akan dibuat otomatis dengan isi default: `SHELL, IRC, ZN, GT, FILTER, JIMCO, LAIN, TOP 1, OLI`.

---

## ⚠️ Catatan Teknis & Perbedaan dari Monitoring AR

Skrip ini dan `Buat_Data.py` (Monitoring AR Outstanding) berbagi input yang sama namun memiliki satu perbedaan kritis pada nama kolom:

| Aspek | `Buat_Analisis_Audit.py` (skrip ini) | `Buat_Data.py` (Monitoring AR) |
|---|---|---|
| Sumber input | `PTM*.xlsx` | `PTM*.xlsx` |
| Output | `Laporan_Analisis_Prosedur_Audit.xlsx` | `Monitoring.xlsx` |
| Filter produk | Dibaca tapi tidak dipakai aktif | Dipakai untuk tabel kontribusi produk |
| Basis data Pareto | Hanya bulan terakhir | Semua bulan (rekap per bulan) |

> Jika file PTM hanya memiliki salah satu dari kedua nama kolom tersebut, salah satu skrip akan menghasilkan metrik aging yang kosong/nol. Pastikan nama kolom di file Accurate konsisten.

---

## 🛠️ Troubleshooting

### ❌ `File PTM*.xlsx tidak ditemukan`
Pastikan file data Accurate ada di folder yang sama dengan skrip dan namanya diawali `PTM`. Skrip mengambil file pertama yang ditemukan oleh `glob("PTM*.xlsx")`.

### ❌ `Tidak ada sheet bulanan (format MMYY) yang valid`
Skrip mencari sheet dengan nama tepat 4 digit (`0125`, `1224`, dst.). Periksa nama sheet di file PTM — tidak boleh ada spasi, huruf, atau karakter tambahan.

### ❌ Seluruh kolom AR 60, Bad Debt, dan customer count bernilai 0
Kemungkinan besar kolom `"Umur AR base on Tgl Faktur"` tidak ditemukan di file PTM. Cek nama kolom umur di file Accurate dan bandingkan dengan nilai `umur_japo_col` di baris:
```python
umur_japo_col = clean_key("Umur AR base on Tgl Faktur")
```
Fungsi `clean_key()` menghapus semua karakter non-alphanumeric dan mengubah ke lowercase, jadi `"Umur AR base on Tgl Faktur"` menjadi `"umarbaseontglfaktur"`. Pastikan nama kolom di file mengandung kata-kata yang sama setelah normalisasi.

### ❌ Sheet 3 (Benford) menampilkan semua `DATA TIDAK CUKUP`
Terjadi jika kolom `Nilai Faktur` tidak ditemukan atau semua nilainya kosong/nol, sehingga tidak ada digit pertama yang bisa diekstrak. Cek keberadaan kolom `nilai_faktur_col` di file PTM.

### ❌ Sheet 4 (Pareto) hanya menampilkan satu atau dua pelanggan
Analisis Pareto menggunakan hanya data bulan **terakhir** yang tersedia. Jika sheet terakhir memiliki sedikit data atau pelanggan, jumlah baris di sheet ini akan sedikit. Ini bukan error — ini mencerminkan kondisi data aktual bulan terakhir.

### ❌ Sheet 5 (Duplikasi) kosong / tidak ada temuan duplikat
Ini adalah kondisi normal dan merupakan hasil yang baik — berarti tidak ada nomor faktur yang diinput ganda dalam bulan yang sama. Sheet tetap dibuat namun berisi header kolom tanpa baris data.

### ❌ `Laporan_Analisis_Prosedur_Audit.xlsx` tertimpa setiap run
Ini adalah perilaku yang disengaja — setiap eksekusi menghasilkan laporan segar. Simpan salinan dengan nama berbeda jika hasil run sebelumnya perlu dipertahankan.

---

## 📌 Catatan Penting

- **Jalankan dari folder yang berisi file PTM** — Skrip menggunakan `glob("PTM*.xlsx")` yang mencari di direktori kerja saat ini, bukan path absolut.
- **File output akan ditimpa setiap run** — Tidak ada pengecekan apakah `Laporan_Analisis_Prosedur_Audit.xlsx` sudah ada; file lama langsung diganti.
- **Z-Score dihitung dari semua data yang tersedia** — Semakin banyak bulan data di file PTM, semakin akurat threshold Z-Score dalam mencerminkan pola historis. Dengan data ≤ 3 bulan, Z-Score kurang bermakna statistik.
- **Hukum Benford paling efektif untuk dataset besar** — Uji Benford memberikan sinyal yang lebih kuat jika jumlah transaksi di seluruh periode cukup besar (idealnya ≥ 1.000 faktur). Hasil pada dataset kecil perlu diinterpretasikan dengan hati-hati.
- **Pareto berbasis snapshot bulan terakhir** — Konsentrasi risiko bisa berubah dramatis dari bulan ke bulan. Untuk tren konsentrasi risiko per bulan, gunakan data dari `Monitoring AR Outstanding`.
- **Duplikasi intra-bulan vs carry-over** — Faktur yang sama muncul di dua bulan berbeda (karena belum lunas) **tidak** dianggap duplikat oleh skrip ini. Hanya kemunculan ganda dalam satu bulan yang sama yang dilaporkan.
- **Kolom `Kontak Pelanggan` tidak dipakai di skrip ini** — Meskipun dibaca sebagai bagian dari `df_micro_master`, kolom ini tidak digunakan dalam kelima sheet output. Berbeda dengan `Buat_Data.py` yang menggunakannya untuk tabel kontribusi produk.

---

## 👤 Author

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
