# BINGUNG-ID · Skrip Pengambil Komentar

Satu skrip (`bingung_scraper.py`) untuk seluruh tahap data BINGUNG-ID, dari mencari video sampai label final siap latih.

```
cari ──► ambil ──► bersihkan ──► sampel ──► (anotasi manual) ──► kappa
```

## 0. Siapkan API key (±5 menit, cukup satu orang)

1. Buka <https://console.cloud.google.com/>, buat project baru (mis. `bingung-id`).
2. **APIs & Services → Library**, cari **YouTube Data API v3**, klik **Enable**.
3. **APIs & Services → Credentials → Create credentials → API key**. Salin kuncinya.
4. Jangan unggah kunci ke GitHub atau grup chat publik.

Kuota gratis: **10.000 unit/hari per project**, reset sekitar **14.00 WIB**.
Gunakan satu project untuk tim. Jangan membuat banyak project hanya untuk menggandakan kuota, karena itu melanggar ketentuan YouTube API.

| Perintah | Biaya kuota |
|---|---|
| `cari` | 100 unit per kata kunci (32 kata kunci ≈ 3.200 unit) |
| `ambil` | ±1 unit per 100 komentar (80 video × 2.000 komentar ≈ 1.600 unit) |
| `bersihkan`, `sampel`, `kappa` | 0 (offline) |

## 1. Jalankan di Google Colab (disarankan)

Simpan semuanya di Google Drive bersama supaya data tidak hilang saat Colab terputus dan bisa diakses seluruh tim.

```python
from google.colab import drive
drive.mount('/content/drive')
%mkdir -p /content/drive/MyDrive/BINGUNG-ID
%cd /content/drive/MyDrive/BINGUNG-ID
# unggah bingung_scraper.py dan kata_kunci.txt ke folder ini (panel Files di kiri)

import os, getpass
os.environ["YT_API_KEY"] = getpass.getpass("API key: ")
```

Setiap perintah di bawah dijalankan di sel Colab dengan awalan `!`.

## 2. Langkah kerja

**a. Cari kandidat video** (pagi, ±3 menit)
```
!python bingung_scraper.py cari --kata-file kata_kunci.txt
```
Hasilnya `kandidat_video.xlsx`, berisi video berkomentar ≥50, sudah terisi mapel dan topik.
Buka di Excel atau Google Sheets, lalu isi kolom **pakai = Y** untuk 15–25 video per mapel. Pilih video:
- yang benar-benar menjelaskan materi (bukan kompilasi atau motivasi),
- dari beberapa kanal berbeda,
- dengan komentar yang ramai.

Simpan sebagai `videos.xlsx`. Daftar sendiri juga bisa: cukup kolom `url`, `mapel`, `jenjang`, `topik`.

**b. Ambil komentar**
```
!python bingung_scraper.py ambil --videos videos.xlsx --balasan
```
Hasilnya `data/komentar_mentah.csv`. Kalau kuota habis atau Colab putus, jalankan ulang perintah yang sama. Video yang sudah selesai akan dilewati.

**c. Bersihkan**
```
!python bingung_scraper.py bersihkan
```
Langkah ini:
- menyamarkan @akun dan URL,
- membuang posting ganda, spam salin-tempel, komentar kreator, dan komentar emoji saja,
- mengekstrak timestamp (`5:32`, `1:02:33`, `menit 5.32`, `menit ke 3`),
- menandai komentar yang mengandung sinyal bingung.

Hasilnya `data/komentar_bersih.csv` dan `data/ringkasan_dataset.csv`. File ringkasan langsung bisa dipakai sebagai Tabel Dataset di makalah.

**d. Buat lembar anotasi**
```
!python bingung_scraper.py sampel --n 1500 --irisan 300 --anotator Rani Budi Sari
```
Hasilnya `anotasi/anotasi_<Nama>.xlsx`, sekitar 700 komentar per orang:
- 300 komentar irisan dinilai **semua** anotator untuk uji kappa, tercampur acak dan tidak ditandai.
- Sisanya dibagi rata.

Rincian sampel:
- Jatah per mapel sama rata.
- 40% berasal dari komentar bersinyal bingung, supaya kelas BINGUNG tidak terlalu sedikit. Cantumkan strategi ini di bagian metode makalah.
- Maksimal 30 komentar per video.

`data/sampel_anotasi.csv` adalah kunci sampel. **Jangan dibagikan ke anotator.**

**e. Hitung kappa dan gabungkan label**
```
!python bingung_scraper.py kappa --lembar anotasi/anotasi_Rani.xlsx anotasi/anotasi_Budi.xlsx anotasi/anotasi_Sari.xlsx
```
Hasilnya:
- `data/laporan_kappa.txt` berisi Cohen's kappa per pasangan, Fleiss' kappa, dan label yang paling sering tertukar.
- `data/ketidaksepakatan.csv` berisi komentar yang perlu didiskusikan.
- `data/dataset_berlabel.csv` berisi label final dengan suara mayoritas.

Setelah sesi diskusi, isi kolom `label_final` (dan `jenis_final` bila BINGUNG) di `ketidaksepakatan.csv`, lalu jalankan:
```
!python bingung_scraper.py kappa --lembar anotasi/anotasi_*.xlsx --putusan data/ketidaksepakatan.csv
```

## 3. Catatan untuk makalah dan etika

- Nama dan ID kanal komentator **tidak disimpan**. Yang disimpan hanya hash SHA-256 bergaram, yaitu kode acak satu arah sehingga ID asli tidak bisa dikembalikan. Tulis ini di bagian Metode sebagai prosedur anonimisasi.
- Teks identik hanya muncul sekali di sampel anotasi, supaya tidak ada kebocoran antara data latih dan data uji.
- Saat membagi train/test nanti, **bagi per `video_id`** (GroupShuffleSplit atau GroupKFold).
- Data diambil lewat API resmi. Dalam presentasi atau lampiran, jangan menampilkan komentar beserta identitas penulisnya.

## 4. Kalau ada galat

| Pesan | Artinya |
|---|---|
| `API key ditolak` | Kunci salah, atau YouTube Data API v3 belum di-Enable di project itu. |
| `kuota habis` | Tunggu reset ±14.00 WIB, lalu jalankan ulang. Progres aman. |
| `komentar_dimatikan` di log | Video itu menutup komentar, jadi otomatis dilewati. |
| Excel menampilkan semua isi CSV di satu kolom | Buka lewat Google Sheets, atau Data → From Text/CSV. |
