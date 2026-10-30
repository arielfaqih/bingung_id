# BINGUNG-ID

**Penambangan Komentar YouTube Berbasis IndoBERT untuk Mendeteksi dan Memetakan Kebingungan Belajar Matematika dan Sains Menuju Education 5.0**

Tim carcircur · Universitas Pendidikan Indonesia · Finalis DIMAS-TI 2026, Divisi Penambangan Data (AMLI · FMIPA Universitas Negeri Gorontalo)

BINGUNG-ID menambang komentar penonton video pembelajaran Matematika, Fisika, Kimia, dan Biologi di YouTube, mengklasifikasikannya menjadi **BINGUNG**, **KOREKSI**, **PAHAM**, atau **LAIN**, lalu menghitung *Indeks Kebingungan* per mata pelajaran, topik, dan posisi dalam video.

## Hasil utama

| Ukuran | Nilai |
| --- | --- |
| Video / komentar mentah | 278 video / 82.661 komentar |
| Komentar penonton setelah pembersihan | 70.829 (63.935 teks unik) |
| Dataset berlabel | 1.498 komentar dari 241 video (BINGUNG 527, LAIN 522, PAHAM 393, KOREKSI 56) |
| Kesepakatan anotator (irisan 300) | Fleiss' kappa 0,852 |
| Model terbaik | IndoBERT-base-p1, macro-F1 0,774 ± 0,010 (3 benih), akurasi 0,826 |
| TF-IDF + Regresi Logistik | macro-F1 0,732 (CV per video 0,724 ± 0,017) |
| Indeks Kebingungan (68.175 komentar) | 27,3% · Matematika 32,3%, Fisika 31,5%, Kimia 27,3%, Biologi 19,8% |
| Topik paling membingungkan | hidrolisis garam 52%, asam basa 48%, sistem imun 44% |

Semua evaluasi memakai pembagian **per video**: komentar dari satu video tidak pernah ada di data latih dan uji sekaligus.

## Struktur repositori

```
bingung_id/
├── bingung_scraper.py                # penambang: cari, ambil, bersihkan, sampel, kappa
├── kata_kunci.txt                    # 32 kata kunci, 8 per mata pelajaran
├── notebooks/
│   ├── 01_penambangan_data.ipynb     # log eksekusi Tahap 1–4 beserta keluarannya
│   └── 02_model_dan_pemetaan.ipynb   # Tahap 5–6: TF-IDF, IndoBERT, evaluasi, pemetaan, simpan model
├── scripts/
│   └── prediksi.py                   # beri label CSV komentar baru (data uji panitia / leaderboard)
├── hasil_ringkas/                    # tabel & gambar agregat (tanpa teks komentar)
├── docs/
│   └── panduan_scraper.md            # panduan lengkap bingung_scraper.py untuk tim
├── requirements.txt
└── README.md
```

Data komentar (`data/`, `anotasi/`) dan bobot model (`hasil/model_transformer/`, ±500 MB) **sengaja tidak diunggah** demi privasi penonton dan ukuran berkas. Keduanya tersimpan di Google Drive tim (`MyDrive/DIMAS_TI_2026`).

## Cara menjalankan

### 1. Penambangan (Google Colab)

```python
from google.colab import drive
drive.mount('/content/drive')
%cd /content/drive/MyDrive/DIMAS_TI_2026
import os, getpass
os.environ["YT_API_KEY"] = getpass.getpass("API key: ")   # kunci tidak pernah ditulis ke berkas
```

```bash
python bingung_scraper.py cari --kata-file kata_kunci.txt          # 32 kata kunci -> kandidat_video.xlsx
# periksa kandidat, isi kolom pakai = Y, simpan sebagai videos.xlsx
python bingung_scraper.py ambil --videos videos.xlsx --balasan     # -> data/komentar_mentah.csv
python bingung_scraper.py bersihkan                                # -> data/komentar_bersih.csv
python bingung_scraper.py sampel --n 1500 --irisan 300 --anotator Rani Budi Sari
python bingung_scraper.py kappa --lembar anotasi/anotasi_*.xlsx    # -> data/dataset_berlabel.csv
```

### 2. Pemodelan dan pemetaan

Buka `notebooks/02_model_dan_pemetaan.ipynb` di Colab, pilih **Runtime → Change runtime type → T4 GPU**, lalu **Run all** (±15–20 menit). Tabel, gambar, dan model tersimpan di `hasil/`.

### 3. Prediksi CSV baru (babak final)

```bash
pip install -r requirements.txt
python scripts/prediksi.py data_uji_panitia.csv --kolom text --keluar hasil/prediksi.csv
# tanpa GPU / tanpa model transformer:
python scripts/prediksi.py data_uji_panitia.csv --kolom text --model tfidf
```

Keluaran menambahkan kolom `prediksi` (BINGUNG/KOREKSI/PAHAM/LAIN) dan `keyakinan`. Skrip menerapkan anonimisasi yang sama dengan saat penambangan (tautan → `<URL>`, sebutan akun → `@pengguna`).

## Etika

Hanya komentar publik yang diambil lewat YouTube Data API v3. Identitas akun diganti hash SHA-256 bergaram sebelum data dibuka anotator, dan hasil dilaporkan per topik, bukan per penonton atau kanal.
