#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BINGUNG-ID · Pengambil & penyiap data komentar YouTube
DIMAS-TI AMLI 2026 — Divisi Penambangan Data

Perintah (jalankan berurutan):
  cari       Cari kandidat video per kata kunci        -> kandidat_video.xlsx       (100 unit kuota / kata kunci)
  ambil      Ambil metadata + komentar video terpilih  -> data/komentar_mentah.csv  (±1 unit / 100 komentar)
  bersihkan  Anonimisasi, buang duplikat/spam, ekstrak timestamp -> data/komentar_bersih.csv
  sampel     Buat lembar anotasi Excel per anotator (+ irisan rahasia untuk uji kappa)
  kappa      Hitung kesepakatan antar-anotator & gabungkan label -> data/dataset_berlabel.csv

Contoh singkat:
  export YT_API_KEY="AIza...."
  python bingung_scraper.py cari --kata-file kata_kunci.txt
  python bingung_scraper.py ambil --videos videos.xlsx
  python bingung_scraper.py bersihkan
  python bingung_scraper.py sampel --n 1500 --irisan 300 --anotator Rani Budi Sari
  python bingung_scraper.py kappa --lembar anotasi/anotasi_Rani.xlsx anotasi/anotasi_Budi.xlsx anotasi/anotasi_Sari.xlsx

Kebutuhan: Python 3.8+, pandas, requests, openpyxl (semuanya sudah ada di Google Colab).
"""
from __future__ import annotations

import argparse
import hashlib
import html
import json
import math
import os
import random
import re
import sys
import time
from collections import Counter
from itertools import combinations
from pathlib import Path

import pandas as pd
import requests

# --------------------------------------------------------------------------------------
# Konstanta
# --------------------------------------------------------------------------------------
API_URL = "https://www.googleapis.com/youtube/v3/"
LABEL = ["BINGUNG", "KOREKSI", "PAHAM", "LAIN"]
JENIS = ["KONSEPTUAL", "PROSEDURAL", "NOTASI", "UMUM"]
SINGKATAN_LABEL = {"B": "BINGUNG", "K": "KOREKSI", "P": "PAHAM", "L": "LAIN"}
SINGKATAN_JENIS = {"K": "KONSEPTUAL", "P": "PROSEDURAL", "N": "NOTASI", "U": "UMUM"}
KUOTA_HABIS = {"quotaExceeded", "dailyLimitExceeded"}
COBA_LAGI = {"rateLimitExceeded", "userRateLimitExceeded", "backendError", "internalError"}

# Garam untuk hash ID penulis. Ganti lewat variabel lingkungan BINGUNG_GARAM bila perlu,
# tapi pakai nilai yang SAMA untuk seluruh tim agar hash konsisten.
GARAM = os.environ.get("BINGUNG_GARAM", "bingung-id-dimasti-2026")

PESAN_TANPA_KUNCI = """
API key YouTube belum diisi.
  1. Buka https://console.cloud.google.com/ -> buat project baru.
  2. Menu "APIs & Services" -> "Library" -> aktifkan "YouTube Data API v3".
  3. "APIs & Services" -> "Credentials" -> "Create credentials" -> "API key".
  4. Jalankan:  export YT_API_KEY="kunci_kalian"   (atau tambahkan --api-key kunci_kalian)
     Di Colab:  import os, getpass; os.environ["YT_API_KEY"] = getpass.getpass("API key: ")
"""

PESAN_KUOTA = """
Kuota harian YouTube API habis (10.000 unit/hari per project).
Progres sudah tersimpan. Video yang selesai tidak akan diambil ulang.
Kuota direset tengah malam waktu Pasifik (sekitar pukul 14.00-15.00 WIB).
Jalankan perintah yang sama lagi setelah reset.
"""

API_KEY = ""
SESSION = requests.Session()
KUOTA_TERPAKAI = 0


def info(pesan: str) -> None:
    print(pesan, flush=True)


# --------------------------------------------------------------------------------------
# Akses API
# --------------------------------------------------------------------------------------
class ApiError(Exception):
    def __init__(self, status: int, alasan: str, pesan: str):
        super().__init__(f"HTTP {status} · {alasan} · {pesan}")
        self.status, self.alasan, self.pesan = status, alasan, pesan


class KuotaHabis(Exception):
    pass


def _alasan_galat(r: requests.Response):
    try:
        err = r.json().get("error", {})
        errs = err.get("errors") or [{}]
        return errs[0].get("reason", "") or str(err.get("status", "")), err.get("message", "")
    except ValueError:
        return "", r.text[:200]


def api_get(endpoint: str, params: dict, biaya: int = 1) -> dict:
    """Panggil YouTube Data API v3 dengan coba-ulang otomatis untuk galat sementara."""
    global KUOTA_TERPAKAI
    if not API_KEY:
        sys.exit(PESAN_TANPA_KUNCI)
    p = dict(params, key=API_KEY)
    for percobaan in range(6):
        try:
            r = SESSION.get(API_URL + endpoint, params=p, timeout=30)
        except requests.RequestException as e:
            tunggu = 2 ** percobaan
            info(f"  ! koneksi gagal ({e.__class__.__name__}), coba lagi dalam {tunggu} detik")
            time.sleep(tunggu)
            continue
        KUOTA_TERPAKAI += biaya  # permintaan gagal pun tetap memakan kuota
        if r.status_code == 200:
            return r.json()
        alasan, pesan = _alasan_galat(r)
        if alasan in KUOTA_HABIS:
            raise KuotaHabis(pesan)
        if r.status_code >= 500 or alasan in COBA_LAGI:
            time.sleep(2 ** percobaan + random.random())
            continue
        raise ApiError(r.status_code, alasan, pesan)
    raise ApiError(0, "gagalBerulang", f"{endpoint} gagal setelah 6 percobaan")


# --------------------------------------------------------------------------------------
# Utilitas
# --------------------------------------------------------------------------------------
RE_ID = re.compile(r"(?:v=|youtu\.be/|shorts/|embed/|live/|/v/)([A-Za-z0-9_-]{11})")
RE_DURASI = re.compile(r"P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?")


def ekstrak_id_video(s) -> str:
    s = str(s).strip()
    m = RE_ID.search(s)
    if m:
        return m.group(1)
    return s if re.fullmatch(r"[A-Za-z0-9_-]{11}", s) else ""


def durasi_detik(iso: str) -> int:
    m = RE_DURASI.fullmatch(iso or "")
    if not m:
        return 0
    d, h, mi, s = (int(x) if x else 0 for x in m.groups())
    return d * 86400 + h * 3600 + mi * 60 + s


def hash_penulis(cid: str) -> str:
    return hashlib.sha256((GARAM + cid).encode("utf-8")).hexdigest()[:12] if cid else ""


def baca_tabel(path) -> pd.DataFrame:
    """Baca .xlsx atau .csv (pemisah koma/titik koma dideteksi otomatis)."""
    path = Path(path)
    if not path.exists():
        sys.exit(f"File tidak ditemukan: {path}")
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str).fillna("")
    return pd.read_csv(path, dtype=str, sep=None, engine="python", encoding="utf-8-sig").fillna("")


def simpan_csv(df: pd.DataFrame, path) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, encoding="utf-8-sig")  # utf-8-sig agar emoji/aksara aman di Excel


def ke_int(df: pd.DataFrame, kolom) -> pd.DataFrame:
    for c in kolom:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    return df


def nama_aman(s: str) -> str:
    return re.sub(r"[^\w-]+", "_", s.strip()) or "anotator"


RE_KONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def isi_sel_aman(ws) -> None:
    """Buang karakter kontrol & cegah teks berawalan '=' dibaca Excel sebagai rumus."""
    for baris in ws.iter_rows(min_row=2):
        for sel in baris:
            if isinstance(sel.value, str):
                v = RE_KONTROL.sub("", sel.value)
                sel.value = v
                if v.startswith("="):
                    sel.data_type = "s"


# --------------------------------------------------------------------------------------
# Metadata video
# --------------------------------------------------------------------------------------
def ambil_metadata(ids) -> list:
    hasil = []
    ids = list(dict.fromkeys(i for i in ids if i))
    for i in range(0, len(ids), 50):
        res = api_get("videos", {"part": "snippet,contentDetails,statistics", "id": ",".join(ids[i:i + 50])})
        for it in res.get("items", []):
            sn, st, cd = it.get("snippet", {}), it.get("statistics", {}), it.get("contentDetails", {})
            komen = st.get("commentCount")
            hasil.append({
                "video_id": it["id"],
                "url": f"https://www.youtube.com/watch?v={it['id']}",
                "judul": sn.get("title", ""),
                "kanal": sn.get("channelTitle", ""),
                "id_kanal": sn.get("channelId", ""),
                "tanggal_unggah": sn.get("publishedAt", "")[:10],
                "durasi_detik": durasi_detik(cd.get("duration", "")),
                "views": int(st.get("viewCount", 0) or 0),
                "jumlah_komentar": int(komen) if komen is not None else -1,  # -1 = disembunyikan/dimatikan
                "bahasa": sn.get("defaultAudioLanguage") or sn.get("defaultLanguage") or "",
            })
    return hasil


# --------------------------------------------------------------------------------------
# 1) CARI
# --------------------------------------------------------------------------------------
def baca_kata_kunci(a) -> list:
    """Kembalikan list (mapel, kata). File boleh memakai baris [Mapel] sebagai penanda bagian."""
    pasangan = [("", k) for k in (a.kata or [])]
    if a.kata_file:
        mapel = ""
        for baris in Path(a.kata_file).read_text(encoding="utf-8").splitlines():
            baris = baris.strip()
            if not baris or baris.startswith("#"):
                continue
            m = re.fullmatch(r"\[(.+)\]", baris)
            if m:
                mapel = m.group(1).strip()
            else:
                pasangan.append((mapel, baris))
    return pasangan


def cmd_cari(a) -> None:
    pasangan = baca_kata_kunci(a)
    if not pasangan:
        sys.exit("Isi --kata \"limit fungsi\" ... atau --kata-file kata_kunci.txt")
    per_kata = math.ceil(a.maks / 50)
    info(f"Mencari {len(pasangan)} kata kunci (perkiraan {len(pasangan) * per_kata * 100} unit kuota)...")
    asal = {}
    for mapel, k in pasangan:
        token, dapat = None, []
        while len(dapat) < a.maks:
            p = {"part": "snippet", "q": k, "type": "video", "maxResults": min(50, a.maks - len(dapat)),
                 "regionCode": "ID", "relevanceLanguage": "id", "order": a.urutan}
            if a.durasi != "any":
                p["videoDuration"] = a.durasi
            if token:
                p["pageToken"] = token
            res = api_get("search", p, biaya=100)
            dapat += [it["id"]["videoId"] for it in res.get("items", []) if it.get("id", {}).get("videoId")]
            token = res.get("nextPageToken")
            if not token:
                break
        for v in dapat:
            asal.setdefault(v, (mapel, k))
        info(f"  · [{mapel or '-'}] {k}: {len(dapat)} video")

    meta = pd.DataFrame(ambil_metadata(list(asal)))
    if meta.empty:
        sys.exit("Tidak ada video ditemukan.")
    meta["mapel"] = meta["video_id"].map(lambda v: asal[v][0])
    meta["topik"] = meta["video_id"].map(lambda v: asal[v][1])
    meta["pakai"] = ""
    meta["jenjang"] = ""
    sebelum = len(meta)
    meta = meta[meta["jumlah_komentar"] >= a.min_komentar]
    meta = meta.sort_values(["mapel", "topik", "jumlah_komentar"], ascending=[True, True, False])
    urutan = ["pakai", "mapel", "jenjang", "topik", "judul", "kanal", "jumlah_komentar", "views",
              "durasi_detik", "tanggal_unggah", "url", "video_id", "id_kanal", "bahasa"]
    meta = meta[urutan]
    simpan_xlsx_rapi(meta, a.keluaran, lebar={"pakai": 7, "mapel": 12, "jenjang": 9, "topik": 22, "judul": 55,
                                              "kanal": 22, "url": 42},
                     dropdown={"pakai": ["Y", "N"], "jenjang": ["SMP", "SMA", "SD"]})
    info(f"\n{len(meta)} video (dari {sebelum}, disaring min. {a.min_komentar} komentar) -> {a.keluaran}")
    info("Langkah berikut: buka file itu, isi kolom 'pakai' = Y untuk video yang dipilih (periksa mapel/jenjang/topik),")
    info("simpan sebagai videos.xlsx, lalu jalankan:  python bingung_scraper.py ambil --videos videos.xlsx")


def simpan_xlsx_rapi(df: pd.DataFrame, path, lebar=None, dropdown=None) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Video"
    ws.append(list(df.columns))
    for r in df.itertuples(index=False):
        ws.append(list(r))
    isi_sel_aman(ws)
    for j, c in enumerate(df.columns, 1):
        sel = ws.cell(row=1, column=j)
        sel.font = Font(bold=True, color="FFFFFF")
        sel.fill = PatternFill("solid", fgColor="1F3A5F")
        ws.column_dimensions[get_column_letter(j)].width = (lebar or {}).get(c, 14)
        if c in (dropdown or {}) and len(df):
            dv = DataValidation(type="list", formula1='"' + ",".join(dropdown[c]) + '"',
                                allow_blank=True, showErrorMessage=False)
            dv.add(f"{get_column_letter(j)}2:{get_column_letter(j)}{len(df) + 1}")
            ws.add_data_validation(dv)
    for baris in ws.iter_rows(min_row=2):
        for sel in baris:
            sel.alignment = Alignment(vertical="top", wrap_text=False)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


# --------------------------------------------------------------------------------------
# 2) AMBIL
# --------------------------------------------------------------------------------------
def baca_daftar_video(path) -> pd.DataFrame:
    df = baca_tabel(path)
    df.columns = [str(c).strip().lower() for c in df.columns]
    kol = next((c for c in ("video", "video_id", "url", "link") if c in df.columns), None)
    if kol is None:
        sys.exit("Daftar video harus punya kolom 'video', 'video_id', 'url', atau 'link'.")
    if "pakai" in df.columns:
        df = df[df["pakai"].str.strip().str.upper().isin(["Y", "YA", "YES", "1", "TRUE", "V", "X"])]
    df = df.copy()
    df["video_id"] = df[kol].map(ekstrak_id_video)
    rusak = df[df["video_id"] == ""]
    if len(rusak):
        info(f"  ! {len(rusak)} baris dilewati karena ID/URL video tidak dikenali: {rusak[kol].head(3).tolist()}")
    for c in ("mapel", "jenjang", "topik"):
        if c not in df.columns:
            df[c] = ""
    return df[df["video_id"] != ""].drop_duplicates("video_id")[["video_id", "mapel", "jenjang", "topik"]]


def rekam_komentar(c: dict, vid: str, induk: str, n_bal: int, id_kanal: str) -> dict:
    s = c.get("snippet", {})
    penulis = (s.get("authorChannelId") or {}).get("value", "")
    return {
        "id_komentar": c.get("id", ""),
        "video_id": vid,
        "id_induk": induk,
        "adalah_balasan": int(bool(induk)),
        "teks": s.get("textDisplay") or s.get("textOriginal") or "",
        "jumlah_like": int(s.get("likeCount", 0) or 0),
        "waktu_terbit": s.get("publishedAt", ""),
        "waktu_ubah": s.get("updatedAt", ""),
        "jumlah_balasan": n_bal,
        "dari_kreator": int(bool(penulis) and penulis == id_kanal),
        "id_penulis_hash": hash_penulis(penulis),  # nama & ID asli penulis TIDAK disimpan
    }


def ambil_balasan(id_induk: str) -> list:
    hasil, token = [], None
    while True:
        p = {"part": "snippet", "parentId": id_induk, "maxResults": 100, "textFormat": "plainText"}
        if token:
            p["pageToken"] = token
        res = api_get("comments", p)
        hasil += res.get("items", [])
        token = res.get("nextPageToken")
        if not token:
            return hasil


def ambil_komentar_video(vid: str, id_kanal: str, a, f) -> int:
    n, token = 0, None
    while n < a.maks_komentar:
        p = {"part": "snippet,replies", "videoId": vid, "maxResults": 100,
             "textFormat": "plainText", "order": a.urutan}
        if token:
            p["pageToken"] = token
        res = api_get("commentThreads", p)
        for th in res.get("items", []):
            sn = th.get("snippet", {})
            top = sn.get("topLevelComment")
            if not top:
                continue
            total_bal = int(sn.get("totalReplyCount", 0) or 0)
            f.write(json.dumps(rekam_komentar(top, vid, "", total_bal, id_kanal), ensure_ascii=False) + "\n")
            n += 1
            if a.balasan and total_bal:
                bal = th.get("replies", {}).get("comments", [])
                if total_bal > len(bal):
                    bal = ambil_balasan(top["id"])
                for c in bal:
                    f.write(json.dumps(rekam_komentar(c, vid, top["id"], 0, id_kanal), ensure_ascii=False) + "\n")
                    n += 1
        token = res.get("nextPageToken")
        if not token:
            break
    return n


def gabungkan_mentah(folder: Path, meta: pd.DataFrame):
    bagian = []
    for vid in meta["video_id"]:
        p = folder / "mentah" / f"{vid}.jsonl"
        if p.exists():
            with open(p, encoding="utf-8") as f:
                bagian += [json.loads(b) for b in f if b.strip()]
    if not bagian:
        return None
    df = pd.DataFrame(bagian)
    kol = ["video_id", "mapel", "jenjang", "topik", "judul", "kanal", "durasi_detik"]
    df = df.merge(meta[kol], on="video_id", how="left")
    simpan_csv(df, folder / "komentar_mentah.csv")
    return df


def cmd_ambil(a) -> int:
    daftar = baca_daftar_video(a.videos)
    if daftar.empty:
        sys.exit("Daftar video kosong. Pastikan kolom 'pakai' berisi Y untuk video yang dipilih.")
    folder = Path(a.data)
    (folder / "mentah").mkdir(parents=True, exist_ok=True)
    info(f"{len(daftar)} video dalam daftar. Mengambil metadata...")
    meta = pd.DataFrame(ambil_metadata(daftar["video_id"].tolist()))
    hilang = set(daftar["video_id"]) - set(meta["video_id"] if not meta.empty else [])
    if hilang:
        info(f"  ! {len(hilang)} video tidak ditemukan/privat: {sorted(hilang)[:5]}")
    if meta.empty:
        sys.exit("Tidak ada video yang bisa diambil.")
    meta = daftar.merge(meta, on="video_id", how="inner")
    simpan_csv(meta, folder / "video_meta.csv")

    log, kuota_habis = [], False
    for i, row in enumerate(meta.itertuples(index=False), 1):
        vid = row.video_id
        tujuan = folder / "mentah" / f"{vid}.jsonl"
        if tujuan.exists() and not a.ulang:
            log.append({"video_id": vid, "status": "sudah_ada", "jumlah": sum(1 for _ in open(tujuan, encoding="utf-8"))})
            continue
        info(f"[{i}/{len(meta)}] {str(row.judul)[:70]}")
        sementara = tujuan.with_suffix(".part")
        try:
            with open(sementara, "w", encoding="utf-8") as f:
                n = ambil_komentar_video(vid, row.id_kanal, a, f)
            sementara.replace(tujuan)
            log.append({"video_id": vid, "status": "ok", "jumlah": n})
            info(f"      {n} komentar")
        except KuotaHabis:
            sementara.unlink(missing_ok=True)
            log.append({"video_id": vid, "status": "kuota_habis", "jumlah": 0})
            kuota_habis = True
            break
        except ApiError as e:
            sementara.unlink(missing_ok=True)
            status = {"commentsDisabled": "komentar_dimatikan", "videoNotFound": "tidak_ditemukan",
                      "forbidden": "terlarang"}.get(e.alasan, "galat")
            log.append({"video_id": vid, "status": status, "jumlah": 0, "pesan": e.pesan[:150]})
            info(f"      dilewati: {status}")

    simpan_csv(pd.DataFrame(log), folder / "log_ambil.csv")
    df = gabungkan_mentah(folder, meta)
    total = 0 if df is None else len(df)
    info(f"\nTotal {total} komentar dari {0 if df is None else df['video_id'].nunique()} video -> {folder / 'komentar_mentah.csv'}")
    info(f"Status per video: {dict(Counter(x['status'] for x in log))}  (detail: {folder / 'log_ambil.csv'})")
    if kuota_habis:
        info(PESAN_KUOTA)
        return 2
    info("Langkah berikut:  python bingung_scraper.py bersihkan")
    return 0


# --------------------------------------------------------------------------------------
# 3) BERSIHKAN
# --------------------------------------------------------------------------------------
RE_URL = re.compile(r"https?://\S+|www\.\S+")
RE_MENTION = re.compile(r"@[\w.\-]+")
RE_TS = re.compile(r"(?<![\d:.,])(?:(\d{1,2}):)?(\d{1,2}):([0-5]\d)(?![\d:])")
RE_MENIT = re.compile(r"\b(?:menit|mnt)\s*(?:ke\s*-?\s*)?(\d{1,3})(?:[.,:]([0-5]\d))?(?!\d)", re.I)
_NEG = r"(?:ga|gak|gk|g|nggak|ngga|ngak|nggk|enggak|engga|tdk|tidak|blm|belum|kurang|krg)"
_MENGERTI = r"(?:ngerti+|mengerti|paham|faham|mudeng|ngeh|nangkep|nangkap|nyambung|ngarti)"
RE_SINYAL = re.compile(
    r"\?|\b" + _NEG + r"\s*" + _MENGERTI + r"\b|"
    r"\b(?:bingung|bingun|binggung|kok\s+bisa|kenapa|knp|mengapa|gimana|gmn|bagaimana|bgmn|"
    r"dari\s*mana|drmn|dapet\s+dari|dapat\s+dari|dpt\s+dari|maksudnya|maksud\s+nya|"
    r"i[jz]in\s+(?:ber)?tanya|mau\s+(?:nanya|tanya)|mo\s+nanya|nanya|bertanya|"
    r"jelasin|jelaskan|teu\s+ngarti|ora\s+(?:mudeng|ngerti)|pusing|puyeng)\b",
    re.I,
)


def bersihkan_teks(t) -> str:
    t = html.unescape(str(t))
    t = RE_KONTROL.sub(" ", t)
    t = RE_URL.sub("<URL>", t)
    t = RE_MENTION.sub("@pengguna", t)  # sebutan akun lain disamarkan
    return re.sub(r"\s+", " ", t).strip()


def ekstrak_timestamp(teks: str, durasi: int) -> list:
    hasil = []
    for m in RE_TS.finditer(teks):
        sebelum = teks[max(0, m.start() - 7):m.start()].lower()
        if re.search(r"(jam|pukul|pkl)\.?\s*$", sebelum):  # "jam 10:30" = jam dinding, bukan posisi video
            continue
        if m.group(1) and int(m.group(2)) >= 60:
            continue
        hasil.append(int(m.group(1) or 0) * 3600 + int(m.group(2)) * 60 + int(m.group(3)))
    for m in RE_MENIT.finditer(teks):
        hasil.append(int(m.group(1)) * 60 + int(m.group(2) or 0))
    if durasi and durasi > 0:
        hasil = [t for t in hasil if t <= durasi]
    return sorted(set(hasil))


def cmd_bersihkan(a) -> None:
    folder = Path(a.data)
    df = baca_tabel(folder / "komentar_mentah.csv")
    df = ke_int(df, ["durasi_detik", "jumlah_like", "adalah_balasan", "dari_kreator", "jumlah_balasan"])
    n0 = len(df)
    df["teks_bersih"] = df["teks"].map(bersihkan_teks)
    df = df[df["teks_bersih"].str.len() > 0].copy()

    df["_norm"] = df["teks_bersih"].str.lower().str.replace(r"[\W_]+", " ", regex=True).str.strip()
    # 1) orang yang sama memposting teks yang sama di video yang sama -> simpan satu
    sebelum = len(df)
    df = df.drop_duplicates(["video_id", "id_penulis_hash", "_norm"]).copy()
    n_dup = sebelum - len(df)
    # 2) spam: teks panjang identik di >=3 video tetapi dari <=2 akun (salin-tempel promosi)
    per_teks = df.groupby("_norm").agg(n_video=("video_id", "nunique"), n_penulis=("id_penulis_hash", "nunique"))
    teks_spam = set(per_teks[(per_teks.n_video >= 3) & (per_teks.n_penulis <= 2)].index)
    df["spam_duplikat"] = (df["_norm"].isin(teks_spam) & (df["_norm"].str.len() >= 25)).astype(int)
    # 3) teks yang sudah pernah muncul (oleh orang lain) tetap dihitung untuk indeks kebingungan,
    #    tetapi tidak ikut disampel agar anotasi tidak berulang & tidak bocor antara data latih dan uji
    df["teks_duplikat"] = df.duplicated("_norm").astype(int)
    df["hanya_emoji"] = (df["_norm"] == "").astype(int)
    df["jumlah_kata"] = df["teks_bersih"].str.split().str.len()

    ts = [ekstrak_timestamp(t, d) for t, d in zip(df["teks_bersih"], df["durasi_detik"])]
    df["timestamp_detik"] = [";".join(map(str, x)) for x in ts]
    df["ada_timestamp"] = [int(bool(x)) for x in ts]
    df["posisi_relatif"] = [round(x[0] / d, 4) if x and d else None for x, d in zip(ts, df["durasi_detik"])]
    df["sinyal_bingung"] = df["teks_bersih"].map(lambda t: int(bool(RE_SINYAL.search(t))))
    df = df.drop(columns="_norm")
    simpan_csv(df, folder / "komentar_bersih.csv")

    layak = df[(df.hanya_emoji == 0) & (df.spam_duplikat == 0) & (df.dari_kreator == 0)]
    ring = layak.groupby(layak["mapel"].replace("", "(kosong)")).agg(
        video=("video_id", "nunique"),
        komentar=("id_komentar", "count"),
        persen_timestamp=("ada_timestamp", lambda s: round(100 * s.mean(), 1)),
        persen_sinyal_bingung=("sinyal_bingung", lambda s: round(100 * s.mean(), 1)),
        median_kata=("jumlah_kata", "median"),
    )
    ring.loc["TOTAL"] = [layak["video_id"].nunique(), len(layak), round(100 * layak["ada_timestamp"].mean(), 1),
                         round(100 * layak["sinyal_bingung"].mean(), 1), layak["jumlah_kata"].median()]
    simpan_csv(ring.reset_index().rename(columns={"index": "mapel"}), folder / "ringkasan_dataset.csv")

    unik = layak[layak.teks_duplikat == 0]
    info(f"Komentar mentah                  : {n0}")
    info(f"Posting ganda (orang sama) dibuang: {n_dup}")
    info(f"Ditandai spam lintas video       : {int(df.spam_duplikat.sum())}")
    info(f"Hanya emoji/tanda baca           : {int(df.hanya_emoji.sum())}")
    info(f"Balasan dari kreator kanal       : {int(df.dari_kreator.sum())}")
    info(f"Komentar penonton (untuk indeks) : {len(layak)}")
    info(f"Teks unik yang bisa dianotasi    : {len(unik)}\n")
    info(ring.to_string())
    info(f"\n-> {folder / 'komentar_bersih.csv'}  dan  {folder / 'ringkasan_dataset.csv'} (untuk Tabel Dataset di makalah)")
    info("Langkah berikut:  python bingung_scraper.py sampel --anotator NamaA NamaB NamaC")


# --------------------------------------------------------------------------------------
# 4) SAMPEL -> lembar anotasi
# --------------------------------------------------------------------------------------
def ambil_sampel_terbatas(df: pd.DataFrame, n: int, cap: int, seed: int) -> pd.DataFrame:
    """Sampel acak n baris, maksimal `cap` komentar per video (dilonggarkan bila kurang)."""
    if n <= 0 or df.empty:
        return df.iloc[0:0].copy()
    acak = df.sample(frac=1, random_state=seed)
    urut = acak.groupby("video_id").cumcount()
    pilih = acak[urut < cap]
    if len(pilih) >= n:
        return pilih.head(n).copy()
    sisa = acak.drop(pilih.index)
    return pd.concat([pilih, sisa.head(n - len(pilih))]).copy()


def susun_sampel(pool: pd.DataFrame, n: int, prop_sinyal: float, cap: int, seed: int,
                 proporsional: bool) -> pd.DataFrame:
    pool = pool.copy()
    pool["mapel"] = pool["mapel"].replace("", "(kosong)")
    ukuran = pool["mapel"].value_counts()
    if proporsional:
        jatah = {m: int(n * c / len(pool)) for m, c in ukuran.items()}
    else:  # default: jatah sama rata per mapel agar analisis antar-mapel seimbang
        jatah = {m: n // len(ukuran) for m in ukuran.index}
    for m in list(ukuran.index)[: n - sum(jatah.values())]:
        jatah[m] += 1

    hasil = []
    for i, (m, k) in enumerate(jatah.items()):
        sub = pool[pool["mapel"] == m]
        k = min(k, len(sub))
        ks = int(round(k * prop_sinyal))
        s1 = ambil_sampel_terbatas(sub[sub["sinyal_bingung"] == 1], ks, cap, seed + i)
        s1["strategi_sampel"] = "pengayaan_sinyal"
        s2 = ambil_sampel_terbatas(sub.drop(s1.index), k - len(s1), cap, seed + 100 + i)
        s2["strategi_sampel"] = "acak"
        hasil += [s1, s2]
    sampel = pd.concat(hasil)
    if len(sampel) < n:  # tutup kekurangan dari mapel lain
        s3 = ambil_sampel_terbatas(pool.drop(sampel.index), n - len(sampel), cap, seed + 999)
        s3["strategi_sampel"] = "acak_tambahan"
        sampel = pd.concat([sampel, s3])
    return sampel.sample(frac=1, random_state=seed)


def tulis_lembar_anotasi(df: pd.DataFrame, path: Path, nama: str) -> None:
    from openpyxl import Workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()
    ws = wb.active
    ws.title = "Anotasi"
    kolom = ["no", "id_komentar", "mapel", "judul_video", "teks", "label", "jenis_bingung", "catatan"]
    ws.append(kolom)
    for i, r in enumerate(df.itertuples(index=False), 1):
        ws.append([i, r.id_komentar, r.mapel, r.judul, r.teks_bersih, "", "", ""])
    isi_sel_aman(ws)
    n = len(df)
    lebar = {"A": 5, "B": 12, "C": 11, "D": 32, "E": 75, "F": 12, "G": 15, "H": 24}
    for kol, w in lebar.items():
        ws.column_dimensions[kol].width = w
    for j in range(1, len(kolom) + 1):
        sel = ws.cell(row=1, column=j)
        sel.font = Font(bold=True, color="FFFFFF")
        sel.fill = PatternFill("solid", fgColor="1F3A5F")
    for baris in ws.iter_rows(min_row=2, max_row=n + 1):
        for sel in baris:
            sel.alignment = Alignment(vertical="top", wrap_text=sel.column_letter in ("D", "E", "H"))
    ws.freeze_panes = "A2"

    dv1 = DataValidation(type="list", formula1='"BINGUNG,KOREKSI,PAHAM,LAIN"', allow_blank=True, showErrorMessage=False)
    dv2 = DataValidation(type="list", formula1='"KONSEPTUAL,PROSEDURAL,NOTASI,UMUM"', allow_blank=True,
                         showErrorMessage=False)
    dv1.add(f"F2:F{n + 1}")
    dv2.add(f"G2:G{n + 1}")
    ws.add_data_validation(dv1)
    ws.add_data_validation(dv2)
    warna = {"BINGUNG": "FFE699", "KOREKSI": "F4B6B6", "PAHAM": "C6E0B4", "LAIN": "D9D9D9"}
    for lab, w in warna.items():
        ws.conditional_formatting.add(
            f"F2:F{n + 1}",
            FormulaRule(formula=[f'OR(UPPER($F2)="{lab}",UPPER($F2)="{lab[0]}")'],
                        fill=PatternFill("solid", fgColor=w, bgColor=w)))

    pd_ws = wb.create_sheet("Panduan Ringkas")
    teks = [
        f"Lembar anotasi BINGUNG-ID untuk: {nama}   ({n} komentar)",
        "",
        "CARA MENGISI",
        "Kolom 'label': pilih dari dropdown, atau ketik huruf awal: B / K / P / L.",
        "Kolom 'jenis_bingung': HANYA diisi bila label = BINGUNG. Ketik K / P / N / U atau pilih dropdown.",
        "Kolom 'catatan': tulis RAGU bila tidak yakin (tetap pilih label terbaik), atau alasan singkat.",
        "",
        "URUTAN KEPUTUSAN (berhenti di jawaban 'ya' pertama)",
        "1. Menunjuk kesalahan ISI video / mengusulkan jawaban lain (yakin/ragu)?    -> KOREKSI",
        "   (masalah teknis seperti suara jelek atau tulisan terpotong = LAIN)",
        "2. Menyatakan belum paham, atau bertanya tentang ISI materi video?          -> BINGUNG",
        "3. Menyatakan paham / terbantu / berterima kasih atas penjelasan?           -> PAHAM",
        "4. Selain itu (salam, absen, spam, minta materi lain, tanya non-materi)    -> LAIN",
        "",
        "JENIS BINGUNG",
        "K = KONSEPTUAL  : 'kenapa / apa maknanya' (alasan prinsip, arti konsep, hubungan antarkonsep)",
        "P = PROSEDURAL  : 'bagaimana / dapat dari mana' (langkah hitung, asal angka, manipulasi rumus)",
        "N = NOTASI      : simbol, istilah, satuan, cara membaca/menulis notasi",
        "U = UMUM        : bingung tanpa menyebut bagian/konsep yang spesifik ('gak ngerti sama sekali')",
        "",
        "ATURAN PENTING",
        "- Kerjakan sendiri. Jangan berdiskusi dengan anotator lain sampai semua lembar selesai.",
        "- Jangan mengurutkan ulang, menghapus, atau menambah baris. Jangan ubah kolom id_komentar.",
        "- Ejaan buruk, bahasa daerah, dan singkatan tetap dinilai dari maksudnya.",
        "- Pedoman lengkap: dokumen 'Pedoman Anotasi BINGUNG-ID'.",
    ]
    for t in teks:
        pd_ws.append([t])
    pd_ws.column_dimensions["A"].width = 110
    for idx in (1, 3, 8, 15, 21):
        pd_ws.cell(row=idx, column=1).font = Font(bold=True)
    wb.save(path)


def cmd_sampel(a) -> None:
    nama = a.anotator
    if len(nama) < 2:
        sys.exit("Minimal 2 anotator (contoh: --anotator Rani Budi Sari).")
    folder = Path(a.data)
    df = baca_tabel(folder / "komentar_bersih.csv")
    df = ke_int(df, ["hanya_emoji", "spam_duplikat", "dari_kreator", "sinyal_bingung", "jumlah_kata",
                     "adalah_balasan", "teks_duplikat"])
    pool = df[(df.hanya_emoji == 0) & (df.spam_duplikat == 0) & (df.dari_kreator == 0)
              & (df.teks_duplikat == 0) & (df.teks_bersih.str.len() >= 2)]
    n = a.n
    if len(pool) < n:
        info(f"  ! Hanya {len(pool)} komentar layak; sampel diperkecil.")
        n = len(pool)
    if a.irisan >= n:
        sys.exit("--irisan harus lebih kecil dari --n.")

    sampel = susun_sampel(pool, n, a.proporsi_sinyal, a.maks_per_video, a.seed, a.proporsional).reset_index(drop=True)
    irisan, sisa = sampel.iloc[: a.irisan], sampel.iloc[a.irisan:]
    keluar = Path(a.keluaran)
    keluar.mkdir(parents=True, exist_ok=True)
    pemilik = ["SEMUA"] * len(irisan)
    jumlah = {}
    for i, nm in enumerate(nama):
        bagian = sisa.iloc[i::len(nama)]
        lembar = pd.concat([irisan, bagian]).sample(frac=1, random_state=a.seed + i + 1)
        path = keluar / f"anotasi_{nama_aman(nm)}.xlsx"
        tulis_lembar_anotasi(lembar, path, nm)
        jumlah[nm] = (len(lembar), path)
    pemilik += [nama[i % len(nama)] for i in range(len(sisa))]
    sampel["irisan"] = [1] * len(irisan) + [0] * len(sisa)
    sampel["anotator"] = pemilik
    simpan_csv(sampel, folder / "sampel_anotasi.csv")

    info(f"Sampel {len(sampel)} komentar  (irisan bersama: {len(irisan)}, maks {a.maks_per_video}/video/strategi)")
    info("Per mapel   : " + ", ".join(f"{k}={v}" for k, v in sampel["mapel"].value_counts().items()))
    info("Per strategi: " + ", ".join(f"{k}={v}" for k, v in sampel["strategi_sampel"].value_counts().items()))
    info(f"Video terwakili: {sampel['video_id'].nunique()}")
    for nm, (j, path) in jumlah.items():
        info(f"  · {nm}: {j} komentar -> {path}")
    info(f"\nKunci sampel (JANGAN dibagikan ke anotator): {folder / 'sampel_anotasi.csv'}")
    info("Setelah semua lembar terisi:  python bingung_scraper.py kappa --lembar anotasi/anotasi_*.xlsx")


# --------------------------------------------------------------------------------------
# 5) KAPPA + gabung label
# --------------------------------------------------------------------------------------
def norm_label(s) -> str:
    s = str(s).strip().upper()
    s = SINGKATAN_LABEL.get(s, s)
    return s if s in LABEL else ("?" if s else "")


def norm_jenis(s) -> str:
    s = str(s).strip().upper()
    s = SINGKATAN_JENIS.get(s, s)
    return s if s in JENIS else ""


def cohen_kappa(a: list, b: list) -> float:
    n = len(a)
    if n == 0:
        return float("nan")
    po = sum(x == y for x, y in zip(a, b)) / n
    ca, cb = Counter(a), Counter(b)
    pe = sum((ca[l] / n) * (cb[l] / n) for l in set(a) | set(b))
    return 1.0 if pe == 1 else (po - pe) / (1 - pe)


def fleiss_kappa(baris: list) -> float:
    """baris: list berisi list label dari setiap anotator (jumlah anotator sama per baris)."""
    if not baris:
        return float("nan")
    r = len(baris[0])
    N = len(baris)
    kategori = sorted({l for b in baris for l in b})
    hitung = [Counter(b) for b in baris]
    p_j = [sum(h[k] for h in hitung) / (N * r) for k in kategori]
    P_i = [(sum(h[k] ** 2 for k in kategori) - r) / (r * (r - 1)) for h in hitung]
    P_bar, P_e = sum(P_i) / N, sum(p ** 2 for p in p_j)
    return 1.0 if P_e == 1 else (P_bar - P_e) / (1 - P_e)


def tafsir_kappa(k: float) -> str:  # Landis & Koch (1977)
    if k != k:
        return "-"
    if k < 0:
        return "buruk"
    for batas, teks in [(0.20, "sangat lemah"), (0.40, "lemah"), (0.60, "sedang"), (0.80, "kuat")]:
        if k <= batas:
            return teks
    return "hampir sempurna"


def mayoritas(nilai: list) -> str:
    nilai = [v for v in nilai if v]
    if not nilai:
        return ""
    c = Counter(nilai).most_common()
    if len(c) > 1 and c[0][1] == c[1][1]:
        return "PERLU_DISKUSI"
    return c[0][0]


def cmd_kappa(a) -> None:
    folder = Path(a.data)
    bagian, usulan = [], []
    for p in a.lembar:
        p = Path(p)
        # anotasi_Rani.xlsx dan anotasi_Rani_tahap2.xlsx = anotator yang sama
        nm = re.sub(r"_tahap\d+$", "", re.sub(r"^anotasi_", "", p.stem))
        df = pd.read_excel(p, sheet_name="Anotasi", dtype=str).fillna("")
        df["anotator"] = nm
        df["label"] = df["label"].map(norm_label)
        df["jenis_bingung"] = df["jenis_bingung"].map(norm_jenis)
        kosong, salah = int((df.label == "").sum()), df[df.label == "?"]
        tanpa_jenis = int(((df.label == "BINGUNG") & (df.jenis_bingung == "")).sum())
        info(f"{p.name}: {len(df)} baris | kosong {kosong} | label tak dikenal {len(salah)} | BINGUNG tanpa jenis {tanpa_jenis}")
        if len(salah):
            info(f"   baris label tak dikenal (kolom 'no'): {salah['no'].tolist()[:15]}")
        df.loc[df.label == "?", "label"] = ""
        bagian.append(df[["id_komentar", "anotator", "label", "jenis_bingung", "catatan"]])
        if "usulan_ai" in df.columns:  # lembar tahap 2: label awal dari AI yang diverifikasi anotator
            u = df[["id_komentar", "usulan_ai", "usulan_jenis_ai", "label", "jenis_bingung"]].copy()
            u["ai_yakin"] = df["ai_yakin"] if "ai_yakin" in df.columns else ""
            usulan.append(u)
    panjang = pd.concat(bagian)
    if panjang.duplicated(["id_komentar", "anotator"]).any():
        ganda = panjang[panjang.duplicated(["id_komentar", "anotator"], keep=False)]
        sys.exit(f"Komentar yang sama muncul dua kali untuk satu anotator ({ganda.anotator.iloc[0]}). "
                 "Pastikan lembar lama (700 baris) sudah diganti dengan lembar tahap 1 + tahap 2.")
    nama = list(dict.fromkeys(panjang["anotator"]))
    lab = panjang.pivot_table(index="id_komentar", columns="anotator", values="label", aggfunc="first").reindex(columns=nama)
    jen = panjang.pivot_table(index="id_komentar", columns="anotator", values="jenis_bingung", aggfunc="first").reindex(columns=nama)
    lab = lab.replace("", pd.NA)

    laporan = ["LAPORAN KESEPAKATAN ANTAR-ANOTATOR · BINGUNG-ID", ""]
    jumlah_isi = lab.notna().sum(axis=1)
    irisan = lab[jumlah_isi >= 2]
    laporan.append(f"Komentar irisan yang dinilai >=2 anotator: {len(irisan)}")
    kappa_pasang = []
    for x, y in combinations(nama, 2):
        m = lab[x].notna() & lab[y].notna()
        if m.sum() == 0:
            continue
        k = cohen_kappa(lab.loc[m, x].tolist(), lab.loc[m, y].tolist())
        setuju = (lab.loc[m, x] == lab.loc[m, y]).mean()
        kappa_pasang.append(k)
        laporan.append(f"  Cohen's kappa {x} vs {y}: {k:.3f} ({tafsir_kappa(k)}) · kesepakatan mentah {setuju:.1%} · n={int(m.sum())}")
    if kappa_pasang:
        rata = sum(kappa_pasang) / len(kappa_pasang)
        laporan.append(f"  Rata-rata Cohen's kappa: {rata:.3f} ({tafsir_kappa(rata)})")
    if len(nama) >= 3:
        lengkap = lab.dropna()
        fk = fleiss_kappa(lengkap.values.tolist())
        laporan.append(f"  Fleiss' kappa ({len(nama)} anotator, n={len(lengkap)}): {fk:.3f} ({tafsir_kappa(fk)})")

    # label final
    final = pd.DataFrame(index=lab.index)
    final["label"] = [mayoritas(r) for r in lab.fillna("").values.tolist()]
    final["n_anotator"] = jumlah_isi
    final["sepakat_penuh"] = [int(len({v for v in r if v}) == 1 and sum(1 for v in r if v) >= 2)
                              for r in lab.fillna("").values.tolist()]
    jenis_final = []
    for idx in final.index:
        if final.at[idx, "label"] != "BINGUNG":
            jenis_final.append("")
            continue
        pilih = [jen.at[idx, nm] for nm in nama if lab.at[idx, nm] == "BINGUNG"]
        jenis_final.append(mayoritas([v for v in pilih if isinstance(v, str)]) or "UMUM")
    final["jenis_bingung"] = jenis_final

    beda = lab[(jumlah_isi >= 2) & (lab.nunique(axis=1) > 1)]
    pasangan_tertukar = Counter()
    for r in beda.values.tolist():
        v = sorted({x for x in r if isinstance(x, str)})
        for p in combinations(v, 2):
            pasangan_tertukar[" vs ".join(p)] += 1
    if pasangan_tertukar:
        laporan.append("  Label yang paling sering tertukar: " +
                       ", ".join(f"{k} ({v})" for k, v in pasangan_tertukar.most_common(4)))

    sampel = baca_tabel(folder / "sampel_anotasi.csv")
    # file ketidaksepakatan untuk sesi diskusi (label beda, atau jenis BINGUNG seri)
    perlu = beda.index.union(final.index[final["jenis_bingung"] == "PERLU_DISKUSI"])
    konteks = sampel.set_index("id_komentar")[["judul", "teks_bersih"]]
    dis = konteks.join(lab.loc[perlu].add_prefix("label_"), how="inner")
    dis = dis.join(jen.loc[perlu].add_prefix("jenis_"), how="left")
    dis["usulan_mayoritas"] = final.loc[dis.index, "label"]
    dis["usulan_jenis"] = final.loc[dis.index, "jenis_bingung"]
    dis["label_final"] = ""
    dis["jenis_final"] = ""
    dis.index.name = "id_komentar"

    diterapkan = 0
    if a.putusan:
        put = baca_tabel(a.putusan)
        if "label_final" not in put.columns:
            sys.exit(f"{a.putusan} tidak punya kolom label_final.")
        for r in put.to_dict("records"):
            idx, lb = r.get("id_komentar", ""), norm_label(r.get("label_final", ""))
            if idx not in final.index or lb not in LABEL:
                continue
            final.at[idx, "label"] = lb
            jf = norm_jenis(r.get("jenis_final", ""))
            final.at[idx, "jenis_bingung"] = (jf or final.at[idx, "jenis_bingung"] or "UMUM") if lb == "BINGUNG" else ""
            diterapkan += 1
        laporan.append(f"  Putusan diskusi diterapkan: {diterapkan} komentar (dari {a.putusan})")
    else:
        simpan_csv(dis.reset_index(), folder / "ketidaksepakatan.csv")

    # sumber label: irisan = dinilai mandiri oleh >=2 manusia; ai_diverifikasi = usulan AI dicek 1 manusia
    final["sumber_label"] = ["manusia_irisan" if n >= 2 else "manusia" for n in final["n_anotator"]]
    if usulan:
        u = pd.concat(usulan).drop_duplicates("id_komentar").set_index("id_komentar")
        ada = final.index.intersection(u.index)
        final.loc[ada, "sumber_label"] = "ai_diverifikasi"
        final["usulan_ai"] = u["usulan_ai"].reindex(final.index).fillna("")
        u = u[u["label"] != ""]
        if u.empty:
            u = u.assign(ai_yakin=pd.Series(dtype=str))
        ubah_label = u["label"] != u["usulan_ai"]
        ubah_jenis = (u["label"] == "BINGUNG") & (u["usulan_ai"] == "BINGUNG") & \
                     (u["jenis_bingung"] != u["usulan_jenis_ai"]) & (u["jenis_bingung"] != "")
        laporan += ["", "VERIFIKASI LABEL AI (lembar tahap 2)",
                    f"  Diperiksa: {len(u)} komentar | label diubah: {int(ubah_label.sum())} ({ubah_label.mean():.1%})"
                    f" | jenis BINGUNG diubah: {int(ubah_jenis.sum())}"]
        for k in ("RAGU", "YAKIN"):
            m = u["ai_yakin"] == k
            if m.any():
                laporan.append(f"  Usulan AI {k}: {int(m.sum())} komentar, {ubah_label[m].mean():.1%} labelnya diubah")
        if ubah_label.any():
            arah = Counter(f"{x}->{y}" for x, y in zip(u.loc[ubah_label, "usulan_ai"], u.loc[ubah_label, "label"]))
            laporan.append("  Perubahan terbanyak: " + ", ".join(f"{k} ({v})" for k, v in arah.most_common(5)))
        if not ubah_label.any() and len(u) > 50:
            laporan.append("  ! Tidak ada satu pun label AI yang diubah. Pastikan lembar tahap 2 benar-benar diperiksa.")

    final = final.reset_index()
    kol = ["id_komentar", "video_id", "mapel", "jenjang", "topik", "judul", "durasi_detik", "teks_bersih",
           "ada_timestamp", "timestamp_detik", "posisi_relatif", "jumlah_like", "adalah_balasan",
           "waktu_terbit", "irisan", "strategi_sampel", "sinyal_bingung"]
    data = sampel[[k for k in kol if k in sampel.columns]].merge(final, on="id_komentar", how="inner")
    simpan_csv(data, folder / "dataset_berlabel.csv")

    siap = data[data["label"].isin(LABEL)]
    laporan += ["", f"Dataset berlabel: {len(data)} komentar | siap dipakai: {len(siap)} | "
                    f"PERLU_DISKUSI: {int((data.label == 'PERLU_DISKUSI').sum())} | kosong: {int((data.label == '').sum())}",
                "Sebaran label: " + ", ".join(f"{k}={v} ({v / max(len(siap), 1):.1%})"
                                              for k, v in siap["label"].value_counts().items())]
    if (siap.label == "BINGUNG").any():
        laporan.append("Sebaran jenis (BINGUNG): " + ", ".join(
            f"{k}={v}" for k, v in siap.loc[siap.label == "BINGUNG", "jenis_bingung"].value_counts().items()))
        n_jd = int((siap.jenis_bingung == "PERLU_DISKUSI").sum())
        if n_jd:
            laporan.append(f"  ! {n_jd} komentar BINGUNG jenisnya masih seri: isi kolom jenis_final di ketidaksepakatan.csv")
    (folder / "laporan_kappa.txt").write_text("\n".join(laporan), encoding="utf-8")
    info("\n" + "\n".join(laporan))
    info(f"\n-> {folder / 'dataset_berlabel.csv'}")
    if not a.putusan:
        info(f"-> {folder / 'ketidaksepakatan.csv'}  ({len(dis)} komentar untuk sesi diskusi; isi kolom label_final, lalu")
        info(f"   jalankan ulang:  python bingung_scraper.py kappa --lembar ... --putusan {folder / 'ketidaksepakatan.csv'})")
    info(f"-> {folder / 'laporan_kappa.txt'}  (angka untuk bagian Dataset/Metode di makalah)")


# --------------------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------------------
def buat_parser() -> argparse.ArgumentParser:
    umum = argparse.ArgumentParser(add_help=False)
    umum.add_argument("--api-key", default=None, help="YouTube Data API key (default: variabel YT_API_KEY)")
    umum.add_argument("--data", default="data", help="folder data (default: data)")

    pa = argparse.ArgumentParser(description="BINGUNG-ID: pengambil & penyiap komentar YouTube untuk DIMAS-TI 2026")
    sub = pa.add_subparsers(dest="perintah", required=True)

    c = sub.add_parser("cari", parents=[umum], help="cari kandidat video per kata kunci")
    c.add_argument("--kata", nargs="*", help='kata kunci, mis. "limit fungsi" "hukum newton"')
    c.add_argument("--kata-file", help="file kata kunci; baris [Mapel] menandai bagian")
    c.add_argument("--maks", type=int, default=25, help="video per kata kunci (default 25; >50 = 2x kuota)")
    c.add_argument("--min-komentar", type=int, default=50, help="buang video dengan komentar di bawah ini")
    c.add_argument("--urutan", default="relevance", choices=["relevance", "viewCount", "date", "rating"])
    c.add_argument("--durasi", default="any", choices=["any", "short", "medium", "long"],
                   help="short <4 mnt, medium 4-20 mnt, long >20 mnt")
    c.add_argument("--keluaran", default="kandidat_video.xlsx")

    m = sub.add_parser("ambil", parents=[umum], help="ambil komentar untuk daftar video")
    m.add_argument("--videos", default="videos.xlsx", help="xlsx/csv dengan kolom video/url (+ mapel, jenjang, topik, pakai)")
    m.add_argument("--maks-komentar", type=int, default=2000, help="batas komentar per video (default 2000)")
    m.add_argument("--urutan", default="time", choices=["time", "relevance"])
    m.add_argument("--balasan", action="store_true", help="ikut ambil semua balasan (lebih banyak kuota)")
    m.add_argument("--ulang", action="store_true", help="ambil ulang video yang sudah pernah diambil")

    sub.add_parser("bersihkan", parents=[umum], help="anonimisasi, dedup, timestamp, fitur")

    s = sub.add_parser("sampel", parents=[umum], help="buat lembar anotasi per anotator")
    s.add_argument("--n", type=int, default=1500, help="jumlah komentar unik yang dianotasi (default 1500)")
    s.add_argument("--irisan", type=int, default=300, help="komentar yang dinilai SEMUA anotator (default 300)")
    s.add_argument("--anotator", nargs="+", default=["A", "B", "C"])
    s.add_argument("--proporsi-sinyal", type=float, default=0.4,
                   help="porsi sampel dari komentar bersinyal bingung (default 0.4)")
    s.add_argument("--maks-per-video", type=int, default=30)
    s.add_argument("--proporsional", action="store_true", help="jatah per mapel proporsional (default: sama rata)")
    s.add_argument("--seed", type=int, default=2026)
    s.add_argument("--keluaran", default="anotasi")

    k = sub.add_parser("kappa", parents=[umum], help="hitung kappa & gabungkan label")
    k.add_argument("--lembar", nargs="+", required=True, help="file anotasi_*.xlsx yang sudah diisi")
    k.add_argument("--putusan", help="ketidaksepakatan.csv yang kolom label_final-nya sudah diisi")
    return pa


def main(argv=None) -> int:
    global API_KEY
    a = buat_parser().parse_args(argv)
    API_KEY = a.api_key or os.environ.get("YT_API_KEY", "")
    perintah = {"cari": cmd_cari, "ambil": cmd_ambil, "bersihkan": cmd_bersihkan,
                "sampel": cmd_sampel, "kappa": cmd_kappa}[a.perintah]
    try:
        hasil = perintah(a)
        return hasil if isinstance(hasil, int) else 0
    except KuotaHabis:
        info(PESAN_KUOTA)
        return 2
    except ApiError as e:
        if "API key" in e.pesan or e.alasan in ("keyInvalid", "API_KEY_INVALID"):
            info("API key ditolak. Cek kunci dan pastikan 'YouTube Data API v3' sudah diaktifkan di project-nya.")
        else:
            info(f"Galat API: {e}")
        return 1
    finally:
        if KUOTA_TERPAKAI:
            info(f"[kuota] perkiraan terpakai sesi ini: {KUOTA_TERPAKAI} unit (batas harian 10.000 per project)")


if __name__ == "__main__":
    sys.exit(main())
