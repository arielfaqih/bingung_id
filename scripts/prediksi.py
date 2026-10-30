"""Prediksi label BINGUNG-ID untuk berkas CSV berisi komentar.

Dipakai untuk menilai data uji panitia (skor leaderboard babak final) atau
memberi label komentar baru dari video lain.

Contoh:
    python scripts/prediksi.py data_uji_panitia.csv --kolom text
    python scripts/prediksi.py komentar.csv --kolom teks --model tfidf --keluar prediksi.csv

Model yang dibaca (hasil Sel 12 di notebooks/02_model_dan_pemetaan.ipynb):
    hasil/model_transformer/   IndoBERT-base hasil fine-tuning (default)
    hasil/model_tfidf.joblib   TF-IDF + Regresi Logistik (cadangan tanpa GPU)
"""
import argparse
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

# Urutan indeks kelas sama persis dengan notebook pelatihan (LABEL di Sel 2).
LABEL = ["BINGUNG", "KOREKSI", "PAHAM", "LAIN"]
MAKS_TOKEN = 128

RE_URL = re.compile(r"https?://\S+|www\.\S+")
RE_MENTION = re.compile(r"@[\w.\-]+")


def anonimkan(t: str) -> str:
    """Samakan dengan pembersihan saat penambangan: tautan -> <URL>, sebutan akun -> @pengguna."""
    t = RE_URL.sub("<URL>", str(t))
    t = RE_MENTION.sub("@pengguna", t)
    return re.sub(r"\s+", " ", t).strip()


def bersih_tfidf(t: str) -> str:
    t = t.lower().replace("<url>", " url ").replace("@pengguna", " pengguna ")
    t = re.sub(r"(.)\1{2,}", r"\1\1", t)          # "bingunggg" -> "bingungg", "???" -> "??"
    t = re.sub(r"\d+(?:[.,]\d+)?", " 0 ", t)       # angka diseragamkan
    return re.sub(r"\s+", " ", t).strip()


def bersih_transformer(t: str) -> str:            # IndoBERT dan IndoBERTweet bersifat uncased
    return t.replace("<URL>", "httpurl").replace("@pengguna", "@user").lower()


def prediksi_transformer(teks, folder, bs=64):
    import torch
    from transformers import AutoModelForSequenceClassification, AutoTokenizer

    perangkat = "cuda" if torch.cuda.is_available() else "cpu"
    tok = AutoTokenizer.from_pretrained(folder)
    model = AutoModelForSequenceClassification.from_pretrained(folder).to(perangkat).eval()
    prob = []
    with torch.no_grad():
        for i in range(0, len(teks), bs):
            enc = tok([bersih_transformer(t) for t in teks[i:i + bs]], truncation=True,
                      max_length=MAKS_TOKEN, padding=True, return_tensors="pt")
            enc = {k: v.to(perangkat) for k, v in enc.items()}
            prob.append(torch.softmax(model(**enc).logits.float(), dim=-1).cpu().numpy())
    return np.concatenate(prob) if prob else np.zeros((0, len(LABEL)))


def prediksi_tfidf(teks, path):
    import joblib

    pipa = joblib.load(path)
    x = [bersih_tfidf(t) for t in teks]
    if hasattr(pipa, "predict_proba"):
        return pipa.predict_proba(x)
    hasil = np.zeros((len(x), len(LABEL)))
    hasil[np.arange(len(x)), pipa.predict(x)] = 1.0
    return hasil


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("masukan", help="CSV berisi komentar")
    ap.add_argument("--kolom", default="teks", help="nama kolom teks komentar (default: teks)")
    ap.add_argument("--model", choices=["transformer", "tfidf"], default="transformer")
    ap.add_argument("--folder-model", default="hasil/model_transformer")
    ap.add_argument("--path-tfidf", default="hasil/model_tfidf.joblib")
    ap.add_argument("--keluar", default="hasil/prediksi.csv")
    a = ap.parse_args()

    df = pd.read_csv(a.masukan, dtype=str).fillna("")
    if a.kolom not in df:
        sys.exit(f"Kolom '{a.kolom}' tidak ada. Kolom yang tersedia: {list(df.columns)}")
    teks = [anonimkan(t) for t in df[a.kolom]]

    if a.model == "transformer":
        prob = prediksi_transformer(teks, a.folder_model)
    else:
        prob = prediksi_tfidf(teks, a.path_tfidf)

    df["prediksi"] = [LABEL[i] for i in prob.argmax(1)]
    df["keyakinan"] = prob.max(1).round(3)
    Path(a.keluar).parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(a.keluar, index=False, encoding="utf-8-sig")
    print(f"{len(df)} baris -> {a.keluar}")
    print("Sebaran prediksi:", df["prediksi"].value_counts().to_dict())


if __name__ == "__main__":
    main()
