# excel_store.py — FINAL SUPER SIMPLE
# Header yang didukung:
# id, nama_produk, hpp, profit, harga_25, harga, sumber, keterangan

import os
import pandas as pd

DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "data.xlsx")

# Kolom standar internal aplikasi
COLUMNS = [
    "id",
    "nama",        # diisi dari nama_produk
    "hpp",
    "profit",
    "harga_25",
    "harga",
    "sumber",
    "keterangan",
    "foto",
    "barcode",
]


def ensure_store():
    """Pastikan folder & file Excel ada. Kalau belum ada, buat dengan contoh data."""
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(DATA_FILE):
        df = pd.DataFrame(
            [
                {
                    "id": "8991234567890",
                    "nama": "Bantal Iskra",
                    "hpp": 30000,
                    "profit": 20000,
                    "harga_25": 0,
                    "harga": 50000,
                    "sumber": "contoh",
                    "keterangan": "Bantal putih empuk",
                    "foto": "",
                    "barcode": "",
                },
                {
                    "id": "8999876543210",
                    "nama": "Kasur Lipat Iskra 6cm",
                    "hpp": 150000,
                    "profit": 100000,
                    "harga_25": 0,
                    "harga": 250000,
                    "sumber": "contoh",
                    "keterangan": "Ringan, mudah dibawa",
                    "foto": "",
                    "barcode": "",
                },
            ],
            columns=COLUMNS,
        )
        df.to_excel(DATA_FILE, index=False)


def _normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalisasi super sederhana:
    - Wajib punya kolom 'id'
    - 'nama' SELALU diisi dari 'nama_produk' kalau ada
    - Angka dibiarkan apa adanya (tidak diubah ke 0)
    """

    # 1) Nama kolom ke lowercase + strip spasi
    lower_map = {c: str(c).lower().strip() for c in df.columns}
    df = df.rename(columns=lower_map)

    # 2) Cek id wajib
    if "id" not in df.columns:
        raise ValueError("Kolom wajib hilang: id. Minimal harus ada: id.")

    # 3) Paksa nama diisi dari nama_produk kalau ada
    if "nama_produk" in df.columns:
        df["nama"] = df["nama_produk"].astype(str)
    elif "nama" in df.columns:
        df["nama"] = df["nama"].astype(str)
    else:
        df["nama"] = ""

    # 4) Tambah kolom lain kalau belum ada
    for col in ["hpp", "profit", "harga_25", "harga"]:
        if col not in df.columns:
            df[col] = 0

    if "sumber" not in df.columns:
        df["sumber"] = ""
    if "keterangan" not in df.columns:
        df["keterangan"] = ""
    if "foto" not in df.columns:
        df["foto"] = ""
    if "barcode" not in df.columns:
        df["barcode"] = ""

    # 5) Rapikan tipe dasar yang penting
    df["id"] = df["id"].astype(str).str.strip()
    df["nama"] = df["nama"].astype(str).str.strip()
    df["keterangan"] = df["keterangan"].astype(str)

    # 6) Urutkan kolom → buang id kosong
    df = df[COLUMNS]
    df = df[df["id"] != ""]
    return df


def load_products() -> pd.DataFrame:
    """Load dari file Excel utama."""
    ensure_store()
    df = pd.read_excel(DATA_FILE)
    return _normalize_df(df)


def save_products(df: pd.DataFrame):
    """Simpan DF (setelah dinormalisasi) ke file Excel."""
    ensure_store()
    df = _normalize_df(df)
    df.to_excel(DATA_FILE, index=False)


def merge_upsert(new_df: pd.DataFrame):
    """
    Gabungkan data baru dengan data lama.
    Jika ada id yang sama, baris baru menimpa yang lama.
    """
    base = load_products()
    new_df = _normalize_df(new_df)
    combined = pd.concat([base, new_df], ignore_index=True)
    combined = combined.drop_duplicates(subset=["id"], keep="last")
    save_products(combined)