# excel_store.py — SIMPLE FINAL (header: id, nama_produk, hpp, profit, harga_25, harga, sumber, keterangan)

import os
import pandas as pd

DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "data.xlsx")

# Struktur internal standar
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


def _clean_number(series: pd.Series) -> pd.Series:
    """
    Bersihkan angka dengan 2 mode:
    - Kalau dari Excel sudah numeric (int/float) -> langsung konversi ke int.
    - Kalau teks (misal 'Rp 50.000', '50.000') -> pakai regex buang simbol.
    """
    # Kalau sudah numeric (int/float) tinggal rapikan saja
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0).round().astype(int)

    # Kalau teks, baru dibersihkan
    s = series.astype(str)
    # buang semua karakter selain digit dan minus
    s = s.str.replace(r"[^0-9\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)


def _normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalisasi sesuai header:
    id, nama_produk, hpp, profit, harga_25, harga, sumber, keterangan
    + kolom opsional foto, barcode.
    """
    # 1) Lowercase & strip nama kolom
    lower_map = {c: str(c).lower().strip() for c in df.columns}
    df = df.rename(columns=lower_map)

    # 2) Mapping langsung dari header kamu + beberapa variasi umum
    rename_map = {
        "id": "id",
        # nama produk
        "nama_produk": "nama",
        "nama produk": "nama",
        "nama": "nama",
        "nama barang": "nama",
        "product_name": "nama",
        "product name": "nama",

        # angka
        "hpp": "hpp",
        "profit": "profit",
        "harga_25": "harga_25",
        "harga 25%": "harga_25",
        "harga": "harga",

        # sumber
        "sumber": "sumber",
        "source": "sumber",

        # keterangan
        "keterangan": "keterangan",
        "deskripsi": "keterangan",
        "notes": "keterangan",

        # foto
        "foto": "foto",
        "gambar": "foto",
        "image": "foto",

        # barcode opsional
        "barcode": "barcode",
    }

    cols_new = {}
    for c in df.columns:
        key = c.strip().lower()
        cols_new[c] = rename_map.get(key, key)
    df = df.rename(columns=cols_new)

    # 3) Kolom wajib: minimal id
    if "id" not in df.columns:
        raise ValueError("Kolom wajib hilang: id. Minimal harus ada: id.")

    # 4) Kalau 'nama' belum ada, coba isi dari nama_produk atau kolom lain yang mirip
    if "nama" not in df.columns:
        if "nama_produk" in df.columns:
            df["nama"] = df["nama_produk"].astype(str)
        else:
            # fallback kasar: cari kolom yang mengandung 'nama' atau 'produk'
            candidate = None
            for c in df.columns:
                k = c.lower()
                if c == "id":
                    continue
                if any(word in k for word in ["nama", "produk", "product", "barang"]):
                    candidate = c
                    break
            if candidate:
                df["nama"] = df[candidate].astype(str)
            else:
                df["nama"] = ""

    # 5) Tambahkan kolom yang belum ada
    for col in COLUMNS:
        if col not in df.columns:
            if col in ["hpp", "profit", "harga_25", "harga"]:
                df[col] = 0
            else:
                df[col] = ""

    # 6) Normalisasi tipe data
    df["id"] = df["id"].astype(str).str.strip()
    df["nama"] = df["nama"].astype(str).str.strip()
    df["keterangan"] = df["keterangan"].astype(str)
    df["sumber"] = df["sumber"].astype(str).str.strip()
    df["foto"] = df["foto"].astype(str).str.strip()
    df["barcode"] = df["barcode"].astype(str).str.strip()

    # angka → pakai _clean_number yang baru
    df["hpp"] = _clean_number(df["hpp"])
    df["profit"] = _clean_number(df["profit"])
    df["harga_25"] = _clean_number(df["harga_25"])
    df["harga"] = _clean_number(df["harga"])

    # 7) Urutkan kolom & buang id kosong
    df = df[COLUMNS]
    df = df[df["id"] != ""]
    return df


def load_products() -> pd.DataFrame:
    """Load dari file Excel utama."""
    ensure_store()
    df = pd.read_excel(DATA_FILE, dtype={"id": str})
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