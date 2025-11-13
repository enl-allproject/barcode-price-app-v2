# excel_store.py
import os
import pandas as pd

DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "data.xlsx")
# Kolom opsional 'barcode' ikut diekspor; pembuatannya tetap di web
COLUMNS = ["id", "nama", "harga", "foto", "keterangan", "barcode"]

def ensure_store():
    """Pastikan folder & file Excel ada. Jika belum ada, buat dengan header + contoh data."""
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(DATA_FILE):
        df = pd.DataFrame([
            {"id": "8991234567890", "nama": "Bantal Iskra", "harga": 50000, "foto": "", "keterangan": "Bantal putih empuk", "barcode": ""},
            {"id": "8999876543210", "nama": "Kasur Lipat Iskra 6cm", "harga": 250000, "foto": "", "keterangan": "Ringan, mudah dibawa", "barcode": ""},
        ], columns=COLUMNS)
        df.to_excel(DATA_FILE, index=False)

def _normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Pastikan kolom & tipe data konsisten."""
    # pastikan nama kolom string lower
    mapping = {c: str(c).lower() for c in df.columns}
    df = df.rename(columns=mapping)

    alias_map = {
        "id": "id", "kode": "id",
        "barcode": "barcode",  # opsional, bukan pengganti id
        "nama": "nama", "nama barang": "nama", "product_name": "nama",
        "harga": "harga", "harga barang": "harga", "price": "harga",
        "foto": "foto", "gambar": "foto", "image": "foto",
        "keterangan": "keterangan", "deskripsi": "keterangan", "notes": "keterangan",
    }
    cols_new = {}
    for c in df.columns:
        key = c.strip().lower()
        cols_new[c] = alias_map.get(key, key)
    df = df.rename(columns=cols_new)

    # kolom wajib
    required = {"id", "nama", "harga"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Kolom wajib hilang: {', '.join(sorted(missing))}. Minimal: id, nama, harga.")

    # tambahkan kolom opsional jika belum ada
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = "" if col != "harga" else 0

    # normalisasi tipe + trimming spasi (pakai .str.strip())
    df["id"] = df["id"].astype(str).str.strip()
    df["nama"] = df["nama"].astype(str).str.strip()
    df["harga"] = pd.to_numeric(df["harga"], errors="coerce").fillna(0).astype(int)
    df["foto"] = df["foto"].astype(str).str.strip()
    df["keterangan"] = df["keterangan"].astype(str)
    df["barcode"] = df["barcode"].astype(str).str.strip()

    # urutkan kolom + drop id kosong
    df = df[COLUMNS]
    df = df[df["id"] != ""]
    return df

def load_products() -> pd.DataFrame:
    ensure_store()
    df = pd.read_excel(DATA_FILE, dtype={"id": str})
    return _normalize_df(df)

def save_products(df: pd.DataFrame):
    ensure_store()
    df = _normalize_df(df)
    df.to_excel(DATA_FILE, index=False)

def merge_upsert(new_df: pd.DataFrame):
    base = load_products()
    new_df = _normalize_df(new_df)
    combined = pd.concat([base, new_df], ignore_index=True)
    combined = combined.drop_duplicates(subset=["id"], keep="last")
    save_products(combined)