# ===========================================================
# excel_store.py — FINAL VERSION (SINKRON DENGAN STRUKTUR BARU)
# ===========================================================

import os
import pandas as pd

DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "data.xlsx")

# Struktur kolom baru (WAJIB)
COLUMNS = [
    "id",
    "nama_produk",
    "hpp",
    "profit",
    "harga_25",
    "harga",
    "sumber",
    "keterangan",
]


def ensure_store():
    """Pastikan folder & file data.xlsx tersedia."""
    os.makedirs(DATA_DIR, exist_ok=True)

    if not os.path.exists(DATA_FILE):
        # Buat contoh data awal
        df = pd.DataFrame([
            {
                "id": "8991234567890",
                "nama_produk": "Bantal Iskra",
                "hpp": 20000,
                "profit": 30000,
                "harga_25": 25000,
                "harga": 50000,
                "sumber": "Default",
                "keterangan": "Bantal putih empuk",
            },
            {
                "id": "8999876543210",
                "nama_produk": "Kasur Lipat Iskra 6cm",
                "hpp": 150000,
                "profit": 100000,
                "harga_25": 180000,
                "harga": 250000,
                "sumber": "Default",
                "keterangan": "Ringan, mudah dibawa",
            },
        ], columns=COLUMNS)

        df.to_excel(DATA_FILE, index=False)


def _normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalisasi kolom dan tipe data agar stabil."""

    # --- Pastikan kolom single-index (bukan multiheader)
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [c[0] for c in df.columns]

    # --- Pastikan semua kolom jadi lowercase tanpa spasi
    df.columns = [str(c).strip().lower() for c in df.columns]

    # --- Mapping alias
    alias = {
        "id": "id",
        "kode": "id",
        "nama": "nama_produk",
        "nama produk": "nama_produk",
        "product_name": "nama_produk",
        "hpp": "hpp",
        "profit": "profit",
        "harga 25": "harga_25",
        "harga_25": "harga_25",
        "harga": "harga",
        "sumber": "sumber",
        "keterangan": "keterangan",
    }

    fixed_cols = {}
    for c in df.columns:
        key = c.strip().lower()
        fixed_cols[c] = alias.get(key, key)

    df = df.rename(columns=fixed_cols)

    # --- Kolom Wajib
    REQUIRED = {"id", "nama_produk", "hpp", "profit", "harga_25", "harga", "sumber", "keterangan"}
    for col in REQUIRED:
        if col not in df.columns:
            df[col] = ""

    # --- Pastikan Series, bukan DataFrame
    df["id"] = df["id"].astype(str)
    df["id"] = df["id"].str.strip()

    df["nama_produk"] = df["nama_produk"].astype(str).str.strip()

    # --- Numerik aman
    for col in ["hpp", "profit", "harga_25", "harga"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    df["sumber"] = df["sumber"].astype(str).str.strip()
    df["keterangan"] = df["keterangan"].astype(str)

    # --- Urutkan kolom
    ORDER = ["id", "nama_produk", "hpp", "profit", "harga_25", "harga", "sumber", "keterangan"]
    df = df[ORDER]

    # --- Drop ID kosong
    df = df[df["id"] != ""]

    return df


def load_products() -> pd.DataFrame:
    """Load data dari Excel."""
    ensure_store()
    df = pd.read_excel(DATA_FILE, dtype={"id": str})
    return _normalize_df(df)


def save_products(df: pd.DataFrame):
    """Simpan ke Excel."""
    ensure_store()
    df = _normalize_df(df)
    df.to_excel(DATA_FILE, index=False)


def merge_upsert(new_df: pd.DataFrame):
    """
    Update + insert (UPSERT):
    - jika id sama → replace
    - jika id baru → tambah
    """
    base = load_products()
    new_df = _normalize_df(new_df)

    combined = pd.concat([base, new_df], ignore_index=True)
    combined = combined.drop_duplicates(subset=["id"], keep="last")

    save_products(combined)