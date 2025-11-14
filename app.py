# app.py — FINAL: baca Excel langsung, harga & angka sebagai TEKS

import os
from io import BytesIO
from datetime import timedelta

import pandas as pd
from flask import (
    Flask, render_template, request, redirect,
    url_for, send_file, flash, abort
)
from flask_login import (
    LoginManager, UserMixin, login_user,
    login_required, logout_user
)
from dotenv import load_dotenv, find_dotenv
from werkzeug.security import check_password_hash

# ======================================================
# ENV LOADER
# ======================================================
load_dotenv(find_dotenv(), override=True)

APP_USER      = os.getenv("APP_USERNAME", "admin")
APP_PASS      = os.getenv("APP_PASSWORD")           # plaintext mode (opsional)
APP_PASS_HASH = os.getenv("APP_PASSWORD_HASH")      # hashed mode (opsional)
SECRET_KEY    = os.getenv("SECRET_KEY", "dev-secret")

# ======================================================
# APP CONFIG
# ======================================================
app = Flask(__name__)
app.secret_key = SECRET_KEY
app.permanent_session_lifetime = timedelta(days=7)

login_manager = LoginManager(app)
login_manager.login_view = "login"

# ======================================================
# SIMPLE EXCEL STORAGE
# ======================================================
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_DIR   = os.path.join(BASE_DIR, "data")
DATA_FILE  = os.path.join(DATA_DIR, "data.xlsx")

DEFAULT_COLUMNS = [
    "id",
    "nama_produk",
    "hpp",
    "profit",
    "harga_25",
    "harga",
    "sumber",
    "keterangan",
]


def ensure_store_exists():
    """Pastikan folder & file Excel ada."""
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(DATA_FILE):
        df = pd.DataFrame(columns=DEFAULT_COLUMNS)
        df.to_excel(DATA_FILE, index=False)


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalisasi nama kolom:
    - lowercase
    - strip spasi pinggir
    - spasi di tengah diganti '_' (harga jual -> harga_jual)
    """
    new_cols = {}
    for c in df.columns:
        key = str(c).strip().lower().replace(" ", "_")
        new_cols[c] = key
    df = df.rename(columns=new_cols)
    return df


def _load_products_df() -> pd.DataFrame:
    """Baca data dari data.xlsx, lalu normalisasi header."""
    ensure_store_exists()
    try:
        df = pd.read_excel(DATA_FILE)
    except Exception:
        df = pd.DataFrame(columns=DEFAULT_COLUMNS)

    df = _normalize_columns(df)

    # pastikan kolom id selalu ada
    if "id" not in df.columns:
        df["id"] = ""

    df["id"] = df["id"].astype(str).str.strip()
    return df


def _save_products_df(df: pd.DataFrame):
    """Simpan DataFrame ke data.xlsx dengan header yang sudah dinormalisasi."""
    os.makedirs(DATA_DIR, exist_ok=True)
    df = _normalize_columns(df)
    df.to_excel(DATA_FILE, index=False)


def _get_first_nonempty(row, candidates, default=""):
    """Ambil nilai pertama yang tidak kosong dari beberapa nama kolom (sudah dinormalisasi)."""
    for name in candidates:
        if name in row.index:
            val = row[name]
            if pd.isna(val):
                continue
            text = str(val).strip()
            if text != "":
                return text
    return default

# ======================================================
# LOGIN MODEL
# ======================================================
class SimpleUser(UserMixin):
    def __init__(self, uid: str):
        self.id = uid


@login_manager.user_loader
def _load_user(uid: str):
    return SimpleUser(uid) if uid == APP_USER else None

# ======================================================
# AUTH ROUTES
# ======================================================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = (request.form.get("username") or "").strip()
        p = request.form.get("password") or ""

        next_url = (
            request.args.get("next")
            or request.form.get("next")
            or url_for("database")
        )

        ok = False

        if u == APP_USER:
            if APP_PASS_HASH:
                ok = check_password_hash(APP_PASS_HASH, p)
            elif APP_PASS is not None:
                ok = (p == APP_PASS)

        if ok:
            login_user(SimpleUser(APP_USER), remember=True,
                       duration=timedelta(days=7))
            return redirect(next_url)

        flash("Login gagal: username atau password salah.", "error")
        return render_template("login.html")

    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("home"))

# ======================================================
# PAGE ROUTES
# ======================================================
@app.route("/")
def home():
    return render_template("home.html")


@app.route("/p/<id_value>")
def product_page(id_value):
    df = _load_products_df()
    row_df = df[df["id"].astype(str) == str(id_value)]

    if row_df.empty:
        abort(404, "Produk tidak ditemukan")

    row = row_df.iloc[0]

    # nama produk: prioritaskan nama_produk, lalu nama, dll
    nama = _get_first_nonempty(
        row,
        ["nama_produk", "nama", "nama_barang", "product_name"],
        default="",
    )

    # harga: apa adanya dari kolom harga (atau harga_jual / price)
    harga_text = _get_first_nonempty(
        row,
        ["harga", "harga_jual", "price"],
        default="",
    )

    keterangan = _get_first_nonempty(
        row,
        ["keterangan", "deskripsi", "notes"],
        default="",
    )

    p = {
        "id": str(row.get("id", "")),
        "nama": nama,
        "harga_text": harga_text,
        "keterangan": keterangan,
    }

    return render_template("product.html", p=p)


@app.route("/database")
@login_required
def database():
    df = _load_products_df()

    records = []
    for _, row in df.iterrows():
        rid = str(row.get("id", "")).strip()
        if not rid:
            continue

        nama_produk = _get_first_nonempty(
            row,
            ["nama_produk", "nama", "nama_barang", "product_name"],
            default="",
        )

        hpp_text      = _get_first_nonempty(row, ["hpp"], default="")
        profit_text   = _get_first_nonempty(row, ["profit"], default="")
        harga25_text  = _get_first_nonempty(row, ["harga_25"], default="")
        harga_text    = _get_first_nonempty(row, ["harga", "harga_jual", "price"], default="")
        sumber_text   = _get_first_nonempty(row, ["sumber", "source"], default="")
        ket_text      = _get_first_nonempty(row, ["keterangan", "deskripsi", "notes"], default="")

        records.append({
            "id": rid,
            "nama_produk": nama_produk,
            "hpp": hpp_text,
            "profit": profit_text,
            "harga_25": harga25_text,
            "harga": harga_text,
            "sumber": sumber_text,
            "keterangan": ket_text,
        })

    return render_template("database.html", products=records)


@app.route("/export-excel")
@login_required
def export_excel():
    df = _load_products_df()
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)

    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="products.xlsx"
    )


@app.route("/import-excel", methods=["GET", "POST"])
@login_required
def import_excel():
    if request.method == "POST":
        f = request.files.get("file")
        if not f:
            flash("Tidak ada file yang diunggah.", "error")
            return redirect(url_for("import_excel"))

        try:
            filename = (f.filename or "").lower()

            # baca file mentah
            if filename.endswith(".csv"):
                df = pd.read_csv(f)
            else:
                df = pd.read_excel(f)

            # normalisasi nama kolom sebelum simpan
            df = _normalize_columns(df)

            # pastikan ada kolom id
            if "id" not in df.columns:
                raise ValueError("File harus punya kolom 'id'.")

            _save_products_df(df)

            flash("Import berhasil!", "ok")
            return redirect(url_for("database"))

        except Exception as e:
            flash(f"Gagal import: {e}", "error")
            return redirect(url_for("import_excel"))

    return render_template("import.html")

# ======================================================
# MAIN
# ======================================================
if __name__ == "__main__":
    app.run(debug=True)