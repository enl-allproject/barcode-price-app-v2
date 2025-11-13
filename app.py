# app.py — FINAL STABLE VERSION FOR PYTHONANYWHERE

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
# ENV LOADER (PythonAnywhere + local development)
# ======================================================
# find_dotenv() memastikan .env ditemukan walau directory berubah
load_dotenv(find_dotenv(), override=True)

APP_USER       = os.getenv("APP_USERNAME", "admin")
APP_PASS       = os.getenv("APP_PASSWORD")           # plaintext mode (opsional)
APP_PASS_HASH  = os.getenv("APP_PASSWORD_HASH")      # hashed mode (opsional)
SECRET_KEY     = os.getenv("SECRET_KEY", "dev-secret")

# ======================================================
# APP CONFIG
# ======================================================
app = Flask(__name__)
app.secret_key = SECRET_KEY
app.permanent_session_lifetime = timedelta(days=7)

login_manager = LoginManager(app)
login_manager.login_view = "login"

# ======================================================
# FILE STORAGE (Excel)
# ======================================================
try:
    from excel_store import load_products
except Exception:
    load_products = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FALLBACK_PATH = os.path.join(BASE_DIR, "products.xlsx")


def _load_products_df() -> pd.DataFrame:
    """Load product list from excel_store or products.xlsx"""
    if load_products:
        df = load_products()
    elif os.path.exists(FALLBACK_PATH):
        if FALLBACK_PATH.lower().endswith(".csv"):
            df = pd.read_csv(FALLBACK_PATH)
        else:
            df = pd.read_excel(FALLBACK_PATH)
    else:
        df = pd.DataFrame(columns=["id", "nama", "harga", "foto", "keterangan"])

    if "id" in df.columns:
        df["id"] = df["id"].astype(str)
    return df


def _save_products_df(df: pd.DataFrame):
    df.to_excel(FALLBACK_PATH, index=False)

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
            # Priority 1 — password hash
            if APP_PASS_HASH:
                ok = check_password_hash(APP_PASS_HASH, p)
            # Priority 2 — plaintext password
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
    row = df[df["id"].astype(str) == str(id_value)]

    if row.empty:
        abort(404, "Produk tidak ditemukan")

    p = row.iloc[0].to_dict()

    try:
        p["harga"] = int(float(p.get("harga", 0)))
    except Exception:
        p["harga"] = 0

    p.setdefault("nama", "")
    p.setdefault("foto", "")
    p.setdefault("keterangan", "")

    return render_template("product.html", p=p)


@app.route("/database")
@login_required
def database():
    df = _load_products_df()
    records = df.to_dict(orient="records")

    for r in records:
        try:
            r["harga"] = int(float(r.get("harga", 0)))
        except Exception:
            r["harga"] = 0

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

            if filename.endswith(".csv"):
                df = pd.read_csv(f)
            else:
                df = pd.read_excel(f)

            required = {"id", "nama", "harga", "foto", "keterangan"}
            missing = [c for c in required if c not in df.columns]

            if missing:
                flash(f"Kolom wajib hilang: {', '.join(missing)}", "error")
                return redirect(url_for("import_excel"))

            df["id"] = df["id"].astype(str)
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