import os, re, hashlib, secrets
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from functools import wraps
from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_from_directory
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_, func
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

load_dotenv()
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "static" / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
ALLOWED_EXCEL = {".xlsx", ".xlsm", ".xls"}
ALLOWED_IMAGES = {".png", ".jpg", ".jpeg", ".webp", ".gif"}

app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "change-this-secret-before-deployment")
db_url = os.getenv("DATABASE_URL", "sqlite:///" + str(BASE_DIR / "local_dev.db"))
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)
if db_url.startswith("postgresql://") and "+psycopg" not in db_url:
    db_url = db_url.replace("postgresql://", "postgresql+psycopg://", 1)
app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
db = SQLAlchemy(app)

class Admin(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

class Product(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    barcode = db.Column(db.String(120), nullable=True, index=True)
    name = db.Column(db.String(300), nullable=False, index=True)
    price = db.Column(db.Numeric(14, 2), nullable=False, default=Decimal("0.00"))
    image = db.Column(db.String(500), nullable=True)
    note = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

class StoreSettings(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    store_name = db.Column(db.String(200), nullable=False, default="اسم المحل")
    logo_path = db.Column(db.String(500), nullable=True)
    label_width_mm = db.Column(db.Integer, nullable=False, default=58)
    label_height_mm = db.Column(db.Integer, nullable=False, default=40)
    print_barcode = db.Column(db.Boolean, nullable=False, default=True)
    currency = db.Column(db.String(20), nullable=False, default="DA")
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

class ImportRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    file_sha256 = db.Column(db.String(64), nullable=False, index=True)
    imported_rows = db.Column(db.Integer, nullable=False, default=0)
    created_rows = db.Column(db.Integer, nullable=False, default=0)
    updated_rows = db.Column(db.Integer, nullable=False, default=0)
    skipped_rows = db.Column(db.Integer, nullable=False, default=0)
    imported_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

def get_settings():
    row = StoreSettings.query.first()
    if row is None:
        row = StoreSettings(store_name=os.getenv("STORE_NAME", "اسم المحل"))
        db.session.add(row)
        db.session.commit()
    return row

def ensure_setup():
    db.create_all()
    username = os.getenv("ADMIN_USERNAME", "admin")
    password = os.getenv("ADMIN_PASSWORD", "ChangeMe123!")
    if not Admin.query.filter_by(username=username).first():
        db.session.add(Admin(username=username, password_hash=generate_password_hash(password)))
    if not StoreSettings.query.first():
        db.session.add(StoreSettings(store_name=os.getenv("STORE_NAME", "اسم المحل")))
    db.session.commit()

with app.app_context():
    ensure_setup()

def login_required(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        if not session.get("admin_id"):
            return redirect(url_for("login", next=request.path))
        return fn(*args, **kwargs)
    return wrapped

def normalize_barcode(value):
    if value is None:
        return ""
    # Preserve Excel string values and avoid common .0 suffix from numeric cells.
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null"}:
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text

def parse_price(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        raise ValueError("السعر فارغ")
    if isinstance(value, (int, float, Decimal)):
        amount = Decimal(str(value))
    else:
        text = str(value).strip().replace("DA", "").replace("دج", "").replace(" ", "").replace(",", ".")
        text = re.sub(r"[^0-9.\-]", "", text)
        if not text:
            raise ValueError("السعر غير صالح")
        amount = Decimal(text)
    if amount < 0:
        raise ValueError("السعر لا يمكن أن يكون سالبًا")
    return amount.quantize(Decimal("0.01"))

def normalized_header(value):
    text = str(value or "").strip().lower()
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    text = re.sub(r"[\s_\-]+", "", text)
    return text

HEADER_ALIASES = {
    "barcode": {"barcode", "bar code", "codebarre", "code a barre", "كودبار", "الباركود", "باركود", "رمزالمنتج", "code"},
    "name": {"name", "product", "productname", "designation", "design", "اسم", "اسم المنتج", "اسم المنتوج", "المنتوج", "المنتج", "التسمية", "designationproduit"},
    "price": {"price", "saleprice", "sellingprice", "prix", "prixvente", "prixdevente", "السعر", "سعرالبيع", "سعر", "سعرالمنتج", "سعرالمنتوج"},
    "note": {"note", "remark", "description", "ملاحظة", "ملاحظات", "الوصف"},
}

def map_columns(columns):
    mapping = {}
    for col in columns:
        h = normalized_header(col)
        for field, aliases in HEADER_ALIASES.items():
            normalized_aliases = {normalized_header(a) for a in aliases}
            if h in normalized_aliases and field not in mapping:
                mapping[field] = col
    return mapping

@app.route("/")
def home():
    if not session.get("admin_id"):
        return redirect(url_for("login"))
    return redirect(url_for("products"))

@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("admin_id"):
        return redirect(url_for("products"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        admin = Admin.query.filter_by(username=username).first()
        if admin and check_password_hash(admin.password_hash, password):
            session.clear()
            session["admin_id"] = admin.id
            session["admin_username"] = admin.username
            return redirect(request.args.get("next") or url_for("products"))
        flash("اسم المستخدم أو كلمة السر غير صحيحة.", "danger")
    return render_template("login.html")

@app.post("/logout")
@login_required
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/products")
@login_required
def products():
    query = request.args.get("q", "").strip()
    mode = request.args.get("mode", "table")
    page = max(int(request.args.get("page", 1) or 1), 1)
    pagination = Product.query
    if query:
        like = f"%{query}%"
        pagination = pagination.filter(or_(Product.name.ilike(like), Product.barcode.ilike(like)))
    pagination = pagination.order_by(Product.name.asc()).paginate(page=page, per_page=48, error_out=False)
    return render_template("products.html", pagination=pagination, query=query, mode=mode, settings=get_settings())

@app.route("/products/new", methods=["GET", "POST"])
@login_required
def product_new():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        barcode = normalize_barcode(request.form.get("barcode"))
        try:
            price = parse_price(request.form.get("price"))
        except Exception:
            flash("أدخل سعرًا صحيحًا.", "danger")
            return render_template("product_form.html", product=None, settings=get_settings())
        if not name:
            flash("اسم المنتج إجباري.", "danger")
            return render_template("product_form.html", product=None, settings=get_settings())
        if barcode and Product.query.filter_by(barcode=barcode).first():
            flash("هذا الباركود موجود مسبقًا. افتح المنتج الموجود لتحديثه.", "warning")
            return render_template("product_form.html", product=None, settings=get_settings())
        image_path = save_image(request.files.get("image"))
        product = Product(barcode=barcode or None, name=name, price=price,
                          note=request.form.get("note", "").strip(), image=image_path)
        db.session.add(product)
        db.session.commit()
        flash("تمت إضافة المنتج إلى قاعدة البيانات.", "success")
        return redirect(url_for("products"))
    return render_template("product_form.html", product=None, settings=get_settings())

@app.route("/products/<int:product_id>/edit", methods=["GET", "POST"])
@login_required
def product_edit(product_id):
    product = Product.query.get_or_404(product_id)
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        barcode = normalize_barcode(request.form.get("barcode"))
        try:
            price = parse_price(request.form.get("price"))
        except Exception:
            flash("أدخل سعرًا صحيحًا.", "danger")
            return render_template("product_form.html", product=product, settings=get_settings())
        duplicate = Product.query.filter(Product.barcode == barcode, Product.id != product.id).first() if barcode else None
        if duplicate:
            flash("هذا الباركود مرتبط بمنتج آخر.", "danger")
            return render_template("product_form.html", product=product, settings=get_settings())
        if not name:
            flash("اسم المنتج إجباري.", "danger")
            return render_template("product_form.html", product=product, settings=get_settings())
        product.name, product.barcode, product.price = name, barcode or None, price
        product.note = request.form.get("note", "").strip()
        new_image = save_image(request.files.get("image"))
        if new_image:
            product.image = new_image
        db.session.commit()
        flash("تم تحديث المنتج.", "success")
        return redirect(url_for("products"))
    return render_template("product_form.html", product=product, settings=get_settings())

@app.post("/products/<int:product_id>/delete")
@login_required
def product_delete(product_id):
    product = Product.query.get_or_404(product_id)
    db.session.delete(product)
    db.session.commit()
    flash("تم حذف المنتج.", "success")
    return redirect(url_for("products"))

def save_image(file_storage):
    if not file_storage or not file_storage.filename:
        return None
    ext = Path(secure_filename(file_storage.filename)).suffix.lower()
    if ext not in ALLOWED_IMAGES:
        raise ValueError("صيغة الصورة غير مدعومة")
    filename = secrets.token_hex(12) + ext
    file_storage.save(UPLOAD_DIR / filename)
    return "uploads/" + filename

@app.route("/settings", methods=["GET", "POST"])
@login_required
def settings():
    settings_row = get_settings()
    if request.method == "POST":
        settings_row.store_name = request.form.get("store_name", "").strip() or "اسم المحل"
        settings_row.currency = request.form.get("currency", "DA").strip()[:20] or "DA"
        try:
            width = int(request.form.get("label_width_mm", 58))
            height = int(request.form.get("label_height_mm", 40))
            if not (25 <= width <= 150 and 20 <= height <= 150):
                raise ValueError()
            settings_row.label_width_mm = width
            settings_row.label_height_mm = height
        except Exception:
            flash("مقاسات البطاقة يجب أن تكون بين 25 و150 مم للعرض وبين 20 و150 مم للارتفاع.", "danger")
            return render_template("settings.html", settings=settings_row, imports=ImportRecord.query.order_by(ImportRecord.imported_at.desc()).limit(10).all())
        settings_row.print_barcode = request.form.get("print_barcode") == "on"
        logo = request.files.get("logo")
        if logo and logo.filename:
            ext = Path(secure_filename(logo.filename)).suffix.lower()
            if ext not in ALLOWED_IMAGES:
                flash("صيغة اللوغو غير مدعومة.", "danger")
                return redirect(url_for("settings"))
            filename = "logo_" + secrets.token_hex(10) + ext
            logo.save(UPLOAD_DIR / filename)
            settings_row.logo_path = "uploads/" + filename
        db.session.commit()
        flash("تم حفظ إعدادات المتجر في قاعدة البيانات.", "success")
        return redirect(url_for("settings"))
    imports = ImportRecord.query.order_by(ImportRecord.imported_at.desc()).limit(10).all()
    return render_template("settings.html", settings=settings_row, imports=imports)

@app.post("/import-excel")
@login_required
def import_excel():
    file = request.files.get("excel_file")
    if not file or not file.filename:
        flash("اختر ملف Excel أولًا.", "danger")
        return redirect(url_for("settings"))
    ext = Path(secure_filename(file.filename)).suffix.lower()
    if ext not in ALLOWED_EXCEL:
        flash("الملف يجب أن يكون Excel بصيغة XLSX أو XLSM أو XLS.", "danger")
        return redirect(url_for("settings"))
    data = file.read()
    digest = hashlib.sha256(data).hexdigest()
    previous = ImportRecord.query.filter_by(file_sha256=digest).first()
    if previous:
        flash(f"هذا الملف مطابق تمامًا لملف سبق استيراده بتاريخ {previous.imported_at:%Y-%m-%d %H:%M}. لم تُكرر العملية.", "info")
        return redirect(url_for("settings"))
    temp_path = UPLOAD_DIR / ("import_" + secrets.token_hex(8) + ext)
    temp_path.write_bytes(data)
    try:
        engine = "openpyxl" if ext in {".xlsx", ".xlsm"} else None
        excel = pd.ExcelFile(temp_path, engine=engine)
        frame = pd.read_excel(excel, sheet_name=0, dtype=object)
        if frame.empty:
            raise ValueError("الورقة الأولى فارغة.")
        mapping = map_columns(frame.columns)
        if "name" not in mapping or "price" not in mapping:
            cols = "، ".join(str(c) for c in frame.columns)
            raise ValueError("لم أتعرف على أعمدة اسم المنتج والسعر. الأعمدة الموجودة: " + cols)
        created = updated = skipped = imported = 0
        # Entire import is transactional: a failure rolls back product changes.
        for _, row in frame.iterrows():
            try:
                name_val = row[mapping["name"]]
                if pd.isna(name_val) or not str(name_val).strip():
                    skipped += 1
                    continue
                name = str(name_val).strip()
                barcode = normalize_barcode(row[mapping["barcode"]]) if "barcode" in mapping else ""
                price = parse_price(row[mapping["price"]])
                note = str(row[mapping["note"]]).strip() if "note" in mapping and not pd.isna(row[mapping["note"]]) else ""
                product = Product.query.filter_by(barcode=barcode).first() if barcode else None
                if product:
                    changed = (product.name != name or product.price != price or (product.note or "") != note)
                    product.name, product.price, product.note = name, price, note
                    if changed:
                        updated += 1
                else:
                    product = Product(barcode=barcode or None, name=name, price=price, note=note)
                    db.session.add(product)
                    created += 1
                imported += 1
            except Exception:
                skipped += 1
        record = ImportRecord(filename=secure_filename(file.filename)[:255], file_sha256=digest,
                              imported_rows=imported, created_rows=created, updated_rows=updated,
                              skipped_rows=skipped)
        db.session.add(record)
        db.session.commit()
        flash(f"اكتمل الاستيراد: {created} منتج جديد، {updated} تحديث، {skipped} سطر متجاوز. تم حفظ البيانات في PostgreSQL.", "success")
    except Exception as exc:
        db.session.rollback()
        flash("لم يتم استيراد الملف: " + str(exc), "danger")
    finally:
        try: temp_path.unlink(missing_ok=True)
        except Exception: pass
    return redirect(url_for("settings"))

@app.get("/api/product-by-barcode")
@login_required
def api_product_by_barcode():
    barcode = normalize_barcode(request.args.get("barcode"))
    product = Product.query.filter_by(barcode=barcode).first()
    if not product:
        return jsonify({"ok": False, "message": "لم يتم العثور على منتج بهذا الباركود."}), 404
    return jsonify({"ok": True, "product": {
        "id": product.id, "barcode": product.barcode, "name": product.name,
        "price": f"{product.price:.2f}", "currency": get_settings().currency,
        "image": url_for("static", filename=product.image) if product.image else None
    }})

@app.get("/print-label/<int:product_id>")
@login_required
def print_label(product_id):
    product = Product.query.get_or_404(product_id)
    return render_template("label.html", product=product, settings=get_settings())

@app.errorhandler(413)
def too_large(_error):
    flash("حجم الملف أكبر من الحد المسموح (20 ميغابايت).", "danger")
    return redirect(url_for("settings"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)), debug=os.getenv("FLASK_DEBUG") == "1")
