from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, date

db = SQLAlchemy()

# -------------------------------------------------------------------
# Model Pengguna
# -------------------------------------------------------------------
class User(db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(128), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="Admin")

    def set_password(self, password: str) -> None:
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)


# -------------------------------------------------------------------
# Model Produk / Data Kain
# -------------------------------------------------------------------
class Product(db.Model):
    __tablename__ = "products"
    id = db.Column(db.Integer, primary_key=True)
    kode_produk = db.Column(db.String(50), unique=True, nullable=True)
    name = db.Column(db.String(120), nullable=False)
    jenis_kain = db.Column(db.String(50), nullable=True)
    warna = db.Column(db.String(50), nullable=True)
    satuan = db.Column(db.String(20), nullable=False, default="Roll")
    harga_beli = db.Column(db.Float, nullable=False, default=0)
    price = db.Column(db.Float, nullable=False)       # Harga jual
    actual_stock = db.Column(db.Float, nullable=False, default=0.0)
    stok_minimum = db.Column(db.Float, nullable=False, default=10.0)
    description = db.Column(db.Text, nullable=True)

    @property
    def stock(self):
        return self.actual_stock

    @property
    def status_stok(self):
        if self.actual_stock <= 0:
            return "Habis"
        elif self.actual_stock <= self.stok_minimum:
            return "Menipis"
        return "Aman"

    # Relasi
    sales = db.relationship("Sale", backref="product", cascade="all, delete-orphan", passive_deletes=True, lazy=True)
    incoming_goods = db.relationship("IncomingGood", backref="product", cascade="all, delete-orphan", passive_deletes=True, lazy=True)


# -------------------------------------------------------------------
# Model Transaksi Penjualan
# -------------------------------------------------------------------
class Sale(db.Model):
    __tablename__ = "sales"
    id = db.Column(db.Integer, primary_key=True)
    nomor_transaksi = db.Column(db.String(30), unique=True, nullable=True)
    product_id = db.Column(db.Integer, db.ForeignKey("products.id", ondelete="CASCADE"), nullable=False)
    nama_pelanggan = db.Column(db.String(120), nullable=True, default="Pelanggan Umum")
    quantity = db.Column(db.Float, nullable=False)
    harga_satuan = db.Column(db.Float, nullable=False)
    total_price = db.Column(db.Float, nullable=False)
    kasir_nama = db.Column(db.String(80), nullable=True, default="Admin")
    timestamp = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)


# -------------------------------------------------------------------
# Model Barang Masuk (Penambahan Stok)
# -------------------------------------------------------------------
class IncomingGood(db.Model):
    __tablename__ = "incoming_goods"
    id = db.Column(db.Integer, primary_key=True)
    nomor_masuk = db.Column(db.String(30), unique=True, nullable=True)
    product_id = db.Column(db.Integer, db.ForeignKey("products.id", ondelete="CASCADE"), nullable=False)
    quantity = db.Column(db.Float, nullable=False)
    supplier = db.Column(db.String(120), nullable=True)
    notes = db.Column(db.Text, nullable=True)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)


# -------------------------------------------------------------------
# Inisialisasi Database
# -------------------------------------------------------------------
def init_db():
    db.create_all()
    if not User.query.filter_by(username="admin").first():
        admin = User(username="admin", role="Admin")
        admin.set_password("admin123")
        db.session.add(admin)
        db.session.commit()
