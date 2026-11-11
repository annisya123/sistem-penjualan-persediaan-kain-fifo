import os
import csv
from io import StringIO
from datetime import datetime, timedelta
from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, session, Response
)
from models import db, init_db, User, Product, Sale, IncomingGood
from functools import wraps

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///data.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = 'cbjayatekstile-secret-key-2026'

db.init_app(app)

with app.app_context():
    init_db()

# ---------------------------------------------------------------------------
# Helper Otentikasi & Context
# ---------------------------------------------------------------------------
def login_required(view):
    @wraps(view)
    def wrapped_view(**kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return view(**kwargs)
    return wrapped_view

@app.context_processor
def inject_global_data():
    current_user = None
    if 'user_id' in session:
        current_user = db.session.get(User, session['user_id'])
    low_stock_count = Product.query.filter(
        Product.actual_stock <= Product.stok_minimum,
        Product.actual_stock > 0
    ).count()
    return dict(current_user=current_user, low_stock_count=low_stock_count)

def generate_nomor_transaksi():
    today_str = datetime.now().strftime('%Y%m%d')
    prefix = f"PJ-{today_str}-"
    last = Sale.query.filter(Sale.nomor_transaksi.like(f"{prefix}%")).order_by(Sale.id.desc()).first()
    seq = 1
    if last and last.nomor_transaksi:
        try:
            seq = int(last.nomor_transaksi.split('-')[-1]) + 1
        except ValueError:
            seq = 1
    return f"{prefix}{seq:03d}"

def generate_nomor_masuk():
    today_str = datetime.now().strftime('%Y%m%d')
    prefix = f"BM-{today_str}-"
    last = IncomingGood.query.filter(IncomingGood.nomor_masuk.like(f"{prefix}%")).order_by(IncomingGood.id.desc()).first()
    seq = 1
    if last and last.nomor_masuk:
        try:
            seq = int(last.nomor_masuk.split('-')[-1]) + 1
        except ValueError:
            seq = 1
    return f"{prefix}{seq:03d}"

# ---------------------------------------------------------------------------
# LOGIN & LOGOUT
# ---------------------------------------------------------------------------
@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('index'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            session['user_id'] = user.id
            session['username'] = user.username
            flash(f'Selamat datang, {user.username}!', 'success')
            return redirect(url_for('index'))
        flash('Username atau password tidak sesuai.', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('Anda telah keluar dari aplikasi.', 'info')
    return redirect(url_for('login'))

# ---------------------------------------------------------------------------
# 1. BERANDA / DASHBOARD
# ---------------------------------------------------------------------------
@app.route('/')
@login_required
def index():
    today = datetime.now().date()
    start_of_month = today.replace(day=1)

    all_products = Product.query.all()
    total_products = len(all_products)
    total_stok = sum(p.actual_stock for p in all_products)
    valuasi_aset = sum(p.actual_stock * p.price for p in all_products)
    perlu_restock = sum(1 for p in all_products if p.actual_stock <= p.stok_minimum)

    # Omset total penjualan
    omset_penjualan = db.session.query(db.func.sum(Sale.total_price)).scalar() or 0
    
    # Penjualan hari ini
    sales_today = Sale.query.filter(db.func.date(Sale.timestamp) == today).all()
    total_sales_today = len(sales_today)
    total_revenue_today = sum(s.total_price for s in sales_today)
    
    # Barang masuk bulan ini
    incoming_month = IncomingGood.query.filter(IncomingGood.timestamp >= start_of_month).all()
    incoming_month_qty = sum(i.quantity for i in incoming_month)

    # Data ringkas untuk 2 kolom di dashboard (sesuai Gambar 4.17 Wireframe)
    recent_sales = Sale.query.order_by(Sale.timestamp.desc()).limit(6).all()
    menipis_products = Product.query.filter(Product.actual_stock <= Product.stok_minimum).order_by(Product.actual_stock.asc()).limit(6).all()
    if not menipis_products:
        menipis_products = Product.query.order_by(Product.actual_stock.asc()).limit(6).all()

    stats = {
        'total_products': total_products,
        'total_stok': total_stok,
        'omset_penjualan': omset_penjualan,
        'valuasi_aset': valuasi_aset,
        'perlu_restock': perlu_restock,
        'total_sales_today': total_sales_today,
        'total_revenue_today': total_revenue_today,
        'incoming_month_qty': incoming_month_qty,
    }

    return render_template(
        'index.html',
        stats=stats,
        recent_sales=recent_sales,
        menipis_products=menipis_products,
        low_stock_count=perlu_restock
    )

# ---------------------------------------------------------------------------
# KALKULATOR KAIN (Sesuai Wireframe Gambar 4.23)
# ---------------------------------------------------------------------------
@app.route('/kalkulator')
@login_required
def kalkulator():
    products = Product.query.order_by(Product.name).all()
    return render_template('kalkulator_kain.html', products=products)

# ---------------------------------------------------------------------------
# 2. DATA & STOK KAIN
# ---------------------------------------------------------------------------
@app.route('/products')
@login_required
def products_page():
    q = request.args.get('q', '').strip()
    kategori = request.args.get('kategori', '').strip()
    
    query = Product.query
    if q:
        query = query.filter(
            (Product.name.ilike(f"%{q}%")) |
            (Product.kode_produk.ilike(f"%{q}%")) |
            (Product.jenis_kain.ilike(f"%{q}%")) |
            (Product.warna.ilike(f"%{q}%"))
        )
    if kategori and kategori != 'ALL' and kategori != 'Semua Kategori':
        query = query.filter(Product.jenis_kain == kategori)
        
    products = query.order_by(Product.id.asc()).all()
    categories = [r[0] for r in db.session.query(Product.jenis_kain).distinct().all() if r[0]]
    return render_template('products.html', products=products, categories=categories, current_q=q, current_kategori=kategori)

@app.route('/product/create', methods=['GET', 'POST'])
@login_required
def create_product():
    if request.method == 'POST':
        stock_val = float(request.form.get('stock', 0))
        new_prod = Product(
            kode_produk=request.form.get('kode_produk', '').strip() or None,
            name=request.form.get('name', '').strip(),
            jenis_kain=request.form.get('jenis_kain', 'Lainnya'),
            warna=request.form.get('warna', '').strip() or None,
            satuan=request.form.get('satuan', 'Roll'),
            harga_beli=float(request.form.get('harga_beli', 0)),
            price=float(request.form.get('price', 0)),
            actual_stock=stock_val,
            stok_minimum=float(request.form.get('stok_minimum', 10)),
            description=request.form.get('description', '').strip() or None
        )
        db.session.add(new_prod)
        db.session.commit()
        flash('Data kain baru berhasil ditambahkan!', 'success')
        return redirect(url_for('products_page'))
    return render_template('create.html')

@app.route('/product/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def edit_product(id):
    product = db.session.get(Product, id)
    if not product:
        flash('Produk tidak ditemukan.', 'danger')
        return redirect(url_for('products_page'))
    
    if request.method == 'POST':
        product.kode_produk = request.form.get('kode_produk', '').strip() or None
        product.name = request.form.get('name', '').strip()
        product.jenis_kain = request.form.get('jenis_kain', 'Lainnya')
        product.warna = request.form.get('warna', '').strip() or None
        product.satuan = request.form.get('satuan', 'Roll')
        product.harga_beli = float(request.form.get('harga_beli', 0))
        product.price = float(request.form.get('price', 0))
        product.actual_stock = float(request.form.get('stock', product.actual_stock))
        product.stok_minimum = float(request.form.get('stok_minimum', 10))
        product.description = request.form.get('description', '').strip() or None
        db.session.commit()
        flash('Data kain berhasil diperbarui!', 'success')
        return redirect(url_for('products_page'))
    return render_template('edit.html', product=product)

@app.route('/product/delete/<int:id>', methods=['POST'])
@login_required
def delete_product_route(id):
    product = db.session.get(Product, id)
    if product:
        db.session.delete(product)
        db.session.commit()
        flash(f'Kain {product.name} berhasil dihapus.', 'info')
    return redirect(url_for('products_page'))

# ---------------------------------------------------------------------------
# 3. KASIR / JUAL KAIN
# ---------------------------------------------------------------------------
@app.route('/transaction', methods=['GET', 'POST'])
@login_required
def create_transaction():
    products = Product.query.filter(Product.actual_stock > 0).order_by(Product.name).all()
    if request.method == 'POST':
        product_id = int(request.form['product_id'])
        qty = float(request.form['quantity'])
        nama_pelanggan = request.form.get('nama_pelanggan', '').strip() or 'Pelanggan Umum'
        product = db.session.get(Product, product_id)

        if not product or qty <= 0:
            flash('Data pembelian tidak valid.', 'danger')
            return redirect(url_for('create_transaction'))

        if qty > product.actual_stock:
            flash(f'Stok tidak cukup! Tersedia: {product.actual_stock} {product.satuan}.', 'danger')
            return redirect(url_for('create_transaction'))

        total = qty * product.price
        nomor = generate_nomor_transaksi()
        kasir_nama = session.get('username', 'Admin')

        # Catat Penjualan
        sale = Sale(
            nomor_transaksi=nomor,
            product_id=product.id,
            nama_pelanggan=nama_pelanggan,
            quantity=qty,
            harga_satuan=product.price,
            total_price=total,
            kasir_nama=kasir_nama
        )
        db.session.add(sale)
        
        # Kurangi stok langsung
        product.actual_stock -= qty
        db.session.commit()

        flash(f'Transaksi {nomor} berhasil dicatat!', 'success')
        return redirect(url_for('sale_detail', sale_id=sale.id))

    return render_template('transaction.html', products=products)

@app.route('/sale/<int:product_id>', methods=['GET', 'POST'])
@login_required
def create_sale(product_id):
    product = db.session.get(Product, product_id)
    if not product:
        flash('Produk tidak ditemukan.', 'danger')
        return redirect(url_for('products_page'))

    if request.method == 'POST':
        qty = float(request.form['quantity'])
        nama_pelanggan = request.form.get('nama_pelanggan', '').strip() or 'Pelanggan Umum'

        if qty <= 0 or qty > product.actual_stock:
            flash(f'Jumlah tidak valid atau melebihi stok ({product.actual_stock}).', 'danger')
            return redirect(url_for('create_sale', product_id=product_id))

        total = qty * product.price
        nomor = generate_nomor_transaksi()
        kasir_nama = session.get('username', 'Admin')

        sale = Sale(
            nomor_transaksi=nomor,
            product_id=product.id,
            nama_pelanggan=nama_pelanggan,
            quantity=qty,
            harga_satuan=product.price,
            total_price=total,
            kasir_nama=kasir_nama
        )
        db.session.add(sale)
        product.actual_stock -= qty
        db.session.commit()

        flash(f'Penjualan {nomor} berhasil!', 'success')
        return redirect(url_for('sale_detail', sale_id=sale.id))

    return render_template('sale.html', product=product)

@app.route('/sales/<int:sale_id>')
@login_required
def sale_detail(sale_id):
    sale = db.session.get(Sale, sale_id)
    if not sale:
        flash('Transaksi tidak ditemukan.', 'danger')
        return redirect(url_for('sales_history'))
    return render_template('transaction_detail.html', sale=sale)

@app.route('/sales')
@login_required
def sales_history():
    sales = Sale.query.order_by(Sale.timestamp.desc()).all()
    return render_template('sales.html', sales=sales)

# ---------------------------------------------------------------------------
# 4. STOK MASUK
# ---------------------------------------------------------------------------
@app.route('/stock/incoming', methods=['GET', 'POST'])
@login_required
def stock_incoming():
    products = Product.query.order_by(Product.name).all()
    if request.method == 'POST':
        product_id = int(request.form['product_id'])
        qty = float(request.form['quantity'])
        supplier = request.form.get('supplier', '').strip() or 'Supplier Resmi'
        notes = request.form.get('notes', '').strip() or None
        product = db.session.get(Product, product_id)

        if not product or qty <= 0:
            flash('Jumlah stok masuk tidak valid.', 'danger')
            return redirect(url_for('stock_incoming'))

        custom_no = request.form.get('nomor_masuk', '').strip()
        nomor = custom_no if custom_no else generate_nomor_masuk()
        entry = IncomingGood(
            nomor_masuk=nomor,
            product_id=product.id,
            quantity=qty,
            supplier=supplier,
            notes=notes
        )
        db.session.add(entry)
        product.actual_stock += qty
        db.session.commit()

        flash(f'Stok {product.name} bertambah +{qty} {product.satuan}!', 'success')
        return redirect(url_for('stock_incoming'))

    incoming = IncomingGood.query.order_by(IncomingGood.timestamp.desc()).all()
    return render_template('stock_incoming.html', products=products, incoming=incoming)

# ---------------------------------------------------------------------------
# 5. LAPORAN & EXPORT CSV
# ---------------------------------------------------------------------------
@app.route('/laporan')
@login_required
def laporan():
    start_date_str = request.args.get('start_date', '')
    end_date_str = request.args.get('end_date', '')
    active_tab = request.args.get('tab', 'penjualan')

    # Penjualan
    sq = Sale.query
    if start_date_str:
        try:
            sq = sq.filter(Sale.timestamp >= datetime.strptime(start_date_str, '%Y-%m-%d'))
        except ValueError:
            pass
    if end_date_str:
        try:
            sq = sq.filter(Sale.timestamp < datetime.strptime(end_date_str, '%Y-%m-%d') + timedelta(days=1))
        except ValueError:
            pass
    sales = sq.order_by(Sale.timestamp.desc()).all()
    total_qty = sum(s.quantity for s in sales)
    total_revenue = sum(s.total_price for s in sales)
    avg_transaction = (total_revenue / len(sales)) if sales else 0

    # Barang Masuk
    iq = IncomingGood.query
    if start_date_str:
        try:
            iq = iq.filter(IncomingGood.timestamp >= datetime.strptime(start_date_str, '%Y-%m-%d'))
        except ValueError:
            pass
    if end_date_str:
        try:
            iq = iq.filter(IncomingGood.timestamp < datetime.strptime(end_date_str, '%Y-%m-%d') + timedelta(days=1))
        except ValueError:
            pass
    incoming = iq.order_by(IncomingGood.timestamp.desc()).all()
    total_incoming_qty = sum(i.quantity for i in incoming)

    # Persediaan
    products = Product.query.order_by(Product.name).all()
    total_stok = sum(p.actual_stock for p in products)
    total_nilai = sum(p.actual_stock * p.price for p in products)
    stok_aman_count = sum(1 for p in products if p.actual_stock > p.stok_minimum)
    stok_menipis_count = sum(1 for p in products if p.actual_stock <= p.stok_minimum and p.actual_stock > 0)
    stok_habis_count = sum(1 for p in products if p.actual_stock <= 0)

    return render_template(
        'laporan.html',
        sales=sales,
        total_qty=total_qty,
        total_revenue=total_revenue,
        avg_transaction=avg_transaction,
        incoming=incoming,
        total_incoming_qty=total_incoming_qty,
        products=products,
        total_stok=total_stok,
        total_nilai=total_nilai,
        stok_aman_count=stok_aman_count,
        stok_menipis_count=stok_menipis_count,
        stok_habis_count=stok_habis_count,
        start_date=start_date_str,
        end_date=end_date_str,
        active_tab=active_tab,
        now=datetime.now()
    )

@app.route('/laporan/export/sales')
@login_required
def export_sales():
    sales = Sale.query.order_by(Sale.timestamp.desc()).all()
    data = StringIO()
    w = csv.writer(data)
    w.writerow(['No. Transaksi', 'Tanggal', 'Nama Kain', 'Kode', 'Jumlah', 'Satuan', 'Harga Satuan (Rp)', 'Total (Rp)', 'Pelanggan', 'Kasir'])
    for s in sales:
        w.writerow([
            s.nomor_transaksi or s.id,
            s.timestamp.strftime('%Y-%m-%d %H:%M'),
            s.product.name if s.product else '-',
            s.product.kode_produk if s.product else '-',
            s.quantity,
            s.product.satuan if s.product else 'Roll',
            int(s.harga_satuan),
            int(s.total_price),
            s.nama_pelanggan or 'Umum',
            s.kasir_nama or 'Admin'
        ])
    return Response(
        data.getvalue(),
        mimetype='text/csv; charset=utf-8',
        headers={"Content-Disposition": "attachment; filename=laporan_penjualan.csv"}
    )

@app.route('/laporan/export/incoming')
@login_required
def export_incoming():
    incoming = IncomingGood.query.order_by(IncomingGood.timestamp.desc()).all()
    data = StringIO()
    w = csv.writer(data)
    w.writerow(['No. Masuk', 'Tanggal', 'Nama Kain', 'Kode', 'Jumlah Masuk', 'Satuan', 'Supplier', 'Catatan'])
    for i in incoming:
        w.writerow([
            i.nomor_masuk or i.id,
            i.timestamp.strftime('%Y-%m-%d %H:%M'),
            i.product.name if i.product else '-',
            i.product.kode_produk if i.product else '-',
            i.quantity,
            i.product.satuan if i.product else 'Roll',
            i.supplier or '-',
            i.notes or '-'
        ])
    return Response(
        data.getvalue(),
        mimetype='text/csv; charset=utf-8',
        headers={"Content-Disposition": "attachment; filename=laporan_barang_masuk.csv"}
    )

@app.route('/laporan/export/stock')
@login_required
def export_stock():
    products = Product.query.order_by(Product.name).all()
    data = StringIO()
    w = csv.writer(data)
    w.writerow(['Kode', 'Nama Kain', 'Jenis', 'Warna', 'Satuan', 'Harga Beli (Rp)', 'Harga Jual (Rp)', 'Sisa Stok', 'Status Stok', 'Nilai Stok (Rp)'])
    for p in products:
        w.writerow([
            p.kode_produk or '-',
            p.name,
            p.jenis_kain or '-',
            p.warna or '-',
            p.satuan,
            int(p.harga_beli),
            int(p.price),
            p.actual_stock,
            p.status_stok,
            int(p.actual_stock * p.price)
        ])
    return Response(
        data.getvalue(),
        mimetype='text/csv; charset=utf-8',
        headers={"Content-Disposition": "attachment; filename=laporan_stok_kain.csv"}
    )

# ---------------------------------------------------------------------------
if __name__ == '__main__':
    app.run(debug=True)
