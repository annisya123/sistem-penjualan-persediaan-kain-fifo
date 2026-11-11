from app import app
from models import db, init_db, User, Product, Sale, IncomingGood
from datetime import datetime, timedelta

with app.app_context():
    db.drop_all()
    init_db()

    # 1. Tambah Data Kain (6 Jenis Kain Populer CB Jaya Textile)
    products_data = [
        {"kode": "KN-001", "name": "Kain Kanvas 14oz", "jenis": "Kanvas", "warna": "Natural White", "satuan": "Roll", "harga_beli": 450000.0, "price": 625000.0, "stock": 48, "min": 10, "desc": "Kain kanvas tenun katun murni berkualitas tinggi dan kokoh."},
        {"kode": "KN-002", "name": "Kain Grey Twill Cotton", "jenis": "Twill", "warna": "Grey", "satuan": "Roll", "harga_beli": 480000.0, "price": 695000.0, "stock": 46, "min": 10, "desc": "Kain grey twill berbahan katun alami untuk celana dan jaket."},
        {"kode": "KN-003", "name": "Kain Marsoto PE", "jenis": "Polyester", "warna": "Putih Tulang", "satuan": "Roll", "harga_beli": 520000.0, "price": 780000.0, "stock": 48, "min": 10, "desc": "Kain marsoto rajut polyester tahan abrasi dan tidak kusut."},
        {"kode": "KN-004", "name": "Kain Kanvas Dabling", "jenis": "Kanvas", "warna": "Broken White", "satuan": "Roll", "harga_beli": 460000.0, "price": 890000.0, "stock": 49, "min": 10, "desc": "Kain kanvas dabling tenun ganda ekstra tebal untuk tas dan craft."},
        {"kode": "KN-005", "name": "Kain Marsoto TC Non Finis", "jenis": "TC", "warna": "Greige Mentah", "satuan": "Roll", "harga_beli": 470000.0, "price": 690000.0, "stock": 45, "min": 10, "desc": "Kain marsoto Tetoron Cotton mentah siap celup warna."},
        {"kode": "KN-006", "name": "Kain Marsoto TC Hitam", "jenis": "TC", "warna": "Hitam Pekat", "satuan": "Roll", "harga_beli": 500000.0, "price": 720000.0, "stock": 47, "min": 10, "desc": "Kain marsoto hitam reaktif tahan luntur."},
    ]

    prod_map = {}
    for p in products_data:
        prod = Product(
            kode_produk=p["kode"],
            name=p["name"],
            jenis_kain=p["jenis"],
            warna=p["warna"],
            satuan=p["satuan"],
            harga_beli=p["harga_beli"],
            price=p["price"],
            actual_stock=p["stock"],
            stok_minimum=p["min"],
            description=p["desc"]
        )
        db.session.add(prod)
        prod_map[p["name"]] = prod

    db.session.commit()

    # 2. Tambah Contoh Riwayat Barang Masuk
    now = datetime.now()
    incoming_data = [
        {"prod": "Kain Kanvas 14oz", "qty": 50, "supplier": "PT. Kain Textile Bandung", "days_ago": 5, "nomor": "BM-20250620-001"},
        {"prod": "Kain Grey Twill Cotton", "qty": 50, "supplier": "PT. Kain Textile Bandung", "days_ago": 5, "nomor": "BM-20250620-002"},
        {"prod": "Kain Marsoto PE", "qty": 50, "supplier": "UD. Sumber Kain", "days_ago": 4, "nomor": "BM-20250621-001"},
        {"prod": "Kain Kanvas Dabling", "qty": 50, "supplier": "PT. Mitra Tekstil", "days_ago": 4, "nomor": "BM-20250621-002"},
        {"prod": "Kain Marsoto TC Non Finis", "qty": 50, "supplier": "UD. Sumber Kain", "days_ago": 3, "nomor": "BM-20250622-001"},
        {"prod": "Kain Marsoto TC Hitam", "qty": 50, "supplier": "UD. Sumber Kain", "days_ago": 3, "nomor": "BM-20250622-002"},
    ]

    for inc in incoming_data:
        p = prod_map[inc["prod"]]
        entry = IncomingGood(
            nomor_masuk=inc["nomor"],
            product_id=p.id,
            quantity=inc["qty"],
            supplier=inc["supplier"],
            notes="Penerimaan stok dari supplier resmi.",
            timestamp=now - timedelta(days=inc["days_ago"])
        )
        db.session.add(entry)

    db.session.commit()

    # 3. Tambah Contoh Transaksi Penjualan
    sales_data = [
        {"prod": "Kain Marsoto TC Non Finis", "cust": "Toko Busana Indah", "qty": 5, "days_ago": 2, "nomor": "PJ-20250623-001"},
        {"prod": "Kain Kanvas 14oz", "cust": "CV. Konveksi Berkah", "qty": 2, "days_ago": 2, "nomor": "PJ-20250623-002"},
        {"prod": "Kain Grey Twill Cotton", "cust": "UD. Celana Maju", "qty": 4, "days_ago": 1, "nomor": "PJ-20250624-001"},
        {"prod": "Kain Marsoto PE", "cust": "Toko Makmur", "qty": 2, "days_ago": 1, "nomor": "PJ-20250624-002"},
        {"prod": "Kain Marsoto TC Hitam", "cust": "Pelanggan Umum", "qty": 3, "days_ago": 0, "nomor": "PJ-20250625-001"},
        {"prod": "Kain Kanvas Dabling", "cust": "CV. Tas Bandung", "qty": 1, "days_ago": 0, "nomor": "PJ-20250625-002"},
    ]

    for s in sales_data:
        p = prod_map[s["prod"]]
        total = s["qty"] * p.price
        sale = Sale(
            nomor_transaksi=s["nomor"],
            product_id=p.id,
            nama_pelanggan=s["cust"],
            quantity=s["qty"],
            harga_satuan=p.price,
            total_price=total,
            kasir_nama="Admin",
            timestamp=now - timedelta(days=s["days_ago"], hours=2)
        )
        db.session.add(sale)

    db.session.commit()
    print("Database CB Jaya Textile berhasil diisi dengan data sederhana!")
