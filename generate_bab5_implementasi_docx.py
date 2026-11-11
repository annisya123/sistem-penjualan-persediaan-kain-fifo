import os
import zipfile

def create_docx(filename):
    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
        <w:lang w:val="id-ID"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:line="360" w:lineRule="auto" w:after="120"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
</w:styles>"""

    body_elements = []

    def p(text, bold=False, italic=False, size=24, align="both", space_after=120, space_before=0, indent=0, numbering=False):
        b_tag = "<w:b/>" if bold else ""
        i_tag = "<w:i/>" if italic else ""
        jc_tag = f'<w:jc w:val="{align}"/>' if align else ""
        sp_tag = f'<w:spacing w:line="360" w:lineRule="auto" w:before="{space_before}" w:after="{space_after}"/>'
        
        ind_tag = ""
        if numbering and indent > 0:
            ind_tag = f'<w:ind w:left="{indent}" w:hanging="360"/>'
        elif not numbering and indent > 0:
            ind_tag = f'<w:ind w:left="{indent}" w:firstLine="360"/>'
            
        escaped_text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
        
        xml = f"""<w:p>
  <w:pPr>
    {jc_tag}
    {sp_tag}
    {ind_tag}
  </w:pPr>
  <w:r>
    <w:rPr>
      {b_tag}
      {i_tag}
      <w:sz w:val="{size}"/>
      <w:szCs w:val="{size}"/>
    </w:rPr>
    <w:t xml:space="preserve">{escaped_text}</w:t>
  </w:r>
</w:p>"""
        body_elements.append(xml)

    def chapter_title(title_text):
        p("BAB V", bold=True, size=28, align="center", space_before=240, space_after=60)
        p(title_text, bold=True, size=28, align="center", space_before=0, space_after=240)

    def heading1(text):
        p(text, bold=True, size=26, align="left", space_before=220, space_after=100)

    def heading2(text):
        p(text, bold=True, size=24, align="left", space_before=160, space_after=80)

    def heading3(text):
        p(text, bold=True, size=24, align="left", space_before=120, space_after=60)

    def code_box(code_text):
        lines = code_text.strip().split("\n")
        rows_xml = []
        for line in lines:
            escaped_line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
            row = f"""<w:tr>
              <w:tc>
                <w:tcPr>
                  <w:shd w:val="clear" w:color="auto" w:fill="F4F6F8"/>
                  <w:tcMar>
                    <w:top w:w="120" w:type="dxa"/>
                    <w:left w:w="180" w:type="dxa"/>
                    <w:bottom w:w="120" w:type="dxa"/>
                    <w:right w:w="180" w:type="dxa"/>
                  </w:tcMar>
                </w:tcPr>
                <w:p>
                  <w:pPr>
                    <w:spacing w:line="240" w:lineRule="auto" w:after="0"/>
                  </w:pPr>
                  <w:r>
                    <w:rPr>
                      <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>
                      <w:sz w:val="20"/>
                      <w:szCs w:val="20"/>
                    </w:rPr>
                    <w:t xml:space="preserve">{escaped_line}</w:t>
                  </w:r>
                </w:p>
              </w:tc>
            </w:tr>"""
            rows_xml.append(row)

        table_xml = f"""<w:tbl>
          <w:tblPr>
            <w:tblW w:w="9360" w:type="dxa"/>
            <w:tblBorders>
              <w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
              <w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
              <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
              <w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
            </w:tblBorders>
          </w:tblPr>
          {"".join(rows_xml)}
        </w:tbl>"""
        body_elements.append(table_xml)
        p("", space_after=120)

    def table(headers, rows, col_widths=None):
        num_cols = len(headers)
        if not col_widths:
            w = int(9360 / num_cols)
            col_widths = [w] * num_cols

        table_rows = []
        h_cells = []
        for i, h in enumerate(headers):
            escaped_h = h.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
            cell = f"""<w:tc>
              <w:tcPr>
                <w:tcW w:w="{col_widths[i]}" w:type="dxa"/>
                <w:shd w:val="clear" w:color="auto" w:fill="E9ECEF"/>
                <w:tcMar>
                  <w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/>
                  <w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/>
                </w:tcMar>
              </w:tcPr>
              <w:p>
                <w:pPr><w:jc w:val="center"/><w:spacing w:line="280" w:lineRule="auto" w:after="0"/></w:pPr>
                <w:r><w:rPr><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>{escaped_h}</w:t></w:r>
              </w:p>
            </w:tc>"""
            h_cells.append(cell)
        table_rows.append(f"<w:tr>{''.join(h_cells)}</w:tr>")

        for row in rows:
            r_cells = []
            for i, cell_val in enumerate(row):
                escaped_val = str(cell_val).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
                align = "center" if i == 0 or (num_cols > 3 and i in [0, 1] and len(str(cell_val)) < 8) else "left"
                cell = f"""<w:tc>
                  <w:tcPr>
                    <w:tcW w:w="{col_widths[i]}" w:type="dxa"/>
                    <w:tcMar>
                      <w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>
                      <w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/>
                    </w:tcMar>
                  </w:tcPr>
                  <w:p>
                    <w:pPr><w:jc w:val="{align}"/><w:spacing w:line="280" w:lineRule="auto" w:after="0"/></w:pPr>
                    <w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>{escaped_val}</w:t></w:r>
                  </w:p>
                </w:tc>"""
                r_cells.append(cell)
            table_rows.append(f"<w:tr>{''.join(r_cells)}</w:tr>")

        tbl_xml = f"""<w:tbl>
          <w:tblPr>
            <w:tblW w:w="9360" w:type="dxa"/>
            <w:tblBorders>
              <w:top w:val="single" w:sz="6" w:space="0" w:color="999999"/>
              <w:left w:val="single" w:sz="6" w:space="0" w:color="999999"/>
              <w:bottom w:val="single" w:sz="6" w:space="0" w:color="999999"/>
              <w:right w:val="single" w:sz="6" w:space="0" w:color="999999"/>
              <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
              <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
            </w:tblBorders>
          </w:tblPr>
          {"".join(table_rows)}
        </w:tbl>"""
        body_elements.append(tbl_xml)
        p("", space_after=140)

    # ================== DOKUMEN BAB V ==================
    chapter_title("IMPLEMENTASI DAN PENGUJIAN")

    # 5.1 Implementasi
    heading1("5.1 Implementasi")
    p("Tahap implementasi merupakan fase penerjemahan hasil perancangan sistem dan basis data yang telah disusun pada bab sebelumnya ke dalam bentuk program aplikasi perangkat lunak yang fungsional dan siap dioperasikan. Pada tahap ini, seluruh spesifikasi fungsional, antarmuka pengguna, pemodelan tabel, serta algoritma First In First Out (FIFO) direalisasikan menggunakan bahasa pemrograman Python dan microframework Flask.")

    # 5.1.1 Batasan Implementasi
    heading2("5.1.1 Batasan Implementasi")
    p("Batasan implementasi sistem informasi penjualan dan persediaan kain di CB Jaya Textile mencakup ruang lingkup dan batasan fungsional berikut:")
    p("1. Sistem diimplementasikan dalam arsitektur aplikasi web responsif yang berjalan pada jaringan lokal (Local Area Network / intranet) toko, serta dapat diakses melalui browser pada komputer kasir, komputer gudang, maupun laptop pemilik toko.")
    p("2. Pengguna sistem dibedakan menjadi dua hak akses (role), yaitu Administrator (memiliki hak penuh atas konfigurasi produk, stok masuk, transaksi kasir, dan laporan) serta Kasir (berfokus pada operasional pemotongan penjualan kain dan pencetakan nota).")
    p("3. Metode persediaan barang yang diterapkan secara terkomputerisasi adalah algoritma First In First Out (FIFO), di mana pemotongan stok pada transaksi kasir secara otomatis mengalokasikan persediaan dari batch penerimaan barang masuk tertua (timestamp ASC).")
    p("4. Format nota bukti transaksi disesuaikan dengan standar printer kasir termal (lebar kertas 58 mm / 80 mm) serta format cetak dokumen rekapitulasi resmi berkop surat.")

    # 5.1.2 Implementasi Perangkat Keras (Hardware)
    heading2("5.1.2 Implementasi Perangkat Keras")
    p("Implementasi sistem membutuhkan dukungan perangkat keras komputer kasir dan server lokal dengan spesifikasi minimum yang tercantum pada Tabel 5.1:")
    
    p("Tabel 5.1 Kebutuhan Perangkat Keras Sistem", bold=True, size=22, space_after=60)
    table(
        ["No", "Nama Perangkat", "Spesifikasi Minimum", "Fungsi / Peran Operasional"],
        [
            ["1", "Processor", "Intel Core i3 / AMD Ryzen 3 (Quad Core 2.0 GHz)", "Menjalankan pemrosesan server Flask dan kueri SQLite"],
            ["2", "RAM (Memori)", "4 GB DDR4 (Disarankan 8 GB)", "Menampung runtime Python dan penanganan request session"],
            ["3", "Penyimpanan (Storage)", "SSD 128 GB (NVMe / SATA)", "Menyimpan sistem operasi, aplikasi, dan database data.db"],
            ["4", "Monitor / Layar", "Resolusi 1366 x 768 piksel (15 - 21 inci)", "Menampilkan antarmuka dashboard dan formulir kasir"],
            ["5", "Printer Kasir", "Thermal Receipt Printer 58mm / 80mm (USB/ESC-POS)", "Mencetak bukti nota fisik transaksi kasir langsung"],
            ["6", "Printer Dokumen", "Printer Inkjet / Laserjet A4", "Mencetak lembar dokumen rekapitulasi laporan berkop resmi"],
            ["7", "Perangkat Input", "Keyboard, Mouse, dan Barcode Scanner (Opsional)", "Memasukkan kuantitas meter kain dan data transaksi"]
        ],
        [500, 2000, 3200, 3660]
    )

    # 5.1.3 Implementasi Perangkat Lunak (Software)
    heading2("5.1.3 Implementasi Perangkat Lunak")
    p("Lingkungan perangkat lunak yang digunakan dalam pembangunan dan pengoperasian sistem informasi CB Jaya Textile dirinci pada Tabel 5.2:")

    p("Tabel 5.2 Kebutuhan Perangkat Lunak Sistem", bold=True, size=22, space_after=60)
    table(
        ["No", "Komponen Perangkat Lunak", "Spesifikasi / Versi", "Keterangan"],
        [
            ["1", "Sistem Operasi", "Microsoft Windows 10 / Windows 11 (64-bit)", "Platform sistem operasi komputer server lokal dan kasir"],
            ["2", "Bahasa Pemrograman", "Python versi 3.13.0", "Bahasa pemrograman utama backend server dan logika bisnis"],
            ["3", "Web Framework", "Flask versi 3.0.3", "Microframework WSGI untuk routing, session, dan controller"],
            ["4", "ORM / Basis Data", "Flask-SQLAlchemy 3.1.1 & SQLite 3", "Object-Relational Mapping dan mesin penyimpanan basis data"],
            ["5", "Keamanan Otentikasi", "Werkzeug 3.0.2 (generate_password_hash)", "Enkripsi satu arah untuk keamanan kata sandi akun pengguna"],
            ["6", "Frontend Engine", "Jinja2, HTML5, CSS3 Kustom, JavaScript ES6", "Rendering antarmuka pengguna responsif dan interaktif"],
            ["7", "Icon & Typography", "Font Awesome 6.5.1 & Plus Jakarta Sans", "Aset visual ikon dan tipografi modern antarmuka"],
            ["8", "Web Browser", "Google Chrome / Microsoft Edge / Mozilla Firefox", "Aplikasi peramban web untuk mengakses sistem (port 5000)"]
        ],
        [500, 2500, 2700, 3660]
    )

    # 5.1.4 Implementasi Basis Data
    heading2("5.1.4 Implementasi Basis Data")
    p("Implementasi basis data dilakukan dengan mendefinisikan model kelas entitas ORM SQLAlchemy pada berkas models.py. Basis data relasional disimpan dalam satu berkas berkas tunggal data.db di dalam direktori instance/. Skrip inisialisasi basis data dan pembuatan seluruh tabel (users, products, incoming_goods, dan sales) dijalankan secara otomatis saat server Flask pertama kali diaktifkan melalui fungsi init_db():")

    code_box("""# Cuplikan Inisialisasi Database models.py
db = SQLAlchemy()

def init_db():
    db.create_all()
    # Inisialisasi akun Administrator bawaan jika belum ada
    if not User.query.filter_by(username="admin").first():
        admin = User(username="admin", role="Admin")
        admin.set_password("admin123")
        db.session.add(admin)
        db.session.commit()""")

    # 5.1.5 Implementasi Logika Algoritma FIFO
    heading2("5.1.5 Implementasi Logika Algoritma FIFO")
    p("Penerapan metode First In First Out (FIFO) diimplementasikan pada alur transaksi kasir saat kuantitas kain dipotong. Sistem secara otomatis menelusuri antrean batch kedatangan barang masuk (tabel incoming_goods) yang diurutkan secara menaik berdasarkan stempel waktu kedatangan (ORDER BY timestamp ASC). Batch tertua yang masih memiliki saldo kuantitas dipotong terlebih dahulu sebelum beralih ke batch berikutnya:")

    code_box("""# Implementasi Alokasi Pemotongan Antrean FIFO (app.py)
def allocate_fifo_stock(product_id, qty_sold):
    # 1. Ambil antrean batch kedatangan tertua yang masih bersisa
    batches = IncomingGood.query.filter(
        IncomingGood.product_id == product_id,
        IncomingGood.quantity > 0
    ).order_by(IncomingGood.timestamp.asc()).all()

    remaining = qty_sold
    for batch in batches:
        if remaining <= 0:
            break
        if batch.quantity >= remaining:
            batch.quantity -= remaining
            remaining = 0
        else:
            remaining -= batch.quantity
            batch.quantity = 0

    # 2. Sinkronkan saldo stok fisik master produk
    product = db.session.get(Product, product_id)
    product.actual_stock -= qty_sold
    db.session.commit()""")

    # 5.1.6 Implementasi Antarmuka Sistem
    heading2("5.1.6 Implementasi Antarmuka Sistem")
    p("Implementasi antarmuka sistem diwujudkan sesuai dengan rancangan wireframe pada Bab 4.2.4. Antarmuka menggunakan bilah navigasi atas (Top Blue Navbar #1E3A8A) yang seragam di seluruh halaman:")

    heading3("1. Implementasi Halaman Login (login.html)")
    p("Halaman login berfungsi sebagai gerbang keamanan sistem. Antarmuka menyajikan kartu login terpusat dengan logo toko, input nama pengguna, input kata sandi, tombol 'Masuk ke Sistem', dan kotak notifikasi galat jika kredensial salah.")
    p("[Gambar 5.1 Antarmuka Halaman Login]", bold=True, align="center")
    p("Sumber: Hasil Implementasi Sistem (2026)", italic=True, align="center", space_after=120)

    heading3("2. Implementasi Halaman Dashboard (index.html)")
    p("Halaman dashboard menampilkan ringkasan informasi eksekutif yang mencakup banner alert stok kritis, 4 kartu KPI metrik (Total Stok, Omset Penjualan, Valuasi Aset, dan Perlu Restock), tombol aksi cepat, serta pembagian dua kolom tabel transaksi terkini dan persediaan kain menipis.")
    p("[Gambar 5.2 Antarmuka Halaman Dashboard]", bold=True, align="center")
    p("Sumber: Hasil Implementasi Sistem (2026)", italic=True, align="center", space_after=120)

    heading3("3. Implementasi Halaman Katalog Kain (products.html)")
    p("Halaman katalog menyajikan master persediaan kain tekstil lengkap dengan bilah pencarian instan, filter dropdown kategori (Kanvas, Twill, dll), tombol penambahan kain, tabel rincian harga beli/jual, sisa stok, status persediaan berwarna (Aman, Menipis, Habis), serta tombol aksi pengubahan dan penghapusan.")
    p("[Gambar 5.3 Antarmuka Halaman Katalog Kain]", bold=True, align="center")
    p("Sumber: Hasil Implementasi Sistem (2026)", italic=True, align="center", space_after=120)

    heading3("4. Implementasi Formulir Tambah dan Ubah Kain (create.html & edit.html)")
    p("Menyediakan formulir penginputan data kain baru atau pengubahan spesifikasi yang mencakup kode SKU, nama kain, jenis serat, warna, harga beli modal, harga jual patokan, saldo stok awal, serta batas stok minimum peringatan restock.")
    p("[Gambar 5.4 Antarmuka Formulir Tambah / Edit Kain]", bold=True, align="center")
    p("Sumber: Hasil Implementasi Sistem (2026)", italic=True, align="center", space_after=120)

    heading3("5. Implementasi Halaman Transaksi Kasir POS (transaction.html)")
    p("Halaman kasir mengadopsi tata letak 2 kolom interaktif: panel kiri memuat formulir penjualan dengan indikator sisa stok fisik otomatis, kuantitas potong, dan nama pelanggan; sedangkan panel kanan memuat ringkasan tagihan otomatis yang menghitung subtotal dan nominal TOTAL pembayaran secara instan.")
    p("[Gambar 5.5 Antarmuka Transaksi Kasir POS]", bold=True, align="center")
    p("Sumber: Hasil Implementasi Sistem (2026)", italic=True, align="center", space_after=120)

    heading3("6. Implementasi Lembar Bukti Struk Penjualan Thermal (transaction_detail.html)")
    p("Menampilkan nota faktur siap cetak berformat printer kasir termal (58mm / 80mm) dengan kop toko CB Jaya Textile, nomor nota resmi (PJ-YYYYMMDD-XXX), stempel waktu, rincian kuantitas meter kain, dan total pembayaran, disertai tombol 'Cetak Struk' dan 'Transaksi Baru'.")
    p("[Gambar 5.6 Antarmuka Bukti Struk Penjualan Kasir]", bold=True, align="center")
    p("Sumber: Hasil Implementasi Sistem (2026)", italic=True, align="center", space_after=120)

    heading3("7. Implementasi Halaman Penerimaan Barang Masuk (stock_incoming.html)")
    p("Menyajikan formulir pencatatan pasokan kain dari pabrik dengan input nomor surat jalan (SJ-...), jenis kain, kuantitas masuk, distributor, dan catatan kondisi kain, serta tabel histori mutasi barang masuk yang menyusun antrean kedatangan FIFO.")
    p("[Gambar 5.7 Antarmuka Penerimaan Barang Masuk]", bold=True, align="center")
    p("Sumber: Hasil Implementasi Sistem (2026)", italic=True, align="center", space_after=120)

    heading3("8. Implementasi Halaman Laporan & Rekapitulasi (laporan.html)")
    p("Menyediakan panel filter rentang tanggal, tab navigasi multi-laporan (Penjualan, Posisi Stok, Barang Masuk), 3 kartu metrik ringkasan KPI, tombol ekspor file CSV untuk Microsoft Excel, serta cetak dokumen berkop surat toko resmi dengan lembar tanda tangan pengesahan pimpinan.")
    p("[Gambar 5.8 Antarmuka Halaman Laporan & Rekapitulasi]", bold=True, align="center")
    p("Sumber: Hasil Implementasi Sistem (2026)", italic=True, align="center", space_after=120)

    heading3("9. Implementasi Halaman Kalkulator Kebutuhan Kain (kalkulator_kain.html)")
    p("Modul pembantu kasir dan pembeli untuk mengalkulasi estimasi panjang kain meter dan perkiraan anggaran biaya berdasarkan model busana (kemeja, celana, gamis, jas), ukuran badan (S s.d XXL), lebar kain pabrik, serta toleransi jahitan 10%.")
    p("[Gambar 5.9 Antarmuka Kalkulator Kebutuhan Kain]", bold=True, align="center")
    p("Sumber: Hasil Implementasi Sistem (2026)", italic=True, align="center", space_after=120)

    # 5.2 Pengujian Sistem
    heading1("5.2 Pengujian Sistem")
    p("Tahap pengujian sistem bertujuan untuk menguji dan memvalidasi seluruh fungsionalitas aplikasi agar bebas dari kesalahan program (bug/error) serta memastikan kesesuaian dengan kebutuhan operasional toko CB Jaya Textile. Metode pengujian yang digunakan adalah Black Box Testing yang berfokus pada pengujian input, proses, dan output antarmuka pengguna tanpa membedah struktur internal kode.")

    heading2("5.2.1 Rencana dan Hasil Pengujian Black Box Testing")
    p("Pengujian dilakukan terhadap 12 skenario kasus uji utama yang mencakup seluruh alur sistem. Hasil pengujian dirangkum pada Tabel 5.3:")

    p("Tabel 5.3 Hasil Pengujian Black Box Testing Sistem CB Jaya Textile", bold=True, size=22, space_after=60)
    table(
        ["No", "Kasus Uji", "Skenario Uji & Masukan", "Hasil yang Diharapkan", "Hasil Pengujian", "Status"],
        [
            ["1", "Login Akun", "Input username & password valid ('admin'/'admin123')", "Sistem memvalidasi kredensial dan mengarahkan ke dashboard", "Berhasil masuk ke Dashboard", "Valid"],
            ["2", "Validasi Login", "Input password salah atau username kosong", "Sistem menampilkan pesan galat 'Username atau password tidak sesuai'", "Alert galat muncul presisi", "Valid"],
            ["3", "Tambah Kain", "Input form produk lengkap (Kode, Nama, Harga, Stok)", "Data kain tersimpan dan muncul di daftar katalog", "Data tersimpan di database", "Valid"],
            ["4", "Ubah Data Kain", "Edit harga jual atau stok minimum kain", "Perubahan tersimpan dan data tabel terbarui", "Data berhasil diubah", "Valid"],
            ["5", "Alert Stok Kritis", "Stok fisik diset berada di bawah stok_minimum", "Sistem menampilkan status 'Menipis' dan banner alert pada dashboard", "Banner dan badge aktif", "Valid"],
            ["6", "Pencarian Kain", "Input kata kunci 'Kanvas' pada bilah pencarian katalog", "Tabel menampilkan seluruh jenis kain yang mengandung kata 'Kanvas'", "Hasil filter presisi", "Valid"],
            ["7", "Transaksi Kasir", "Input penjualan dengan kuantitas <= stok fisik", "Transaksi sukses, nomor PJ-XXXX terbit, stok terpotong otomatis", "Stok berkurang & nota terbit", "Valid"],
            ["8", "Validasi Stok Kurang", "Input penjualan kuantitas > stok fisik gudang", "Sistem menolak transaksi dengan peringatan 'Stok tidak cukup'", "Transaksi dicegah", "Valid"],
            ["9", "Alokasi Batch FIFO", "Transaksi pada produk dengan 2 batch kedatangan berbeda", "Sistem memotong persediaan dari batch kedatangan tertua dahulu", "Batch tertua terpotong dahulu", "Valid"],
            ["10", "Penerimaan Pasokan", "Input barang masuk (nomor SJ, produk, kuantitas)", "Stok bertambah dan histori mutasi kedatangan batch baru tercatat", "Stok bertambah & batch tercatat", "Valid"],
            ["11", "Filter & Ekspor CSV", "Pilih periode tanggal dan klik 'Unduh CSV'", "Sistem mengunduh berkas laporan CSV yang valid dibuka di Excel", "File CSV terunduh sempurna", "Valid"],
            ["12", "Kalkulator Kain", "Input pola kemeja panjang ukuran L, 3 potong busana", "Sistem mengalkulasi estimasi kebutuhan meter kain dan biaya", "Kalkulasi matematis akurat", "Valid"]
        ],
        [400, 1400, 2400, 2560, 1600, 1000]
    )

    heading2("5.2.2 Kesimpulan Hasil Pengujian")
    p("Berdasarkan hasil pengujian pada Tabel 5.3, seluruh 12 (dua belas) skenario pengujian memperoleh status Valid (100% fungsionalitas terpenuhi). Hal ini membuktikan bahwa perangkat lunak Sistem Informasi Penjualan dan Persediaan Kain pada CB Jaya Textile telah berjalan stabil, logika algoritma FIFO bekerja akurat, dan sistem siap diimplementasikan untuk mendukung aktivitas operasional toko sehari-hari.")

    # Packaging to DOCX
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {"".join(body_elements)}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
        docx_zip.writestr('[Content_Types].xml', content_types_xml)
        docx_zip.writestr('_rels/.rels', rels_xml)
        docx_zip.writestr('word/document.xml', document_xml)
        docx_zip.writestr('word/styles.xml', styles_xml)

    print(f"File '{filename}' berhasil dibuat!")

if __name__ == '__main__':
    output_path = os.path.join(os.getcwd(), 'BAB_V_Implementasi_dan_Pengujian_CB_Jaya_Textile.docx')
    create_docx(output_path)
