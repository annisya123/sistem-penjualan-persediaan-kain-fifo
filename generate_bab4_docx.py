import os
import zipfile

def create_docx(filename):
    # XML templates for a standard DOCX file
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
        p("BAB IV", bold=True, size=28, align="center", space_before=240, space_after=60)
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

    # ================== DOKUMEN BAB IV ==================
    chapter_title("HASIL DAN PEMBAHASAN")

    # 4.1 Deskripsi Umum Sistem
    heading1("4.1 Deskripsi Umum Sistem")
    p("Sistem Informasi Penjualan dan Persediaan Kain pada CB Jaya Textile merupakan aplikasi berbasis web yang dibangun untuk mengatasi kendala pencatatan manual, keterlambatan laporan, dan risiko penurunan mutu kain akibat tertimbunnya pasokan lama. Sistem ini mengintegrasikan modul Point of Sale (POS) kasir, pencatatan penerimaan pasokan barang masuk dari pabrik, monitoring batas stok minimum (reorder point), serta penerapan algoritma First In First Out (FIFO) secara otomatis.")

    # 4.2 Perancangan
    heading1("4.2 Perancangan")
    p("Tahap perancangan merupakan fase transformasi dari hasil analisis kebutuhan fungsional dan operasional toko ke dalam cetak biru (blueprint) teknis sebelum dilakukan implementasi perangkat lunak. Perancangan sistem dibagi ke dalam empat tahapan terstruktur: Unified Modeling Language (UML), Entity Relationship Diagram (ERD), Struktur Tabel Basis Data, dan Perancangan Desain Antarmuka (Wireframe).")

    # 4.2.1 Unified Modeling Language (UML)
    heading2("4.2.1 Unified Modeling Language (UML)")
    p("Unified Modeling Language (UML) digunakan untuk memodelkan proses bisnis, interaksi pengguna, pertukaran pesan dinamis, dan struktur kelas perangkat lunak berbasis Model-View-Controller (MVC).")

    heading3("1. Use Case Diagram")
    p("Use Case Diagram menggambarkan interaksi fungsional antara sistem dengan 2 (dua) aktor utama, yaitu Administrator (Owner/Gudang) dan Kasir (Staf Penjualan Toko).")
    p("[Gambar 4.14 Use Case Diagram Sistem Informasi CB Jaya Textile]", bold=True, align="center")
    p("Sumber: Diolah oleh Penulis (2026)", italic=True, align="center", space_after=140)

    p("Tabel 4.3 Deskripsi Use Case Sistem Informasi CB Jaya Textile", bold=True, size=22, space_after=60)
    table(
        ["Kode", "Nama Use Case", "Aktor", "Deskripsi Fungsionalitas"],
        [
            ["UC-01", "Login & Logout", "Admin, Kasir", "Memvalidasi kredensial nama pengguna dan kata sandi untuk mengamankan sesi kerja."],
            ["UC-02", "Kelola Data Produk Kain", "Admin", "Menambah jenis kain baru, memperbarui harga jual/modal, mengatur batas stok minimum."],
            ["UC-03", "Pencarian & Cek Stok", "Admin, Kasir", "Mencari kain secara cepat berdasarkan kode SKU atau nama, serta cek sisa stok aktual."],
            ["UC-04", "Pencatatan Barang Masuk", "Admin", "Merekam pasokan kain baru, nomor surat jalan (BM-XXXX), serta menyusun antrean batch FIFO."],
            ["UC-05", "Transaksi Penjualan Kasir", "Admin, Kasir", "Memproses penjualan kain, input meter pemotongan, dan hitung total tagihan otomatis."],
            ["UC-06", "Eksekusi Algoritma FIFO", "Sistem (Include)", "Memotong stok kain secara otomatis dari batch kedatangan tertua yang masih bersisa."],
            ["UC-07", "Cetak Nota Penjualan", "Admin, Kasir (Extend)", "Menerbitkan dan mencetak lembar bukti struk/faktur transaksi penjualan resmi."],
            ["UC-08", "Monitoring Audit FIFO", "Admin", "Menelusuri mutasi pengeluaran per batch untuk memastikan rotasi kain tertua berjalan tertib."],
            ["UC-09", "Kelola Mitra Bisnis", "Admin", "Mengelola direktori data pabrik/supplier kain dan data pelanggan/konveksi."],
            ["UC-10", "Rekapitulasi Laporan", "Admin", "Menyajikan laporan rekap penjualan, mutasi barang masuk, valuasi aset, dan ekspor CSV."]
        ],
        [800, 2200, 1800, 4560]
    )

    heading3("2. Activity Diagram")
    p("Activity Diagram memodelkan aliran aktivitas operasional bisnis secara prosedural. Pada alur transaksi penjualan kasir, kasir memilih kain dan menginput kuantitas meter. Sistem memverifikasi kecukupan stok fisik. Jika cukup, sistem mengeksekusi algoritma FIFO dengan memotong sisa stok dari batch kedatangan tertua (timestamp ASC), memutakhirkan saldo total master produk, mencatat penjualan, dan menerbitkan nota faktur.")
    p("[Gambar 4.15 Activity Diagram Transaksi Penjualan Kasir Berbasis FIFO]", bold=True, align="center")
    p("Sumber: Diolah oleh Penulis (2026)", italic=True, align="center", space_after=140)

    heading3("3. Sequence Diagram")
    p("Sequence Diagram memodelkan interaksi kronologis antara Aktor (Kasir), antarmuka formulir (Boundary: transaction.html), pengontrol server (Controller: app.py), dan basis data (Model: Product, IncomingGood, Sale) saat pemotongan stok FIFO dan penyimpanan transaksi dijalankan.")
    p("[Gambar 4.16 Sequence Diagram Transaksi Penjualan Berbasis FIFO]", bold=True, align="center")
    p("Sumber: Diolah oleh Penulis (2026)", italic=True, align="center", space_after=140)

    heading3("4. Class Diagram")
    p("Class Diagram memodelkan struktur kelas entitas Object-Relational Mapping (ORM SQLAlchemy) yang meliputi kelas User, Product, IncomingGood, dan Sale lengkap dengan atribut dan operasi metodenya.")
    p("[Gambar 4.17 Class Diagram Sistem Informasi CB Jaya Textile]", bold=True, align="center")
    p("Sumber: Diolah oleh Penulis (2026)", italic=True, align="center", space_after=140)

    # 4.2.2 Entity Relationship Diagram (ERD)
    heading2("4.2.2 Entity Relationship Diagram (ERD)")
    p("Entity Relationship Diagram (ERD) digunakan untuk menggambarkan struktur basis data dan hubungan antarentitas yang digunakan pada sistem informasi CB Jaya Textile. Basis data terdiri dari empat tabel utama: users, products, incoming_goods, dan sales.")
    p("[Gambar 4.18 Entity Relationship Diagram (ERD)]", bold=True, align="center")
    p("Sumber: Diolah oleh Penulis (2026)", italic=True, align="center", space_after=140)
    p("Tabel users menyimpan data otentikasi login pengguna. Tabel products menyimpan katalog master kain beserta saldo stok gudang. Tabel incoming_goods mencatat pasokan kain dari supplier berdasarkan nomor surat jalan dan tanggal masuk sebagai dasar urutan batch FIFO. Tabel sales mencatat transaksi penjualan kasir yang terhubung dengan products melalui kunci tamu product_id. Setiap transaksi langsung memutakhirkan saldo stok dan dialokasikan pengurangannya dari batch tertua.")

    # 4.2.3 Struktur Tabel
    heading2("4.2.3 Struktur Tabel")
    p("Struktur tabel menjelaskan rincian atribut kolom, tipe data, dan aturan batasan yang digunakan dalam basis data sistem:")

    p("a. Tabel users", bold=True)
    p("Tabel users digunakan untuk menyimpan data pengguna dan kredensial hak akses login.")
    p("Tabel 4.4 Struktur Tabel Users", bold=True, size=22, space_after=60)
    table(
        ["Nama Kolom", "Tipe Data", "Keterangan"],
        [
            ["id", "INTEGER", "Primary Key (PK), Auto Increment, identitas akun"],
            ["username", "VARCHAR(80)", "Not Null, Unique, nama pengguna untuk login"],
            ["password_hash", "VARCHAR(128)", "Not Null, hash enkripsi keamanan kata sandi"],
            ["role", "VARCHAR(20)", "Not Null, peran hak akses pengguna (Admin / Kasir)"]
        ],
        [2200, 2200, 4960]
    )

    p("b. Tabel products", bold=True)
    p("Tabel products digunakan untuk menyimpan katalog master persediaan kain tekstil.")
    p("Tabel 4.5 Struktur Tabel Products", bold=True, size=22, space_after=60)
    table(
        ["Nama Kolom", "Tipe Data", "Keterangan"],
        [
            ["id", "INTEGER", "Primary Key (PK), Auto Increment, identitas produk"],
            ["kode_produk", "VARCHAR(50)", "Unique, kode registrasi kain (SKU / Barcode)"],
            ["name", "VARCHAR(120)", "Not Null, nama jenis kain tekstil"],
            ["jenis_kain", "VARCHAR(50)", "Kategori serat kain (Kanvas, Twill, TC, PE)"],
            ["warna", "VARCHAR(50)", "Varian warna kain"],
            ["satuan", "VARCHAR(20)", "Not Null, satuan hitung fisik (Roll / Meter)"],
            ["harga_beli", "FLOAT", "Not Null, harga beli modal dari supplier"],
            ["price", "FLOAT", "Not Null, harga jual patokan transaksi kasir"],
            ["actual_stock", "FLOAT", "Not Null, saldo sisa kuantitas stok di gudang"],
            ["stok_minimum", "FLOAT", "Not Null, batas minimum peringatan restock"],
            ["description", "TEXT", "Deskripsi spesifikasi serat dan gramasi kain"]
        ],
        [2200, 2200, 4960]
    )

    p("c. Tabel incoming_goods", bold=True)
    p("Tabel incoming_goods digunakan untuk mencatat riwayat penerimaan pasokan kain dari supplier per batch kedatangan (FIFO).")
    p("Tabel 4.6 Struktur Tabel Incoming Goods", bold=True, size=22, space_after=60)
    table(
        ["Nama Kolom", "Tipe Data", "Keterangan"],
        [
            ["id", "INTEGER", "Primary Key (PK), Auto Increment, identitas barang masuk"],
            ["nomor_masuk", "VARCHAR(30)", "Unique, nomor surat jalan (BM-YYYYMMDD-XXX)"],
            ["product_id", "INTEGER", "Foreign Key (FK), relasi ke tabel products(id)"],
            ["quantity", "FLOAT", "Not Null, kuantitas pasokan kain yang diterima"],
            ["supplier", "VARCHAR(120)", "Nama pabrik atau distributor pengirim"],
            ["notes", "TEXT", "Catatan kondisi fisik kain saat penerimaan"],
            ["timestamp", "DATETIME", "Not Null, waktu kedatangan sebagai acuan urutan FIFO"]
        ],
        [2200, 2200, 4960]
    )

    p("d. Tabel sales", bold=True)
    p("Tabel sales digunakan untuk mencatat transaksi penjualan kain di kasir yang memotong persediaan.")
    p("Tabel 4.7 Struktur Tabel Sales", bold=True, size=22, space_after=60)
    table(
        ["Nama Kolom", "Tipe Data", "Keterangan"],
        [
            ["id", "INTEGER", "Primary Key (PK), Auto Increment, identitas transaksi"],
            ["nomor_transaksi", "VARCHAR(30)", "Unique, nomor nota bukti penjualan (PJ-YYYYMMDD-XXX)"],
            ["product_id", "INTEGER", "Foreign Key (FK), relasi ke tabel products(id)"],
            ["nama_pelanggan", "VARCHAR(120)", "Nama konsumen / UMKM konveksi pemesan"],
            ["quantity", "FLOAT", "Not Null, kuantitas kain yang terjual/dipotong"],
            ["harga_satuan", "FLOAT", "Not Null, harga satuan kain saat transaksi"],
            ["total_price", "FLOAT", "Not Null, total nominal tagihan pembayaran"],
            ["kasir_nama", "VARCHAR(80)", "Nama staf kasir yang memproses transaksi"],
            ["timestamp", "DATETIME", "Not Null, waktu pencatatan transaksi penjualan"]
        ],
        [2200, 2200, 4960]
    )

    # 4.2.4 Desain
    heading2("4.2.4 Desain")
    p("Tahap Desain merupakan proses perancangan tampilan antarmuka visual (wireframe) sebagai representasi tata letak dan komponen sebelum tahap implementasi dilakukan:")
    p("1. Wireframe Halaman Login: Tata letak kartu formulir terpusat di tengah layar dengan input username, password, dan tombol masuk.")
    p("2. Wireframe Halaman Dashboard: Beranda ringkasan informasi yang menyajikan banner alert stok kritis, empat kartu indikator metrik (Stok, Omset, Valuasi, Restock), tombol aksi cepat, serta tabel transaksi terbaru.")
    p("3. Wireframe Halaman Katalog Kain: Menampilkan tabel daftar master kain tekstil, bilah pencarian instan, filter jenis kain, status persediaan berwarna, dan tombol CRUD.")
    p("4. Wireframe Halaman Transaksi Kasir (POS): Formulir pemotongan kain dua panel dengan indikator stok sisa langsung, kalkulasi otomatis subtotal, dan validasi kuantitas.")
    p("5. Wireframe Halaman Bukti Struk Penjualan: Tata letak nota penjualan formal berkop toko CB Jaya Textile yang siap dicetak ke printer kasir 58mm/80mm.")
    p("6. Wireframe Halaman Penerimaan Barang Masuk: Formulir pencatatan surat jalan supplier dan penyusunan batch FIFO baru.")
    p("7. Wireframe Halaman Laporan & Rekapitulasi: Panel filter tanggal, tab multi-laporan (Penjualan, Stok, Masuk), tombol unduh CSV, serta cetak dokumen berkop surat resmi.")
    p("8. Wireframe Halaman Kalkulator Kain: Modul interaktif konversi model busana (kemeja, celana, gamis) terhadap estimasi panjang kain dan anggaran biaya.")

    # 4.3 Implementasi Sistem
    heading1("4.3 Implementasi Sistem")
    p("Implementasi sistem merupakan tahap penerjemahan rancangan yang telah dimodelkan ke dalam bentuk perangkat lunak fungsional yang berjalan di komputer operasional CB Jaya Textile.")

    heading2("4.3.1 Lingkungan Implementasi Perangkat Keras dan Perangkat Lunak")
    p("Sistem diimplementasikan pada komputer server lokal dengan spesifikasi perangkat keras Intel Core i5, RAM 8 GB DDR4, SSD NVMe 512 GB, serta thermal printer nota 58mm. Di sisi perangkat lunak, sistem berjalan di atas sistem operasi Windows 11, bahasa pemrograman Python 3.13.0, microframework Flask 3.0.3, database SQLite 3 terintegrasi Flask-SQLAlchemy 3.1.1, serta antarmuka HTML5, CSS3 kustom, dan Vanilla JavaScript.")

    heading2("4.3.2 Implementasi Basis Data")
    p("Struktur tabel users, products, incoming_goods, dan sales dibuat secara otomatis menggunakan skrip DDL SQLAlchemy (db.create_all()) pada berkas models.py. Akun administrator default otomatis diinisialisasi dengan kata sandi ter-hash.")

    heading2("4.3.3 Implementasi Logika Algoritma FIFO")
    p("Penerapan algoritma FIFO diimplementasikan pada fungsi pemotongan persediaan saat kasir memproses transaksi penjualan:")
    code_box("""# Implementasi Pemotongan Antrean Stok Berbasis FIFO
def allocate_fifo_stock(product_id, qty_sold):
    # 1. Ambil antrean batch masuk tertua yang masih memiliki saldo stok
    batches = IncomingGood.query.filter(
        IncomingGood.product_id == product_id,
        IncomingGood.quantity > 0
    ).order_by(IncomingGood.timestamp.asc()).all()

    remaining_to_deduct = qty_sold
    for batch in batches:
        if remaining_to_deduct <= 0:
            break
        if batch.quantity >= remaining_to_deduct:
            batch.quantity -= remaining_to_deduct
            remaining_to_deduct = 0
        else:
            remaining_to_deduct -= batch.quantity
            batch.quantity = 0

    # 2. Sinkronisasi saldo stok total produk
    product = db.session.get(Product, product_id)
    product.actual_stock -= qty_sold
    db.session.commit()""")

    heading2("4.3.4 Implementasi Antarmuka Sistem")
    p("Seluruh rancangan antarmuka telah berhasil diimplementasikan ke dalam kode HTML dan CSS nyata pada aplikasi web CB Jaya Textile:")
    p("1. Implementasi Halaman Login (login.html): Mengamankan akses sistem melalui otentikasi session terenkripsi.")
    p("2. Implementasi Halaman Dashboard (index.html): Menampilkan peringatan visual stok kritis, kartu metrik real-time, dan pemantauan perputaran stok.")
    p("3. Implementasi Halaman Katalog Kain (products.html): Menyajikan katalog master kain lengkap dengan pencarian cepat dan status stok dinamis.")
    p("4. Implementasi Halaman Transaksi Kasir (transaction.html): Menyediakan antarmuka kasir cepat dengan validasi pencegahan stok minus.")
    p("5. Implementasi Lembar Bukti Struk Penjualan (sale_detail.html): Menyediakan struk nota resmi siap cetak ke printer kasir.")
    p("6. Implementasi Halaman Penerimaan Barang Masuk (stock_incoming.html): Mencatat pasokan supplier dan nomor batch BM-YYYYMMDD-XXX.")
    p("7. Implementasi Halaman Laporan & Rekapitulasi (laporan.html): Menyajikan 4 kartu ringkasan KPI, ekspor CSV, serta cetak dokumen berkop surat resmi dan lembar pengesahan tanda tangan.")
    p("8. Implementasi Halaman Kalkulator Kebutuhan Kain (kalkulator_kain.html): Membantu kasir dan pembeli mengestimasi panjang meter kain sesuai pola pakaian.")

    # 4.4 Pengujian Sistem (Testing)
    heading1("4.4 Pengujian Sistem (Testing)")
    p("Pengujian sistem dilakukan guna memastikan setiap fungsionalitas dan logika bisnis bekerja bebas dari kesalahan (bug/error). Metode pengujian yang digunakan adalah Black Box Testing yang menguji kesesuaian input dan output antarmuka pengguna tanpa melihat struktur kode internal.")

    heading2("4.4.1 Hasil Pengujian Black Box Testing")
    p("Tabel 4.8 Hasil Pengujian Black Box Testing Sistem CB Jaya Textile", bold=True, size=22, space_after=60)
    table(
        ["No", "Kasus Uji", "Skenario Uji & Masukan", "Hasil yang Diharapkan", "Hasil Pengujian", "Status"],
        [
            ["1", "Login Sistem", "Input username & password benar ('admin'/'admin123')", "Sistem memvalidasi kredensial dan mengarahkan ke dashboard utama", "Masuk ke Dashboard", "Valid"],
            ["2", "Login Validasi", "Input password salah", "Muncul notifikasi error 'Username atau password salah!'", "Muncul alert peringatan", "Valid"],
            ["3", "Tambah Kain", "Input form produk lengkap (Kode, Nama, Harga, Stok)", "Data kain tersimpan dan tampil pada daftar master produk", "Data tersimpan di database", "Valid"],
            ["4", "Alert Stok Kritis", "Stok diset <= stok_minimum (<= 10 Roll)", "Sistem menampilkan status 'Menipis' dan banner alert pada dashboard", "Badge dan alert aktif", "Valid"],
            ["5", "Transaksi Kasir", "Input penjualan dengan kuantitas <= stok fisik", "Transaksi sukses, nomor PJ-XXXX terbit, stok berkurang otomatis", "Stok berkurang & nota terbit", "Valid"],
            ["6", "Validasi Stok Kasir", "Input penjualan kuantitas > stok fisik yang ada", "Sistem menolak transaksi dan menampilkan pesan 'Stok tidak cukup'", "Transaksi dicegah", "Valid"],
            ["7", "Alokasi FIFO", "Transaksi pada kain yang memiliki 2 batch masuk berbeda", "Sistem memprioritaskan pengurangan dari batch tertua", "Batch terlama terpotong dahulu", "Valid"],
            ["8", "Barang Masuk", "Input penerimaan pasokan dan kuantitas baru", "Nomor BM-XXXX tercatat dan stok bertambah sejumlah roll yang masuk", "Stok bertambah & batch tercatat", "Valid"],
            ["9", "Pencarian Produk", "Input kata kunci 'Kanvas' pada kotak pencarian", "Sistem menampilkan seluruh kain dengan nama/jenis 'Kanvas'", "Hasil filter akurat", "Valid"],
            ["10", "Filter Laporan", "Pilih rentang tanggal 01/06/2025 s.d 30/06/2025", "Tabel laporan hanya memuat transaksi pada rentang tanggal tersebut", "Filter tanggal presisi", "Valid"],
            ["11", "Ekspor CSV", "Klik tombol 'Unduh CSV Penjualan'", "Browser mengunduh file CSV yang dapat dibuka di Microsoft Excel", "File CSV terunduh valid", "Valid"],
            ["12", "Kalkulator Kain", "Input pola kemeja pria dan jumlah 5 potong pakaian", "Sistem mengalkulasi estimasi kebutuhan meter kain dan biaya", "Kalkulasi akurat", "Valid"]
        ],
        [500, 1500, 2400, 2500, 1500, 960]
    )
    p("Seluruh 12 (dua belas) skenario pengujian memperoleh status Valid, membuktikan bahwa perangkat lunak telah siap dioperasikan.")

    # 4.5 Pembahasan Hasil Penelitian
    heading1("4.5 Pembahasan Hasil Penelitian")
    p("Pembahasan hasil penelitian mengkaji dampak dan perbandingan operasional toko sebelum dan sesudah penerapan sistem informasi berbasis FIFO di CB Jaya Textile:")

    heading2("4.5.1 Efektivitas Penerapan Algoritma FIFO terhadap Pengendalian Kualitas Kain")
    p("Sebelum adanya sistem, pengambilan gulungan kain dari rak gudang tidak memiliki mekanisme kontrol waktu kedatangan sehingga kain lama kerap tertimbun kain baru. Akibatnya timbul dead stock dan penurunan mutu fisik (kusam atau rapuh). Dengan diterapkannya algoritma FIFO secara otomatis, sistem memastikan kain yang pertama masuk adalah kain yang pertama kali dialokasikan untuk dipotong saat transaksi kasir. Hal ini menjamin sirkulasi barang berjalan sehat dan menekan kerugian akibat kain usang.")

    heading2("4.5.2 Peningkatan Akurasi Pencatatan Stok dan Pencegahan Selisih Fisik")
    p("Pencatatan manual pada buku stok sering kali mengalami keterlambatan pembaharuan saldo (delay recording) dan kekeliruan hitung. Sistem terkomputerisasi ini menerapkan transaksi atomik basis data, di mana setiap penambahan melalui barang masuk dan pemotongan melalui kasir langsung memperbarui saldo stok aktual secara real-time, menghilangkan selisih stok (discrepancy) saat stock opname.")

    heading2("4.5.3 Efisiensi Waktu Pembuatan Laporan dan Pengambilan Keputusan")
    p("Proses rekapitulasi nota fisik yang sebelumnya memakan waktu 2 hingga 4 hari setiap akhir bulan kini dapat diselesaikan secara instan dalam hitungan detik. Pemilik toko dapat melihat omset kotor, mutasi persediaan, serta estimasi valuasi aset secara real-time, sekaligus memanfaatkan peringatan stok minimum untuk melakukan pemesanan kembali ke supplier secara tepat waktu tanpa risiko kehabisan barang (stockout).")

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

    print(f"File '{filename}' berhasil diperbarui!")

if __name__ == '__main__':
    output_path = os.path.join(os.getcwd(), 'Laporan_Skripsi_CB_Jaya_Textile_Bab4.docx')
    create_docx(output_path)
    output_path2 = os.path.join(os.getcwd(), 'BAB_IV_Hasil_dan_Pembahasan_CB_Jaya_Textile.docx')
    create_docx(output_path2)
