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

    # Helper functions to construct document.xml
    body_elements = []

    def p(text, bold=False, italic=False, size=24, align="both", space_after=120, space_before=0):
        b_tag = "<w:b/>" if bold else ""
        i_tag = "<w:i/>" if italic else ""
        jc_tag = f'<w:jc w:val="{align}"/>' if align else ""
        sp_tag = f'<w:spacing w:line="360" w:lineRule="auto" w:before="{space_before}" w:after="{space_after}"/>'
        
        # Escape XML entities
        escaped_text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
        
        xml = f"""<w:p>
          <w:pPr>{jc_tag}{sp_tag}</w:pPr>
          <w:r>
            <w:rPr>{b_tag}{i_tag}<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>
            <w:t xml:space="preserve">{escaped_text}</w:t>
          </w:r>
        </w:p>"""
        body_elements.append(xml)

    def heading1(text):
        p(text, bold=True, size=28, align="left", space_before=240, space_after=140)

    def heading2(text):
        p(text, bold=True, size=26, align="left", space_before=200, space_after=100)

    def heading3(text):
        p(text, bold=True, size=24, align="left", space_before=160, space_after=80)

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
        
        # Header Row
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

        # Data Rows
        for row in rows:
            r_cells = []
            for i, cell_val in enumerate(row):
                escaped_val = str(cell_val).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
                align = "center" if i == 0 or (num_cols > 3 and i in [0, 2, 3]) else "left"
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

    # ================== CONTENT GENERATION ==================

    heading1("2.2.8 Sequence Diagram")
    p("Sequence Diagram menggambarkan interaksi dinamis antar objek dalam sistem yang disusun berdasarkan urutan waktu (kronologis). Diagram ini memperlihatkan pertukaran pesan (message) antara Aktor (Kasir / Admin), Antarmuka Web (Browser), Pengendali Sistem (Flask Server), dan Lapisan Penyimpanan Data (Database) pada setiap proses bisnis utama di CB Jaya Tekstil.")

    heading2("2.2.8.1 Sequence Diagram Transaksi Penjualan (Metode FIFO)")
    p("Proses transaksi penjualan memodelkan eksekusi algoritma First In First Out (FIFO) secara otomatis di mana stok kain yang dijual diambil dari batch penerimaan barang masuk tertua yang masih memiliki sisa stok. Rincian tahapan alur interaksi terbagi menjadi empat fase sebagai berikut:")
    
    heading3("1. Fase Membuka Halaman Transaksi")
    p("Kasir mengakses alamat URL /transaction pada browser. Browser mengirimkan permintaan GET /transaction ke Flask Server. Server mengeksekusi kueri ke Database untuk memuat katalog master produk kain serta antrean batch FIFO yang aktif, kemudian me-render dan menampilkan formulir penjualan ke antarmuka kasir.")

    heading3("2. Fase Memilih Produk & Kuantitas")
    p("Kasir memilih item kain dan memasukkan jumlah meter pembelian. Browser memicu permintaan asynchronous (AJAX GET /api/batches/<product_id>) ke server. Server mengambil daftar batch aktif dari tabel incoming_goods yang diurutkan berdasarkan tanggal masuk paling awal (timestamp ASC). Data JSON dikembalikan ke browser untuk memberikan pratinjau urutan batch yang akan dialokasikan.")

    heading3("3. Fase Memproses Transaksi (Eksekusi Algoritma FIFO)")
    p("Kasir menekan tombol 'Proses Transaksi (FIFO)'. Browser mengirimkan data form via POST /transaction. Server memulai transaksi database (BEGIN TRANSACTION) dan mengevaluasi ketersediaan stok:")
    p("• Kondisi Alternatif 1 [Stok Kurang]: Jika total sisa stok tidak mencukupi kuantitas yang diminta, server membatalkan proses (ROLLBACK) dan menampilkan notifikasi kesalahan 'Stok tidak mencukupi'.")
    p("• Kondisi Alternatif 2 [Stok Mencukupi]: Sistem mengeksekusi perulangan alokasi FIFO (Loop Allocate FIFO Batch). Server mengambil batch tertua (LIMIT 1), menghitung kuantitas pengurangan, memperbarui kolom sisa_stok pada incoming_goods, dan mencatat rincian pemotongan ke tabel log audit fifo_transactions. Proses berulang hingga seluruh kuantitas meter terpenuhi. Selanjutnya, data induk penjualan disimpan ke tabel sales, total persediaan produk diperbarui pada tabel products, dan transaksi diselesaikan secara permanen (COMMIT).")

    heading3("4. Fase Menampilkan & Mencetak Nota")
    p("Browser menampilkan pratinjau nota transaksi penjualan beserta nomor batch kain yang dialokasikan. Kasir menekan tombol 'Cetak Nota' untuk memicu perintah pencetakan window.print() guna mencetak struk fisik atau menyimpannya sebagai dokumen digital PDF.")

    heading2("2.2.8.2 Sequence Diagram Transaksi Barang Masuk (Batch FIFO)")
    p("Pencatatan pasokan kain dari pemasok merupakan basis pembentukan antrean batch FIFO baru. Alur proses barang masuk adalah sebagai berikut:")
    p("1. Admin / Petugas Gudang membuka menu Barang Masuk (GET /incoming).")
    p("2. Admin mengisi data pasokan (produk, supplier, kuantitas meter, dan catatan) kemudian menekan tombol 'Simpan Pasokan Masuk'.")
    p("3. Browser mengirimkan data melalui permintaan POST /incoming/create.")
    p("4. Server membentuk nomor batch masuk unik (format: BM-YYYYMMDD-XXX) dan melakukan transaksi database:")
    p("   a. Menyimpan record batch baru ke tabel incoming_goods dengan menginisialisasi nilai sisa_stok = quantity.")
    p("   b. Menambahkan nilai actual_stock pada tabel products sejumlah kuantitas pasokan baru.")
    p("5. Database meng-commit transaksi dan server mengembalikan pesan sukses ke browser.")

    heading2("2.2.8.3 Sequence Diagram Manajemen Master Data Produk")
    p("Proses pengelolaan data master produk kain meliputi:")
    p("1. Admin menginput data kain baru (kode produk, nama kain, jenis kain, warna, satuan, harga beli, harga jual, stok minimum, serta file gambar kain).")
    p("2. Admin menekan tombol 'Simpan'. Browser mengirimkan data form dan berkas gambar secara multipart ke server melalui POST /products/create.")
    p("3. Server memvalidasi integritas data dan menyimpan file gambar ke direktori static/uploads.")
    p("4. Server menyimpan entitas produk baru ke tabel products pada database SQLite.")
    p("5. Server me-redirect tampilan ke daftar produk dengan notifikasi keberhasilan penambahan produk.")

    heading2("2.2.8.4 Sequence Diagram Pembuatan & Unduh Laporan")
    p("Proses rekapitulasi data laporan operasional meliputi:")
    p("1. Manajer / Admin memilih jenis laporan (Laporan Penjualan, Laporan Stok Kain, atau Laporan Audit FIFO) dan menentukan rentang tanggal filter.")
    p("2. Pengguna menekan tombol 'Export CSV' atau 'Cetak'.")
    p("3. Browser mengirimkan permintaan GET /reports/export dengan parameter filter.")
    p("4. Server melakukan kueri gabungan (JOIN) antar tabel sales, products, users, dan fifo_transactions.")
    p("5. Server memformat dataset hasil kueri ke dalam buffer CSV (StringIO) dan mengirimkannya sebagai file attachment yang langsung diunduh otomatis oleh browser pengguna.")

    # ================== CLASS DIAGRAM ==================
    heading1("2.2.8.5 Class Diagram")
    p("Class Diagram menggambarkan struktur kelas data, atribut, tipe data, method/operasi, serta relasi asosiasi dan kardinalitas antar entitas pada sistem CB Jaya Tekstil.")

    heading2("Tabel Spesifikasi Kelas dan Atribut")
    table(
        ["No", "Nama Kelas", "Deskripsi & Peran", "Atribut Utama"],
        [
            ["1", "User", "Mengelola akun pengguna dan otentikasi hak akses.", "id, username, password_hash, role"],
            ["2", "Product", "Menyimpan data master kain, harga, dan stok aktual.", "kode_produk, name, jenis_kain, warna, satuan, harga_beli, price, actual_stock, stok_minimum"],
            ["3", "IncomingGood", "Mencatat batch penerimaan kain dan sisa stok FIFO.", "nomor_masuk, product_id, quantity, sisa_stok, supplier, notes, timestamp"],
            ["4", "Sale", "Menyimpan transaksi penjualan kain ke pelanggan.", "nomor_transaksi, product_id, user_id, nama_pelanggan, quantity, harga_satuan, total_price, timestamp"],
            ["5", "FifoTransaction", "Log audit alokasi pemotongan stok batch FIFO.", "sale_id, product_id, incoming_good_id, jumlah_keluar, tanggal_keluar"],
            ["6", "Supplier", "Mengelola master data pemasok pasokan kain.", "kode_supplier, nama, kontak, telepon, alamat, created_at"],
            ["7", "Customer", "Mengelola master direktori data pelanggan.", "kode_pelanggan, nama, kontak, telepon, alamat, created_at"],
            ["8", "StoreProfile", "Menyimpan profil identitas toko untuk kop faktur.", "nama_toko, tagline, alamat, telepon, email, deskripsi"]
        ],
        [400, 1600, 3600, 3760]
    )

    heading2("Tabel Relasi dan Kardinalitas Antar Kelas")
    table(
        ["Relasi Antar Kelas", "Kardinalitas", "Penjelasan Makna Bisnis"],
        [
            ["User ke Sale", "1 ke 0..*", "Satu kasir/admin dapat melayani banyak transaksi penjualan, namun setiap transaksi penjualan wajib dicatat oleh satu kasir penanggung jawab."],
            ["Product ke IncomingGood", "1 ke 0..*", "Satu produk kain dapat menerima banyak batch pasokan dari waktu ke waktu yang membentuk antrean FIFO."],
            ["Product ke Sale", "1 ke 0..*", "Satu jenis kain dapat terjual pada banyak transaksi penjualan."],
            ["IncomingGood ke FifoTransaction", "1 ke 0..*", "Satu batch masuk dapat dipotong berkali-kali oleh transaksi penjualan hingga sisa stoknya habis."],
            ["Sale ke FifoTransaction", "1 ke 1..*", "Satu transaksi penjualan menghasilkan satu atau lebih pemotongan batch FIFO sesuai kecukupan stok batch tertua."],
            ["Product ke FifoTransaction", "1 ke 0..*", "Menghubungkan log alokasi langsung ke master produk untuk mempermudah audit kartu persediaan."]
        ],
        [2400, 1400, 5560]
    )

    # ================== BASIS DATA ==================
    heading1("2.2.9 Basis Data")
    p("Perancangan basis data memodelkan skema penyimpanan data yang terstruktur, konsisten, dan terintegrasi untuk mendukung proses bisnis penjualan serta pencatatan stok metode FIFO.")

    heading2("2.2.9.1 Logical Record Structure (LRS)")
    p("Relasi tabel pada basis data sistem didefinisikan sebagai berikut:")
    p("1. users (id [PK], username, password_hash, role)")
    p("2. products (id [PK], kode_produk, name, jenis_kain, warna, satuan, harga_beli, price, initial_stock, actual_stock, stok_minimum, image_filename)")
    p("3. incoming_goods (id [PK], nomor_masuk, product_id [FK], quantity, sisa_stok, supplier, notes, timestamp)")
    p("4. sales (id [PK], nomor_transaksi, product_id [FK], user_id [FK], nama_pelanggan, quantity, harga_satuan, total_price, timestamp)")
    p("5. fifo_transactions (id [PK], sale_id [FK], product_id [FK], incoming_good_id [FK], jumlah_keluar, tanggal_keluar)")
    p("6. suppliers (id [PK], kode_supplier, nama, kontak, telepon, alamat, created_at)")
    p("7. customers (id [PK], kode_pelanggan, nama, kontak, telepon, alamat, created_at)")
    p("8. store_profile (id [PK], nama_toko, tagline, alamat, telepon, email, deskripsi)")

    heading2("2.2.9.2 Kamus Data / Spesifikasi Tabel Basis Data")

    p("1. Tabel users", bold=True)
    table(
        ["No", "Nama Kolom", "Tipe Data", "Ukuran", "Keterangan / Constraint"],
        [
            ["1", "id", "Integer", "-", "Primary Key, Auto Increment"],
            ["2", "username", "Varchar", "80", "Unique, Not Null"],
            ["3", "password_hash", "Varchar", "128", "Not Null (Terenkripsi Hash)"],
            ["4", "role", "Varchar", "20", "Not Null, Default: 'Kasir'"]
        ],
        [500, 2200, 1600, 1200, 3860]
    )

    p("2. Tabel products", bold=True)
    table(
        ["No", "Nama Kolom", "Tipe Data", "Ukuran", "Keterangan / Constraint"],
        [
            ["1", "id", "Integer", "-", "Primary Key, Auto Increment"],
            ["2", "kode_produk", "Varchar", "50", "Unique, Nullable"],
            ["3", "name", "Varchar", "120", "Not Null (Nama Kain)"],
            ["4", "jenis_kain", "Varchar", "50", "Nullable (Katun, Rayon, dll)"],
            ["5", "warna", "Varchar", "50", "Nullable"],
            ["6", "satuan", "Varchar", "20", "Not Null, Default: 'Meter'"],
            ["7", "description", "Text", "-", "Nullable"],
            ["8", "harga_beli", "Float", "-", "Not Null, Default: 0"],
            ["9", "price", "Float", "-", "Not Null (Harga Jual)"],
            ["10", "initial_stock", "Integer", "-", "Not Null, Default: 0"],
            ["11", "actual_stock", "Integer", "-", "Not Null (Stok Fisik Tersedia)"],
            ["12", "stok_minimum", "Integer", "-", "Not Null, Default: 10 (Reorder Point)"],
            ["13", "image_filename", "Varchar", "255", "Nullable (File Gambar)"]
        ],
        [500, 2200, 1600, 1200, 3860]
    )

    p("3. Tabel incoming_goods (Batch FIFO)", bold=True)
    table(
        ["No", "Nama Kolom", "Tipe Data", "Ukuran", "Keterangan / Constraint"],
        [
            ["1", "id", "Integer", "-", "Primary Key, Auto Increment"],
            ["2", "nomor_masuk", "Varchar", "30", "Unique, Format: BM-YYYYMMDD-XXX"],
            ["3", "product_id", "Integer", "-", "Foreign Key -> products.id (CASCADE)"],
            ["4", "quantity", "Integer", "-", "Not Null (Jumlah Masuk Awal)"],
            ["5", "sisa_stok", "Integer", "-", "Not Null (Kunci FIFO: Sisa Stok Batch)"],
            ["6", "supplier", "Varchar", "120", "Nullable (Nama Pemasok)"],
            ["7", "notes", "Text", "-", "Nullable"],
            ["8", "timestamp", "DateTime", "-", "Not Null, Default: Current UTC"]
        ],
        [500, 2200, 1600, 1200, 3860]
    )

    p("4. Tabel sales", bold=True)
    table(
        ["No", "Nama Kolom", "Tipe Data", "Ukuran", "Keterangan / Constraint"],
        [
            ["1", "id", "Integer", "-", "Primary Key, Auto Increment"],
            ["2", "nomor_transaksi", "Varchar", "30", "Unique, Format: PJ-YYYYMMDD-XXX"],
            ["3", "product_id", "Integer", "-", "Foreign Key -> products.id (CASCADE)"],
            ["4", "user_id", "Integer", "-", "Foreign Key -> users.id (SET NULL)"],
            ["5", "nama_pelanggan", "Varchar", "120", "Nullable"],
            ["6", "quantity", "Integer", "-", "Not Null (Meter Terjual)"],
            ["7", "harga_satuan", "Float", "-", "Not Null"],
            ["8", "total_price", "Float", "-", "Not Null (Total Nilai Transaksi)"],
            ["9", "timestamp", "DateTime", "-", "Not Null, Default: Current UTC"]
        ],
        [500, 2200, 1600, 1200, 3860]
    )

    p("5. Tabel fifo_transactions (Log Audit FIFO)", bold=True)
    table(
        ["No", "Nama Kolom", "Tipe Data", "Ukuran", "Keterangan / Constraint"],
        [
            ["1", "id", "Integer", "-", "Primary Key, Auto Increment"],
            ["2", "sale_id", "Integer", "-", "Foreign Key -> sales.id (CASCADE)"],
            ["3", "product_id", "Integer", "-", "Foreign Key -> products.id (CASCADE)"],
            ["4", "incoming_good_id", "Integer", "-", "Foreign Key -> incoming_goods.id (CASCADE)"],
            ["5", "jumlah_keluar", "Integer", "-", "Not Null (Kuantitas Dipotong dari Batch)"],
            ["6", "tanggal_keluar", "DateTime", "-", "Not Null, Default: Current UTC"]
        ],
        [500, 2200, 1600, 1200, 3860]
    )

    p("6. Tabel suppliers", bold=True)
    table(
        ["No", "Nama Kolom", "Tipe Data", "Ukuran", "Keterangan / Constraint"],
        [
            ["1", "id", "Integer", "-", "Primary Key, Auto Increment"],
            ["2", "kode_supplier", "Varchar", "50", "Unique, Nullable"],
            ["3", "nama", "Varchar", "120", "Not Null"],
            ["4", "kontak", "Varchar", "100", "Nullable"],
            ["5", "telepon", "Varchar", "50", "Nullable"],
            ["6", "alamat", "Text", "-", "Nullable"],
            ["7", "created_at", "DateTime", "-", "Not Null, Default: Current UTC"]
        ],
        [500, 2200, 1600, 1200, 3860]
    )

    p("7. Tabel customers", bold=True)
    table(
        ["No", "Nama Kolom", "Tipe Data", "Ukuran", "Keterangan / Constraint"],
        [
            ["1", "id", "Integer", "-", "Primary Key, Auto Increment"],
            ["2", "kode_pelanggan", "Varchar", "50", "Unique, Nullable"],
            ["3", "nama", "Varchar", "120", "Not Null"],
            ["4", "kontak", "Varchar", "100", "Nullable"],
            ["5", "telepon", "Varchar", "50", "Nullable"],
            ["6", "alamat", "Text", "-", "Nullable"],
            ["7", "created_at", "DateTime", "-", "Not Null, Default: Current UTC"]
        ],
        [500, 2200, 1600, 1200, 3860]
    )

    p("8. Tabel store_profile", bold=True)
    table(
        ["No", "Nama Kolom", "Tipe Data", "Ukuran", "Keterangan / Constraint"],
        [
            ["1", "id", "Integer", "-", "Primary Key, Auto Increment"],
            ["2", "nama_toko", "Varchar", "120", "Not Null"],
            ["3", "tagline", "Varchar", "255", "Nullable"],
            ["4", "alamat", "Text", "-", "Nullable"],
            ["5", "telepon", "Varchar", "50", "Nullable"],
            ["6", "email", "Varchar", "100", "Nullable"],
            ["7", "deskripsi", "Text", "-", "Nullable"]
        ],
        [500, 2200, 1600, 1200, 3860]
    )

    # ================== SQLITE ==================
    heading1("2.2.10 SQLite")
    
    heading2("A. Pengertian dan Karakteristik SQLite")
    p("SQLite adalah Relational Database Management System (RDBMS) berbasis berkas (file-based) yang bersifat serverless, self-contained, dan zero-configuration. Karakteristik utama SQLite yang diterapkan pada sistem CB Jaya Tekstil meliputi:")
    p("1. Serverless: Mesin basis data berjalan langsung di dalam proses memori aplikasi Flask tanpa memerlukan instalasi dan pengelolaan server database terpisah.")
    p("2. Self-Contained: Seluruh skema, tabel, indeks, relasi, dan data tersimpan dalam satu berkas tunggal (data.db) di dalam direktori instance/.")
    p("3. Zero-Configuration: Tidak memerlukan konfigurasi port jaringan atau hak akses pengguna database level server.")
    p("4. Transaksional ACID: Memenuhi prinsip Atomicity, Consistency, Isolation, dan Durability sehingga proses pemotongan antrean FIFO terjamin konsisten dan aman dari kegagalan proses.")

    heading2("B. Alasan Pemilihan SQLite")
    p("1. Portabilitas Tinggi: Berkas data.db dapat dipindahkan, dicadangkan (backup), dan dipulihkan dengan mudah tanpa instalasi dependensi rumit.")
    p("2. Kecepatan Akses Tinggi: Menghilangkan latensi komunikasi jaringan (in-process direct I/O) sehingga iterasi pemotongan stok FIFO berlangsung seketika.")
    p("3. Ringan dan Hemat Sumber Daya: Membutuhkan konsumsi CPU dan RAM yang sangat minimal untuk operasional harian toko kasir.")

    heading2("C. Konfigurasi SQLite pada Flask-SQLAlchemy")
    code_box("""# Konfigurasi Koneksi SQLite pada Flask
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///data.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db.init_app(app)

with app.app_context():
    init_db()  # Inisialisasi skema tabel otomatis""")

    heading2("D. Implementasi Query SQL Utama Algoritma FIFO")
    p("1. Kueri Pengambilan Antrean Batch FIFO Tertua:")
    code_box("""SELECT id, nomor_masuk, quantity, sisa_stok, timestamp 
FROM incoming_goods 
WHERE product_id = :product_id AND sisa_stok > 0 
ORDER BY timestamp ASC;""")

    p("2. Kueri Pembaruan Sisa Stok Batch secara Atomik:")
    code_box("""UPDATE incoming_goods 
SET sisa_stok = sisa_stok - :qty_alokasi 
WHERE id = :batch_id;""")

    p("3. Kueri Pencatatan Log Audit Alokasi FIFO:")
    code_box("""INSERT INTO fifo_transactions (sale_id, product_id, incoming_good_id, jumlah_keluar, tanggal_keluar) 
VALUES (:sale_id, :product_id, :batch_id, :qty_alokasi, CURRENT_TIMESTAMP);""")

    # Construct complete document.xml
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

    # Package into a valid DOCX file (ZIP archive)
    with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
        docx_zip.writestr('[Content_Types].xml', content_types_xml)
        docx_zip.writestr('_rels/.rels', rels_xml)
        docx_zip.writestr('word/document.xml', document_xml)
        docx_zip.writestr('word/styles.xml', styles_xml)

    print(f"File '{filename}' successfully generated!")

if __name__ == '__main__':
    output_path = os.path.join(os.getcwd(), 'Laporan_Skripsi_CB_Jaya_Tekstil_Bab2.docx')
    create_docx(output_path)
