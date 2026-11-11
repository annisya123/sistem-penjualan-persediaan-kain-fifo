import os
import zipfile

def create_bab3_docx(filename):
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

    def p(text, bold=False, italic=False, size=24, align="both", space_after=120, space_before=0):
        b_tag = "<w:b/>" if bold else ""
        i_tag = "<w:i/>" if italic else ""
        jc_tag = f'<w:jc w:val="{align}"/>' if align else ""
        sp_tag = f'<w:spacing w:line="360" w:lineRule="auto" w:before="{space_before}" w:after="{space_after}"/>'
        escaped_text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
        
        xml = f"""<w:p>
          <w:pPr>{jc_tag}{sp_tag}</w:pPr>
          <w:r>
            <w:rPr>{b_tag}{i_tag}<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>
            <w:t xml:space="preserve">{escaped_text}</w:t>
          </w:r>
        </w:p>"""
        body_elements.append(xml)

    def chapter_title(title_text):
        p("BAB III", bold=True, size=28, align="center", space_before=240, space_after=60)
        p(title_text, bold=True, size=28, align="center", space_before=0, space_after=240)

    def heading1(text):
        p(text, bold=True, size=26, align="left", space_before=220, space_after=100)

    def heading2(text):
        p(text, bold=True, size=24, align="left", space_before=160, space_after=80)

    def heading3(text):
        p(text, bold=True, size=24, align="left", space_before=120, space_after=60)

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

    # ================== DOKUMEN BAB III ==================
    chapter_title("METODOLOGI PENELITIAN")

    # 3.1 Objek Penelitian
    heading1("3.1 Objek Penelitian")
    p("Objek penelitian merupakan tempat dan entitas yang menjadi fokus kajian dalam pengumpulan data, analisis masalah, serta perancangan dan implementasi sistem informasi. Pada penelitian ini, objek yang diteliti adalah CB Jaya Textile, dengan fokus kajian pada aktivitas transaksi penjualan dan manajemen persediaan kain tekstil berbasis metode First In First Out (FIFO).")

    heading2("3.1.1 Profil Singkat CB Jaya Textile")
    p("CB Jaya Textile adalah unit usaha dagang yang bergerak di bidang distribusi, penjualan grosir, dan eceran aneka jenis produk kain tekstil. Usaha ini mendistribusikan berbagai bahan tekstil berkualitas untuk kebutuhan industri konveksi, penjahit busana rumahan, desainer, dan konsumen individu. Produk kain yang disediakan terdiri dari beragam jenis, motif, dan ketebalan, seperti katun (rayon, jepang, toyobo), sutra velvet, satin silk, kain drill seragam, linen, denim, dan wolfis dengan satuan penjualan per meter maupun per gulungan (roll).")

    heading2("3.1.2 Lokasi Penelitian")
    p("Penelitian dan pengumpulan data dilakukan secara langsung pada lokasi operasional toko dengan data entitas sebagai berikut:")
    p("• Nama Toko / Usaha : CB Jaya Textile")
    p("• Alamat Operasional : Pusat Grosir Textile Blok A No. 12, Bandung, Jawa Barat")
    p("• Telepon / WhatsApp : (+62) 812-3456-7890")
    p("• Email Resmi : info@cbjayatextile.com")
    p("• Waktu Operasional : Senin – Sabtu, Pukul 08.00 – 17.00 WIB")
    p("Pemilihan lokasi ini didasari oleh tingginya volume perputaran barang masuk dari pemasok serta frekuensi transaksi penjualan harian yang memerlukan pengelolaan antrean stok secara sistematis.")

    heading2("3.1.3 Kegiatan Penjualan Kain")
    p("Aktivitas penjualan kain pada CB Jaya Textile mencakup tahapan operasional berikut:")
    p("1. Pemilihan dan Pemesanan Produk: Pelanggan memilih jenis kain, varian warna, serta menentukan kuantitas panjang meter kain yang dibutuhkan.")
    p("2. Pengukuran dan Pemotongan Fisik: Petugas kasir/toko mengambil gulungan kain dari rak penyimpanan untuk diukur dan dipotong secara presisi sesuai permintaan pelanggan.")
    p("3. Pencatatan Transaksi Kasir: Kasir menghitung total harga berdasarkan kuantitas meter dikalikan harga jual satuan per meter, mencatat identitas pembeli, dan memproses pembayaran (tunai atau transfer).")
    p("4. Penerbitan Bukti Transaksi: Kasir menerbitkan nota/struk penjualan resmi sebagai bukti sah transaksi bagi pelanggan dan arsip penjualan toko.")

    heading2("3.1.4 Kegiatan Pengelolaan Persediaan Kain")
    p("Aktivitas manajemen persediaan kain di CB Jaya Textile meliputi:")
    p("1. Penerimaan Pasokan (Inbound Batch): Toko menerima pengiriman kain dari pabrik atau supplier yang datang secara bertahap dalam bentuk batch kedatangan.")
    p("2. Pencatatan Barang Masuk: Petugas gudang mencatat tanggal masuk, nomor surat jalan, nama supplier, jenis kain, dan volume meter yang diterima.")
    p("3. Penyimpanan dan Pengaturan Gudang: Gulungan kain ditempatkan pada rak gudang untuk siap dijual.")
    p("4. Pengurangan Persediaan: Setiap kali terjadi penjualan di kasir, persediaan fisik kain berkurang sejumlah meter yang terpotong.")
    p("5. Pemeriksaan Fisik Persediaan (Stock Opname): Petugas melakukan penghitungan fisik sisa meter kain secara periodik untuk mencocokkan kondisi fisik dengan catatan pembukuan.")

    heading2("3.1.5 Permasalahan pada Proses yang Sedang Berjalan")
    p("Berdasarkan hasil observasi langsung dan wawancara pada operasional toko, diidentifikasi beberapa kendala utama:")
    p("1. Belum Menerapkan Pengendalian Antrean FIFO: Kain dari batch pasokan lama kerap tertumpuk oleh batch pasokan kain yang baru datang. Akibatnya, kain yang lebih lama mengendap di gudang berisiko mengalami penurunan kualitas fisik (warna memudar, kusam, atau serat kain rapuh).")
    p("2. Pencatatan Persediaan Rentan Selisih (Discrepancy): Pencatatan sisa stok meter kain per batch masih dilakukan secara konvensional, sehingga rawan terjadi kesalahan hitung (human error) dan ketidaksesuaian antara catatan buku dengan stok riil di gudang.")
    p("3. Ketiadaan Peringatan Dini Stok Kritis (Low Stock Alert): Tidak adanya sistem pemberitahuan otomatis saat stok kain tertentu telah mencapai batas minimum (reorder point), yang mengakibatkan keterlambatan pemesanan kembali ke supplier (stockout).")
    p("4. Pembuatan Laporan Operasional yang Lambat: Rekapitulasi laporan transaksi penjualan harian, perputaran stok, dan perhitungan laba membutuhkan waktu lama karena pemilik toko harus merekap manual dari tumpukan nota kertas.")

    # 3.2 Metode Penelitian
    heading1("3.2 Metode Penelitian")
    p("Metode penelitian menjelaskan alur kerja dan pendekatan sistematis yang digunakan oleh peneliti untuk menyelesaikan permasalahan pada CB Jaya Textile. Penelitian ini menerapkan pendekatan Rekayasa Perangkat Lunak dengan model System Development Life Cycle (SDLC) yang mencakup tahapan terstruktur:")
    p("1. Identifikasi Masalah: Menemukan dan membatasi ruang lingkup permasalahan operasional, seperti ketiadaan algoritma FIFO dan pencatatan manual.")
    p("2. Pengumpulan Data: Mengumpulkan informasi primer dan sekunder melalui observasi langsung, wawancara mendalam, dan studi kepustakaan.")
    p("3. Analisis Kebutuhan Sistem: Mengidentifikasi kebutuhan fungsional (seperti fitur POS kasir, alokasi batch FIFO, peringatan stok) dan kebutuhan nonfungsional (keamanan, performa, dan keandalan data ACID).")
    p("4. Perancangan Sistem: Merancang model visual UML (Use Case, Activity, Sequence, dan Class Diagram), skema basis data relasional (ERD, LRS, kamus data), serta perancangan antarmuka (UI/UX).")
    p("5. Implementasi Sistem: Menerjemahkan rancangan ke dalam kode program menggunakan Python (Flask), basis data SQLite (SQLAlchemy), dan frontend HTML/CSS/JavaScript.")
    p("6. Pengujian Sistem: Menguji fungsionalitas sistem menggunakan metode Black Box Testing guna memastikan algoritma FIFO dan seluruh modul berjalan bebas dari kesalahan.")

    # 3.3 Teknik Pengumpulan Data
    heading1("3.3 Teknik Pengumpulan Data")
    p("Teknik pengumpulan data merupakan langkah strategis untuk memperoleh data yang valid, akurat, dan relevan guna mendukung analisis dan perancangan sistem.")

    heading2("3.3.1 Observasi")
    p("Pengamatan secara langsung dilakukan pada 6 (enam) aktivitas proses kerja operasional:")
    p("1. Proses Barang Masuk: Mengamati prosedur penerimaan gulungan kain dari supplier dan pencocokan surat jalan pengiriman.")
    p("2. Proses Penyimpanan Kain: Mengamati tata letak penempatan kain di rak gudang serta risiko tertimbunnya pasokan lama.")
    p("3. Proses Pencatatan Stok: Mengamati pencatatan data kain yang masih menggunakan buku besar manual.")
    p("4. Proses Transaksi Penjualan: Mengamati pelayanan pembeli, pengukuran meter kain, dan perhitungan tagihan di meja kasir.")
    p("5. Proses Pengurangan Stok: Mengamati mekanisme pengurangan sisa stok kain setelah transaksi pemotongan fisik selesai.")
    p("6. Proses Pembuatan Laporan: Mengamati penghitungan manual nota-nota fisik untuk membuat rekapitulasi penjualan dan stok bulanan.")

    heading2("3.3.2 Wawancara")
    p("Wawancara dilakukan kepada pihak internal yang terlibat langsung dalam proses penjualan dan persediaan:")
    p("1. Pemilik Toko (Owner / Manajer): Membahas kebijakan harga jual/beli, pengadaan barang, dan kendala pembuatan laporan keuangan serta evaluasi laba kotor.")
    p("2. Bagian Administrasi / Gudang (Admin): Membahas kendala pelacakan urutan kedatangan kain, kesulitan menjalankan FIFO secara manual, dan terjadinya selisih stok.")
    p("3. Bagian Penjualan / Kasir: Membahas alur transaksi kasir, perhitungan harga total, antrean jam sibuk, dan pembuatan nota.")

    p("Tabel 3.1 Ringkasan Narasumber dan Hasil Wawancara", bold=True)
    table(
        ["No", "Narasumber", "Peran / Jabatan", "Kebutuhan Utama terhadap Sistem Baru"],
        [
            ["1", "Pemilik Toko", "Pemilik Usaha / Manajer", "Rekapitulasi laporan penjualan, laba kotor, dan monitoring stok secara otomatis dan real-time."],
            ["2", "Administrator / Gudang", "Pengelola Persediaan & Pasokan", "Pencatatan nomor batch masuk, otomatisasi antrean FIFO, dan notifikasi batas stok minimum."],
            ["3", "Bagian Penjualan", "Kasir Toko", "Form kasir yang cepat, validasi stok fisik secara instan, dan pencetakan nota penjualan otomatis."]
        ],
        [500, 2000, 2200, 4660]
    )

    heading2("3.3.3 Studi Pustaka")
    p("Mengkaji literatur ilmiah mengenai manajemen persediaan, algoritma First In First Out (FIFO), framework Flask, basis data SQLite, dan perancangan UML.")

    # 3.4 Analisis Sistem yang Berjalan
    heading1("3.4 Analisis Sistem yang Berjalan")
    p("Menjelaskan bagaimana proses operasional di CB Jaya Textile dilakukan sebelum sistem informasi diterapkan.")

    heading2("3.4.1 Proses Pendataan Kain")
    p("Data kain baru dicatat manual di buku master barang (nama, jenis bahan, warna, harga beli, dan harga jual). Jika ada perubahan harga atau varian warna, petugas mencoret catatan lama sehingga buku menjadi tidak rapi.")

    heading2("3.4.2 Proses Barang Masuk")
    p("Pemasok mengirimkan kain beserta surat jalan. Petugas memeriksa fisik gulungan kain dan mencatat tanggal masuk, supplier, jenis kain, serta total meter pada Buku Barang Masuk tanpa penomoran batch kedatangan khusus.")

    heading2("3.4.3 Proses Penjualan")
    p("Pelanggan memilih kain dan kuantitas meter. Kasir mengambil gulungan kain, mengukur, dan memotong secara fisik. Kasir menghitung total harga dengan kalkulator dan menuliskan nota manual 2 rangkap (Lembar 1 ke pembeli, Lembar 2 untuk arsip toko).")

    heading2("3.4.4 Proses Pengelolaan Persediaan")
    p("Kasir tidak dapat mengetahui stok secara instan di meja kasir. Pengurangan stok baru dicatat di akhir hari atau saat waktu luang. Untuk mengetahui sisa stok riil, petugas harus mengukur ulang kain secara manual saat stock opname.")

    heading2("3.4.5 Permasalahan Sistem Berjalan")
    p("1. Pencatatan Masih Manual: Menggunakan media kertas yang rentan rusak, hilang, atau salah catat (human error).")
    p("2. Sulit Mengetahui Stok secara Cepat: Kasir harus memeriksa langsung ke gudang sehingga memperlambat pelayanan.")
    p("3. Sulit Mengetahui Urutan Stok Berdasarkan Barang Masuk (Ketiadaan FIFO): Kain pasokan lama tertimbun kain baru sehingga berisiko rusak atau kusam.")
    p("4. Pencatatan Transaksi dan Stok Belum Terintegrasi: Catatan buku gudang tidak terpotong otomatis saat transaksi kasir terjadi sehingga timbul selisih stok.")
    p("5. Pembuatan Laporan Membutuhkan Waktu Lama: Rekapitulasi nota kertas satu per satu memakan waktu berhari-hari dan rawan salah hitung.")

    heading2("3.4.6 Flowmap Sistem Berjalan")
    p("Flowmap menggambarkan aliran dokumen fisik antara Pelanggan, Kasir, Bagian Gudang, dan Pemilik Toko:")
    p("• Pelanggan memesan kain -> Kasir cek fisik ke gudang -> Kasir potong kain & buat Nota Penjualan rangkap 2 -> Lembar 1 ke Pelanggan, Lembar 2 disimpan di map arsip.")
    p("• Pemasok kirim kain + Surat Jalan -> Petugas Gudang cek fisik -> Catat ke Buku Barang Masuk -> Simpan di rak.")
    p("• Di akhir periode, Pemilik Toko mengumpulkan arsip nota dan buku stok -> Rekapitulasi manual -> Membuat Laporan Penjualan & Laporan Stok.")

    # 3.5 Analisis Kebutuhan Sistem
    heading1("3.5 Analisis Kebutuhan Sistem")

    heading2("3.5.1 Kebutuhan Fungsional")
    p("Tabel 3.2 Matriks Kebutuhan Fungsional Sistem", bold=True)
    table(
        ["Kode", "Modul / Fitur", "Deskripsi Kebutuhan Fungsional", "Hak Akses"],
        [
            ["KF-01", "Autentikasi Pengguna", "Memverifikasi login dengan username/password, kelola session, dan batasi hak akses sesuai role.", "Admin, Kasir"],
            ["KF-02", "Kelola Master Produk", "Menambah, mengubah, melihat detail, menghapus, serta unggah foto kain beserta harga dan stok minimum.", "Admin"],
            ["KF-03", "Kelola Supplier & Customer", "Mencatat dan memperbarui direktori data pemasok kain dan pelanggan.", "Admin, Kasir"],
            ["KF-04", "Pencatatan Barang Masuk (FIFO)", "Mencatat pasokan kain ke dalam batch masuk baru (nomor batch unik BM-YYYYMMDD-XXX) dan sisa_stok per batch.", "Admin"],
            ["KF-05", "Transaksi Penjualan Kasir", "Memproses transaksi penjualan kain, kalkulasi total harga otomatis, validasi stok, dan cetak struk.", "Admin, Kasir"],
            ["KF-06", "Eksekusi Algoritma FIFO", "Memotong stok kain otomatis dari batch masuk tertua (timestamp ASC) saat transaksi penjualan.", "Sistem"],
            ["KF-07", "Log Audit FIFO", "Merekam hubungan nomor nota dengan batch asal dan jumlah meter yang dipotong ke fifo_transactions.", "Sistem"],
            ["KF-08", "Peringatan Stok Minimum", "Menampilkan peringatan visual (low stock alert) saat stok kain mencapai ambang batas minimum.", "Admin, Kasir"],
            ["KF-09", "Cetak Nota & Faktur", "Menampilkan pratinjau nota transaksi dan mencetaknya langsung ke printer struk atau unduh PDF.", "Admin, Kasir"],
            ["KF-10", "Laporan & Rekapitulasi", "Menyajikan laporan penjualan, laporan stok, dan riwayat audit FIFO dengan filter tanggal dan ekspor CSV.", "Admin"],
            ["KF-11", "Kelola Profil Toko", "Mengubah data toko (nama, slogan, alamat, telepon) untuk kop nota resmi.", "Admin"]
        ],
        [900, 2100, 4860, 1500]
    )

    heading2("3.5.2 Kebutuhan Nonfungsional")
    p("Tabel 3.3 Matriks Kebutuhan Nonfungsional Sistem", bold=True)
    table(
        ["Aspek", "Parameter", "Spesifikasi Kebutuhan Nonfungsional"],
        [
            ["Usability", "Antarmuka (UI/UX)", "Antarmuka web responsif, intuitif, dan mudah dioperasikan oleh kasir tanpa pelatihan rumit."],
            ["Performance", "Waktu Respons", "Eksekusi pemotongan stok FIFO dan penyimpanan transaksi kasir berlangsung cepat (< 2 detik)."],
            ["Reliability", "Integritas Data (ACID)", "Menerapkan transaksi atomik SQLite dengan mekanisme Rollback jika stok kurang, serta Foreign Key CASCADE."],
            ["Security", "Keamanan Akses", "Enkripsi kata sandi menggunakan hash PBKDF2/SHA-256 dan proteksi rute halaman via @login_required."],
            ["Portability", "Portabilitas Sistem", "Berbasis SQLite file-based (data.db) dan Python Flask yang mandiri dan mudah dipindahkan (portable)."]
        ],
        [1600, 2000, 5760]
    )

    # 3.6 Perancangan Sistem
    heading1("3.6 Perancangan Sistem")

    heading2("3.6.1 Arsitektur Sistem")
    p("Sistem CB Jaya Textile dirancang dengan arsitektur 3-Tier berbasis Model-View-Controller (MVC):")
    p("1. Presentation Layer (View): Antarmuka web pengguna berbasis HTML5, CSS3, JavaScript, dan Jinja2 Template.")
    p("2. Application Logic Layer (Controller): Server Flask (app.py) yang mengelola routing, autentikasi, dan mesin FIFO (proses_fifo).")
    p("3. Data Storage Layer (Model): Pustaka Flask-SQLAlchemy (models.py) dan database relasional SQLite (data.db).")

    heading2("3.6.2 Use Case Diagram")
    p("Sistem melibatkan 2 aktor: Admin dan Kasir. Admin memiliki hak akses penuh terhadap seluruh modul (kelola produk, supplier, customer, barang masuk, penjualan, laporan, dan profil toko). Kasir memiliki hak akses pada modul penjualan kasir, cetak nota, dan pemantauan stok.")

    p("Tabel 3.4 Deskripsi Use Case Sistem", bold=True)
    table(
        ["No", "Nama Use Case", "Aktor", "Deskripsi Fungsionalitas"],
        [
            ["1", "Login / Logout", "Admin, Kasir", "Memvalidasi kredensial pengguna dan mengakhiri sesi kerja."],
            ["2", "Kelola Data Produk", "Admin", "Menambah, mengubah, melihat detail, menghapus, dan upload foto kain."],
            ["3", "Kelola Supplier & Customer", "Admin", "Mencatat dan mengelola data pemasok kain dan pelanggan."],
            ["4", "Transaksi Barang Masuk", "Admin", "Mencatat pasokan kain baru ke dalam entitas batch masuk FIFO."],
            ["5", "Transaksi Penjualan Kasir", "Admin, Kasir", "Memilih kain, input meter, dan alokasi pemotongan stok otomatis dari batch tertua."],
            ["6", "Cetak Nota Penjualan", "Admin, Kasir", "Mencetak struk transaksi penjualan resmi ke printer atau PDF."],
            ["7", "Kelola Laporan", "Admin", "Melihat dan mengekspor rekapitulasi penjualan, stok, dan audit FIFO ke CSV."],
            ["8", "Kelola Profil Toko", "Admin", "Memperbarui identitas toko untuk kop faktur."]
        ],
        [500, 2400, 1600, 4860]
    )

    heading2("3.6.3 Activity Diagram")
    p("Menggambarkan alur aktivitas transaksi kasir: Kasir memilih kain -> Input kuantitas meter -> Sistem cek sisa stok -> Jika cukup, sistem mengambil batch masuk tertua (timestamp ASC) -> Kurangi sisa_stok batch -> Catat ke fifo_transactions -> Simpan transaksi sales -> Kurangi actual_stock -> Commit database -> Cetak nota.")

    heading2("3.6.4 Sequence Diagram")
    p("Memodelkan pertukaran pesan secara kronologis antara Aktor (Kasir), Browser (UI), Server Flask (Controller), dan Database (Model) saat mengeksekusi algoritma pemotongan antrean batch FIFO.")

    heading2("3.6.5 Class Diagram")
    p("Memodelkan kelas entitas: User, Product, IncomingGood, Sale, FifoTransaction, Supplier, Customer, dan StoreProfile beserta atribut, metode, dan relasi kardinalitasnya.")

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

    print(f"File '{filename}' successfully created!")

if __name__ == '__main__':
    output_file = os.path.join(os.getcwd(), 'Laporan_Skripsi_CB_Jaya_Textile_Bab3.docx')
    create_bab3_docx(output_file)
