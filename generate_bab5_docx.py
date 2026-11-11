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

    def p(text, bold=False, italic=False, size=24, align="both", space_after=120, space_before=0, indent=0, numbering=False):
        b_tag = "<w:b/>" if bold else ""
        i_tag = "<w:i/>" if italic else ""
        jc_tag = f'<w:jc w:val="{align}"/>' if align else ""
        sp_tag = f'<w:spacing w:line="360" w:lineRule="auto" w:before="{space_before}" w:after="{space_after}"/>'
        
        ind_tag = f'<w:ind w:left="{indent}" w:firstLine="0"/>' if indent > 0 else ""
        if numbering and indent > 0:
            ind_tag = f'<w:ind w:left="{indent}" w:hanging="360"/>'
        elif not numbering and indent > 0:
             ind_tag = f'<w:ind w:left="{indent}" w:firstLine="360"/>'
             
        # Escape XML entities
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
    <w:t>{escaped_text}</w:t>
  </w:r>
</w:p>"""
        body_elements.append(xml)

    def heading(text, level=1):
        if level == 1:
            p(text, bold=True, size=28, align="center", space_before=240, space_after=240)
        elif level == 2:
            p(text, bold=True, size=24, align="left", space_before=120, space_after=120)
        else:
            p(text, bold=True, size=24, align="left", space_before=120, space_after=120)

    def br():
        body_elements.append("<w:p><w:pPr><w:spacing w:after=\"120\"/></w:pPr></w:p>")
        
    def page_break():
        body_elements.append("""<w:p>
  <w:r>
    <w:br w:type="page"/>
  </w:r>
</w:p>""")

    # ---------------------------------------------------------
    # Document Content Generation
    # ---------------------------------------------------------

    # BAB V KESIMPULAN DAN SARAN
    heading("BAB V", level=1)
    heading("KESIMPULAN DAN SARAN", level=1)

    heading("5.1 Kesimpulan", level=2)
    p("Berisi jawaban langsung terhadap rumusan masalah dan tujuan penelitian.", indent=360)
    p("Untuk penelitian Anda, isi kesimpulan mencakup:", indent=360)
    
    p("1. Hasil pembangunan sistem penjualan dan persediaan kain berbasis web di CB Jaya Textile.", indent=720, numbering=True)
    p("2. Hasil penerapan algoritma FIFO dalam pengeluaran stok berdasarkan urutan batch masuk.", indent=720, numbering=True)
    p("3. Hasil penerapan metode Waterfall dalam proses pengembangan sistem.", indent=720, numbering=True)
    p("4. Hasil pengujian sistem, khususnya Black Box Testing.", indent=720, numbering=True)

    br()
    
    heading("5.2 Saran", level=2)
    p("Berisi pengembangan yang dapat dilakukan pada penelitian berikutnya, misalnya:", indent=360)
    
    p("1. Pengembangan notifikasi stok minimum.", indent=720, numbering=True)
    p("2. Pengembangan laporan penjualan dan persediaan yang lebih lengkap.", indent=720, numbering=True)
    p("3. Pengembangan pengelolaan supplier dan pelanggan.", indent=720, numbering=True)
    p("4. Pengembangan penerapan FIFO untuk jumlah produk dan transaksi yang lebih kompleks.", indent=720, numbering=True)
    p("5. Perbandingan FIFO dengan metode persediaan lainnya.", indent=720, numbering=True)

    page_break()

    # DAFTAR PUSTAKA
    heading("DAFTAR PUSTAKA", level=1)
    
    p("Al Fatta, H. (2007). Analisis dan Perancangan Sistem Informasi untuk Keunggulan Bersaing Perusahaan dan Organisasi Modern. Yogyakarta: Andi Offset.", indent=360)
    p("Kendall, K. E., & Kendall, J. E. (2010). Analisis dan Perancangan Sistem (Edisi 5). Jakarta: Indeks.", indent=360)
    p("Pressman, R. S. (2012). Rekayasa Perangkat Lunak: Pendekatan Praktisi (Buku 1). Yogyakarta: Andi Offset.", indent=360)
    p("Sukamto, R. A., & Shalahuddin, M. (2018). Rekayasa Perangkat Lunak Terstruktur dan Berorientasi Objek. Bandung: Informatika.", indent=360)
    p("Romney, M. B., & Steinbart, P. J. (2015). Sistem Informasi Akuntansi (Edisi 13). Jakarta: Salemba Empat.", indent=360)

    page_break()

    # LAMPIRAN
    heading("LAMPIRAN", level=1)
    p("Lampiran 1. Surat Izin Penelitian", align="left")
    p("Lampiran 2. Surat Keterangan Selesai Penelitian dari CB Jaya Textile", align="left")
    p("Lampiran 3. Form Hasil Wawancara", align="left")
    p("Lampiran 4. Hasil Pengujian Black Box Testing", align="left")
    p("Lampiran 5. Dokumentasi Kegiatan Penelitian", align="left")

    # ---------------------------------------------------------
    # Compile document.xml
    # ---------------------------------------------------------
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {''.join(body_elements)}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    # ---------------------------------------------------------
    # Write to ZIP
    # ---------------------------------------------------------
    with zipfile.ZipFile(filename, 'w', compression=zipfile.ZIP_DEFLATED) as docx:
        docx.writestr('[Content_Types].xml', content_types_xml)
        docx.writestr('_rels/.rels', rels_xml)
        docx.writestr('word/document.xml', document_xml)
        docx.writestr('word/styles.xml', styles_xml)
    
    print(f"Created: {filename}")

if __name__ == '__main__':
    create_docx('BAB_V_Kesimpulan_Saran_CB_Jaya_Textile.docx')
