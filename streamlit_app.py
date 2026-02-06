import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import io
import os
from PIL import Image

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(page_title="Generator Kartu Ujian (Word)", page_icon="🖨️", layout="wide")

# CSS Modern
st.markdown("""
<style>
    .stApp { background-color: #f8fafc; }
    .main-header { 
        font-size: 2rem; font-weight: 800; color: #1e3a8a; margin-bottom: 10px; 
    }
    div.stButton > button {
        background-color: #166534; color: white; border-radius: 8px; font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

if 'jadwal_ujian' not in st.session_state:
    st.session_state['jadwal_ujian'] = []

# ==========================================
# 2. LOGIC WORD GENERATOR
# ==========================================
def set_cell_color(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def generate_word_doc(df, config, template_type, logo_bytes, ttd_bytes, jadwal_list):
    doc = Document()
    
    # Setup Halaman A4 Landscape
    section = doc.sections[0]
    section.orientation = 1 # Landscape
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(1.27); section.right_margin = Cm(1.27)
    section.top_margin = Cm(1.27); section.bottom_margin = Cm(1.27)

    # Konfigurasi Template
    header_bg = "FFFFFF"
    text_hdr_color = RGBColor(0,0,0)
    judul_kartu = "KARTU PESERTA UJIAN"
    judul_color = RGBColor(0,0,0)

    if template_type == "Modern Blue": 
        header_bg = "1E3A8A"; text_hdr_color = RGBColor(255,255,255)
    elif template_type == "Islamic Green": 
        header_bg = "14532D"; text_hdr_color = RGBColor(255,215,0)
    elif template_type == "Kartu Hilang (Emergency)":
        header_bg = "B91C1C"; text_hdr_color = RGBColor(255,255,255)
        judul_kartu = "DUPLIKAT KARTU UJIAN"
        judul_color = RGBColor(255,0,0)

    # Looping Data Siswa
    count = 0
    # Pastikan nama kolom sesuai (Upper case)
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    for index, row in df.iterrows():
        # Tabel Utama (2 Kolom)
        main_tbl = doc.add_table(rows=1, cols=2)
        main_tbl.style = 'Table Grid'; main_tbl.autofit = False
        
        # === KIRI: BIODATA ===
        cell_l = main_tbl.cell(0, 0); cell_l.width = Cm(13.5)
        p = cell_l.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Logo
        if logo_bytes:
            run = p.add_run()
            run.add_picture(io.BytesIO(logo_bytes), width=Cm(1.8))
        
        p.add_run(f"\nYAYASAN PENDIDIKAN ISLAM AL-GHOZALI\n").font.size = Pt(9)
        r = p.add_run(f"{config['sekolah']}\n")
        r.font.bold = True; r.font.size = Pt(11)
        p.add_run("_____________________________________________").font.size = Pt(6)
        
        # Judul
        p2 = cell_l.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p2.add_run(judul_kartu)
        r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = judul_color
        
        if template_type == "Kartu Hilang (Emergency)":
            r = p2.add_run("\n(DICETAK ULANG OLEH PANITIA)")
            r.font.size = Pt(8); r.font.color.rgb = RGBColor(255,0,0)

        # Tabel Biodata Nested
        bio_tbl = cell_l.add_table(rows=4, cols=3); bio_tbl.autofit = False
        c_foto = bio_tbl.cell(0,0); c_foto.merge(bio_tbl.cell(3,0)); c_foto.width = Cm(3.0)
        c_foto.text = "FOTO\n3x4"; c_foto.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Isi Data
        nisn_val = str(row.get('NISN', '')).replace('.0','')
        items = [
            ("No Peserta", row.get('NOMOR PESERTA', '-')), 
            ("Nama", row.get('NAMA PESERTA', '-')), 
            ("NISN", nisn_val), 
            ("Ruang", row.get('RUANG', '-'))
        ]
        
        for i, (lbl, val) in enumerate(items):
            bio_tbl.cell(i, 1).text = lbl
            bio_tbl.cell(i, 1).paragraphs[0].runs[0].font.size = Pt(9)
            bio_tbl.cell(i, 2).text = ": " + str(val)
            bio_tbl.cell(i, 2).paragraphs[0].runs[0].font.size = Pt(9)
            bio_tbl.cell(i, 2).paragraphs[0].runs[0].font.bold = True

        # TTD Area
        pttd = cell_l.add_paragraph(); pttd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pttd.add_run(f"\n{config['tanggal']}\nKepala Sekolah,\n").font.size = Pt(9)
        
        if ttd_bytes:
            pttd.add_run().add_picture(io.BytesIO(ttd_bytes), width=Cm(2.0))
            pttd.add_run("\n")
        else:
            pttd.add_run("\n\n\n")
            
        pttd.add_run(config['kepsek']).font.bold = True

        # === KANAN: JADWAL ===
        cell_r = main_tbl.cell(0, 1); cell_r.width = Cm(13.5)
        pr = cell_r.paragraphs[0]; pr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = pr.add_run("JADWAL UJIAN"); r.font.bold = True; r.font.size = Pt(12)
        if template_type == "Kartu Hilang (Emergency)": r.font.color.rgb = RGBColor(255,0,0)

        if jadwal_list:
            j_tbl = cell_r.add_table(rows=1, cols=4); j_tbl.style = 'Table Grid'
            hdr = j_tbl.rows[0]
            
            # Header
            col_names = ["HARI", "JAM", "WAKTU", "MAPEL"]
            for idx, txt in enumerate(col_names):
                set_cell_color(hdr.cells[idx], header_bg)
                hdr.cells[idx].text = txt
                run = hdr.cells[idx].paragraphs[0].runs[0]
                run.font.bold = True
                run.font.color.rgb = text_hdr_color
                run.font.size = Pt(8)

            # Isi
            for jdata in jadwal_list:
                cells = j_tbl.add_row().cells
                # jdata = [Hari, Jam, Waktu, Mapel]
                for idx, val in enumerate(jdata):
                    if idx < 4:
                        cells[idx].text = str(val)
                        cells[idx].paragraphs[0].runs[0].font.size = Pt(8)
        else:
            cell_r.add_paragraph("(Jadwal Tidak Diatur)")

        doc.add_paragraph("\n")
        count += 1
        if count % 2 == 0: doc.add_page_break()

    return doc

# ==========================================
# 3. UI LAYOUT
# ==========================================
st.markdown('<div class="main-header">Al-Ghozali Word Generator (Web)</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Konfigurasi")
    template_sel = st.selectbox("Template Kartu:", ["Standard", "Modern Blue", "Islamic Green", "Kartu Hilang (Emergency)"])
    st.markdown("---")
    in_sekolah = st.text_input("Nama Sekolah", "SMA ISLAM AL-GHOZALI")
    in_kepsek = st.text_input("Kepala Sekolah", "Antoni Firdaus, SHI, M.Pd.")
    in_tgl = st.text_input("Tanggal TTD", "Gunung Sindur, 20 Mei 2026")

tab1, tab2, tab3 = st.tabs(["📂 1. Upload Data", "📅 2. Atur Jadwal", "🖨️ 3. Download Word"])

# --- TAB 1 ---
with tab1:
    st.subheader("Upload Data Siswa & Aset")
    col1, col2 = st.columns(2)
    with col1:
        upl_excel = st.file_uploader("File Excel Data Siswa (.xlsx)", type=['xlsx'])
        if upl_excel:
            df = pd.read_excel(upl_excel)
            st.session_state['df_siswa'] = df
            st.success(f"Berhasil memuat {len(df)} data siswa.")
            
    with col2:
        upl_logo = st.file_uploader("Logo Sekolah", type=['png','jpg'])
        upl_ttd = st.file_uploader("Scan TTD", type=['png','jpg'])
        
        # Simpan di session state sebagai bytes
        if upl_logo: st.session_state['logo_bytes'] = upl_logo.getvalue()
        if upl_ttd: st.session_state['ttd_bytes'] = upl_ttd.getvalue()

# --- TAB 2 ---
with tab2:
    st.subheader("Pengaturan Jadwal")
    
    # Cek apakah ada jadwal di Excel
    if upl_excel:
        try:
            df_jadwal = pd.read_excel(upl_excel, sheet_name="JADWAL")
            # Ambil 4 kolom pertama konversi ke list
            jadwal_excel = df_jadwal.iloc[:, :4].astype(str).values.tolist()
            if jadwal_excel:
                st.session_state['jadwal_ujian'] = jadwal_excel
                st.success(f"{len(jadwal_excel)} Mapel dimuat otomatis dari Excel.")
        except:
            st.info("Tidak ditemukan sheet 'JADWAL' di Excel. Silakan input manual di bawah.")

    # Input Manual
    c1, c2, c3, c4 = st.columns(4)
    with c1: t_hari = st.text_input("Hari")
    with c2: t_jam = st.text_input("Jam Ke")
    with c3: t_waktu = st.text_input("Waktu")
    with c4: 
        t_mapel = st.text_input("Mapel")
        if st.button("➕ Tambah"):
            st.session_state['jadwal_ujian'].append([t_hari, t_jam, t_waktu, t_mapel])
            st.rerun()

    if st.session_state['jadwal_ujian']:
        st.write("Preview Jadwal:")
        st.table(pd.DataFrame(st.session_state['jadwal_ujian'], columns=["HARI", "JAM", "WAKTU", "MAPEL"]))
        if st.button("Hapus Semua Jadwal"):
            st.session_state['jadwal_ujian'] = []
            st.rerun()

# --- TAB 3 ---
with tab3:
    st.subheader("Eksekusi")
    
    if 'df_siswa' in st.session_state:
        st.info("Klik tombol di bawah untuk membuat file Word (.docx).")
        
        if st.button("🚀 GENERATE FILE WORD"):
            config = {'sekolah': in_sekolah, 'kepsek': in_kepsek, 'tanggal': in_tgl}
            logo_b = st.session_state.get('logo_bytes')
            ttd_b = st.session_state.get('ttd_bytes')
            
            # Generate Doc
            doc = generate_word_doc(st.session_state['df_siswa'], config, template_sel, logo_b, ttd_b, st.session_state['jadwal_ujian'])
            
            # Simpan ke Memory Buffer
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("Selesai! Silakan download.")
            st.download_button(
                label="📥 DOWNLOAD KARTU (.DOCX)",
                data=bio.getvalue(),
                file_name=f"Kartu_Ujian_{template_sel}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("Silakan upload data Excel dulu di Tab 1.")
