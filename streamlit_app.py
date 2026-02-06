import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import io
import os
from PIL import Image
import zipfile

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(page_title="Generator Kartu Ujian (Word)", page_icon="🖨️", layout="wide")

# CSS Modern & Tabel Full
st.markdown("""
<style>
    .stApp { background-color: #f1f5f9; }
    .main-header { 
        font-size: 2rem; font-weight: 800; color: #1e3a8a; margin-bottom: 10px; 
        border-bottom: 2px solid #cbd5e1; padding-bottom: 10px;
    }
    /* Tombol Hijau */
    div.stButton > button {
        background-color: #15803d; color: white; border-radius: 6px; font-weight: bold; height: 3em; width: 100%;
    }
    div.stButton > button:hover { background-color: #166534; border-color: #166534; color: white; }
    
    /* Tabel Dataframe */
    [data-testid="stDataFrame"] { width: 100%; }
</style>
""", unsafe_allow_html=True)

# Inisialisasi Session State
if 'jadwal_ujian' not in st.session_state: st.session_state['jadwal_ujian'] = []
if 'photos' not in st.session_state: st.session_state['photos'] = {}

# ==========================================
# 2. LOGIC FUNCTIONS
# ==========================================
def process_images(uploaded_files, is_zip=False):
    """Memproses gambar baik dari ZIP maupun Upload Biasa"""
    new_photos = {}
    
    if is_zip and uploaded_files:
        with zipfile.ZipFile(uploaded_files) as z:
            for f in z.namelist():
                if f.lower().endswith(('.png','.jpg','.jpeg')) and not f.startswith('__'):
                    try:
                        base = f.split('/')[-1].rsplit('.',1)[0]
                        new_photos[base] = Image.open(io.BytesIO(z.read(f)))
                    except: continue
    elif uploaded_files:
        for up_file in uploaded_files:
            try:
                base = up_file.name.rsplit('.',1)[0]
                new_photos[base] = Image.open(up_file)
            except: continue
            
    return new_photos

def set_cell_color(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def generate_word_doc(df, config, template_type, logo_bytes, ttd_bytes, jadwal_list, photos):
    doc = Document()
    
    section = doc.sections[0]
    section.orientation = 1
    section.page_width = Cm(29.7); section.page_height = Cm(21.0)
    section.left_margin = Cm(1.27); section.right_margin = Cm(1.27)
    section.top_margin = Cm(1.27); section.bottom_margin = Cm(1.27)

    header_bg = "FFFFFF"; text_hdr_color = RGBColor(0,0,0)
    judul_kartu = "KARTU PESERTA UJIAN"; judul_color = RGBColor(0,0,0)

    if template_type == "Modern Blue": header_bg = "1E3A8A"; text_hdr_color = RGBColor(255,255,255)
    elif template_type == "Islamic Green": header_bg = "14532D"; text_hdr_color = RGBColor(255,215,0)
    elif template_type == "Kartu Hilang (Emergency)":
        header_bg = "B91C1C"; text_hdr_color = RGBColor(255,255,255)
        judul_kartu = "DUPLIKAT KARTU UJIAN"; judul_color = RGBColor(255,0,0)

    df.columns = [str(c).strip().upper() for c in df.columns]
    
    for index, row in df.iterrows():
        main_tbl = doc.add_table(rows=1, cols=2)
        main_tbl.style = 'Table Grid'; main_tbl.autofit = False
        
        # === KIRI ===
        cell_l = main_tbl.cell(0, 0); cell_l.width = Cm(13.5)
        p = cell_l.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        if logo_bytes:
            run = p.add_run(); run.add_picture(io.BytesIO(logo_bytes), width=Cm(1.8))
        
        p.add_run(f"\nYAYASAN PENDIDIKAN ISLAM AL-GHOZALI\n").font.size = Pt(9)
        r = p.add_run(f"{config['sekolah']}\n")
        r.font.bold = True; r.font.size = Pt(11)
        p.add_run("_____________________________________________").font.size = Pt(6)
        
        p2 = cell_l.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p2.add_run(judul_kartu)
        r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = judul_color
        if template_type == "Kartu Hilang (Emergency)":
            r = p2.add_run("\n(DICETAK ULANG OLEH PANITIA)"); r.font.size = Pt(8); r.font.color.rgb = RGBColor(255,0,0)

        bio_tbl = cell_l.add_table(rows=4, cols=3); bio_tbl.autofit = False
        c_foto = bio_tbl.cell(0,0); c_foto.merge(bio_tbl.cell(3,0)); c_foto.width = Cm(3.0)
        
        # LOGIC FOTO
        nis_key = str(row.get('NISN', '')).replace('.0','')
        if nis_key not in photos: nis_key = str(row.get('NIS', '')).replace('.0','') # Fallback ke NIS
        
        if nis_key in photos:
            try:
                img_pil = photos[nis_key]
                img_byte = io.BytesIO()
                img_pil.save(img_byte, format='JPEG')
                p_f = c_foto.paragraphs[0]; p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_f.add_run().add_picture(io.BytesIO(img_byte.getvalue()), width=Cm(2.5), height=Cm(3.2))
            except: c_foto.text = "Error Foto"
        else:
            c_foto.text = "FOTO\n3x4"; c_foto.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        items = [("No Peserta", row.get('NOMOR PESERTA', '-')), ("Nama", row.get('NAMA PESERTA', '-')), ("NISN", str(row.get('NISN','-')).replace('.0','')), ("Ruang", row.get('RUANG', '-'))]
        for i, (lbl, val) in enumerate(items):
            bio_tbl.cell(i, 1).text = lbl; bio_tbl.cell(i, 1).paragraphs[0].runs[0].font.size = Pt(9)
            bio_tbl.cell(i, 2).text = ": " + str(val); bio_tbl.cell(i, 2).paragraphs[0].runs[0].font.size = Pt(9); bio_tbl.cell(i, 2).paragraphs[0].runs[0].font.bold = True

        pttd = cell_l.add_paragraph(); pttd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pttd.add_run(f"\n{config['tanggal']}\nKepala Sekolah,\n").font.size = Pt(9)
        if ttd_bytes: pttd.add_run().add_picture(io.BytesIO(ttd_bytes), width=Cm(2.0)); pttd.add_run("\n")
        else: pttd.add_run("\n\n\n")
        pttd.add_run(config['kepsek']).font.bold = True

        # === KANAN ===
        cell_r = main_tbl.cell(0, 1); cell_r.width = Cm(13.5)
        pr = cell_r.paragraphs[0]; pr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = pr.add_run("JADWAL UJIAN"); r.font.bold = True; r.font.size = Pt(12)
        if template_type == "Kartu Hilang (Emergency)": r.font.color.rgb = RGBColor(255,0,0)

        if jadwal_list:
            j_tbl = cell_r.add_table(rows=1, cols=5); j_tbl.style = 'Table Grid'
            hdr = j_tbl.rows[0]
            widths = [Cm(2.0), Cm(1.0), Cm(2.5), Cm(6.0), Cm(1.5)]
            for idx, txt in enumerate(["HARI", "JAM", "WAKTU", "MAPEL", "PARAF"]):
                set_cell_color(hdr.cells[idx], header_bg)
                hdr.cells[idx].text = txt; hdr.cells[idx].width = widths[idx]
                run = hdr.cells[idx].paragraphs[0].runs[0]
                run.font.bold = True; run.font.color.rgb = text_hdr_color; run.font.size = Pt(7)

            for jdata in jadwal_list:
                cells = j_tbl.add_row().cells
                for idx, val in enumerate(jdata):
                    if idx < 4: cells[idx].text = str(val); cells[idx].paragraphs[0].runs[0].font.size = Pt(8); cells[idx].width = widths[idx]
                cells[4].text = ""; cells[4].width = widths[4]
        else: cell_r.add_paragraph("(Jadwal Tidak Diatur)")

        doc.add_paragraph("\n")
        if (index + 1) % 2 == 0: doc.add_page_break()

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

tab1, tab2, tab3 = st.tabs(["📂 1. Data Siswa & Foto", "📅 2. Atur Jadwal", "🖨️ 3. Download Word"])

# --- TAB 1: DATA SISWA & FOTO ---
with tab1:
    col_up_1, col_up_2 = st.columns([1, 1])
    
    with col_up_1:
        st.info("1. Upload Excel Data Siswa (.xlsx)")
        upl_excel = st.file_uploader("Pilih File Excel", type=['xlsx'])
        
    with col_up_2:
        st.info("2. Upload Foto Siswa (Bisa Banyak Sekaligus)")
        # UPLOAD BISA FILE BIASA (Multiple) ATAU ZIP
        upl_photos = st.file_uploader("Pilih Foto (.jpg/.png) atau ZIP", type=['jpg','png','jpeg','zip'], accept_multiple_files=True)
        
        c1, c2 = st.columns(2)
        upl_logo = c1.file_uploader("Logo Sekolah", type=['png','jpg'])
        upl_ttd = c2.file_uploader("Scan TTD", type=['png','jpg'])

        # PROSES FOTO LANGSUNG
        if upl_photos:
            # Cek apakah itu list file (Direct Upload) atau 1 File ZIP
            is_zip = len(upl_photos) == 1 and upl_photos[0].name.endswith('.zip')
            
            if is_zip:
                new_photos = extract_photos(upl_photos[0])
            else:
                new_photos = process_images(upl_photos, is_zip=False)
            
            # Update Session State
            st.session_state['photos'].update(new_photos)
            st.success(f"✅ {len(new_photos)} Foto berhasil ditambahkan ke sistem!")

        if upl_logo: st.session_state['logo_bytes'] = upl_logo.getvalue()
        if upl_ttd: st.session_state['ttd_bytes'] = upl_ttd.getvalue()

    # --- TABEL PREVIEW DATA (FULL) ---
    st.write("---")
    st.subheader("👁️ Preview Data & Cek Foto")
    
    if upl_excel:
        try:
            df = pd.read_excel(upl_excel)
            # Normalisasi Header
            df.columns = [str(c).strip().upper() for c in df.columns]
            st.session_state['df_siswa'] = df
            
            # Logic Status Foto
            if not df.empty:
                def cek_status_foto(row):
                    nisn = str(row.get('NISN','')).replace('.0','')
                    nis = str(row.get('NIS','')).replace('.0','')
                    
                    if nisn in st.session_state['photos']: return "✅ OKE"
                    if nis in st.session_state['photos']: return "✅ OKE"
                    return "❌ KOSONG"

                # Buat Dataframe baru untuk Display agar tidak merusak data asli
                df_display = df.copy()
                df_display.insert(0, "STATUS FOTO", df.apply(cek_status_foto, axis=1))
                
                # Tampilkan Full Data
                st.dataframe(
                    df_display, 
                    use_container_width=True, 
                    height=500, # Tinggi tabel agar muat banyak
                    hide_index=True
                )
                
                total_ok = df_display['STATUS FOTO'].value_counts().get("✅ OKE", 0)
                st.info(f"📊 Total Siswa: {len(df)} | Foto Lengkap: {total_ok} | Foto Kurang: {len(df)-total_ok}")

        except Exception as e:
            st.error(f"Error membaca Excel: {e}")
    else:
        st.warning("Silakan upload Excel data siswa terlebih dahulu.")

# --- TAB 2: JADWAL ---
with tab2:
    st.subheader("Pengaturan Jadwal")
    
    if upl_excel:
        try:
            df_jadwal = pd.read_excel(upl_excel, sheet_name="JADWAL")
            jadwal_excel = df_jadwal.iloc[:, :4].astype(str).values.tolist()
            if jadwal_excel and not st.session_state['jadwal_ujian']: # Load only if empty
                st.session_state['jadwal_ujian'] = jadwal_excel
                st.success(f"Jadwal dimuat dari Excel.")
        except: pass

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
        st.write("Preview Tabel Jadwal (+ Kolom Paraf Otomatis):")
        df_show = pd.DataFrame(st.session_state['jadwal_ujian'], columns=["HARI", "JAM", "WAKTU", "MAPEL"])
        st.table(df_show)
        if st.button("Hapus Semua Jadwal"):
            st.session_state['jadwal_ujian'] = []; st.rerun()

# --- TAB 3: CETAK ---
with tab3:
    st.subheader("Download Hasil")
    
    if 'df_siswa' in st.session_state:
        st.info(f"Siap mencetak {len(st.session_state['df_siswa'])} kartu ujian dalam format Word (.docx).")
        
        if st.button("🚀 GENERATE FILE WORD"):
            config = {'sekolah': in_sekolah, 'kepsek': in_kepsek, 'tanggal': in_tgl}
            logo_b = st.session_state.get('logo_bytes')
            ttd_b = st.session_state.get('ttd_bytes')
            photos = st.session_state.get('photos', {})
            
            try:
                doc = generate_word_doc(st.session_state['df_siswa'], config, template_sel, logo_b, ttd_b, st.session_state['jadwal_ujian'], photos)
                bio = io.BytesIO()
                doc.save(bio)
                st.success("Berhasil! Silakan download file di bawah ini.")
                st.download_button(
                    label="📥 DOWNLOAD KARTU (.DOCX)",
                    data=bio.getvalue(),
                    file_name=f"Kartu_Ujian_{template_sel}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Gagal generate: {e}")
    else:
        st.warning("Belum ada data siswa. Upload di Tab 1 dulu.")
