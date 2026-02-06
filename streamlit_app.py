import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import io
from PIL import Image
import zipfile

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(page_title="Generator Kartu Ujian (Word)", page_icon="🖨️", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #f8fafc; }
    .main-header { 
        font-size: 2rem; font-weight: 800; color: #1e3a8a; margin-bottom: 10px; 
        border-bottom: 2px solid #cbd5e1; padding-bottom: 10px;
    }
    div.stButton > button {
        background-color: #15803d; color: white; border-radius: 6px; font-weight: bold; width: 100%;
    }
    .student-row {
        background-color: white; padding: 10px; border-radius: 8px; margin-bottom: 5px; border: 1px solid #e2e8f0;
        display: flex; align-items: center; 
    }
    .status-badge {
        padding: 4px 12px; border-radius: 20px; font-size: 0.85em; font-weight: bold; text-align: center; display: inline-block; width: 100px;
    }
    .ok { background-color: #dcfce7; color: #166534; border: 1px solid #86efac; }
    .err { background-color: #fee2e2; color: #991b1b; border: 1px solid #fca5a5; }
</style>
""", unsafe_allow_html=True)

if 'jadwal_ujian' not in st.session_state: st.session_state['jadwal_ujian'] = []
if 'photos' not in st.session_state: st.session_state['photos'] = {}

# ==========================================
# 2. LOGIC FUNCTIONS
# ==========================================
def clean_str(val):
    """Membersihkan format angka float (1.0 -> 1)"""
    return str(val).replace('.0', '').strip()

def compress_image(image_file):
    try:
        img = Image.open(image_file)
        if img.mode in ("RGBA", "P"): img = img.convert("RGB")
        
        # Resize jika lebar > 300px
        max_width = 300
        if img.width > max_width:
            w_percent = (max_width / float(img.width))
            h_size = int((float(img.height) * float(w_percent)))
            img = img.resize((max_width, h_size), Image.Resampling.LANCZOS)
        return img
    except: return None

def set_cell_color(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def generate_word_doc(df, config, template_type, logo_bytes, ttd_bytes, jadwal_list, photos):
    doc = Document()
    section = doc.sections[0]; section.orientation = 1
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

    # Bersihkan Header Kolom
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    for index, row in df.iterrows():
        main_tbl = doc.add_table(rows=1, cols=2); main_tbl.style = 'Table Grid'; main_tbl.autofit = False
        
        # --- KIRI: BIODATA ---
        cell_l = main_tbl.cell(0, 0); cell_l.width = Cm(13.5)
        p = cell_l.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        if logo_bytes: run = p.add_run(); run.add_picture(io.BytesIO(logo_bytes), width=Cm(1.8))
        
        p.add_run(f"\nYAYASAN PENDIDIKAN ISLAM AL-GHOZALI\n").font.size = Pt(9)
        r = p.add_run(f"{config['sekolah']}\n"); r.font.bold = True; r.font.size = Pt(11)
        p.add_run("_____________________________________________").font.size = Pt(6)
        
        p2 = cell_l.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p2.add_run(judul_kartu); r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = judul_color
        if template_type == "Kartu Hilang (Emergency)":
            r = p2.add_run("\n(DICETAK ULANG OLEH PANITIA)"); r.font.size = Pt(8); r.font.color.rgb = RGBColor(255,0,0)

        bio_tbl = cell_l.add_table(rows=4, cols=3); bio_tbl.autofit = False
        c_foto = bio_tbl.cell(0,0); c_foto.merge(bio_tbl.cell(3,0)); c_foto.width = Cm(3.0)
        
        # Logic Foto
        nis_key = clean_str(row.get('NISN', ''))
        if nis_key not in photos: nis_key = clean_str(row.get('NIS', ''))
        
        if nis_key in photos:
            try:
                img_pil = photos[nis_key]
                img_byte = io.BytesIO(); img_pil.save(img_byte, format='JPEG')
                p_f = c_foto.paragraphs[0]; p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_f.add_run().add_picture(io.BytesIO(img_byte.getvalue()), width=Cm(2.5), height=Cm(3.2))
            except: c_foto.text = "Error"
        else:
            c_foto.text = "FOTO\n3x4"; c_foto.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Data Siswa (Bersihkan .0 di sini juga)
        items = [
            ("No Peserta", clean_str(row.get('NOMOR PESERTA', '-'))), 
            ("Nama", str(row.get('NAMA PESERTA', '-'))), 
            ("NISN", clean_str(row.get('NISN', '-'))), 
            ("Ruang", clean_str(row.get('RUANG', '-')))
        ]
        
        for i, (lbl, val) in enumerate(items):
            bio_tbl.cell(i, 1).text = lbl; bio_tbl.cell(i, 1).paragraphs[0].runs[0].font.size = Pt(9)
            bio_tbl.cell(i, 2).text = ": " + val; bio_tbl.cell(i, 2).paragraphs[0].runs[0].font.size = Pt(9); bio_tbl.cell(i, 2).paragraphs[0].runs[0].font.bold = True

        pttd = cell_l.add_paragraph(); pttd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pttd.add_run(f"\n{config['tanggal']}\nKepala Sekolah,\n").font.size = Pt(9)
        if ttd_bytes: pttd.add_run().add_picture(io.BytesIO(ttd_bytes), width=Cm(2.0)); pttd.add_run("\n")
        else: pttd.add_run("\n\n\n")
        pttd.add_run(config['kepsek']).font.bold = True

        # --- KANAN: JADWAL ---
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
                # jdata = [HARI, JAM, WAKTU, MAPEL]
                # Bersihkan JAM KE (Index 1) dari .0
                hari = str(jdata[0])
                jam = clean_str(jdata[1]) # <-- FIX DISINI
                waktu = str(jdata[2])
                mapel = str(jdata[3])
                
                cleaned_row = [hari, jam, waktu, mapel]
                
                for idx, val in enumerate(cleaned_row):
                    cells[idx].text = val; cells[idx].paragraphs[0].runs[0].font.size = Pt(8); cells[idx].width = widths[idx]
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

# --- TAB 1: DATA ---
with tab1:
    col_up_1, col_up_2 = st.columns([1, 1])
    with col_up_1:
        st.info("1. Upload Excel Data Siswa")
        upl_excel = st.file_uploader("Pilih File Excel", type=['xlsx'])
    with col_up_2:
        st.info("2. Upload Aset (Logo & TTD)")
        c1, c2 = st.columns(2)
        upl_logo = c1.file_uploader("Logo", type=['png','jpg']); upl_ttd = c2.file_uploader("TTD", type=['png','jpg'])
        if upl_logo: st.session_state['logo_bytes'] = upl_logo.getvalue()
        if upl_ttd: st.session_state['ttd_bytes'] = upl_ttd.getvalue()

    st.write("---")
    st.markdown("### 📸 Upload Foto Massal (Auto-Compress)")
    bulk_photos = st.file_uploader("Drop Banyak Foto Disini (jpg/png/zip)", type=['jpg','png','jpeg','zip'], accept_multiple_files=True)
    
    if bulk_photos:
        count_new = 0
        is_zip = len(bulk_photos) == 1 and bulk_photos[0].name.endswith('.zip')
        
        if is_zip:
            with zipfile.ZipFile(bulk_photos[0]) as z:
                for f in z.namelist():
                    if f.lower().endswith(('.png','.jpg','.jpeg')) and not f.startswith('__'):
                        try:
                            # Nama file "12345.jpg" -> "12345"
                            fname = f.split('/')[-1].rsplit('.',1)[0]
                            img = compress_image(io.BytesIO(z.read(f)))
                            if img: st.session_state['photos'][fname] = img; count_new += 1
                        except: continue
        else:
            for p in bulk_photos:
                try:
                    fname = p.name.rsplit('.', 1)[0]
                    img = compress_image(p)
                    if img: st.session_state['photos'][fname] = img; count_new += 1
                except: continue
        
        if count_new > 0: st.success(f"✅ {count_new} Foto berhasil diproses!")

    if upl_excel:
        try:
            df = pd.read_excel(upl_excel)
            df.columns = [str(c).strip().upper() for c in df.columns]
            st.session_state['df_siswa'] = df
            
            st.write("---")
            st.subheader(f"👁️ Status Foto Siswa ({len(df)} Data)")
            
            cols = st.columns([3, 2, 2, 2])
            cols[0].markdown("**Nama Siswa**")
            cols[1].markdown("**NISN**")
            cols[2].markdown("**Status Foto**")
            cols[3].markdown("**Upload Satuan**")
            
            for idx, row in df.iterrows():
                nama = str(row.get('NAMA PESERTA', '-'))
                # BERSIHKAN NISN (12345.0 -> 12345)
                nisn = clean_str(row.get('NISN', ''))
                
                status_ok = False
                if nisn in st.session_state['photos']: status_ok = True
                else:
                    nis = clean_str(row.get('NIS', ''))
                    if nis in st.session_state['photos']: status_ok = True
                
                with st.container():
                    c_row = st.columns([3, 2, 2, 2])
                    c_row[0].write(nama); c_row[1].write(nisn)
                    if status_ok: c_row[2].markdown('<div class="status-badge ok">✅ ADA</div>', unsafe_allow_html=True)
                    else: c_row[2].markdown('<div class="status-badge err">❌ KOSONG</div>', unsafe_allow_html=True)
                    
                    with c_row[3]:
                        up_single = st.file_uploader("Upload", type=['jpg','png'], key=f"s_{idx}", label_visibility="collapsed")
                        if up_single:
                            img_ok = compress_image(up_single)
                            if img_ok:
                                # Paksa simpan pakai NISN dari Excel
                                st.session_state['photos'][nisn] = img_ok
                                st.rerun()
        except Exception as e: st.error(f"Error: {e}")

# --- TAB 2: JADWAL ---
with tab2:
    st.subheader("Pengaturan Jadwal")
    if upl_excel:
        try:
            df_jadwal = pd.read_excel(upl_excel, sheet_name="JADWAL")
            # BERSIHKAN DATA JADWAL DARI AWAL LOAD
            raw_jadwal = df_jadwal.iloc[:, :4].values.tolist()
            clean_jadwal = []
            for row in raw_jadwal:
                # [Hari, Jam, Waktu, Mapel] -> Jam dibersihkan dari .0
                clean_jadwal.append([str(row[0]), clean_str(row[1]), str(row[2]), str(row[3])])
            
            if clean_jadwal and not st.session_state['jadwal_ujian']:
                st.session_state['jadwal_ujian'] = clean_jadwal
                st.success("Jadwal dimuat dari Excel (Format Jam diperbaiki).")
        except: pass

    c1, c2, c3, c4 = st.columns(4)
    with c1: t_hari = st.text_input("Hari")
    with c2: t_jam = st.text_input("Jam Ke")
    with c3: t_waktu = st.text_input("Waktu")
    with c4: 
        t_mapel = st.text_input("Mapel")
        if st.button("➕ Tambah"):
            st.session_state['jadwal_ujian'].append([t_hari, t_jam, t_waktu, t_mapel]); st.rerun()

    if st.session_state['jadwal_ujian']:
        st.table(pd.DataFrame(st.session_state['jadwal_ujian'], columns=["HARI", "JAM", "WAKTU", "MAPEL"]))
        if st.button("Hapus Semua Jadwal"): st.session_state['jadwal_ujian'] = []; st.rerun()

# --- TAB 3: DOWNLOAD ---
with tab3:
    if 'df_siswa' in st.session_state:
        if st.button("🚀 GENERATE FILE WORD"):
            config = {'sekolah': in_sekolah, 'kepsek': in_kepsek, 'tanggal': in_tgl}
            try:
                doc = generate_word_doc(st.session_state['df_siswa'], config, template_sel, st.session_state.get('logo_bytes'), st.session_state.get('ttd_bytes'), st.session_state['jadwal_ujian'], st.session_state['photos'])
                bio = io.BytesIO(); doc.save(bio)
                st.success("Selesai!"); st.download_button("📥 DOWNLOAD KARTU (.DOCX)", bio.getvalue(), f"Kartu_Ujian.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e: st.error(f"Error: {e}")
