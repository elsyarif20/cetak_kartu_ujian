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
        background-color: white; padding: 12px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #e2e8f0;
        display: flex; align-items: center; justify-content: space-between;
    }
    .status-badge {
        padding: 4px 10px; border-radius: 4px; font-size: 0.9em; font-weight: bold;
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

    df.columns = [str(c).strip().upper() for c in df.columns]
    
    for index, row in df.iterrows():
        main_tbl = doc.add_table(rows=1, cols=2); main_tbl.style = 'Table Grid'; main_tbl.autofit = False
        
        # KIRI
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
        
        # LOGIC FOTO
        nis_key = str(row.get('NISN', '')).replace('.0','')
        # Coba cari NISN, kalau gagal cari NIS
        if nis_key not in photos: nis_key = str(row.get('NIS', '')).replace('.0','')
        
        if nis_key in photos:
            try:
                img_pil = photos[nis_key]
                img_byte = io.BytesIO(); img_pil.save(img_byte, format='JPEG')
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

        # KANAN
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

        doc.add_paragraph("\n"); 
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
        st.info("1. Upload Excel Data Siswa")
        upl_excel = st.file_uploader("Pilih File Excel", type=['xlsx'])
        
    with col_up_2:
        st.info("2. Upload Aset (Logo & TTD)")
        c1, c2 = st.columns(2)
        upl_logo = c1.file_uploader("Logo", type=['png','jpg'])
        upl_ttd = c2.file_uploader("TTD", type=['png','jpg'])
        if upl_logo: st.session_state['logo_bytes'] = upl_logo.getvalue()
        if upl_ttd: st.session_state['ttd_bytes'] = upl_ttd.getvalue()

    st.write("---")
    
    # === FITUR UPLOAD MASSAL (KUNCI UTAMA) ===
    st.markdown("### 📸 Upload Foto (Bisa Banyak Sekaligus)")
    st.caption("Tips: Blok 50 foto di folder Anda, lalu tarik ke kotak di bawah ini. Nama file harus NISN (contoh: `12345.jpg`).")
    
    # `accept_multiple_files=True` adalah kuncinya
    bulk_photos = st.file_uploader("Drop Foto Siswa Disini (Bisa Banyak)", type=['jpg','png','jpeg'], accept_multiple_files=True)
    
    if bulk_photos:
        count_new = 0
        for p in bulk_photos:
            # Ambil nama file tanpa ekstensi sebagai key (contoh: "12345.jpg" -> "12345")
            nis_key = p.name.rsplit('.', 1)[0]
            st.session_state['photos'][nis_key] = Image.open(p)
            count_new += 1
        
        st.success(f"✅ Berhasil menambahkan {count_new} foto baru! Cek tabel status di bawah.")

    st.write("---")
    
    # === TABEL DATA & STATUS ===
    if upl_excel:
        try:
            df = pd.read_excel(upl_excel)
            df.columns = [str(c).strip().upper() for c in df.columns]
            st.session_state['df_siswa'] = df
            
            st.subheader("👁️ Cek Status Foto Siswa")
            
            # HEADER TABEL MANUAL (Agar lebih rapi dari dataframe biasa)
            cols = st.columns([3, 2, 2, 3])
            cols[0].markdown("**Nama Siswa**")
            cols[1].markdown("**NISN**")
            cols[2].markdown("**Status Foto**")
            cols[3].markdown("**Upload Satuan (Jika Kurang)**")
            
            # LOOPING BARIS
            for idx, row in df.iterrows():
                nama = str(row.get('NAMA PESERTA', '-'))
                # Bersihkan NISN (hilangkan .0 jika float)
                nisn = str(row.get('NISN', '')).replace('.0','')
                
                # Cek apakah foto ada di session_state
                status_ok = False
                if nisn in st.session_state['photos']:
                    status_ok = True
                else:
                    # Cek fallback ke NIS jika NISN tidak ketemu
                    nis = str(row.get('NIS', '')).replace('.0','')
                    if nis in st.session_state['photos']:
                        status_ok = True
                
                # Tampilan Baris
                with st.container():
                    col_row = st.columns([3, 2, 2, 3])
                    
                    col_row[0].write(nama)
                    col_row[1].write(nisn)
                    
                    if status_ok:
                        col_row[2].markdown('<span class="status-badge ok">✅ ADA</span>', unsafe_allow_html=True)
                    else:
                        col_row[2].markdown('<span class="status-badge err">❌ KOSONG</span>', unsafe_allow_html=True)
                    
                    # Tombol Upload Per Baris (Satuan)
                    with col_row[3]:
                        # Key unik per baris agar tidak bentrok
                        up_single = st.file_uploader("Upload", type=['jpg','png'], key=f"u_{idx}_{nisn}", label_visibility="collapsed")
                        if up_single:
                            # Simpan foto dengan key NISN siswa ini
                            st.session_state['photos'][nisn] = Image.open(up_single)
                            st.rerun() # Refresh agar status berubah jadi Hijau

        except Exception as e:
            st.error(f"Error membaca Excel: {e}")
    else:
        st.info("Menunggu upload Excel...")

# --- TAB 2: JADWAL ---
with tab2:
    st.subheader("Pengaturan Jadwal")
    if upl_excel:
        try:
            df_jadwal = pd.read_excel(upl_excel, sheet_name="JADWAL")
            jadwal_excel = df_jadwal.iloc[:, :4].astype(str).values.tolist()
            if jadwal_excel and not st.session_state['jadwal_ujian']:
                st.session_state['jadwal_ujian'] = jadwal_excel
                st.success("Jadwal dimuat dari Excel.")
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

# --- TAB 3: CETAK ---
with tab3:
    st.subheader("Download Hasil")
    if 'df_siswa' in st.session_state:
        if st.button("🚀 GENERATE FILE WORD"):
            config = {'sekolah': in_sekolah, 'kepsek': in_kepsek, 'tanggal': in_tgl}
            logo_b = st.session_state.get('logo_bytes')
            ttd_b = st.session_state.get('ttd_bytes')
            photos = st.session_state.get('photos', {})
            try:
                doc = generate_word_doc(st.session_state['df_siswa'], config, template_sel, logo_b, ttd_b, st.session_state['jadwal_ujian'], photos)
                bio = io.BytesIO(); doc.save(bio)
                st.success("Selesai!"); st.download_button("📥 DOWNLOAD KARTU (.DOCX)", bio.getvalue(), f"Kartu_Ujian_{template_sel}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e: st.error(f"Gagal generate: {e}")
    else: st.warning("Upload data dulu di Tab 1.")
