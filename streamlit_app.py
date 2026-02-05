import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(page_title="Cetak Kartu Ujian - Al Ghozali", layout="wide")

# Inisialisasi Session State untuk menyimpan Jadwal sementara
if 'jadwal_ujian' not in st.session_state:
    st.session_state['jadwal_ujian'] = []

# ==========================================
# 2. FUNGSI LOGIC (GENERATOR GAMBAR)
# ==========================================
def extract_photos_from_zip(uploaded_zip):
    """Membaca file ZIP dan menjadikannya Dictionary {NamaFile: Gambar}"""
    photo_dict = {}
    if uploaded_zip is not None:
        with zipfile.ZipFile(uploaded_zip) as z:
            for filename in z.namelist():
                # Filter hanya file gambar, abaikan folder (/)
                if filename.lower().endswith(('.png', '.jpg', '.jpeg')) and not filename.startswith('__MACOSX'):
                    try:
                        with z.open(filename) as f:
                            img_data = f.read()
                            img = Image.open(io.BytesIO(img_data))
                            # Simpan dengan kunci nama file tanpa ekstensi (misal '12345')
                            # Kita ambil nama file paling dasar (basename)
                            base_name = filename.split('/')[-1]
                            name_key = base_name.rsplit('.', 1)[0]
                            photo_dict[name_key] = img
                    except:
                        continue
    return photo_dict

def generate_card_image(siswa, config, logo_img, ttd_img, photo_dict, jadwal_list):
    W, H = 1200, 500
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)
    
    # Load Fonts (Fallback ke default jika di server tidak ada Arial)
    try:
        f_header = ImageFont.truetype("arialbd.ttf", 20)
        f_norm = ImageFont.truetype("arial.ttf", 16)
        f_bold = ImageFont.truetype("arialbd.ttf", 16)
        f_small = ImageFont.truetype("arial.ttf", 13)
    except:
        f_header = f_norm = f_bold = f_small = ImageFont.load_default()

    # === BAGIAN KIRI: KARTU SISWA ===
    draw.rectangle([10, 10, 590, 490], outline="black", width=2)
    
    # Kop
    draw.text((300, 30), "YAYASAN PENDIDIKAN ISLAM AL-GHOZALI", font=f_small, fill="black", anchor="mm")
    draw.text((300, 55), config['sekolah'], font=f_header, fill="black", anchor="mm")
    draw.text((300, 80), config['alamat'], font=f_small, fill="black", anchor="mm")
    draw.line([20, 100, 580, 100], fill="black", width=2)
    
    draw.text((300, 130), "KARTU PESERTA UJIAN", font=f_header, fill="black", anchor="mm")

    # Logo
    if logo_img:
        try:
            logo_resized = logo_img.resize((80, 80))
            img.paste(logo_resized, (30, 20))
        except: pass

    # Data Siswa
    y = 170
    # Bersihkan NIS dari float (jika excel membaca angka sebagai 12345.0)
    nis_raw = str(siswa.get('NIS', '')).replace('.0', '').strip()
    
    fields = [
        ("No Peserta", str(siswa.get('NO PESERTA',''))),
        ("Nama", str(siswa.get('NAMA','')).upper()),
        ("NIS", nis_raw),
        ("Ruang", str(siswa.get('RUANG','')))
    ]
    
    for k, v in fields:
        draw.text((30, y), k, font=f_norm, fill="black")
        draw.text((150, y), ":", font=f_norm, fill="black")
        draw.text((160, y), v, font=f_bold, fill="black")
        y += 35

    # --- FOTO SISWA (MATCHING DARI ZIP) ---
    foto_x, foto_y = 30, 330
    draw.rectangle([foto_x, foto_y, foto_x+100, foto_y+130], outline="black")
    
    if nis_raw in photo_dict:
        try:
            foto_siswa = photo_dict[nis_raw].resize((100, 130))
            img.paste(foto_siswa, (foto_x, foto_y))
        except: pass
    else:
        draw.text((foto_x+25, foto_y+60), "FOTO", fill="gray", font=f_small)

    # TTD
    draw.text((350, 350), "Kepala Sekolah,", font=f_norm, fill="black")
    if ttd_img:
        try:
            ttd_resized = ttd_img.resize((120, 60))
            # Gunakan mask jika PNG transparan
            if ttd_resized.mode == 'RGBA':
                img.paste(ttd_resized, (350, 370), ttd_resized)
            else:
                img.paste(ttd_resized, (350, 370))
        except: pass
    draw.text((350, 430), config['kepsek'], font=f_bold, fill="black")

    # === BAGIAN KANAN: JADWAL ===
    draw.rectangle([600, 10, 1190, 490], outline="black", width=2)
    draw.rectangle([600, 10, 1190, 50], fill="#eee", outline="black")
    draw.text((900, 30), "JADWAL UJIAN", font=f_header, fill="black", anchor="mm")

    # Header Tabel
    ty = 60
    draw.text((610, ty), "HARI/TGL", font=f_bold, fill="black")
    draw.text((750, ty), "JAM", font=f_bold, fill="black")
    draw.text((850, ty), "MATA PELAJARAN", font=f_bold, fill="black")
    draw.line([600, ty+20, 1190, ty+20], fill="black", width=2)

    # Isi Jadwal
    curr_y = ty + 30
    for item in jadwal_list:
        # item = {'hari': ..., 'jam': ..., 'mapel': ...}
        draw.text((610, curr_y), str(item['hari']), font=f_small, fill="black")
        draw.text((750, curr_y), str(item['jam']), font=f_small, fill="black")
        draw.text((850, curr_y), str(item['mapel']), font=f_small, fill="black")
        draw.line([600, curr_y+15, 1190, curr_y+15], fill="#ddd", width=1)
        curr_y += 25

    return img

# ==========================================
# 3. USER INTERFACE (STREAMLIT)
# ==========================================
st.title("🖨️ Aplikasi Cetak Kartu Ujian V2 (Web)")
st.markdown("---")

# Gunakan Tabs agar rapi
tab1, tab2, tab3 = st.tabs(["1. Data & Aset", "2. Atur Jadwal", "3. Preview & Download"])

# --- TAB 1: DATA SISWA & ASET ---
with tab1:
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.subheader("Data Siswa")
        file_excel = st.file_uploader("Upload Excel Data Siswa (.xlsx)", type=['xlsx'])
        if file_excel:
            df = pd.read_excel(file_excel)
            # Normalisasi Header
            df.columns = [str(c).strip().upper() for c in df.columns]
            st.success(f"Berhasil load {len(df)} siswa.")
            
        st.markdown("---")
        st.subheader("Foto Siswa")
        st.info("Upload 1 file ZIP berisi semua foto siswa. Nama file foto harus sesuai NIS (Contoh: 12345.jpg)")
        file_zip = st.file_uploader("Upload ZIP Foto", type=['zip'])
        
        # Proses ZIP
        photo_dict = {}
        if file_zip:
            photo_dict = extract_photos_from_zip(file_zip)
            st.success(f"Berhasil mengekstrak {len(photo_dict)} foto dari ZIP.")

    with col_b:
        st.subheader("Identitas Sekolah")
        sekolah_nama = st.text_input("Nama Sekolah", "SMA ISLAM AL-GHOZALI")
        sekolah_alamat = st.text_area("Alamat", "Jl. Permata No. 19 Curug Gunungsindur")
        kepsek_nama = st.text_input("Kepala Sekolah", "Antoni Firdaus, SHI, M.Pd.")
        
        st.markdown("---")
        st.subheader("Logo & TTD")
        upl_logo = st.file_uploader("Logo Sekolah", type=['png', 'jpg'])
        upl_ttd = st.file_uploader("Scan TTD", type=['png', 'jpg'])
        
        # Load Images
        logo_img = Image.open(upl_logo) if upl_logo else None
        ttd_img = Image.open(upl_ttd) if upl_ttd else None

# --- TAB 2: JADWAL UJIAN ---
with tab2:
    st.subheader("Input Jadwal Ujian")
    
    c1, c2, c3, c4 = st.columns([2, 1, 2, 1])
    with c1: in_hari = st.text_input("Hari/Tanggal", "Senin, 18 Mei")
    with c2: in_jam = st.text_input("Jam", "07.30 - 09.00")
    with c3: in_mapel = st.text_input("Mata Pelajaran")
    with c4: 
        st.write("") # Spacer
        btn_add = st.button("➕ Tambah")

    if btn_add and in_mapel:
        st.session_state['jadwal_ujian'].append({
            'hari': in_hari, 'jam': in_jam, 'mapel': in_mapel
        })
        st.success("Jadwal ditambahkan!")

    # Tampilkan Tabel Jadwal
    if st.session_state['jadwal_ujian']:
        st.write("Daftar Jadwal:")
        st.table(st.session_state['jadwal_ujian'])
        
        if st.button("Hapus Semua Jadwal"):
            st.session_state['jadwal_ujian'] = []
            st.rerun()
    else:
        st.info("Belum ada jadwal diinput.")

# --- TAB 3: PREVIEW & CETAK ---
with tab3:
    st.subheader("Cetak Kartu")
    
    config = {'sekolah': sekolah_nama, 'alamat': sekolah_alamat, 'kepsek': kepsek_nama}
    
    if file_excel is not None:
        # Preview Tombol
        if st.button("👁️ Preview Kartu Pertama"):
            first_student = df.iloc[0].to_dict()
            img_preview = generate_card_image(
                first_student, config, logo_img, ttd_img, photo_dict, st.session_state['jadwal_ujian']
            )
            st.image(img_preview, caption=f"Preview: {first_student.get('NAMA')}")

        st.markdown("---")
        
        # Generator ZIP Download
        if st.button("⬇️ Download Semua Kartu (.ZIP)"):
            if df is None:
                st.error("Data Excel belum diupload!")
            else:
                # Membuat ZIP di Memory (RAM)
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    progress_bar = st.progress(0)
                    total = len(df)
                    
                    for i, row in df.iterrows():
                        student = row.to_dict()
                        # Generate Gambar
                        img = generate_card_image(
                            student, config, logo_img, ttd_img, photo_dict, st.session_state['jadwal_ujian']
                        )
                        
                        # Simpan gambar ke Bytes
                        img_byte_arr = io.BytesIO()
                        img.save(img_byte_arr, format='JPEG', quality=95)
                        
                        # Masukkan ke dalam ZIP
                        nama_file = f"{student.get('NAMA', 'Siswa').strip()}.jpg"
                        zf.writestr(nama_file, img_byte_arr.getvalue())
                        
                        progress_bar.progress((i + 1) / total)
                
                # Siapkan tombol download
                st.success("Selesai! Silakan download file di bawah.")
                st.download_button(
                    label="📥 Klik Disini untuk Download ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="Kartu_Ujian_Lengkap.zip",
                    mime="application/zip"
                )
    else:
        st.warning("Silakan upload data Excel di Tab 1 terlebih dahulu.")