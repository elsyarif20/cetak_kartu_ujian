import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile
import qrcode

# ==========================================
# 1. KONFIGURASI HALAMAN & CSS MODERN
# ==========================================
st.set_page_config(
    page_title="Al-Ghozali CardPro", 
    page_icon="🎓", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- INJECT CSS MODERN ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');

    /* BASE STYLE */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        color: #1e293b;
    }
    .stApp {
        background-color: #f1f5f9; /* Slate-100 */
    }

    /* CARDS (KOTAK KONTEN) */
    .css-card {
        background-color: white;
        padding: 2rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        margin-bottom: 1.5rem;
        border: 1px solid #e2e8f0;
    }

    /* HEADER STYLE */
    .main-header {
        font-size: 2.5rem;
        font-weight: 800;
        background: -webkit-linear-gradient(45deg, #0f172a, #334155);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        color: #64748b;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }

    /* CUSTOM BUTTONS */
    div.stButton > button {
        width: 100%;
        background-color: #2563eb;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.6rem 1rem;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    div.stButton > button:hover {
        background-color: #1d4ed8;
        box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.3);
        transform: translateY(-2px);
    }
    
    /* SIDEBAR */
    section[data-testid="stSidebar"] {
        background-color: white;
        border-right: 1px solid #e2e8f0;
    }

    /* TABS */
    .stTabs [data-baseweb="tab-list"] {
        gap: 20px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: white;
        border-radius: 8px 8px 0 0;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: white;
        border-bottom: 3px solid #2563eb;
        color: #2563eb;
        font-weight: bold;
    }

</style>
""", unsafe_allow_html=True)

if 'jadwal_ujian' not in st.session_state:
    st.session_state['jadwal_ujian'] = []

# ==========================================
# 2. LOGIC FUNCTIONS (Sama seperti sebelumnya)
# ==========================================
# (Saya ringkas bagian font loading agar kode tidak terlalu panjang, logic tetap sama)
def load_fonts():
    try:
        return {
            "h1": ImageFont.truetype("arialbd.ttf", 28),
            "h2": ImageFont.truetype("arialbd.ttf", 22),
            "body_b": ImageFont.truetype("arialbd.ttf", 18),
            "body": ImageFont.truetype("arial.ttf", 18),
            "small": ImageFont.truetype("arial.ttf", 14)
        }
    except:
        d = ImageFont.load_default()
        return {"h1":d, "h2":d, "body_b":d, "body":d, "small":d}

def extract_photos(zip_file):
    photo_dict = {}
    if zip_file:
        with zipfile.ZipFile(zip_file) as z:
            for f in z.namelist():
                if f.lower().endswith(('.png','.jpg','.jpeg')) and not f.startswith('__'):
                    try:
                        base = f.split('/')[-1].rsplit('.',1)[0]
                        photo_dict[base] = Image.open(io.BytesIO(z.read(f)))
                    except: continue
    return photo_dict

def draw_template(template_type, siswa, config, logo, ttd, photos, jadwal):
    # --- LOGIC GAMBAR (SAMA PERSIS SEPERTI SEBELUMNYA) ---
    # Saya gunakan template Islamic Green & Modern Blue sebagai default logic
    W, H = 1200, 500
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)
    fonts = load_fonts()
    
    nis = str(siswa.get('NIS','')).replace('.0','').strip()
    data = {"nopes": str(siswa.get('NO PESERTA','')), "nama": str(siswa.get('NAMA','')).upper(), "nis": nis, "ruang": str(siswa.get('RUANG',''))}
    foto_img = photos.get(nis).resize((110, 140)) if nis in photos else None
    qr = qrcode.make(nis).resize((90, 90))

    # --- STYLE SELECTION ---
    colors = {"header": "#333", "text": "black", "accent": "black"}
    
    if template_type == "Modern Blue":
        draw.rectangle([0, 0, 600, 110], fill="#0f172a") # Dark Slate
        if logo: img.paste(logo.resize((90,90)), (20, 10))
        draw.text((130, 30), config['sekolah'].upper(), font=fonts['h1'], fill="white")
        draw.text((130, 70), "KARTU PESERTA UJIAN", font=fonts['body'], fill="#94a3b8")
        draw.rectangle([20, 130, 580, 300], outline="#0f172a", width=2)
        y=140
        for k,v in [("NO PESERTA", data['nopes']), ("NAMA", data['nama']), ("NIS", data['nis']), ("RUANG", data['ruang'])]:
            draw.text((40,y), k, font=fonts['small'], fill="#64748b")
            draw.text((200,y), ": "+v, font=fonts['body_b'], fill="#0f172a")
            y+=40
        if foto_img: img.paste(foto_img, (40, 330)); draw.rectangle([40, 330, 150, 470], outline="#0f172a", width=2)
        else: draw.rectangle([40, 330, 150, 470], outline="gray")
        img.paste(qr, (480, 340))
        draw.text((250, 350), "Mengetahui,", font=fonts['small'], fill="black")
        if ttd: img.paste(ttd.resize((120,60)), (250, 380), ttd if ttd.mode=='RGBA' else None)
        draw.text((250, 450), config['kepsek'], font=fonts['body_b'], fill="black")
        colors = {"header": "#0f172a", "text": "white", "accent": "#0f172a"}

    elif template_type == "Islamic Green":
        draw.rectangle([5,5,W-5,H-5], outline="#14532d", width=5)
        draw.rectangle([15,15,W-15,H-15], outline="#eab308", width=2)
        if logo: img.paste(logo.resize((80,80)), (250, 25))
        draw.text((340, 35), config['sekolah'].upper(), font=fonts['h1'], fill="#14532d")
        draw.line([100, 100, 500, 100], fill="#eab308", width=3)
        draw.text((300, 110), "KARTU UJIAN", font=fonts['h2'], fill="black", anchor="mm")
        if foto_img: img.paste(foto_img, (40, 150)); draw.rectangle([40, 150, 150, 290], outline="#14532d", width=2)
        else: draw.rectangle([40, 150, 150, 290], outline="gray")
        dy=150
        for k,v in [("No Peserta", data['nopes']), ("Nama", data['nama']), ("NIS", data['nis']), ("Ruang", data['ruang'])]:
            draw.text((170,dy), k, font=fonts['body'], fill="black"); draw.text((315,dy), ": "+v, font=fonts['body_b'], fill="#14532d"); dy+=35
        draw.text((400, 320), "Kepala Sekolah,", font=fonts['small'], fill="black")
        if ttd: img.paste(ttd.resize((120,60)), (400, 350), ttd if ttd.mode=='RGBA' else None)
        draw.text((400, 420), config['kepsek'], font=fonts['body_b'], fill="black")
        img.paste(qr, (50, 350))
        colors = {"header": "#14532d", "text": "white", "accent": "#14532d"}
        
    else: # Default Classic
        draw.rectangle([10,10,W-10,H-10], outline="black", width=3)
        if logo: img.paste(logo.resize((80,80)), (30, 20))
        draw.text((W//4, 30), "YAYASAN PENDIDIKAN", font=fonts['small'], fill="black", anchor="mm")
        draw.text((W//4, 55), config['sekolah'].upper(), font=fonts['h1'], fill="black", anchor="mm")
        draw.line([20, 100, 580, 100], fill="black", width=2)
        draw.text((W//4, 130), "KARTU PESERTA", font=fonts['h2'], fill="black", anchor="mm")
        y=170
        for k,v in [("No Peserta", data['nopes']), ("Nama", data['nama']), ("NIS", data['nis']), ("Ruang", data['ruang'])]:
            draw.text((30,y), k, font=fonts['body'], fill="black"); draw.text((160,y), ": "+v, font=fonts['body_b'], fill="black"); y+=35
        if foto_img: img.paste(foto_img, (30, 320)); draw.rectangle([30, 320, 140, 460], outline="black")
        else: draw.rectangle([30, 320, 140, 460], outline="black")
        img.paste(qr, (480, 20))
        draw.text((400, 350), "Kepala Sekolah,", font=fonts['small'], fill="black")
        if ttd: img.paste(ttd.resize((120,60)), (400, 370), ttd if ttd.mode=='RGBA' else None)
        draw.text((400, 440), config['kepsek'], font=fonts['body_b'], fill="black")

    # JADWAL (KANAN)
    draw.rectangle([600, 0, 1200, 500], fill="white"); draw.rectangle([600, 0, 605, 500], fill="#f1f5f9")
    draw.rectangle([620, 20, 1180, 70], fill=colors["header"])
    draw.text((900, 45), "JADWAL UJIAN", font=fonts['h2'], fill=colors["text"], anchor="mm")
    ty = 90; cols = [630, 780, 900]
    draw.text((cols[0], ty), "HARI", font=fonts['body_b'], fill="black"); draw.text((cols[1], ty), "JAM", font=fonts['body_b'], fill="black"); draw.text((cols[2], ty), "MAPEL", font=fonts['body_b'], fill="black")
    draw.line([620, ty+25, 1180, ty+25], fill="black", width=2)
    curr_y = ty + 35
    for item in jadwal:
        draw.text((cols[0], curr_y), str(item['hari']), font=fonts['small'], fill="black")
        draw.text((cols[1], curr_y), str(item['jam']), font=fonts['small'], fill="black")
        draw.text((cols[2], curr_y), str(item['mapel']), font=fonts['small'], fill="black")
        draw.line([620, curr_y+20, 1180, curr_y+20], fill="#f1f5f9", width=1)
        curr_y += 25
    return img

# ==========================================
# 3. UI LAYOUT (SIDEBAR)
# ==========================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=50) # Placeholder Icon
    st.title("Admin Panel")
    st.caption("Al-Ghozali Card Generator v2.0")
    
    st.markdown("---")
    
    st.markdown("### 🎨 Tampilan Kartu")
    template_option = st.selectbox("Pilih Template", ["Islamic Green", "Modern Blue", "Classic Formal"])
    
    st.markdown("### 🏫 Identitas Sekolah")
    conf_sekolah = st.text_input("Nama Sekolah", "SMA ISLAM AL-GHOZALI")
    conf_alamat = st.text_area("Alamat", "Jl. Permata No. 19 Curug Gunungsindur")
    conf_kepsek = st.text_input("Kepala Sekolah", "Antoni Firdaus, SHI, M.Pd.")
    
    st.markdown("---")
    st.info("💡 **Tips:** Gunakan template 'Islamic Green' untuk branding Al-Ghozali.")

# ==========================================
# 4. MAIN CONTENT
# ==========================================
st.markdown('<div class="main-header">Al-Ghozali CardPro</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Sistem Generator Kartu Ujian Terpadu & Profesional</div>', unsafe_allow_html=True)

# --- DASHBOARD METRICS ---
if 'data_siswa' in st.session_state and st.session_state['data_siswa'] is not None:
    tot_siswa = len(st.session_state['data_siswa'])
else:
    tot_siswa = 0
    
if 'photos' in st.session_state:
    tot_foto = len(st.session_state['photos'])
else:
    tot_foto = 0

col_m1, col_m2, col_m3 = st.columns(3)
with col_m1:
    st.metric("Total Siswa", f"{tot_siswa} Siswa", delta="Data Excel")
with col_m2:
    st.metric("Total Foto", f"{tot_foto} File", delta="Data ZIP")
with col_m3:
    st.metric("Total Jadwal", f"{len(st.session_state['jadwal_ujian'])} Mapel")

st.markdown("<br>", unsafe_allow_html=True)

# --- TABS NAVIGATION ---
tab_data, tab_jadwal, tab_cetak = st.tabs(["📂 1. Upload Data", "📅 2. Atur Jadwal", "🖨️ 3. Cetak Kartu"])

# --- TAB 1: DATA ---
with tab_data:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    
    with c1:
        st.subheader("📄 Data Akademik")
        upl_excel = st.file_uploader("Upload File Excel (.xlsx)", type=['xlsx'])
        if upl_excel:
            df = pd.read_excel(upl_excel)
            df.columns = [str(c).strip().upper() for c in df.columns]
            st.session_state['data_siswa'] = df
            st.success(f"✅ Berhasil memuat {len(df)} data siswa.")
            with st.expander("Lihat Data Tabel"):
                st.dataframe(df.head())

    with c2:
        st.subheader("🖼️ Aset Gambar")
        upl_zip = st.file_uploader("Foto Siswa (ZIP)", type=['zip'], help="Nama file = NIS")
        c_logo, c_ttd = st.columns(2)
        with c_logo: upl_logo = st.file_uploader("Logo (PNG)", type=['png','jpg'])
        with c_ttd: upl_ttd = st.file_uploader("TTD (PNG)", type=['png','jpg'])
        
        if upl_zip:
            st.session_state['photos'] = extract_photos(upl_zip)
            st.success(f"✅ {len(st.session_state['photos'])} Foto diekstrak.")
        
        # Simpan Logo & TTD ke session state biar gak ilang pas refresh
        if upl_logo: st.session_state['logo'] = Image.open(upl_logo)
        if upl_ttd: st.session_state['ttd'] = Image.open(upl_ttd)
            
    st.markdown('</div>', unsafe_allow_html=True)

# --- TAB 2: JADWAL ---
with tab_jadwal:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    st.subheader("📅 Input Jadwal Pelajaran")
    
    col_input, col_table = st.columns([1, 2])
    
    with col_input:
        st.markdown("#### Tambah Mapel")
        in_hari = st.text_input("Hari / Tanggal", placeholder="Senin, 10 Mei 2026")
        in_jam = st.text_input("Waktu", placeholder="07.30 - 09.00")
        in_mapel = st.text_input("Mata Pelajaran", placeholder="Matematika Wajib")
        
        if st.button("➕ Tambah Jadwal"):
            if in_hari and in_mapel:
                st.session_state['jadwal_ujian'].append({'hari':in_hari, 'jam':in_jam, 'mapel':in_mapel})
                st.success("Ditambahkan!")
            else:
                st.error("Isi data dengan lengkap.")
        
        if st.button("🗑️ Reset Jadwal", type="primary"):
            st.session_state['jadwal_ujian'] = []
            st.rerun()

    with col_table:
        st.markdown("#### Preview Tabel")
        if st.session_state['jadwal_ujian']:
            st.table(pd.DataFrame(st.session_state['jadwal_ujian']))
        else:
            st.info("Belum ada jadwal yang diinput.")
            
    st.markdown('</div>', unsafe_allow_html=True)

# --- TAB 3: CETAK ---
with tab_cetak:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    st.subheader("🖨️ Preview & Eksekusi")
    
    if 'data_siswa' in st.session_state and st.session_state['data_siswa'] is not None:
        
        col_preview, col_action = st.columns([2, 1])
        
        config = {'sekolah':conf_sekolah, 'alamat':conf_alamat, 'kepsek':conf_kepsek}
        logo = st.session_state.get('logo')
        ttd = st.session_state.get('ttd')
        photos = st.session_state.get('photos', {})
        
        with col_preview:
            st.markdown("##### Live Preview (Siswa Pertama)")
            first_row = st.session_state['data_siswa'].iloc[0].to_dict()
            img_prev = draw_template(template_option, first_row, config, logo, ttd, photos, st.session_state['jadwal_ujian'])
            st.image(img_prev, caption=f"Preview Template: {template_option}", use_container_width=True)

        with col_action:
            st.markdown("##### Download")
            st.write("Siap untuk men-generate seluruh kartu?")
            
            if st.button("🚀 GENERATE SEMUA KARTU"):
                with st.status("Sedang memproses...", expanded=True) as status:
                    st.write("Membaca data siswa...")
                    mem_zip = io.BytesIO()
                    with zipfile.ZipFile(mem_zip, "w") as zf:
                        df = st.session_state['data_siswa']
                        total = len(df)
                        progress_bar = st.progress(0)
                        
                        for i, row in df.iterrows():
                            s_data = row.to_dict()
                            st.write(f"Memproses: {s_data.get('NAMA')}")
                            img = draw_template(template_option, s_data, config, logo, ttd, photos, st.session_state['jadwal_ujian'])
                            
                            img_byte = io.BytesIO()
                            img.save(img_byte, format="JPEG", quality=95)
                            fname = f"{s_data.get('NAMA','Siswa').strip()}.jpg"
                            zf.writestr(fname, img_byte.getvalue())
                            progress_bar.progress((i + 1) / total)
                            
                    status.update(label="Selesai!", state="complete", expanded=False)
                
                st.balloons()
                st.success("Kartu siap didownload!")
                st.download_button(
                    label="📥 DOWNLOAD ZIP FILE",
                    data=mem_zip.getvalue(),
                    file_name="Kartu_Ujian_AlGhozali_Pro.zip",
                    mime="application/zip",
                    type="primary"
                )
    else:
        st.warning("⚠️ Data siswa belum ada. Silakan upload Excel di Tab 1.")
        
    st.markdown('</div>', unsafe_allow_html=True)
