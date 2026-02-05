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
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    
    /* SIDEBAR */
    section[data-testid="stSidebar"] {
        background-color: white;
        border-right: 1px solid #e2e8f0;
    }
</style>
""", unsafe_allow_html=True)

if 'jadwal_ujian' not in st.session_state:
    st.session_state['jadwal_ujian'] = []

# ==========================================
# 2. LOGIC FUNCTIONS
# ==========================================
def load_fonts():
    """Load font dengan fallback aman untuk server Cloud"""
    try:
        return {
            "h1": ImageFont.truetype("arialbd.ttf", 28),
            "h2": ImageFont.truetype("arialbd.ttf", 22),
            "body_b": ImageFont.truetype("arialbd.ttf", 18),
            "body": ImageFont.truetype("arial.ttf", 18),
            "small": ImageFont.truetype("arial.ttf", 14)
        }
    except:
        # Fallback ke default jika Arial tidak ada di server
        d = ImageFont.load_default()
        return {"h1":d, "h2":d, "body_b":d, "body":d, "small":d}

def extract_photos(zip_file):
    """Ekstrak foto dari ZIP ke Memory"""
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
    """
    Fungsi Utama Menggambar Kartu.
    SUDAH DIPERBAIKI: Masalah resize gambar transparan (ValueError) sudah fix.
    """
    W, H = 1200, 500
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)
    fonts = load_fonts()
    
    # Data Siswa
    nis = str(siswa.get('NIS','')).replace('.0','').strip()
    data = {
        "nopes": str(siswa.get('NO PESERTA','')),
        "nama": str(siswa.get('NAMA','')).upper(),
        "nis": nis,
        "ruang": str(siswa.get('RUANG',''))
    }

    # Foto Siswa
    foto_img = None
    if nis in photos:
        foto_img = photos[nis].resize((110, 140))

    # QR Code
    qr = qrcode.make(nis).resize((90, 90))

    # --- TEMPLATE SELECTION ---

    # 1. TEMPLATE MODERN BLUE
    if template_type == "Modern Blue":
        # Header Block
        draw.rectangle([0, 0, 600, 110], fill="#0f172a") # Dark Slate
        
        if logo: 
            l_res = logo.resize((90,90))
            img.paste(l_res, (20, 10), l_res if l_res.mode=='RGBA' else None)
            
        draw.text((130, 30), config['sekolah'].upper(), font=fonts['h1'], fill="white")
        draw.text((130, 70), "KARTU PESERTA UJIAN", font=fonts['body'], fill="#94a3b8")

        # Body
        draw.rectangle([20, 130, 580, 300], outline="#0f172a", width=2)
        y = 140
        labels = [("NO PESERTA", data['nopes']), ("NAMA SISWA", data['nama']), ("NIS / NISN", data['nis']), ("RUANG UJIAN", data['ruang'])]
        for k, v in labels:
            draw.text((40, y), k, font=fonts['small'], fill="#64748b")
            draw.text((200, y), ": "+v, font=fonts['body_b'], fill="#0f172a")
            y += 40

        # Footer
        if foto_img: 
            img.paste(foto_img, (40, 330))
            draw.rectangle([40, 330, 150, 470], outline="#0f172a", width=2)
        else:
            draw.rectangle([40, 330, 150, 470], outline="gray")
        
        img.paste(qr, (480, 340))
        
        # TTD Fixed
        draw.text((250, 350), "Mengetahui,", font=fonts['small'], fill="black")
        if ttd: 
            t_res = ttd.resize((120,60))
            img.paste(t_res, (250, 390), t_res if t_res.mode=='RGBA' else None)
            
        draw.text((250, 450), config['kepsek'], font=fonts['body_b'], fill="black")
        
        # Jadwal Color
        header_color = "#0f172a"
        text_color = "white"

    # 2. TEMPLATE ISLAMIC GREEN
    elif template_type == "Islamic Green":
        draw.rectangle([5, 5, W-5, H-5], outline="#14532d", width=5) 
        draw.rectangle([15, 15, W-15, H-15], outline="#eab308", width=2) 
        
        if logo: 
            l_res = logo.resize((80,80))
            img.paste(l_res, (250, 25), l_res if l_res.mode=='RGBA' else None)
            
        draw.text((340, 35), config['sekolah'].upper(), font=fonts['h1'], fill="#14532d")
        draw.line([100, 100, 500, 100], fill="#eab308", width=3) 
        draw.text((300, 110), "KARTU UJIAN", font=fonts['h2'], fill="black", anchor="mm")
        
        if foto_img: 
            img.paste(foto_img, (40, 150))
            draw.rectangle([40, 150, 150, 290], outline="#14532d", width=2)
        else:
            draw.rectangle([40, 150, 150, 290], outline="gray")
            draw.text((70, 200), "FOTO", font=fonts['small'], fill="gray")

        dy = 150
        labels = [("No Peserta", data['nopes']), ("Nama", data['nama']), ("NIS", data['nis']), ("Ruang", data['ruang'])]
        for k, v in labels:
            draw.text((170, dy), k, font=fonts['body'], fill="black")
            draw.text((300, dy), ":", font=fonts['body'], fill="black")
            draw.text((315, dy), v, font=fonts['body_b'], fill="#14532d")
            dy += 35
            
        # TTD Fixed
        draw.text((400, 320), "Kepala Sekolah,", font=fonts['small'], fill="black")
        if ttd: 
            t_res = ttd.resize((120,60))
            img.paste(t_res, (400, 350), t_res if t_res.mode=='RGBA' else None)
        draw.text((400, 420), config['kepsek'], font=fonts['body_b'], fill="black")
        
        img.paste(qr, (50, 350)) 
        
        # Jadwal Color
        header_color = "#14532d"
        text_color = "white"

    # 3. TEMPLATE CLASSIC FORMAL (DEFAULT)
    else:
        draw.rectangle([10, 10, W-10, H-10], outline="black", width=3)
        if logo: 
            l_res = logo.resize((80,80))
            img.paste(l_res, (30, 20), l_res if l_res.mode=='RGBA' else None)
            
        draw.text((W//4, 30), "YAYASAN PENDIDIKAN", font=fonts['small'], fill="black", anchor="mm")
        draw.text((W//4, 55), config['sekolah'].upper(), font=fonts['h1'], fill="black", anchor="mm")
        draw.line([20, 100, 580, 100], fill="black", width=2)
        draw.text((W//4, 130), "KARTU PESERTA", font=fonts['h2'], fill="black", anchor="mm")
        
        y = 170
        labels = [("No Peserta", data['nopes']), ("Nama", data['nama']), ("NIS", data['nis']), ("Ruang", data['ruang'])]
        for k, v in labels:
            draw.text((30, y), k, font=fonts['body'], fill="black")
            draw.text((160, y), ": "+v, font=fonts['body_b'], fill="black")
            y += 35
            
        if foto_img: 
            img.paste(foto_img, (30, 320))
            draw.rectangle([30, 320, 140, 460], outline="black")
        else:
            draw.rectangle([30, 320, 140, 460], outline="black")
        
        img.paste(qr, (480, 20))
        
        # TTD Fixed
        draw.text((400, 350), "Kepala Sekolah,", font=fonts['small'], fill="black")
        if ttd: 
            t_res = ttd.resize((120,60))
            img.paste(t_res, (400, 370), t_res if t_res.mode=='RGBA' else None)
        draw.text((400, 440), config['kepsek'], font=fonts['body_b'], fill="black")
        
        header_color = "#f1f5f9"
        text_color = "black"

    # --- GAMBAR JADWAL (KANAN) ---
    draw.rectangle([600, 0, 1200, 500], fill="white")
    draw.rectangle([600, 0, 605, 500], fill="#ddd") # Lipatan

    draw.rectangle([620, 20, 1180, 70], fill=header_color)
    draw.text((900, 45), "JADWAL UJIAN", font=fonts['h2'], fill=text_color, anchor="mm")

    ty = 90
    cols = [630, 780, 900]
    draw.text((cols[0], ty), "HARI / TANGGAL", font=fonts['body_b'], fill="black")
    draw.text((cols[1], ty), "JAM", font=fonts['body_b'], fill="black")
    draw.text((cols[2], ty), "MATA PELAJARAN", font=fonts['body_b'], fill="black")
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
# 3. UI LAYOUT (SIDEBAR & MAIN)
# ==========================================
with st.sidebar:
    st.title("Admin Panel")
    st.caption("Al-Ghozali CardPro v2.1")
    st.markdown("---")
    
    st.markdown("### 🎨 Tampilan")
    template_option = st.selectbox("Template Desain", ["Islamic Green", "Modern Blue", "Classic Formal"])
    
    st.markdown("### 🏫 Identitas")
    conf_sekolah = st.text_input("Nama Sekolah", "SMA ISLAM AL-GHOZALI")
    conf_alamat = st.text_area("Alamat", "Jl. Permata No. 19 Curug Gunungsindur")
    conf_kepsek = st.text_input("Kepala Sekolah", "Antoni Firdaus, SHI, M.Pd.")

# HEADER
st.markdown('<div class="main-header">Al-Ghozali CardPro</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Sistem Cetak Kartu Ujian Profesional</div>', unsafe_allow_html=True)

# DASHBOARD METRICS
c1, c2, c3 = st.columns(3)
tot_siswa = len(st.session_state.get('data_siswa', [])) if 'data_siswa' in st.session_state else 0
tot_foto = len(st.session_state.get('photos', [])) if 'photos' in st.session_state else 0
c1.metric("Total Siswa", tot_siswa)
c2.metric("Foto Terupload", tot_foto)
c3.metric("Jadwal Ujian", len(st.session_state['jadwal_ujian']))

st.markdown("<br>", unsafe_allow_html=True)

# TABS
tab1, tab2, tab3 = st.tabs(["📂 Upload Data", "📅 Atur Jadwal", "🖨️ Cetak Kartu"])

# --- TAB 1: DATA ---
with tab1:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. Data Siswa (Excel)")
        upl_excel = st.file_uploader("Upload Excel (.xlsx)", type=['xlsx'])
        if upl_excel:
            df = pd.read_excel(upl_excel)
            df.columns = [str(c).strip().upper() for c in df.columns]
            st.session_state['data_siswa'] = df
            st.success(f"✅ {len(df)} Siswa dimuat.")
    
    with col2:
        st.subheader("2. Aset Gambar")
        upl_zip = st.file_uploader("Foto Siswa (.zip)", type=['zip'], help="Nama file harus sesuai NIS")
        c_l, c_t = st.columns(2)
        upl_logo = c_l.file_uploader("Logo Sekolah", type=['png','jpg'])
        upl_ttd = c_t.file_uploader("TTD Kepsek", type=['png','jpg'])
        
        if upl_zip:
            st.session_state['photos'] = extract_photos(upl_zip)
            st.success(f"✅ {len(st.session_state['photos'])} Foto diekstrak.")
        
        if upl_logo: st.session_state['logo'] = Image.open(upl_logo)
        if upl_ttd: st.session_state['ttd'] = Image.open(upl_ttd)
    
    st.markdown('</div>', unsafe_allow_html=True)

# --- TAB 2: JADWAL ---
with tab2:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    st.subheader("Input Jadwal Ujian")
    c_in, c_tbl = st.columns([1, 2])
    
    with c_in:
        in_hari = st.text_input("Hari", placeholder="Senin, 10 Mei")
        in_jam = st.text_input("Jam", placeholder="07.30 - 09.00")
        in_mapel = st.text_input("Mapel", placeholder="Matematika")
        
        if st.button("➕ Tambah"):
            if in_mapel:
                st.session_state['jadwal_ujian'].append({'hari':in_hari, 'jam':in_jam, 'mapel':in_mapel})
                st.success("Ditambahkan")
        
        if st.button("🗑️ Reset", type="primary"):
            st.session_state['jadwal_ujian'] = []
            st.rerun()
            
    with c_tbl:
        if st.session_state['jadwal_ujian']:
            st.table(pd.DataFrame(st.session_state['jadwal_ujian']))
        else:
            st.info("Belum ada jadwal.")
    st.markdown('</div>', unsafe_allow_html=True)

# --- TAB 3: CETAK ---
with tab3:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    st.subheader("Preview & Download")
    
    if 'data_siswa' in st.session_state:
        # Config Data
        config = {'sekolah':conf_sekolah, 'alamat':conf_alamat, 'kepsek':conf_kepsek}
        logo = st.session_state.get('logo')
        ttd = st.session_state.get('ttd')
        photos = st.session_state.get('photos', {})
        
        # Preview
        col_p, col_d = st.columns([2, 1])
        with col_p:
            st.markdown("##### Live Preview")
            first_row = st.session_state['data_siswa'].iloc[0].to_dict()
            img = draw_template(template_option, first_row, config, logo, ttd, photos, st.session_state['jadwal_ujian'])
            st.image(img, caption=f"Desain: {template_option}", use_container_width=True)
            
        with col_d:
            st.markdown("##### Download")
            st.write("Generate semua kartu siswa dalam satu file ZIP.")
            
            if st.button("🚀 GENERATE ZIP FILE", type="primary"):
                with st.status("Sedang memproses...", expanded=True):
                    mem_zip = io.BytesIO()
                    with zipfile.ZipFile(mem_zip, "w") as zf:
                        df = st.session_state['data_siswa']
                        prog = st.progress(0)
                        for i, row in df.iterrows():
                            s_data = row.to_dict()
                            img = draw_template(template_option, s_data, config, logo, ttd, photos, st.session_state['jadwal_ujian'])
                            
                            img_byte = io.BytesIO()
                            img.save(img_byte, format="JPEG", quality=95)
                            fname = f"{s_data.get('NAMA','Siswa').strip()}.jpg"
                            zf.writestr(fname, img_byte.getvalue())
                            prog.progress((i+1)/len(df))
                            
                st.success("Selesai!")
                st.download_button("📥 DOWNLOAD ZIP", mem_zip.getvalue(), "Kartu_Ujian.zip", "application/zip")
                
    else:
        st.warning("Upload Data Excel dulu di Tab 1.")
    st.markdown('</div>', unsafe_allow_html=True)
