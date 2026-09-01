import os
import io
import zipfile
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Sistem Berkas Penyidikan", layout="wide")

# ==========================================
# 1. DATABASE AKUN PENGGUNA & ADMIN
# ==========================================
DATA_AKUN = {
    "user": {
        "password": "user123",
        "role": "pengguna",
        "nama": "Petugas Reskrim"
    },
    "admin": {
        "password": "admin123",
        "role": "admin",
        "nama": "Administrator Pemilik"
    }
}

# ==========================================
# 2. PENGELOLAAN TEMPLATE OTOMATIS (AUTO-SCAN)
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

if not os.path.exists(TEMPLATE_DIR):
    os.makedirs(TEMPLATE_DIR)

def load_templates():
    templates = {}
    if os.path.exists(TEMPLATE_DIR):
        for file in os.listdir(TEMPLATE_DIR):
            # Abaikan file temp Word dan hanya ambil .docx
            if file.endswith(".docx") and not file.startswith("~$"):
                # Buat nama tampilan rapi berdasarkan nama filenya
                label = file.replace(".docx", "").replace("_", " ").upper()
                templates[label] = file
    return templates

mapping_template = load_templates()

# ==========================================
# 3. KONTROL SESSION LOGIN
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""
if "role" not in st.session_state:
    st.session_state["role"] = ""
if "nama_user" not in st.session_state:
    st.session_state["nama_user"] = ""

# ==========================================
# 4. HALAMAN LOGIN (JIKA BELUM LOGIN)
# ==========================================
if not st.session_state["logged_in"]:
    st.title("🔒 Login Sistem Berkas Penyidikan")
    st.caption("Silakan masukkan Username dan Password Anda.")

    col_a, col_b, col_c = st.columns([1, 2, 1])
    with col_b:
        with st.form("form_login"):
            username_input = st.text_input("Username")
            password_input = st.text_input("Password", type="password")
            btn_login = st.form_submit_button("🔑 Masuk / Login", use_container_width=True)

            if btn_login:
                user_key = username_input.strip().lower()
                if user_key in DATA_AKUN and DATA_AKUN[user_key]["password"] == password_input:
                    st.session_state["logged_in"] = True
                    st.session_state["username"] = user_key
                    st.session_state["role"] = DATA_AKUN[user_key]["role"]
                    st.session_state["nama_user"] = DATA_AKUN[user_key]["nama"]
                    st.rerun()
                else:
                    st.error("❌ Username atau Password salah!")
    st.stop()

# ==========================================
# 5. SIDEBAR / NAVIGASI SETELAH LOGIN
# ==========================================
st.sidebar.title("👤 Profil Pengguna")
st.sidebar.write(f"**Nama:** {st.session_state['nama_user']}")
st.sidebar.write(f"**Role:** `{st.session_state['role'].upper()}`")

if st.sidebar.button("🚪 Keluar / Logout", use_container_width=True):
    for key in ["logged_in", "username", "role", "nama_user"]:
        st.session_state[key] = ""
    st.session_state["logged_in"] = False
    st.rerun()

st.sidebar.write("---")
st.sidebar.title("📌 Menu Utama")

if st.session_state["role"] == "admin":
    pilihan_halaman = st.sidebar.radio("Pilih Halaman:", ["👤 Pengguna (Cetak Surat)", "⚙️ Admin (Kelola Template)"])
else:
    pilihan_halaman = st.sidebar.radio("Pilih Halaman:", ["👤 Pengguna (Cetak Surat)"])

# Hubungkan ke Google Sheets
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = conn.read(ttl=0)
    df = df.dropna(how="all")
except Exception as e:
    st.error(f"Gagal terhubung ke Google Sheets: {e}")
    st.stop()

# ==========================================
# HALAMAN 1: PENGGUNA (INPUT & CETAK SURAT)
# ==========================================
if pilihan_halaman == "👤 Pengguna (Cetak Surat)":
    st.title("📂 Sistem Otomasi Berkas Penyidikan")
    
    mode = st.radio(
        "Pilih Mode Operasi:", 
        ["➕ Input Berkas Baru (Form)", "🖨️ Cetak dari Database yang Sudah Ada"], 
        horizontal=True
    )
    st.write("---")

    # --- FORM INPUT ---
    if mode == "➕ Input Berkas Baru (Form)":
        st.subheader("📝 Formulir Input Data Penyidikan")
        with st.form("form_input_kasus", clear_on_submit=False):
            col1, col2 = st.columns(2)
            with col1:
                kop_polsek = st.text_input("Kop Polsek", value="POLSEK URBAN BANDAR LAMPUNG")
                nomor_sp_sidik = st.text_input("Nomor SP.Sidik", placeholder="SP.Sidik/12/VIII/2026/Reskrim")
                lp = st.text_input("Nomor LP")
                perkara = st.text_input("Perkara / Tindak Pidana")
                pasal = st.text_input("Pasal")
                hari_tanggal_e = st.text_input("Hari & Tanggal Kejadian")
                jam = st.text_input("Jam Kejadian")

            with col2:
                tanggal_bulan_ = st.text_input("Tanggal Surat Dibuat")
                nama_petugas = st.text_input("Nama Petugas / Pelapor")
                nrp_petugas = st.text_input("NRP Petugas")
                nama_penyidik = st.text_input("Nama Penyidik / Kanit")
                nrp_penyidik = st.text_input("NRP Penyidik")
                list_nama_petugas = st.text_area("List Nama Tim Petugas")

            submit_form = st.form_submit_button("💾 Simpan Data ke Google Sheets")

        if submit_form:
            if not nomor_sp_sidik.strip():
                st.warning("⚠️ Nomor SP.Sidik wajib diisi!")
            else:
                data_baru = {
                    "Kop_Polsek": kop_polsek, "Nomor_Sp_sidik": nomor_sp_sidik, "LP": lp,
                    "Perkara": perkara, "Pasal": pasal, "Hari_tanggal_E": hari_tanggal_e,
                    "Jam": jam, "tanggal_Bulan_": tanggal_bulan_, "Nama_Petugas": nama_petugas,
                    "Nrp_Petugas": nrp_petugas, "Nama_Penyidik": nama_penyidik,
                    "Nrp_Penyidik": nrp_penyidik, "List_Nama_petugas": list_nama_petugas
                }
                try:
                    df_baru = pd.DataFrame([data_baru])
                    df_updated = pd.concat([df, df_baru], ignore_index=True)
                    conn.update(data=df_updated)
                    st.success("✅ Data berhasil disimpan ke database!")
                    st.session_state["data_aktif"] = data_baru
                    st.session_state["nomor_aktif"] = nomor_sp_sidik
                except Exception as e:
                    st.error(f"Gagal menyimpan data: {e}")

    # --- CETAK SURAT ---
    st.write("---")
    st.subheader("📋 Ceklis & Cetak Surat")
    data_terpilih, nomor_kasus = None, ""

    if mode == "🖨️ Cetak dari Database yang Sudah Ada":
        df_bersih = df[~df["Nomor_Sp_sidik"].isna() & (df["Nomor_Sp_sidik"].astype(str).str.strip() != "")]
        if df_bersih.empty:
            st.info("ℹ️ Belum ada data di Google Sheets.")
        else:
            nomor_kasus = st.selectbox("Pilih Nomor SP.Sidik dari Database:", df_bersih["Nomor_Sp_sidik"].tolist())
            data_terpilih = df_bersih[df_bersih["Nomor_Sp_sidik"] == nomor_kasus].iloc[0].to_dict()
    else:
        if "data_aktif" in st.session_state:
            data_terpilih, nomor_kasus = st.session_state["data_aktif"], st.session_state["nomor_aktif"]
            st.info(f"📍 Menyiapkan dokumen untuk: **{nomor_kasus}**")

    if data_terpilih:
        data_clean = {k: ("" if pd.isna(v) or str(v).lower() == "nan" else str(v).strip()) for k, v in data_terpilih.items()}

        if not mapping_template:
            st.warning("⚠️ Belum ada template surat di folder `templates`.")
        else:
            surat_pilihan = []
            cols = st.columns(3)
            for idx, label_surat in enumerate(mapping_template.keys()):
                with cols[idx % 3]:
                    if st.checkbox(label_surat, value=True if idx == 0 else False):
                        surat_pilihan.append(label_surat)

            st.write("")
            if st.button("🚀 Proses & Buat Berkas"):
                if not surat_pilihan:
                    st.warning("⚠️ Silakan centang minimal satu dokumen.")
                else:
                    try:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for jenis in surat_pilihan:
                                filename_template = mapping_template[jenis]
                                template_path = os.path.join(TEMPLATE_DIR, filename_template)

                                if os.path.exists(template_path):
                                    doc = DocxTemplate(template_path)
                                    doc.render(data_clean)
                                    doc_buffer = io.BytesIO()
                                    doc.save(doc_buffer)
                                    doc_buffer.seek(0)
                                    
                                    nomor_clean = str(nomor_kasus).replace("/", "_").replace("\\", "_")
                                    zip_file.writestr(f"{jenis}_{nomor_clean}.docx", doc_buffer.getvalue())

                        zip_buffer.seek(0)
                        st.success("✅ Berkas berhasil dibuat!")
                        st.download_button(
                            label="📥 Unduh Kumpulan Dokumen (.ZIP)",
                            data=zip_buffer,
                            file_name=f"Berkas_{str(nomor_kasus).replace('/', '_')}.zip",
                            mime="application/zip"
                        )
                    except Exception as e:
                        st.error(f"Terjadi kesalahan: {e}")

# ==========================================
# HALAMAN 2: ADMIN (KELOLA TEMPLATE SURAT)
# ==========================================
elif pilihan_halaman == "⚙️ Admin (Kelola Template)":
    st.title("⚙️ Pengelola Template")
    
    tab1, tab2 = st.tabs(["➕ Upload Template Baru", "🗑️ Hapus Template"])

    with tab1:
        st.write("Sistem otomatis membaca file `.docx` dari folder. Jika nama file Anda `Surat Tugas.docx`, akan muncul di menu sebagai **SURAT TUGAS**.")
        nama_file_kustom = st.text_input("Ganti Nama File (Opsional, tanpa .docx):", placeholder="Contoh: SP SIDIK")
        uploaded_file = st.file_uploader("Upload File `.docx`:", type=["docx"])

        if st.button("💾 Simpan Template"):
            if uploaded_file is None:
                st.warning("⚠️ Pilih file `.docx` terlebih dahulu.")
            else:
                nama_akhir = f"{nama_file_kustom.strip()}.docx" if nama_file_kustom.strip() else uploaded_file.name
                file_path = os.path.join(TEMPLATE_DIR, nama_akhir)
                
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                st.success(f"✅ Template **{nama_akhir}** berhasil ditambahkan!")
                st.rerun()

    with tab2:
        if not mapping_template:
            st.info("Belum ada template surat.")
        else:
            target_hapus = st.selectbox("Pilih Template yang Dihapus:", list(mapping_template.keys()))
            if st.button("❌ Hapus Template"):
                file_to_delete = mapping_template[target_hapus]
                file_path = os.path.join(TEMPLATE_DIR, file_to_delete)
                if os.path.exists(file_path):
                    os.remove(file_path)
                st.success(f"✅ Template **{target_hapus}** berhasil dihapus!")
                st.rerun()