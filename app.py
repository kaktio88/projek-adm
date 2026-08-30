import streamlit as st
import pandas as pd
from st_gsheets_connection import GSheetsConnection
import google.generativeai as genai
from docx import Document
import io

# 1. Konfigurasi Halaman
st.set_page_config(
    page_title="Sistem Administrasi Otomatis",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Sistem Administrasi & Pembuat Surat Otomatis")
st.write("---")

# 2. Inisialisasi Google Sheets
st.subheader("📊 Status Koneksi Google Sheets")
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = conn.read(ttl=0)
    st.success("✅ Berhasil terhubung ke Google Sheets!")
    
    with st.expander("Lihat Data Google Sheets"):
        st.dataframe(df)

except Exception as e:
    st.warning("⚠️ Belum terhubung ke Google Sheets atau format Secrets belum sesuai.")
    st.caption(f"Detail kendala: {e}")

# 3. Inisialisasi Gemini AI
st.write("---")
st.subheader("🤖 Generator Teks & AI Assistant (Gemini)")

gemini_key = st.secrets.get("GEMINI_API_KEY") if "GEMINI_API_KEY" in st.secrets else None

if gemini_key:
    try:
        genai.configure(api_key=gemini_key)
        model = genai.GenerativeModel("gemini-1.5-flash")
        
        prompt_user = st.text_area("Masukkan instruksi pembuatan surat / draf:", "Buatkan draf surat permohonan izin kegiatan...")
        
        if st.button("Generate Teks Surat"):
            with st.spinner("Sedang memproses draf dengan Gemini AI..."):
                response = model.generate_content(prompt_user)
                st.session_state["draft_surat"] = response.text
                st.success("✅ Draf berhasil dibuat!")
                st.write(response.text)
    except Exception as e:
        st.error(f"❌ Terjadi kesalahan pada Gemini AI: {e}")
else:
    st.warning("⚠️ GEMINI_API_KEY belum ditambahkan pada Secrets di Streamlit Cloud.")

# 4. Fitur Ekspor ke File Word (.docx)
st.write("---")
st.subheader("📄 Unduh Dokumen Word (.docx)")

if "draft_surat" in st.session_state and st.session_state["draft_surat"]:
    if st.button("Buat File Word (.docx)"):
        doc = Document()
        doc.add_heading("DRAF SURAT RESMI", level=1)
        doc.add_paragraph(st.session_state["draft_surat"])
        
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            label="📥 Download Surat (.docx)",
            data=buffer,
            file_name="Surat_Otomatis.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.info("Buat draf surat menggunakan Gemini AI terlebih dahulu untuk mengunduh file Word.")