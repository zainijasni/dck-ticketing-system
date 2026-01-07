import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
import gspread
from google.oauth2.service_account import Credentials
import cloudinary
import cloudinary.uploader

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="DCK Tech Management", layout="wide")

# Cloudinary Config
cloudinary.config( 
  cloud_name = "dmnk9vbyj", 
  api_key = "382361922884785", 
  api_secret = "qZBl9CL1886_Ohvv_5JWPxzUQz8",
  secure = True
)

SHEET_ID = "1ssuZ3BzAih5goP5m_XsgAPjCj1OeDX_CdE--S0h-xek"

# Initialize Session State
if 'page' not in st.session_state: st.session_state.page = "📊 DASHBOARD"
if 'selected_ticket' not in st.session_state: st.session_state.selected_ticket = None
if 'last_ticket_pdf' not in st.session_state: st.session_state.last_ticket_pdf = None

# --- DATABASE FUNCTIONS ---
@st.cache_resource
def get_gspread_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_info = st.secrets["google_creds"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
    return gspread.authorize(creds)

def load_data(tab_name):
    try:
        client = get_gspread_client()
        data = client.open_by_key(SHEET_ID).worksheet(tab_name).get_all_records()
        return pd.DataFrame(data)
    except: return pd.DataFrame()

def add_row(tab_name, row_data):
    try:
        get_gspread_client().open_by_key(SHEET_ID).worksheet(tab_name).append_row(row_data)
        st.cache_data.clear() 
    except: st.error("Gagal simpan!")

def update_status_sheet(ticket_id, stat, note, harga_jual, kos_part):
    try:
        sheet = get_gspread_client().open_by_key(SHEET_ID).worksheet("Tickets")
        cell = sheet.find(str(ticket_id))
        if cell:
            sheet.update_cell(cell.row, 9, stat)
            sheet.update_cell(cell.row, 10, kos_part)
            sheet.update_cell(cell.row, 11, harga_jual)
            sheet.update_cell(cell.row, 13, note)
            st.cache_data.clear()
            return True
        return False
    except: return False

# --- PDF GENERATOR (WITH T&C & SIGNATURE) ---
def generate_pdf(t):
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4

    # Header
    p.setFont("Helvetica-Bold", 22); p.drawString(50, h-60, "DCK TECH")
    p.setFont("Helvetica", 10); p.drawString(50, h-75, "Solusi Komputer & Gadget Anda")
    p.line(50, h-85, w-50, h-85)

    # Info Tiket
    p.setFont("Helvetica-Bold", 14); p.drawString(50, h-110, f"RESIT PENERIMAAN PERANTI: {t['ID']}")
    p.setFont("Helvetica", 11)
    p.drawString(50, h-140, f"Nama Pelanggan: {t['Customer']}")
    p.drawString(350, h-140, f"Tarikh: {t['Tarikh']}")
    p.drawString(50, h-155, f"No. Telefon: {t['Phone']}")
    p.drawString(350, h-155, f"Model: {t['Model']}")
    
    # Masalah
    p.setFont("Helvetica-Bold", 11); p.drawString(50, h-190, "KEROSAKAN / ADUAN:")
    p.setFont("Helvetica", 11); p.drawString(60, h-205, f"{t['Masalah']}")

    # Terma & Syarat (T&C)
    p.setDash(1, 2)
    p.line(50, h-250, w-50, h-250)
    p.setDash(1, 0)
    p.setFont("Helvetica-Bold", 12); p.drawString(50, h-270, "TERMA & SYARAT PERKHIDMATAN:")
    p.setFont("Helvetica", 9)
    tc_list = [
        "1. Pihak DCK TECH tidak bertanggungjawab atas sebarang kehilangan data semasa proses repair.",
        "2. Peranti yang tidak dituntut dalam tempoh 90 hari (3 bulan) akan menjadi hak milik kedai.",
        "3. Warranty hanya sah untuk alat ganti (parts) yang ditukar sahaja. Kerosakan lain tidak termasuk.",
        "4. Deposit tidak akan dikembalikan sekiranya pelanggan membatalkan repair selepas persetujuan.",
        "5. Pelanggan wajib membawa resit ini semasa urusan pengambilan peranti."
    ]
    y_pos = h-285
    for line in tc_list:
        p.drawString(60, y_pos, line)
        y_pos -= 15

    # Ruangan Tandatangan
    p.setFont("Helvetica-Bold", 10)
    p.drawString(50, y_pos-50, "Tandatangan Pelanggan,")
    p.drawString(350, y_pos-50, "Tandatangan DCK TECH,")
    p.line(50, y_pos-100, 200, y_pos-100)
    p.line(350, y_pos-100, 500, y_pos-100)
    p.setFont("Helvetica", 8)
    p.drawString(50, y_pos-110, "(Nama: ____________________)")
    
    p.showPage(); p.save()
    buffer.seek(0)
    return buffer

def switch_page(page, tid=None):
    st.session_state.page = page
    st.session_state.selected_ticket = tid
    st.rerun()

# --- APP UI ---
menu = st.sidebar.radio("MENU UTAMA", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"], 
                        index=["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"].index(st.session_state.page))

if menu == "📊 DASHBOARD":
    st.title("📊 Dashboard DCK Tech")
    df = load_data("Tickets")
    if not df.empty:
        # Statistik Card
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Pending ⏳", len(df[df['Status'] == 'Pending']))
        s2.metric("Checking 🛠️", len(df[df['Status'] == 'Checking']))
        s3.metric("Done ✅", len(df[df['Status'] == 'Done']))
        s4.metric("Collected 📦", len(df[df['Status'] == 'Collected']))

        st.divider()
        search = st.text_input("🔍 Cari Ticket/Nama/Model:")
        if search:
            df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        
        for _, row in df.iloc[::-1].iterrows():
            with st.expander(f"📌 {row['ID']} - {row['Customer']} ({row['Status']})"):
                st.write(f"**Model:** {row['Model']} | **Masalah:** {row['Masalah']}")
                if st.button("🔧 Update", key=f"d_{row['ID']}"): switch_page("🔧 UPDATE STATUS", row['ID'])

elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Pendaftaran Tiket Baru")
    # INTERFACE CANTIK: Guna container & columns
    with st.container(border=True):
        st.subheader("👤 Maklumat Pelanggan")
        c1, c2 = st.columns(2)
        nama = c1.text_input("Nama Penuh", placeholder="Contoh: Ahmad Abu")
        phone = c2.text_input("No. WhatsApp", placeholder="0123456789")
        
        st.subheader("💻 Maklumat Peranti")
        c3, c4 = st.columns(2)
        model = c3.text_input("Model / Jenama", placeholder="Contoh: MacBook Pro 2017")
        sn = c4.text_input("Serial No (S/N)", placeholder="Optional")
        masalah = st.text_area("Diagnosis / Masalah Peranti")
        
        st.subheader("📸 Keadaan Fizikal")
        gambar = st.camera_input("Snap Gambar Peranti (Sangat Digalakkan)")
        
        st.info("⚠️ Pastikan pelanggan membaca Terma & Syarat sebelum mendaftar.")
        tnc = st.checkbox("SAYA MENGAKU PELANGGAN TELAH BERSETUJU DENGAN T&C")

        if st.button("🚀 SIMPAN & CETAK TIKET", use_container_width=True):
            if nama and model and tnc:
                with st.spinner("Processing..."):
                    tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                    img = cloudinary.uploader.upload(gambar)["secure_url"] if gambar else "No Image"
                    row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, sn, "", masalah, "Pending", 0, 0, img, ""]
                    add_row("Tickets", row)
                    st.session_state.last_ticket_pdf = {"ID": tid, "Customer": nama, "Model": model, "Phone": phone, "Masalah": masalah, "Tarikh": row[1]}
                    st.success("Tiket Berjaya Disimpan!"); st.balloons()
            else: st.error("Lengkapkan Nama, Model & Tick T&C!")

    if st.session_state.last_ticket_pdf:
        st.divider()
        st.download_button("📥 DOWNLOAD RESIT PDF SEKARANG", generate_pdf(st.session_state.last_ticket_pdf), f"Resit_{st.session_state.last_ticket_pdf['ID']}.pdf", use_container_width=True)

elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Kerja Technician")
    df_t = load_data("Tickets")
    if not df_t.empty:
        ids = df_t['ID'].astype(str).tolist()
        target = str(st.session_state.selected_ticket) if st.session_state.selected_ticket else ids[-1]
        pilih_id = st.selectbox("Pilih Job:", ids, index=ids.index(target) if target in ids else 0)
        job = df_t[df_t['ID'].astype(str) == str(pilih_id)].iloc[0]
        
        st.info(f"**PELANGGAN:** {job['Customer']} | **MODEL:** {job['Model']}")
        
        col_a, col_b = st.columns([1, 2])
        with col_a:
            if "http" in str(job['Image_Link']): st.image(job['Image_Link'], caption="Kondisi Awal")
            st.download_button("🖨️ Cetak Semula Resit", generate_pdf({"ID": job['ID'], "Customer": job['Customer'], "Model": job['Model'], "Phone": job['Phone'], "Masalah": job['Masalah'], "Tarikh": job['Tarikh']}), f"Resit_{job['ID']}.pdf", use_container_width=True)

        with col_b:
            with st.form("upd"):
                stat = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"], index=["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"].index(job['Status']) if job['Status'] in ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"] else 0)
                note = st.text_area("Nota Tech", value=str(job['Tech_Note']))
                k_m = st.number_input("Kos Modal Part (RM)", value=float(job['Kos_Part']) if job['Kos_Part'] != "" else 0.0)
                h_c = st.number_input("Harga Jual (RM)", value=float(job['Harga_Jual']) if job['Harga_Jual'] != "" else 0.0)
                if st.form_submit_button("UPDATE JOB"):
                    if update_status_sheet(pilih_id, stat, note, h_c, k_m): st.success("Updated!"); st.rerun()

elif menu == "📦 INVENTORY":
    st.title("📦 Inventory Log")
    st.dataframe(load_data("Parts"), use_container_width=True)
