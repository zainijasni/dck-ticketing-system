import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import gspread
from google.oauth2.service_account import Credentials
import cloudinary
import cloudinary.uploader

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="DCK Tech System", layout="wide")

# Cloudinary Config
cloudinary.config( 
  cloud_name = "dmnk9vbyj", 
  api_key = "382361922884785", 
  api_secret = "qZBl9CL1886_Ohvv_5JWPxzUQz8",
  secure = True
)

SHEET_ID = "1ssuZ3BzAih5goP5m_XsgAPjCj1OeDX_CdE--S0h-xek"

# Initialize Session State untuk Navigasi & PDF
if 'page' not in st.session_state: st.session_state.page = "📊 DASHBOARD"
if 'selected_ticket' not in st.session_state: st.session_state.selected_ticket = None
if 'last_ticket_pdf' not in st.session_state: st.session_state.last_ticket_pdf = None

def connect_google():
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        creds_info = st.secrets["google_creds"]
        creds = Credentials.from_service_account_info(creds_info, scopes=scope)
        return gspread.authorize(creds)
    except: return None

def load_data(tab_name):
    try:
        client = connect_google()
        return pd.DataFrame(client.open_by_key(SHEET_ID).worksheet(tab_name).get_all_records())
    except: return pd.DataFrame()

def add_row(tab_name, row_data):
    try:
        connect_google().open_by_key(SHEET_ID).worksheet(tab_name).append_row(row_data)
    except: st.error("Gagal simpan!")

def update_status_sheet(ticket_id, stat, note, harga_jual, kos_part):
    try:
        sheet = connect_google().open_by_key(SHEET_ID).worksheet("Tickets")
        cell = sheet.find(str(ticket_id))
        if cell:
            sheet.update_cell(cell.row, 9, stat)
            sheet.update_cell(cell.row, 10, kos_part)
            sheet.update_cell(cell.row, 11, harga_jual)
            sheet.update_cell(cell.row, 13, note)
            return True
        return False
    except: return False

def generate_pdf(tiket_data):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.setFont("Helvetica-Bold", 25); c.drawString(50, 800, "DCK TECH")
    c.line(50, 775, 550, 775)
    c.setFont("Helvetica-Bold", 14); c.drawString(50, 750, f"TIKET REPAIR: {tiket_data['ID']}")
    c.setFont("Helvetica", 11)
    c.drawString(50, 720, f"Customer: {tiket_data['Customer']}"); c.drawString(300, 720, f"Model: {tiket_data['Model']}")
    c.drawString(50, 705, f"Phone: {tiket_data.get('Phone', '-')}")
    c.drawString(50, 670, "Masalah:"); c.drawString(50, 655, f"{tiket_data['Masalah']}")
    c.line(50, 600, 550, 600)
    c.setFont("Helvetica-Bold", 10); c.drawString(50, 585, "TERMA & SYARAT:")
    tc = ["1. Data loss bukan tanggungjawab kedai.", "2. Hak milik kedai selepas 3 bulan.", "3. Warranty sparepart baru sahaja."]
    y = 570
    for line in tc: c.drawString(60, y, line); y -= 15
    c.save(); buffer.seek(0)
    return buffer

# --- SIDEBAR NAV ---
st.sidebar.title("DCK TECH")
menu = st.sidebar.radio("NAVIGASI", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"], index=["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"].index(st.session_state.page))

# Function to switch page
def switch_page(page_name, ticket_id=None):
    st.session_state.page = page_name
    st.session_state.selected_ticket = ticket_id
    st.rerun()

# --- 1. DASHBOARD ---
if menu == "📊 DASHBOARD":
    st.title("📊 Dashboard")
    df = load_data("Tickets")
    if not df.empty:
        st.subheader("📋 Senarai Job Terkini")
        # Layout dashboard dengan butang Edit/Teleport
        for index, row in df.iloc[::-1].iterrows(): # Tunjuk yang terbaru dulu
            with st.expander(f"📌 {row['ID']} - {row['Customer']} ({row['Status']})"):
                col_a, col_b = st.columns([3, 1])
                col_a.write(f"**Model:** {row['Model']} | **Masalah:** {row['Masalah']}")
                if col_b.button("🔧 Update / Edit", key=f"btn_{row['ID']}"):
                    switch_page("🔧 UPDATE STATUS", row['ID'])
        
        st.divider()
        st.write("Semua Data (Raw):")
        st.dataframe(df, use_container_width=True)

# --- 2. DAFTAR TIKET ---
elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Daftar Baru")
    with st.form("reg_form"):
        col1, col2 = st.columns(2)
        nama = col1.text_input("Nama Customer")
        phone = col2.text_input("WhatsApp")
        model = col1.text_input("Model Laptop")
        sn = col2.text_input("S/N")
        masalah = st.text_area("Masalah")
        gambar = st.camera_input("Snap Gambar")
        tnc = st.checkbox("Setuju T&C")
        
        if st.form_submit_button("SIMPAN"):
            if nama and model and tnc:
                tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                img_url = cloudinary.uploader.upload(gambar)["secure_url"] if gambar else "No Image"
                row_data = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, sn, "", masalah, "Pending", 0, 0, img_url, ""]
                add_row("Tickets", row_data)
                # Store PDF data dalam session supaya butang tak hilang
                st.session_state.last_ticket_pdf = {"ID": tid, "Customer": nama, "Model": model, "Phone": phone, "Masalah": masalah, "Tarikh": row_data[1]}
                st.success(f"Berjaya! ID: {tid}")
            else: st.error("Isi Nama, Model & T&C!")

    if st.session_state.last_ticket_pdf:
        st.divider()
        st.subheader("🖨️ Cetak Resit Terakhir")
        pdf_file = generate_pdf(st.session_state.last_ticket_pdf)
        st.download_button(label="📥 Download Resit PDF", data=pdf_file, file_name=f"Resit_{st.session_state.last_ticket_pdf['ID']}.pdf", mime="application/pdf")

# --- 3. UPDATE STATUS ---
elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Technician")
    df_t = load_data("Tickets")
    if not df_t.empty:
        ids = df_t['ID'].tolist()
        # Jika datang dari teleport dashboard, auto-select ID tu
        target_id = st.session_state.selected_ticket if st.session_state.selected_ticket in ids else ids[0]
        pilih_id = st.selectbox("Pilih Job ID:", ids, index=ids.index(target_id))
        
        job = df_t[df_t['ID'] == pilih_id].iloc[0]
        st.info(f"JOB: {job['ID']} | CUSTOMER: {job['Customer']}")
        
        if "http" in str(job['Image_Link']): st.image(job['Image_Link'], width=300)

        with st.form("upd_form"):
            colA, colB = st.columns(2)
            stat = colA.selectbox("Status", ["Checking", "Waiting Part", "Repairing", "Done", "Collected"], index=["Checking", "Waiting Part", "Repairing", "Done", "Collected"].index(job['Status']) if job['Status'] in ["Checking", "Waiting Part", "Repairing", "Done", "Collected"] else 0)
            note = st.text_area("Nota Technician", value=str(job['Tech_Note']))
            k_part = colA.number_input("Kos Modal Part (RM)", value=float(job['Kos_Part']) if job['Kos_Part'] != "" else 0.0)
            h_jual = colB.number_input("Harga Caj Customer (RM)", value=float(job['Harga_Jual']) if job['Harga_Jual'] != "" else 0.0)
            
            if st.form_submit_button("SIMPAN PERUBAHAN"):
                if update_status_sheet(pilih_id, stat, note, h_jual, k_part):
                    st.success("Telah Dikemaskini!"); st.rerun()

        st.divider()
        st.subheader("🔩 Alat Ganti")
        df_p = load_data("Parts")
        if not df_p.empty:
            curr = df_p[df_p['TicketID'].astype(str) == str(pilih_id)]
            if not curr.empty: st.table(curr[['NamaPart', 'Supplier', 'TarikhExpire']])
        
        with st.expander("➕ Tambah Part"):
            with st.form("add_p"):
                pn = st.text_input("Nama Part"); ps = st.text_input("Supplier"); pw = st.number_input("Warranty (Bln)", 0)
                if st.form_submit_button("REKOD"):
                    exp = (datetime.now() + pd.DateOffset(months=pw)).strftime("%Y-%m-%d")
                    add_row("Parts", [f"P-{datetime.now().strftime('%M%S')}", pilih_id, pn, ps, datetime.now().strftime("%Y-%m-%d"), pw, exp])
                    st.rerun()

elif menu == "📦 INVENTORY":
    st.title("📦 Inventory")
    st.dataframe(load_data("Parts"), use_container_width=True)
