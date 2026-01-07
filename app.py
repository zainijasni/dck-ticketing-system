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

# --- CACHING UNTUK LAJU ---
@st.cache_resource
def get_gspread_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_info = st.secrets["google_creds"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
    return gspread.authorize(creds)

def load_data(tab_name):
    try:
        client = get_gspread_client()
        # Gunakan cache_data untuk speed, clear setiap 10 saat
        data = client.open_by_key(SHEET_ID).worksheet(tab_name).get_all_records()
        return pd.DataFrame(data)
    except: return pd.DataFrame()

def add_row(tab_name, row_data):
    try:
        client = get_gspread_client()
        client.open_by_key(SHEET_ID).worksheet(tab_name).append_row(row_data)
        st.cache_data.clear() # Clear cache supaya data baru muncul
    except: st.error("Gagal simpan!")

def update_status_sheet(ticket_id, stat, note, harga_jual, kos_part):
    try:
        client = get_gspread_client()
        sheet = client.open_by_key(SHEET_ID).worksheet("Tickets")
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

# --- PDF GENERATOR ---
def generate_pdf(tiket_data):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.setFont("Helvetica-Bold", 25); c.drawString(50, 800, "DCK TECH")
    c.line(50, 775, 550, 775)
    c.setFont("Helvetica-Bold", 14); c.drawString(50, 750, f"RESIT TIKET: {tiket_data['ID']}")
    c.setFont("Helvetica", 11)
    c.drawString(50, 720, f"Customer: {tiket_data['Customer']}")
    c.drawString(300, 720, f"Model: {tiket_data['Model']}")
    c.drawString(50, 705, f"Phone: {tiket_data.get('Phone', '-')}")
    c.drawString(300, 705, f"Tarikh: {tiket_data.get('Tarikh', '-')}")
    c.drawString(50, 670, "Masalah:"); c.drawString(50, 655, f"{tiket_data['Masalah']}")
    c.save(); buffer.seek(0)
    return buffer

def switch_page(page_name, ticket_id=None):
    st.session_state.page = page_name
    st.session_state.selected_ticket = ticket_id
    st.rerun()

# --- SIDEBAR ---
menu = st.sidebar.radio("NAVIGASI", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"], 
                        index=["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"].index(st.session_state.page))

# --- 1. DASHBOARD ---
if menu == "📊 DASHBOARD":
    st.title("📊 Dashboard")
    df = load_data("Tickets")
    if not df.empty:
        # Metrik Kewangan
        df['Harga_Jual'] = pd.to_numeric(df['Harga_Jual'], errors='coerce').fillna(0)
        df['Kos_Part'] = pd.to_numeric(df['Kos_Part'], errors='coerce').fillna(0)
        profit = df['Harga_Jual'].sum() - df['Kos_Part'].sum()
        
        c1, c2 = st.columns(2)
        c1.metric("Total Jualan", f"RM {df['Harga_Jual'].sum():.2f}")
        c2.metric("Untung Bersih", f"RM {profit:.2f}")

        st.divider()
        search = st.text_input("🔍 Cari Ticket/Customer:")
        if search:
            df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        
        for index, row in df.iloc[::-1].iterrows():
            with st.expander(f"📌 {row['ID']} - {row['Customer']} [{row['Status']}]"):
                st.write(f"**Model:** {row['Model']} | **Masalah:** {row['Masalah']}")
                if st.button("🔧 Update Kerja", key=f"dash_{row['ID']}"):
                    switch_page("🔧 UPDATE STATUS", row['ID'])

# --- 2. DAFTAR TIKET ---
elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Daftar Baru")
    with st.form("reg"):
        c1, c2 = st.columns(2)
        nama = c1.text_input("Nama Customer")
        phone = c2.text_input("WhatsApp")
        model = c1.text_input("Model Laptop")
        sn = c2.text_input("S/N")
        masalah = st.text_area("Masalah")
        gambar = st.camera_input("Snap Gambar")
        tnc = st.checkbox("Setuju T&C DCK Tech")
        
        if st.form_submit_button("SIMPAN"):
            if nama and model and tnc:
                tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                img = cloudinary.uploader.upload(gambar)["secure_url"] if gambar else "No Image"
                row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, sn, "", masalah, "Pending", 0, 0, img, ""]
                add_row("Tickets", row)
                st.session_state.last_ticket_pdf = {"ID": tid, "Customer": nama, "Model": model, "Phone": phone, "Masalah": masalah, "Tarikh": row[1]}
                st.success("Tersimpan!"); st.rerun()

    if st.session_state.last_ticket_pdf:
        pdf = generate_pdf(st.session_state.last_ticket_pdf)
        st.download_button("📥 Download PDF", pdf, f"Resit_{st.session_state.last_ticket_pdf['ID']}.pdf")

# --- 3. UPDATE STATUS ---
elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Kerja")
    df_t = load_data("Tickets")
    if not df_t.empty:
        ids = df_t['ID'].astype(str).tolist()
        # Ambil ID yang dipilih dari session atau default pertama
        target = str(st.session_state.selected_ticket) if st.session_state.selected_ticket else ids[-1]
        
        if target in ids:
            curr_idx = ids.index(target)
        else:
            curr_idx = 0
            
        pilih_id = st.selectbox("Cari ID Tiket:", ids, index=curr_idx)
        
        # DAPATKAN INFO CUSTOMER (DISPLAY FIX)
        job_data = df_t[df_t['ID'].astype(str) == str(pilih_id)].iloc[0]
        
        # Tunjuk info besar-besar supaya tak silap laptop
        st.success(f"**PELANGGAN:** {job_data['Customer']} | **MODEL:** {job_data['Model']}")
        
        col_img, col_form = st.columns([1, 2])
        with col_img:
            if "http" in str(job_data['Image_Link']): 
                st.image(job_data['Image_Link'], caption="Kondisi Asal", use_container_width=True)
            
            # Butang Print Sentiasa Ada
            pdf_info = {"ID": job_data['ID'], "Customer": job_data['Customer'], "Model": job_data['Model'], "Phone": job_data['Phone'], "Masalah": job_data['Masalah'], "Tarikh": job_data['Tarikh']}
            st.download_button("🖨️ Cetak Semula Resit", generate_pdf(pdf_info), f"Resit_{job_data['ID']}.pdf")

        with col_form:
            with st.form("upd_job"):
                stat = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"], 
                                   index=["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"].index(job_data['Status']) if job_data['Status'] in ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"] else 0)
                note = st.text_area("Nota Technician", value=str(job_data['Tech_Note']))
                k_mod = st.number_input("Kos Modal Part (RM)", value=float(job_data['Kos_Part']) if job_data['Kos_Part'] != "" else 0.0)
                h_caj = st.number_input("Harga Caj Customer (RM)", value=float(job_data['Harga_Jual']) if job_data['Harga_Jual'] != "" else 0.0)
                if st.form_submit_button("SIMPAN UPDATE"):
                    if update_status_sheet(pilih_id, stat, note, h_caj, k_mod):
                        st.success("Update Berjaya!"); st.rerun()

        st.divider()
        st.subheader("Alat Ganti (Parts)")
        df_p = load_data("Parts")
        if not df_p.empty:
            curr_p = df_p[df_p['TicketID'].astype(str) == str(pilih_id)]
            st.table(curr_p[['NamaPart', 'Supplier', 'TarikhExpire']])
        
        with st.expander("Tambah Part"):
            with st.form("add_p"):
                pn = st.text_input("Nama Part"); ps = st.text_input("Supplier"); pw = st.number_input("Warranty (Bln)", 0)
                if st.form_submit_button("REKOD"):
                    exp = (datetime.now() + pd.DateOffset(months=pw)).strftime("%Y-%m-%d")
                    add_row("Parts", [f"P-{datetime.now().strftime('%M%S')}", pilih_id, pn, ps, datetime.now().strftime("%Y-%m-%d"), pw, exp])
                    st.rerun()

elif menu == "📦 INVENTORY":
    st.title("📦 Inventory")
    st.dataframe(load_data("Parts"), use_container_width=True)
