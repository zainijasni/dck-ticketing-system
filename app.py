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
st.set_page_config(page_title="DCK Tech Cloud System", layout="wide")

# ==========================================
# ✅ CLOUDINARY CONFIG (GUNA SECRETS)
# ==========================================
# Kita simpan maklumat Cloudinary dalam Secrets juga supaya selamat
cloudinary.config( 
  cloud_name = "dmnk9vbyj", 
  api_key = "382361922884785", 
  api_secret = "qZBl9CL1886_Ohvv_5JWPxzUQz8",
  secure = True
)

# ID Google Sheet Boss
SHEET_ID = "1ssuZ3BzAih5goP5m_XsgAPjCj1OeDX_CdE--S0h-xek"

# --- FUNGSI CONNECT GOOGLE (VERSI CLOUD) ---
def connect_google():
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        # AMBIL DATA DARI STREAMLIT SECRETS (BUKAN FAIL .JSON)
        creds_info = st.secrets["google_creds"]
        creds = Credentials.from_service_account_info(creds_info, scopes=scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Sistem gagal hubungi Google: {e}")
        return None

def upload_to_cloudinary(file_obj):
    try:
        upload_result = cloudinary.uploader.upload(file_obj)
        return upload_result["secure_url"]
    except: return "Error Upload"

def load_data(tab_name):
    try:
        client = connect_google()
        if client:
            sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
            return pd.DataFrame(sheet.get_all_records())
        return pd.DataFrame()
    except: return pd.DataFrame()

def add_row(tab_name, row_data):
    try:
        client = connect_google()
        sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
        sheet.append_row(row_data)
    except: pass

def update_status_sheet(ticket_id, new_status, new_note, new_cost):
    try:
        client = connect_google()
        sheet = client.open_by_key(SHEET_ID).worksheet("Tickets")
        cell = sheet.find(str(ticket_id))
        if cell:
            row = cell.row
            sheet.update_cell(row, 9, new_status)
            sheet.update_cell(row, 11, new_cost)
            sheet.update_cell(row, 13, new_note)
            return True
        return False
    except: return False

def generate_pdf(tiket_data):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.setFont("Helvetica-Bold", 25); c.drawString(50, 800, "DCK TECH")
    c.setFont("Helvetica", 10); c.drawString(50, 785, "COMPUTER REPAIR & SERVICES")
    c.line(50, 775, 550, 775)
    c.setFont("Helvetica-Bold", 14); c.drawString(50, 750, f"TIKET: {tiket_data['ID']}")
    c.setFont("Helvetica", 11)
    c.drawString(50, 720, f"Customer: {tiket_data['Customer']}"); c.drawString(300, 720, f"Model: {tiket_data['Model']}")
    c.drawString(50, 680, "Masalah:"); c.drawString(50, 665, f"{tiket_data['Masalah']}")
    c.line(50, 600, 550, 600)
    c.setFont("Helvetica-Bold", 10); c.drawString(50, 585, "T&C:")
    tc = ["1. Data loss bukan tanggungjawab kedai.", "2. Hak milik kedai >3 bulan.", "3. Warranty sparepart sahaja."]
    y = 570
    for line in tc:
        c.drawString(60, y, line); y -= 15
    c.save(); buffer.seek(0)
    return buffer

# --- NAVIGATION ---
menu = st.sidebar.radio("NAVIGASI", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"])

if menu == "📊 DASHBOARD":
    st.title("📊 Dashboard")
    df = load_data("Tickets")
    if not df.empty:
        search = st.text_input("Cari Nama/ID:")
        if search: df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        st.dataframe(df[['ID','Tarikh','Customer','Model','Status','Kos']], use_container_width=True)

elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Daftar Tiket")
    if 'last_ticket' not in st.session_state: st.session_state['last_ticket'] = None
    with st.form("f"):
        nama = st.text_input("Nama"); phone = st.text_input("Phone")
        model = st.text_input("Model"); masalah = st.text_area("Masalah")
        gambar = st.camera_input("Snap")
        t = st.checkbox("Setuju T&C")
        if st.form_submit_button("SIMPAN"):
            if t and nama:
                tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                img = upload_to_cloudinary(gambar) if gambar else "No Image"
                row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, "", "", masalah, "Pending", "-", 0.0, img, ""]
                add_row("Tickets", row)
                st.session_state['last_ticket'] = {"ID": tid, "Tarikh": row[1], "Customer": nama, "Model": model, "Masalah": masalah, "Phone": phone}
                st.success("Success!"); st.rerun()
    if st.session_state['last_ticket']:
        resit = generate_pdf(st.session_state['last_ticket'])
        st.download_button("Download PDF", resit, "resit.pdf")

elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Technician")
    df = load_data("Tickets")
    if not df.empty:
        pilih_id = st.selectbox("Pilih ID", df['ID'].tolist())
        job = df[df['ID'] == pilih_id].iloc[0]
        if "http" in str(job['Image_Link']): st.image(job['Image_Link'], width=300)
        with st.form("u"):
            stat = st.selectbox("Status", ["Pending", "Checking", "Done"])
            kos = st.number_input("Kos", value=float(job['Kos']) if job['Kos'] != "" else 0.0)
            note = st.text_area("Nota", value=str(job['Tech_Note']))
            if st.form_submit_button("UPDATE"):
                if update_status_sheet(pilih_id, stat, note, kos): st.rerun()

elif menu == "📦 INVENTORY":
    st.title("📦 Parts")
    st.dataframe(load_data("Parts"), use_container_width=True)
