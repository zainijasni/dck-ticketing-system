import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
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
    except: st.error("Gagal simpan data ke Google Sheets.")

def update_status_sheet(ticket_id, stat, note, harga_jual, kos_part):
    try:
        sheet = connect_google().open_by_key(SHEET_ID).worksheet("Tickets")
        cell = sheet.find(str(ticket_id))
        if cell:
            # Update ikut susunan kolum: I=9, J=10, K=11, M=13
            sheet.update_cell(cell.row, 9, stat)        # Status
            sheet.update_cell(cell.row, 10, kos_part)   # Kos_Part (Modal)
            sheet.update_cell(cell.row, 11, harga_jual) # Harga_Jual (Caj)
            sheet.update_cell(cell.row, 13, note)       # Tech_Note
            return True
        return False
    except: return False

# --- NAVIGATION ---
menu = st.sidebar.radio("NAVIGASI", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"])

# --- 1. DASHBOARD ---
if menu == "📊 DASHBOARD":
    st.title("📊 Dashboard DCK Tech")
    df = load_data("Tickets")
    if not df.empty:
        # Convert data kpd nombor untuk kira profit
        df['Harga_Jual'] = pd.to_numeric(df['Harga_Jual'], errors='coerce').fillna(0)
        df['Kos_Part'] = pd.to_numeric(df['Kos_Part'], errors='coerce').fillna(0)
        df['Untung'] = df['Harga_Jual'] - df['Kos_Part']
        
        c1, c2 = st.columns(2)
        c1.metric("Total Jualan (Caj)", f"RM {df['Harga_Jual'].sum():.2f}")
        c2.metric("Total Untung Bersih", f"RM {df['Untung'].sum():.2f}")
        
        st.divider()
        search = st.text_input("🔍 Cari (Nama/Model/Phone/ID):")
        if search:
            df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        st.dataframe(df[['ID','Customer','Model','Status','Harga_Jual']], use_container_width=True)

# --- 2. DAFTAR TIKET ---
elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Pendaftaran Baru")
    with st.form("reg_form"):
        col1, col2 = st.columns(2)
        nama = col1.text_input("Nama Customer")
        phone = col2.text_input("No WhatsApp")
        model = col1.text_input("Model Peranti")
        sn = col2.text_input("Serial No (S/N)")
        pwd = col1.text_input("Password")
        masalah = st.text_area("Masalah")
        gambar = st.camera_input("Snap Gambar")
        tnc = st.checkbox("Pelanggan setuju T&C DCK Tech")
        
        if st.form_submit_button("SIMPAN"):
            if not nama or not model or not tnc:
                st.error("Nama, Model & T&C wajib!")
            else:
                with st.spinner("Menyimpan..."):
                    tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                    img_url = upload_result = cloudinary.uploader.upload(gambar)["secure_url"] if gambar else "No Image"
                    # Susunan Row: ID, Tarikh, Customer, Phone, Model, SN, Pwd, Masalah, Status, Kos_Part, Harga_Jual, Image, Note
                    row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, sn, pwd, masalah, "Pending", 0, 0, img_url, ""]
                    add_row("Tickets", row)
                    st.success(f"Tiket {tid} Berjaya!"); st.balloons()

# --- 3. UPDATE STATUS ---
elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Technician")
    df_t = load_data("Tickets")
    if not df_t.empty:
        search_id = st.text_input("🔍 Masukkan Ticket ID (Cth: DCK-1234):")
        ids = df_t['ID'].tolist()
        idx = ids.index(search_id) if search_id in ids else 0
        pilih_id = st.selectbox("Atau Pilih Dari Senarai:", ids, index=idx)
        
        job = df_t[df_t['ID'] == pilih_id].iloc[0]
        st.info(f"JOB: {job['ID']} | CUSTOMER: {job['Customer']} | MODEL: {job['Model']}")
        
        if "http" in str(job['Image_Link']): st.image(job['Image_Link'], width=300)

        with st.form("upd_job"):
            colA, colB = st.columns(2)
            stat = colA.selectbox("Status", ["Checking", "Waiting Part", "Repairing", "Done", "Collected"], 
                                 index=["Checking", "Waiting Part", "Repairing", "Done", "Collected"].index(job['Status']) if job['Status'] in ["Checking", "Waiting Part", "Repairing", "Done", "Collected"] else 0)
            note = st.text_area("Nota Tindakan", value=str(job['Tech_Note']))
            
            st.divider()
            k_part = colA.number_input("Kos Modal Part (RM)", value=float(job['Kos_Part']) if job['Kos_Part'] != "" else 0.0)
            h_jual = colB.number_input("Harga Caj Customer (RM)", value=float(job['Harga_Jual']) if job['Harga_Jual'] != "" else 0.0)
            
            if st.form_submit_button("KEMASKINI"):
                if update_status_sheet(pilih_id, stat, note, h_jual, k_part):
                    st.success("Berjaya diupdate!"); st.rerun()

        st.divider()
        st.subheader("🔩 Alat Ganti")
        df_p = load_data("Parts")
        if not df_p.empty:
            curr = df_p[df_p['TicketID'].astype(str) == str(pilih_id)]
            if not curr.empty: st.table(curr[['NamaPart', 'Supplier', 'TarikhExpire']])
        
        with st.expander("➕ Tambah Part"):
            with st.form("add_p"):
                pn = st.text_input("Nama Part"); ps = st.text_input("Supplier"); pw = st.number_input("Warranty (Bln)", 0)
                if st.form_submit_button("SIMPAN PART"):
                    exp = (datetime.now() + pd.DateOffset(months=pw)).strftime("%Y-%m-%d")
                    add_row("Parts", [f"P-{datetime.now().strftime('%M%S')}", pilih_id, pn, ps, datetime.now().strftime("%Y-%m-%d"), pw, exp])
                    st.rerun()

elif menu == "📦 INVENTORY":
    st.title("📦 Inventory")
    st.dataframe(load_data("Parts"), use_container_width=True)
