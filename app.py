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

# Cloudinary (Data Boss dah ada dalam Secrets atau kod)
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
    except: pass

def update_status_sheet(ticket_id, stat, note, harga_jual, kos_part):
    try:
        sheet = connect_google().open_by_key(SHEET_ID).worksheet("Tickets")
        cell = sheet.find(str(ticket_id))
        if cell:
            # Kolum 9=Status, 11=HargaJual, 12=KosPart, 13=Note
            sheet.update_cell(cell.row, 9, stat)
            sheet.update_cell(cell.row, 11, harga_jual)
            sheet.update_cell(cell.row, 12, kos_part)
            sheet.update_cell(cell.row, 13, note)
            return True
        return False
    except: return False

# --- NAVIGATION ---
menu = st.sidebar.radio("NAVIGASI", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"])

# --- DASHBOARD ---
if menu == "📊 DASHBOARD":
    st.title("📊 DCK Tech Dashboard")
    df = load_data("Tickets")
    if not df.empty:
        # Tunjuk Untung Kasar di Dashboard
        df['Untung'] = pd.to_numeric(df['Kos'], errors='coerce').fillna(0) - pd.to_numeric(df['Kos_Part'], errors='coerce').fillna(0)
        total_profit = df['Untung'].sum()
        st.metric("Total Profit Keseluruhan", f"RM {total_profit:.2f}")
        
        search = st.text_input("🔍 Cari Nama/ID/Model:")
        if search:
            df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        st.dataframe(df[['ID','Customer','Model','Status','Kos']], use_container_width=True)

# --- DAFTAR TIKET ---
elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Daftar Tiket Baru")
    with st.form("daftar"):
        c1, c2 = st.columns([1,1]) # Layout Mobile Friendly
        nama = c1.text_input("Nama Pelanggan")
        phone = c2.text_input("No WhatsApp")
        model = c1.text_input("Model Peranti")
        masalah = st.text_area("Masalah")
        gambar = st.camera_input("Snap Gambar")
        tnc = st.checkbox("Setuju T&C (Data loss/3 bulan/Warranty part)")
        
        if st.form_submit_button("SIMPAN TIKET"):
            if nama and model and tnc:
                tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                img = cloudinary.uploader.upload(gambar)["secure_url"] if gambar else "No Image"
                # Row: ID, Tarikh, Nama, Phone, Model, SN, Pwd, Masalah, Status, KosPart, HargaJual, Image, Note
                add_row("Tickets", [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, "", "", masalah, "Pending", 0, 0, img, ""])
                st.success(f"Tiket {tid} Berjaya Disimpan!"); st.balloons()
            else: st.error("Isi Nama, Model & Tick T&C!")

# --- UPDATE STATUS ---
elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Technician")
    df_t = load_data("Tickets")
    if not df_t.empty:
        search_tid = st.text_input("🔍 Taip Ticket ID:")
        t_list = df_t['ID'].tolist()
        idx = t_list.index(search_tid) if search_tid in t_list else 0
        pilih_id = st.selectbox("Pilih Job:", t_list, index=idx)
        job = df_t[df_t['ID'] == pilih_id].iloc[0]
        
        st.warning(f"JOB: {job['ID']} | {job['Customer']}")
        
        with st.form("update_job"):
            col_a, col_b = st.columns(2)
            stat = col_a.selectbox("Status Terkini", ["Checking", "Waiting Part", "Repairing", "Done", "Collected"])
            note = st.text_area("Nota Kerosakan/Tindakan", value=str(job['Tech_Note']))
            
            st.divider()
            st.subheader("💰 Pengiraan Harga")
            # Sini Boss letak harga kos dan harga jual
            k_part = col_a.number_input("Kos Part (RM) - Modal Boss", value=float(job['Kos_Part']) if job['Kos_Part'] != "" else 0.0)
            h_jual = col_b.number_input("Harga Jual/Servis (RM) - Caj Customer", value=float(job['Kos']) if job['Kos'] != "" else 0.0)
            
            profit = h_jual - k_part
            st.write(f"**Untung Bersih Job Ini: RM {profit:.2f}**")
            
            if st.form_submit_button("KEMASKINI SEMUA"):
                if update_status_sheet(pilih_id, stat, note, h_jual, k_part):
                    st.success("Data Berjaya Dikemaskini!"); st.rerun()

        st.divider()
        # --- RUANGAN PARTS (FIXED) ---
        st.subheader("🔩 Alat Ganti (Parts)")
        df_p = load_data("Parts")
        if not df_p.empty:
            curr = df_p[df_p['TicketID'].astype(str) == str(pilih_id)]
            if not curr.empty: 
                st.table(curr[['NamaPart', 'Supplier', 'TarikhExpire']])
            else:
                st.info("Belum ada part direkodkan untuk job ini.")
        
        with st.expander("➕ Rekod Part Baru"):
            with st.form("rekod_p"):
                pn = st.text_input("Nama Part"); ps = st.text_input("Supplier"); pw = st.number_input("Warranty (Bulan)", 0)
                if st.form_submit_button("SIMPAN PART"):
                    exp = (datetime.now() + pd.DateOffset(months=pw)).strftime("%Y-%m-%d")
                    add_row("Parts", [f"P-{datetime.now().strftime('%M%S')}", pilih_id, pn, ps, datetime.now().strftime("%Y-%m-%d"), pw, exp])
                    st.success("Part Direkod!"); st.rerun()

# --- INVENTORY ---
elif menu == "📦 INVENTORY":
    st.title("📦 Log Parts")
    st.dataframe(load_data("Parts"), use_container_width=True)
