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
    except: st.error("Gagal simpan data!")

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
    c.setFont("Helvetica", 10); c.drawString(50, 785, "SERVICES & REPAIRS")
    c.line(50, 775, 550, 775)
    c.setFont("Helvetica-Bold", 14); c.drawString(50, 750, f"RESIT TIKET: {tiket_data['ID']}")
    c.setFont("Helvetica", 11)
    c.drawString(50, 720, f"Customer: {tiket_data['Customer']}"); c.drawString(300, 720, f"Model: {tiket_data['Model']}")
    c.drawString(50, 705, f"Phone: {tiket_data.get('Phone', '-')}"); c.drawString(300, 705, f"Tarikh: {tiket_data.get('Tarikh', '-')}")
    c.drawString(50, 670, "Masalah:"); c.drawString(50, 655, f"{tiket_data['Masalah']}")
    c.line(50, 600, 550, 600)
    c.setFont("Helvetica-Bold", 10); c.drawString(50, 585, "TERMA & SYARAT:")
    tc = ["1. Data loss bukan tanggungjawab kedai.", "2. Hak milik kedai selepas 3 bulan jika tidak dituntut.", "3. Warranty hanya pada part yang diganti sahaja."]
    y = 570
    for line in tc: c.drawString(60, y, line); y -= 15
    c.setFont("Helvetica-Bold", 10); c.drawString(50, y-30, "Tandatangan Pelanggan: ________________________")
    c.save(); buffer.seek(0)
    return buffer

def switch_page(page_name, ticket_id=None):
    st.session_state.page = page_name
    st.session_state.selected_ticket = ticket_id
    st.rerun()

# --- SIDEBAR ---
st.sidebar.title("🚀 DCK TECH")
menu = st.sidebar.radio("NAVIGASI", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"], 
                        index=["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"].index(st.session_state.page))

# --- 1. DASHBOARD ---
if menu == "📊 DASHBOARD":
    st.title("📊 DCK Business Dashboard")
    df = load_data("Tickets")
    if not df.empty:
        # Convert numeric
        df['Harga_Jual'] = pd.to_numeric(df['Harga_Jual'], errors='coerce').fillna(0)
        df['Kos_Part'] = pd.to_numeric(df['Kos_Part'], errors='coerce').fillna(0)
        df['Profit'] = df['Harga_Jual'] - df['Kos_Part']

        # Statistik Kaunter
        st.subheader("⚙️ Status Kerja")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Pending ⏳", len(df[df['Status'] == 'Pending']))
        c2.metric("Repairing 🛠️", len(df[df['Status'].isin(['Checking', 'Repairing'])]))
        c3.metric("Done ✅", len(df[df['Status'] == 'Done']))
        c4.metric("Collected 📦", len(df[df['Status'] == 'Collected']))

        # Statistik Duit
        st.subheader("💰 Ringkasan Kewangan")
        d1, d2, d3 = st.columns(3)
        d1.metric("Total Jualan", f"RM {df['Harga_Jual'].sum():.2f}")
        d2.metric("Total Kos Part", f"RM {df['Kos_Part'].sum():.2f}")
        d3.metric("Untung Bersih", f"RM {df['Profit'].sum():.2f}", delta=f"{df['Profit'].sum():.2f}")

        st.divider()
        search = st.text_input("🔍 Cari Ticket/Customer/Model:")
        if search:
            df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        
        for index, row in df.iloc[::-1].iterrows():
            with st.expander(f"📌 {row['ID']} - {row['Customer']} [{row['Status']}]"):
                col_x, col_y = st.columns([4, 1])
                col_x.write(f"**Model:** {row['Model']} | **Masalah:** {row['Masalah']}")
                if col_y.button("🔧 Update", key=f"dash_{row['ID']}"):
                    switch_page("🔧 UPDATE STATUS", row['ID'])

# --- 2. DAFTAR TIKET ---
elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Pendaftaran Tiket Baru")
    with st.form("reg"):
        col1, col2 = st.columns(2)
        nama = col1.text_input("Nama Customer")
        phone = col2.text_input("WhatsApp")
        model = col1.text_input("Model Laptop/PC")
        sn = col2.text_input("S/N / Service Tag")
        masalah = st.text_area("Aduan Masalah")
        gambar = st.camera_input("Snap Gambar Kondisi")
        tnc = st.checkbox("Pelanggan setuju dengan T&C DCK Tech")
        
        if st.form_submit_button("SIMPAN & GENERATE TIKET"):
            if nama and model and tnc:
                tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                img = cloudinary.uploader.upload(gambar)["secure_url"] if gambar else "No Image"
                row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, sn, "", masalah, "Pending", 0, 0, img, ""]
                add_row("Tickets", row)
                st.session_state.last_ticket_pdf = {"ID": tid, "Customer": nama, "Model": model, "Phone": phone, "Masalah": masalah, "Tarikh": row[1]}
                st.success("Tiket Berjaya Disimpan!"); st.balloons()
            else: st.error("Lengkapkan maklumat & T&C!")

    if st.session_state.last_ticket_pdf:
        pdf = generate_pdf(st.session_state.last_ticket_pdf)
        st.download_button("📥 Download Resit PDF", pdf, f"Resit_{st.session_state.last_ticket_pdf['ID']}.pdf", "application/pdf")

# --- 3. UPDATE STATUS ---
elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Update & Print Semula")
    df_t = load_data("Tickets")
    if not df_t.empty:
        ids = df_t['ID'].tolist()
        target = st.session_state.selected_ticket if st.session_state.selected_ticket in ids else ids[0]
        pilih_id = st.selectbox("Cari ID Tiket:", ids, index=ids.index(target))
        
        job = df_t[df_t['ID'] == pilih_id].iloc[0]
        
        # Section Cetak Semula
        st.subheader("🖨️ Reprinter")
        pdf_data = {"ID": job['ID'], "Customer": job['Customer'], "Model": job['Model'], "Phone": job['Phone'], "Masalah": job['Masalah'], "Tarikh": job['Tarikh']}
        btn_pdf = generate_pdf(pdf_data)
        st.download_button("📥 Cetak Semula Resit Ini", btn_pdf, f"Resit_{job['ID']}.pdf", "application/pdf")
        
        st.divider()
        if "http" in str(job['Image_Link']): st.image(job['Image_Link'], width=300, caption="Gambar Awal")

        with st.form("upd"):
            cA, cB = st.columns(2)
            stat = cA.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"], 
                               index=["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"].index(job['Status']) if job['Status'] in ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"] else 0)
            note = st.text_area("Nota Technician", value=str(job['Tech_Note']))
            k_mod = cA.number_input("Kos Modal Part (RM)", value=float(job['Kos_Part']) if job['Kos_Part'] != "" else 0.0)
            h_caj = cB.number_input("Harga Caj Customer (RM)", value=float(job['Harga_Jual']) if job['Harga_Jual'] != "" else 0.0)
            
            if st.form_submit_button("SIMPAN KEMASKINI"):
                if update_status_sheet(pilih_id, stat, note, h_caj, k_mod):
                    st.success("Updated!"); st.rerun()

        st.divider()
        st.subheader("🔩 Alat Ganti Digunakan")
        df_p = load_data("Parts")
        if not df_p.empty:
            curr = df_p[df_p['TicketID'].astype(str) == str(pilih_id)]
            if not curr.empty: st.table(curr[['NamaPart', 'Supplier', 'TarikhExpire']])
        
        with st.expander("Tambah Part Baru"):
            with st.form("p"):
                pn = st.text_input("Nama Part"); ps = st.text_input("Supplier"); pw = st.number_input("Warranty (Bln)", 0)
                if st.form_submit_button("REKOD PART"):
                    exp = (datetime.now() + pd.DateOffset(months=pw)).strftime("%Y-%m-%d")
                    add_row("Parts", [f"P-{datetime.now().strftime('%M%S')}", pilih_id, pn, ps, datetime.now().strftime("%Y-%m-%d"), pw, exp])
                    st.rerun()

elif menu == "📦 INVENTORY":
    st.title("📦 Inventory Log")
    st.dataframe(load_data("Parts"), use_container_width=True)
