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
# ✅ CLOUDINARY CONFIG (DCK TECH)
# ==========================================
cloudinary.config( 
  cloud_name = "dmnk9vbyj", 
  api_key = "382361922884785", 
  api_secret = "qZBl9CL1886_Ohvv_5JWPxzUQz8",
  secure = True
)

# ID Google Sheet Boss
SHEET_ID = "1ssuZ3BzAih5goP5m_XsgAPjCj1OeDX_CdE--S0h-xek"

def connect_google():
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file("secrets.json", scopes=scope)
        client = gspread.authorize(creds)
        return client
    except: return None

def upload_to_cloudinary(file_obj):
    try:
        upload_result = cloudinary.uploader.upload(file_obj)
        return upload_result["secure_url"]
    except Exception as e:
        return f"Error: {str(e)[:20]}"

def load_data(tab_name):
    try:
        client = connect_google()
        sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
        return pd.DataFrame(sheet.get_all_records())
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
            sheet.update_cell(row, 9, new_status)  # Status
            sheet.update_cell(row, 11, new_cost)   # Kos
            sheet.update_cell(row, 13, new_note)   # Note
            return True
        return False
    except: return False

def generate_pdf(tiket_data):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    
    # Header
    c.setFont("Helvetica-Bold", 25); c.drawString(50, 800, "DCK TECH")
    c.setFont("Helvetica", 10); c.drawString(50, 785, "COMPUTER REPAIR & SERVICES")
    c.line(50, 775, 550, 775)
    
    # Info Tiket
    c.setFont("Helvetica-Bold", 14); c.drawString(50, 750, f"TIKET REPAIR: {tiket_data['ID']}")
    c.setFont("Helvetica", 11)
    c.drawString(50, 720, f"Customer: {tiket_data['Customer']}")
    c.drawString(50, 705, f"Phone: {tiket_data['Phone']}")
    c.drawString(300, 720, f"Model: {tiket_data['Model']}")
    c.drawString(300, 705, f"Tarikh: {tiket_data['Tarikh']}")
    
    c.drawString(50, 670, "Masalah / Aduan:")
    c.setFont("Helvetica-Oblique", 11)
    c.drawString(50, 655, f"{tiket_data['Masalah']}")
    
    # --- T&C PENUH DALAM PDF ---
    c.line(50, 600, 550, 600)
    c.setFont("Helvetica-Bold", 10); c.drawString(50, 585, "TERMA & SYARAT (T&C):")
    c.setFont("Helvetica", 9)
    tc_list = [
        "1. Semua maklumat yang diberikan adalah benar dan sah.",
        "2. DCK Tech TIDAK bertanggungjawab ke atas sebarang kehilangan data peranti.",
        "3. Peranti yang tidak dituntut melebihi 3 bulan akan menjadi hak milik kedai.",
        "4. Warranty hanya sah untuk alat ganti (sparepart) yang baru dipasang sahaja.",
        "5. Kerosakan disebabkan kecuaian sendiri (masuk air, jatuh) akan membatalkan warranty.",
        "6. Bayaran deposit tidak akan dikembalikan jika pelanggan membatalkan repair."
    ]
    y = 570
    for line in tc_list:
        c.drawString(60, y, line)
        y -= 15
    
    c.setFont("Helvetica-Bold", 10); c.drawString(50, y-30, "Tandatangan Pelanggan: ________________________")
    c.save(); buffer.seek(0)
    return buffer

# --- MENU ---
menu = st.sidebar.radio("NAVIGASI", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"])

# --- 1. DASHBOARD ---
if menu == "📊 DASHBOARD":
    st.title("📊 Dashboard Live DCK Tech")
    df = load_data("Tickets")
    if not df.empty:
        c1, c2, c3, c4, c5 = st.columns(5)
        stats = ["Pending", "Checking", "Waiting Part", "Repairing Process", "Done"]
        for i, s in enumerate(stats):
            with [c1,c2,c3,c4,c5][i]: st.metric(s, len(df[df['Status'] == s]))
        st.divider()
        search = st.text_input("🔍 Cari (Nama/ID/Model):")
        if search:
            df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        st.dataframe(df[['ID','Tarikh','Customer','Model','Status','Kos']], use_container_width=True)

# --- 2. DAFTAR TIKET ---
elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Daftar Pelanggan Baru")
    if 'last_ticket' not in st.session_state: st.session_state['last_ticket'] = None
    with st.form("frm_reg"):
        st.subheader("👤 Maklumat Pelanggan")
        c1, c2 = st.columns(2); nama = c1.text_input("Nama"); phone = c2.text_input("WhatsApp")
        st.subheader("💻 Maklumat Peranti")
        c3, c4 = st.columns(2); model = c3.text_input("Model Laptop"); sn = c4.text_input("Serial No")
        pwd = st.text_input("Password Device"); masalah = st.text_area("Masalah")
        gambar = st.camera_input("Snap Gambar Device")
        st.subheader("⚖️ Persetujuan T&C")
        st.info("Sila pastikan pelanggan faham: Tiada ganti rugi data & hak milik kedai selepas 3 bulan.")
        t = st.checkbox("Saya bersetuju dengan Terma & Syarat DCK Tech.")
        if st.form_submit_button("SIMPAN & CETAK RESIT"):
            if not t or not nama: st.error("Sila isi nama & setuju T&C!")
            else:
                with st.spinner("Uploading Image..."):
                    tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                    img_url = upload_to_cloudinary(gambar) if gambar else "No Image"
                    row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, sn, pwd, masalah, "Pending", "-", 0.0, img_url, ""]
                    add_row("Tickets", row)
                    st.session_state['last_ticket'] = {"ID": tid, "Tarikh": row[1], "Customer": nama, "Phone": phone, "Model": model, "Masalah": masalah}
                    st.success("Berjaya!"); st.balloons()
    if st.session_state['last_ticket']:
        resit = generate_pdf(st.session_state['last_ticket'])
        st.download_button("🖨️ Download Resit PDF", resit, f"Resit_{st.session_state['last_ticket']['ID']}.pdf")

# --- 3. UPDATE STATUS ---
elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Kerja Technician")
    df_t = load_data("Tickets")
    if not df_t.empty:
        search_tid = st.text_input("🔍 Cari Ticket ID:")
        t_list = df_t['ID'].tolist()
        idx = t_list.index(search_tid) if search_tid in t_list else 0
        pilih_id = st.selectbox("Pilih Job:", t_list, index=idx)
        job = df_t[df_t['ID'] == pilih_id].iloc[0]
        st.info(f"CUSTOMER: {job['Customer']} | MODEL: {job['Model']}")
        if "http" in str(job['Image_Link']): st.image(job['Image_Link'], width=400, caption="Keadaan Device Asal")
        with st.form("upd"):
            stat = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Repairing Process", "Done", "Collected"])
            kos = st.number_input("Kos Repair", value=float(job['Kos']) if job['Kos'] != "" else 0.0)
            note = st.text_area("Nota Tech", value=str(job['Tech_Note']))
            if st.form_submit_button("UPDATE PROGRESS"):
                if update_status_sheet(pilih_id, stat, note, kos): st.success("Updated!"); st.rerun()
        st.divider(); st.subheader("🔩 Alat Ganti (Parts)")
        df_p = load_data("Parts")
        if not df_p.empty:
            curr = df_p[df_p['TicketID'].astype(str) == str(pilih_id)]
            if not curr.empty: st.table(curr[['NamaPart', 'Supplier', 'TarikhExpire']])
        with st.expander("➕ Tambah Rekod Part"):
            with st.form("ap"):
                pn = st.text_input("Nama Part"); ps = st.text_input("Supplier"); pw = st.number_input("Warranty (Bln)", 0)
                if st.form_submit_button("REKOD PART"):
                    exp = (datetime.now() + pd.DateOffset(months=pw)).strftime("%Y-%m-%d")
                    add_row("Parts", [f"P-{datetime.now().strftime('%M%S')}", pilih_id, pn, ps, datetime.now().strftime("%Y-%m-%d"), pw, exp])
                    st.rerun()

# --- 4. INVENTORY ---
elif menu == "📦 INVENTORY":
    st.title("📦 Log Alat Ganti & Warranty")
    df_p = load_data("Parts")
    if not df_p.empty:
        st.dataframe(df_p, use_container_width=True)