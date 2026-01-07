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

# --- DATABASE ---
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

# --- PDF GENERATOR (WITH FULL T&C & SIGNATURE) ---
def generate_pdf(t):
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    p.setFont("Helvetica-Bold", 22); p.drawString(50, h-60, "DCK TECH")
    p.setFont("Helvetica", 10); p.drawString(50, h-75, "Solusi Komputer & Gadget Anda")
    p.line(50, h-85, w-50, h-85)
    
    p.setFont("Helvetica-Bold", 14); p.drawString(50, h-110, f"TIKET: {t.get('ID', 'N/A')}")
    p.setFont("Helvetica", 11)
    p.drawString(50, h-140, f"Customer: {t.get('Customer', 'N/A')}"); p.drawString(350, h-140, f"Tarikh: {t.get('Tarikh', '-')}")
    p.drawString(50, h-155, f"No. Tel: {t.get('Phone', '-')}"); p.drawString(350, h-155, f"Model: {t.get('Model', '-')}")
    
    p.setFont("Helvetica-Bold", 11); p.drawString(50, h-190, "ADUAN:")
    p.setFont("Helvetica", 11); p.drawString(60, h-205, f"{t.get('Masalah', '-')}")
    
    p.line(50, h-250, w-50, h-250)
    p.setFont("Helvetica-Bold", 10); p.drawString(50, h-270, "TERMA & SYARAT:")
    tc = [
        "1. Data loss bukan tanggungjawab kedai.",
        "2. Peranti tidak dituntut >90 hari menjadi hak milik kedai.",
        "3. Warranty sah untuk sparepart yang diganti sahaja."
    ]
    y = h-285
    for line in tc: p.drawString(60, y, line); y -= 15
    
    p.drawString(50, h-400, "Tandatangan Pelanggan: ________________________")
    p.drawString(350, h-400, "Tandatangan Kedai: ________________________")
    p.showPage(); p.save(); buffer.seek(0)
    return buffer

def switch_page(page, tid=None):
    st.session_state.page = page
    st.session_state.selected_ticket = tid
    st.rerun()

# --- SIDEBAR NAV ---
menu = st.sidebar.radio("MENU", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"], 
                        index=["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"].index(st.session_state.page))

# --- 1. DASHBOARD ---
if menu == "📊 DASHBOARD":
    st.title("📊 DCK Business Dashboard")
    df = load_data("Tickets")
    if not df.empty:
        # 1. STATISTIK STATUS (MUNCUL SEMULA)
        st.subheader("⚙️ Status Kerja")
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Pending ⏳", len(df[df['Status'] == 'Pending']))
        s2.metric("Checking 🛠️", len(df[df['Status'] == 'Checking']))
        s3.metric("Done ✅", len(df[df['Status'] == 'Done']))
        s4.metric("Collected 📦", len(df[df['Status'] == 'Collected']))

        # 2. KEWANGAN
        df['Harga_Jual'] = pd.to_numeric(df['Harga_Jual'], errors='coerce').fillna(0)
        df['Kos_Part'] = pd.to_numeric(df['Kos_Part'], errors='coerce').fillna(0)
        profit = df['Harga_Jual'].sum() - df['Kos_Part'].sum()
        st.subheader("💰 Kewangan")
        d1, d2 = st.columns(2)
        d1.metric("Total Sales", f"RM {df['Harga_Jual'].sum():.2f}")
        d2.metric("Total Profit", f"RM {profit:.2f}")

        st.divider()
        search = st.text_input("🔍 Cari (Nama/ID/Model):")
        if search:
            df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        
        for _, row in df.iloc[::-1].iterrows():
            with st.expander(f"📌 {row['ID']} - {row['Customer']} ({row['Status']})"):
                st.write(f"**Model:** {row['Model']} | **Masalah:** {row['Masalah']}")
                if st.button("🔧 Update", key=f"d_{row['ID']}"): switch_page("🔧 UPDATE STATUS", row['ID'])

# --- 2. DAFTAR TIKET (BEAUTIFUL INTERFACE) ---
elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Pendaftaran Tiket Baru")
    with st.container(border=True):
        st.subheader("👤 Maklumat Pelanggan")
        c1, c2 = st.columns(2)
        nama = c1.text_input("Nama Customer", placeholder="Nama Penuh")
        phone = c2.text_input("No. WhatsApp", placeholder="0123456789")
        
        st.subheader("💻 Peranti & Diagnosis")
        c3, c4 = st.columns(2)
        model = c3.text_input("Model / Jenama", placeholder="Contoh: Dell Latitude 5400")
        sn = c4.text_input("Serial Number", placeholder="S/N Tag")
        masalah = st.text_area("Masalah / Aduan Pelanggan")
        
        st.subheader("📷 Keadaan Peranti")
        gambar = st.camera_input("Snap Gambar Kondisi")
        
        st.warning("⚠️ Sila pastikan T&C telah diterangkan kepada pelanggan.")
        tnc = st.checkbox("Pelanggan bersetuju dengan Terma & Syarat DCK Tech.")

        if st.button("🚀 SIMPAN & TERBITKAN RESIT", use_container_width=True):
            if nama and model and tnc:
                tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                img = cloudinary.uploader.upload(gambar)["secure_url"] if gambar else "No Image"
                add_row("Tickets", [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, sn, "", masalah, "Pending", 0, 0, img, ""])
                st.session_state.last_ticket_pdf = {"ID": tid, "Customer": nama, "Model": model, "Phone": phone, "Masalah": masalah, "Tarikh": datetime.now().strftime("%Y-%m-%d")}
                st.success("Tiket Berjaya Disimpan!"); st.rerun()
            else: st.error("Lengkapkan borang & tanda T&C!")

    if st.session_state.last_ticket_pdf:
        st.divider()
        st.download_button("📥 DOWNLOAD RESIT PDF", generate_pdf(st.session_state.last_ticket_pdf), f"Resit_{st.session_state.last_ticket_pdf['ID']}.pdf", use_container_width=True)

# --- 3. UPDATE STATUS (WITH PARTS LIST) ---
elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Kerja Technician")
    df_t = load_data("Tickets")
    if not df_t.empty:
        ids = df_t['ID'].astype(str).tolist()
        target = str(st.session_state.selected_ticket) if st.session_state.selected_ticket else ids[-1]
        pilih_id = st.selectbox("Pilih Job:", ids, index=ids.index(target) if target in ids else 0)
        job = df_t[df_t['ID'].astype(str) == str(pilih_id)].iloc[0]
        
        st.success(f"**CUSTOMER:** {job['Customer']} | **MODEL:** {job['Model']}")
        
        col_x, col_y = st.columns([1, 2])
        with col_x:
            if "http" in str(job['Image_Link']): st.image(job['Image_Link'], use_container_width=True)
            st.download_button("🖨️ Cetak Semula Resit", generate_pdf(job), f"Resit_{job['ID']}.pdf")

        with col_y:
            with st.form("upd"):
                stat = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"], 
                                   index=["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"].index(job['Status']) if job['Status'] in ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"] else 0)
                note = st.text_area("Nota Tech", value=str(job['Tech_Note']))
                km = st.number_input("Kos Modal Part (RM)", value=float(job['Kos_Part']) if job['Kos_Part'] != "" else 0.0)
                hj = st.number_input("Harga Jual (RM)", value=float(job['Harga_Jual']) if job['Harga_Jual'] != "" else 0.0)
                if st.form_submit_button("SAVE UPDATE"):
                    if update_status_sheet(pilih_id, stat, note, hj, km): st.rerun()

        st.divider()
        st.subheader("🔩 Parts Terlibat")
        df_p = load_data("Parts")
        if not df_p.empty:
            curr_p = df_p[df_p['TicketID'].astype(str) == str(pilih_id)]
            if not curr_p.empty: st.table(curr_p[['NamaPart', 'Supplier', 'TarikhExpire']])
        
        with st.expander("➕ Tambah Rekod Part"):
            with st.form("ap"):
                pn = st.text_input("Nama Part"); ps = st.text_input("Supplier"); pw = st.number_input("Warranty (Bln)", 0)
                if st.form_submit_button("SIMPAN PART"):
                    exp = (datetime.now() + pd.DateOffset(months=pw)).strftime("%Y-%m-%d")
                    add_row("Parts", [f"P-{datetime.now().strftime('%M%S')}", pilih_id, pn, ps, datetime.now().strftime("%Y-%m-%d"), pw, exp])
                    st.rerun()

# --- 4. INVENTORY (WITH TELEPORT) ---
elif menu == "📦 INVENTORY":
    st.title("📦 Inventory Log")
    df_p = load_data("Parts")
    if not df_p.empty:
        for i, p_row in df_p.iloc[::-1].iterrows():
            with st.container(border=True):
                c1, c2, c3 = st.columns([2, 2, 1])
                c1.write(f"**Part:** {p_row.get('NamaPart', 'N/A')}\n**Supp:** {p_row.get('Supplier', 'N/A')}")
                c2.write(f"**Ticket:** {p_row.get('TicketID', 'N/A')}\n**Expire:** {p_row.get('TarikhExpire', 'N/A')}")
                if c3.button("🔧 Update Job", key=f"inv_btn_{i}"):
                    switch_page("🔧 UPDATE STATUS", p_row.get('TicketID'))
