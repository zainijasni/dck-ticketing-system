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
st.set_page_config(page_title="DCK Tech - Professional System", layout="wide")

# Cloudinary
cloudinary.config( 
  cloud_name = "dmnk9vbyj", 
  api_key = "382361922884785", 
  api_secret = "qZBl9CL1886_Ohvv_5JWPxzUQz8",
  secure = True
)

SHEET_ID = "1ssuZ3BzAih5goP5m_XsgAPjCj1OeDX_CdE--S0h-xek"

# --- LIST DATA (DARI BORANG BOSS) ---
LIST_MASALAH = ["Slow", "Screen Pecah", "Hinge Rosak", "Keyboard Rosak", "Tiada Display", "Tiada Power", "Tak Boleh Charge", "Auto Shutdown", "Format", "BSOD", "Water Damage", "Cleaning", "Upgrade RAM/SSD", "Battery Rosak"]
LIST_FIZIKAL = ["Calar Biasa", "Calar Teruk", "Skru Hilang", "Palm Rest Pecah", "Case Pecah", "I/O Port Rosak", "Keyboard Rosak"]
LIST_AKSESORI = ["Beg", "Charger", "Mouse", "Power Cable", "Lain-lain"]

# --- DATABASE ---
@st.cache_resource
def get_gspread_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(st.secrets["google_creds"], scopes=scope)
    return gspread.authorize(creds)

def load_data(tab_name):
    try:
        client = get_gspread_client()
        return pd.DataFrame(client.open_by_key(SHEET_ID).worksheet(tab_name).get_all_records())
    except: return pd.DataFrame()

def add_row(tab_name, row_data):
    try:
        get_gspread_client().open_by_key(SHEET_ID).worksheet(tab_name).append_row(row_data)
        st.cache_data.clear()
    except: st.error("Database Error")

# --- PDF GENERATOR (WITH PASSWORD & T&C) ---
def generate_service_form(t):
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    p.setFont("Helvetica-Bold", 20); p.drawString(50, h-50, "DCK TECH SERVICES")
    p.line(50, h-75, w-50, h-75)
    p.setFont("Helvetica-Bold", 12); p.drawString(50, h-100, f"TIKET ID: {t['ID']}")
    p.setFont("Helvetica", 10)
    p.drawString(50, h-120, f"Customer: {t['Customer']} ({t['Phone']})")
    p.drawString(300, h-120, f"Model: {t['Model']}")
    p.drawString(50, h-135, f"Password: {t['Password']}") # Password muncul kat resit
    p.drawString(300, h-135, f"S/N: {t['SN']}")
    
    y = h-170
    p.setFont("Helvetica-Bold", 10); p.drawString(50, y, "CHECKLIST PENERIMAAN:"); y-=15
    p.setFont("Helvetica", 9); p.drawString(60, y, f"Masalah: {t['Masalah']}"); y-=12
    p.drawString(60, y, f"Fizikal: {t['Fizikal']}"); y-=12
    p.drawString(60, y, f"Aksesori: {t['Aksesori']}"); y-=30
    
    p.setFont("Helvetica-Bold", 10); p.drawString(50, y, "TERMA & SYARAT:"); y-=15
    p.setFont("Helvetica", 8)
    tc = ["1. Data loss bukan tanggungjawab kedai.", "2. Tuntut peranti dalam tempoh 90 hari.", "3. Waranti hanya pada alat ganti baru."]
    for line in tc: p.drawString(60, y, line); y-=12
    
    p.drawString(50, y-50, "Tandatangan Pelanggan: ________________________")
    p.save(); buffer.seek(0)
    return buffer

# --- NAVIGATION ---
if 'page' not in st.session_state: st.session_state.page = "📊 DASHBOARD"
menu = st.sidebar.radio("MENU", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"], 
                        index=["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"].index(st.session_state.page))

# --- DASHBOARD ---
if menu == "📊 DASHBOARD":
    st.title("📊 Business Intelligence")
    df = load_data("Tickets")
    if not df.empty:
        # Statistik Jualan
        df['Harga_Jual'] = pd.to_numeric(df['Harga_Jual'], errors='coerce').fillna(0)
        df['Kos_Part'] = pd.to_numeric(df['Kos_Part'], errors='coerce').fillna(0)
        profit = df['Harga_Jual'].sum() - df['Kos_Part'].sum()
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Jualan", f"RM {df['Harga_Jual'].sum():.2f}")
        c2.metric("Profit", f"RM {profit:.2f}")
        c3.metric("Pending", len(df[df['Status'] == 'Pending']))
        
        st.divider()
        search = st.text_input("🔍 Cari Ticket/Customer/Model:")
        if search:
            df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        
        for _, row in df.iloc[::-1].iterrows():
            with st.expander(f"📌 {row['ID']} - {row['Customer']} ({row['Status']})"):
                st.write(f"**Model:** {row['Model']} | **Phone:** {row['Phone']}")
                if st.button("🔧 Update", key=f"d_{row['ID']}"):
                    st.session_state.page = "🔧 UPDATE STATUS"
                    st.session_state.selected_ticket = row['ID']
                    st.rerun()

# --- DAFTAR TIKET (BEAUTIFUL UI) ---
elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Borang Servis Baru")
    with st.container(border=True):
        # Seksyen 1: Pelanggan
        st.markdown("#### 👤 Maklumat Pelanggan")
        c1, c2 = st.columns(2)
        nama = c1.text_input("Nama Pelanggan")
        phone = c2.text_input("No. WhatsApp")
        
        # Seksyen 2: Peranti
        st.markdown("#### 💻 Maklumat Peranti")
        c3, c4, c5 = st.columns(3)
        model = c3.text_input("Model Device")
        sn = c4.text_input("Serial Number (S/N)")
        pwd = c5.text_input("Password / Pin", help="Penting untuk technician buat testing")
        
        # Seksyen 3: Checklist (DARI GAMBAR BOSS)
        st.markdown("#### 📋 Checklist Diagnosis")
        mslh = st.multiselect("Kenalpasti Masalah:", LIST_MASALAH)
        fiz = st.multiselect("Keadaan Fizikal:", LIST_FIZIKAL)
        acc = st.multiselect("Aksesori Diterima:", LIST_AKSESORI)
        custom_note = st.text_area("Nota Tambahan")
        
        gambar = st.camera_input("Snap Gambar Kondisi")
        tnc = st.checkbox("Pelanggan bersetuju dengan Terma & Syarat")

        if st.button("🚀 SIMPAN & CETAK TIKET", use_container_width=True):
            if nama and model and tnc:
                tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                img = cloudinary.uploader.upload(gambar)["secure_url"] if gambar else "No Image"
                
                # Row alignment: ID, Tarikh, Customer, Phone, Model, SN, PWD, Masalah, Fizikal, Aksesori, Status, Kos, Harga, Img, Note
                row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, sn, pwd, ", ".join(mslh), ", ".join(fiz), ", ".join(acc), "Pending", 0, 0, img, custom_note]
                add_row("Tickets", row)
                
                st.session_state.last_ticket_pdf = {"ID": tid, "Customer": nama, "Phone": phone, "Model": model, "SN": sn, "Password": pwd, "Masalah": ", ".join(mslh), "Fizikal": ", ".join(fiz), "Aksesori": ", ".join(acc), "Tarikh": row[1]}
                st.success("Tiket Berjaya Disimpan!"); st.rerun()

    if st.session_state.last_ticket_pdf:
        st.download_button("📥 DOWNLOAD SERVICE FORM", generate_service_form(st.session_state.last_ticket_pdf), f"Service_{st.session_state.last_ticket_pdf['ID']}.pdf", use_container_width=True)

# --- UPDATE STATUS ---
elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Update & Bill")
    df_t = load_data("Tickets")
    if not df_t.empty:
        ids = df_t['ID'].astype(str).tolist()
        target = str(st.session_state.selected_ticket) if st.session_state.selected_ticket else ids[-1]
        pilih_id = st.selectbox("Pilih Job:", ids, index=ids.index(target) if target in ids else 0)
        job = df_t[df_t['ID'].astype(str) == str(pilih_id)].iloc[0]
        
        st.success(f"**PELANGGAN:** {job['Customer']} | **MODEL:** {job['Model']} | **PWD:** {job['Password']}")
        
        col_img, col_form = st.columns([1, 2])
        with col_img:
            if "http" in str(job['Image_Link']): st.image(job['Image_Link'], use_container_width=True)
            st.download_button("🖨️ Cetak Semula Tiket", generate_service_form(job), f"Tiket_{job['ID']}.pdf")

        with col_form:
            # Auto-calculate cost based on Parts Tab
            df_p = load_data("Parts")
            curr_parts = df_p[df_p['TicketID'].astype(str) == str(pilih_id)]
            total_kos = pd.to_numeric(curr_parts['HargaBeli'], errors='coerce').sum() if 'HargaBeli' in curr_parts.columns else 0.0
            
            st.write(f"**Total Modal Part (Auto): RM {total_kos:.2f}**")
            
            with st.form("upd"):
                stat = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"], index=["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"].index(job['Status']) if job['Status'] in ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"] else 0)
                note = st.text_area("Nota Technician", value=str(job['Tech_Note']))
                h_jual = st.number_input("Harga Jual (Caj Customer)", value=float(job['Harga_Jual']) if job['Harga_Jual'] else 0.0)
                
                if st.form_submit_button("UPDATE"):
                    sheet = get_gspread_client().open_by_key(SHEET_ID).worksheet("Tickets")
                    cell = sheet.find(str(pilih_id))
                    if cell:
                        sheet.update_cell(cell.row, 11, stat)
                        sheet.update_cell(cell.row, 12, total_kos)
                        sheet.update_cell(cell.row, 13, h_jual)
                        sheet.update_cell(cell.row, 15, note)
                        st.cache_data.clear(); st.success("Updated!"); st.rerun()

        st.divider()
        st.subheader("Alat Ganti (Input Modal)")
        if not curr_parts.empty: st.table(curr_parts[['NamaPart', 'Supplier', 'HargaBeli']])
        
        with st.expander("Tambah Part & Kos Modal"):
            with st.form("ap"):
                pn = st.text_input("Part"); ps = st.text_input("Supplier"); hb = st.number_input("Harga Beli (Modal)", 0.0)
                if st.form_submit_button("SAVE PART"):
                    add_row("Parts", [f"P-{datetime.now().strftime('%M%S')}", pilih_id, pn, ps, datetime.now().strftime("%Y-%m-%d"), 0, "", hb])
                    st.rerun()

elif menu == "📦 INVENTORY":
    st.title("📦 Log Parts")
    st.dataframe(load_data("Parts"), use_container_width=True)
