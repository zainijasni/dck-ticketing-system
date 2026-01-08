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
st.set_page_config(page_title="DCK Tech - Pro System", layout="wide")

cloudinary.config( 
  cloud_name = "dmnk9vbyj", 
  api_key = "382361922884785", 
  api_secret = "qZBl9CL1886_Ohvv_5JWPxzUQz8",
  secure = True
)

SHEET_ID = "1ssuZ3BzAih5goP5m_XsgAPjCj1OeDX_CdE--S0h-xek"

LIST_MASALAH = ["Slow", "Screen Pecah", "Hinge Rosak", "Keyboard Rosak", "Tiada Display", "Tiada Power", "Tak Boleh Charge", "Format", "Upgrade RAM/SSD"]
LIST_FIZIKAL = ["Calar Biasa", "Calar Teruk", "Skru Hilang", "Case Pecah", "I/O Port Rosak"]
LIST_AKSESORI = ["Beg", "Charger", "Mouse", "Lain-lain"]

# --- DATABASE ---
@st.cache_resource
def get_gspread_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(st.secrets["google_creds"], scopes=scope)
    return gspread.authorize(creds)

def load_data(tab_name):
    try:
        client = get_gspread_client()
        sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
        # Ambil data tanpa convert jadi Dictionary terus untuk elak KeyError nama kolum
        data = sheet.get_all_values()
        if len(data) > 1:
            return pd.DataFrame(data[1:], columns=data[0])
        return pd.DataFrame()
    except: return pd.DataFrame()

def add_row(tab_name, row_data):
    try:
        get_gspread_client().open_by_key(SHEET_ID).worksheet(tab_name).append_row(row_data)
        st.cache_data.clear()
    except: st.error("Database Error")

# --- PDF GENERATOR ---
def generate_service_form(t):
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    p.setFont("Helvetica-Bold", 20); p.drawString(50, h-50, "DCK TECH SERVICES")
    p.line(50, h-75, w-50, h-75)
    p.setFont("Helvetica", 10)
    p.drawString(50, h-100, f"TIKET ID: {t.get('ID', 'N/A')}")
    p.drawString(50, h-115, f"Customer: {t.get('Customer', 'N/A')}")
    p.drawString(300, h-115, f"Model: {t.get('Model', 'N/A')}")
    p.drawString(50, h-130, f"Password: {t.get('Password', 'N/A')}")
    p.line(50, h-150, w-50, h-150)
    p.drawString(50, h-170, f"Masalah: {t.get('Masalah', '-')}")
    p.drawString(50, h-185, f"Fizikal: {t.get('Fizikal', '-')}")
    p.drawString(50, h-200, f"Aksesori: {t.get('Aksesori', '-')}")
    p.save(); buffer.seek(0)
    return buffer

# --- NAVIGATION ---
if 'page' not in st.session_state: st.session_state.page = "📊 DASHBOARD"
menu = st.sidebar.radio("MENU", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"], 
                        index=["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"].index(st.session_state.page))

# --- 1. DASHBOARD ---
if menu == "📊 DASHBOARD":
    st.title("📊 Dashboard Utama")
    df = load_data("Tickets")
    if not df.empty:
        # Gunakan iloc (nombor kolum) untuk keselamatan
        # I=8 (Status), J=9 (KosPart), K=10 (HargaJual) - ikut susunan 0,1,2...
        col_status = df.columns[10] # Status
        col_jual = df.columns[12]   # Harga_Jual
        
        c1, c2 = st.columns(2)
        total_sales = pd.to_numeric(df[col_jual], errors='coerce').sum()
        c1.metric("Total Jualan", f"RM {total_sales:.2f}")
        c2.metric("Total Jobs", len(df))
        
        st.divider()
        search = st.text_input("🔍 Cari Ticket/Customer:")
        if search:
            df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        
        for _, row in df.iloc[::-1].iterrows():
            with st.expander(f"📌 {row.iloc[0]} - {row.iloc[2]} ({row.iloc[10]})"):
                st.write(f"**Model:** {row.iloc[4]} | **Masalah:** {row.iloc[7]}")
                if st.button("🔧 Manage", key=f"d_{row.iloc[0]}"):
                    st.session_state.page = "🔧 UPDATE STATUS"
                    st.session_state.selected_ticket = row.iloc[0]
                    st.rerun()

# --- 2. DAFTAR TIKET ---
elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Pendaftaran Tiket Baru")
    with st.container(border=True):
        c1, c2 = st.columns(2)
        nama = c1.text_input("Nama Customer")
        phone = c2.text_input("WhatsApp")
        c3, c4, c5 = st.columns(3)
        model = c3.text_input("Model Device")
        sn = c4.text_input("S/N Tag")
        pwd = c5.text_input("Password")
        mslh = st.multiselect("Masalah:", LIST_MASALAH)
        fiz = st.multiselect("Fizikal:", LIST_FIZIKAL)
        acc = st.multiselect("Aksesori:", LIST_AKSESORI)
        note = st.text_area("Nota")
        gambar = st.camera_input("Snap")
        tnc = st.checkbox("Setuju T&C")

        if st.button("🚀 SIMPAN", use_container_width=True):
            if nama and model and tnc:
                tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                img = cloudinary.uploader.upload(gambar)["secure_url"] if gambar else "No Image"
                # Row: ID, Tarikh, Cust, Phone, Model, SN, PWD, Masalah, Fizikal, Aksesori, Status, Kos, Harga, Img, Note
                row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, sn, pwd, ", ".join(mslh), ", ".join(fiz), ", ".join(acc), "Pending", 0, 0, img, note]
                add_row("Tickets", row)
                st.session_state.last_ticket_pdf = {"ID": tid, "Customer": nama, "Phone": phone, "Model": model, "Password": pwd, "Masalah": ", ".join(mslh), "Fizikal": ", ".join(fiz), "Aksesori": ", ".join(acc), "Tarikh": row[1]}
                st.success("Tersimpan!"); st.rerun()

    if st.session_state.last_ticket_pdf:
        st.download_button("📥 DOWNLOAD PDF", generate_service_form(st.session_state.last_ticket_pdf), "ServiceForm.pdf")

# --- 3. UPDATE STATUS ---
elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Update Kerja")
    df_t = load_data("Tickets")
    if not df_t.empty:
        ids = df_t.iloc[:, 0].astype(str).tolist() # Kolum 0 = ID
        target = str(st.session_state.selected_ticket) if st.session_state.selected_ticket else ids[-1]
        pilih_id = st.selectbox("Pilih Job:", ids, index=ids.index(target) if target in ids else 0)
        
        # Ambil row data
        job = df_t[df_t.iloc[:, 0].astype(str) == str(pilih_id)].iloc[0]
        
        # Display Info guna index supaya tak KeyError
        st.success(f"**PELANGGAN:** {job.iloc[2]} | **MODEL:** {job.iloc[4]} | **PWD:** {job.iloc[6]}")
        
        col_img, col_form = st.columns([1, 2])
        with col_img:
            if "http" in str(job.iloc[13]): st.image(job.iloc[13], use_container_width=True)
            st.download_button("🖨️ Cetak Tiket", generate_service_form({
                "ID": job.iloc[0], "Customer": job.iloc[2], "Model": job.iloc[4], 
                "Password": job.iloc[6], "Masalah": job.iloc[7], "Fizikal": job.iloc[8], "Aksesori": job.iloc[9], "Tarikh": job.iloc[1]
            }), "Tiket.pdf")

        with col_form:
            # Auto-calculate cost
            df_p = load_data("Parts")
            total_kos = 0.0
            if not df_p.empty:
                curr_parts = df_p[df_p.iloc[:, 1].astype(str) == str(pilih_id)]
                if not curr_parts.empty:
                    # Ambil kolum terakhir (HargaBeli)
                    total_kos = pd.to_numeric(curr_parts.iloc[:, -1], errors='coerce').sum()
                    st.table(curr_parts.iloc[:, [2, 3, 7]]) # NamaPart, Supp, HargaBeli
            
            st.write(f"**Total Modal Part (Auto): RM {total_kos:.2f}**")
            
            with st.form("upd_form"):
                # Indeks kolum: 10=Status, 12=HargaJual, 14=Note
                stat = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"], 
                                   index=["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"].index(job.iloc[10]) if job.iloc[10] in ["Pending", "Checking", "Waiting Part", "Repairing", "Done", "Collected"] else 0)
                note = st.text_area("Nota Tech", value=str(job.iloc[14]))
                h_jual = st.number_input("Harga Jual (Caj Customer)", value=float(job.iloc[12]) if job.iloc[12] else 0.0)
                
                if st.form_submit_button("UPDATE"):
                    sheet = get_gspread_client().open_by_key(SHEET_ID).worksheet("Tickets")
                    cell = sheet.find(str(pilih_id))
                    if cell:
                        sheet.update_cell(cell.row, 11, stat)
                        sheet.update_cell(cell.row, 12, total_kos)
                        sheet.update_cell(cell.row, 13, h_jual)
                        sheet.update_cell(cell.row, 15, note)
                        st.cache_data.clear(); st.success("Success!"); st.rerun()

        with st.expander("➕ Tambah Part & Kos"):
            with st.form("ap"):
                pn = st.text_input("Part"); ps = st.text_input("Supplier"); hb = st.number_input("Modal", 0.0)
                if st.form_submit_button("SIMPAN PART"):
                    add_row("Parts", [f"P-{datetime.now().strftime('%M%S')}", pilih_id, pn, ps, datetime.now().strftime("%Y-%m-%d"), 0, "", hb])
                    st.rerun()

elif menu == "📦 INVENTORY":
    st.title("📦 Log Parts")
    st.dataframe(load_data("Parts"), use_container_width=True)
