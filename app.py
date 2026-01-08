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
import time

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="DCK Tech System", layout="wide")

cloudinary.config( 
  cloud_name = "dmnk9vbyj", 
  api_key = "382361922884785", 
  api_secret = "qZBl9CL1886_Ohvv_5JWPxzUQz8",
  secure = True
)

SHEET_ID = "1ssuZ3BzAih5goP5m_XsgAPjCj1OeDX_CdE--S0h-xek"

# --- LIST DATA ---
LIST_MASALAH = ["Slow", "Screen Pecah", "Hinge Rosak", "Keyboard Rosak", "Tiada Display", "Tiada Power", "Tak Boleh Charge", "Format", "Upgrade RAM/SSD", "Lain-lain"]
LIST_FIZIKAL = ["Calar Biasa", "Calar Teruk", "Skru Hilang", "Case Pecah", "I/O Port Rosak", "Sempurna"]
LIST_AKSESORI = ["Beg", "Charger", "Mouse", "Tiada"]

# --- 2. DATABASE ENGINE (SMART & RETRY) ---
@st.cache_resource
def get_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(st.secrets["google_creds"], scopes=scope)
    return gspread.authorize(creds)

def init_db():
    """Jalan SEKALI sahaja bila app start untuk jimat kuota API"""
    if 'db_checked' in st.session_state: return
    
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        
        # Tickets Header
        try: ws = sh.worksheet("Tickets")
        except: ws = sh.add_worksheet("Tickets", 1000, 20)
        header_t = ["ID", "Tarikh", "Customer", "Phone", "Model", "SN", "Password", "Masalah", "Fizikal", "Aksesori", "Status", "Kos_Part", "Harga_Jual", "Image_Link", "Tech_Note"]
        if ws.row_values(1) != header_t: ws.update("A1:O1", [header_t])
            
        # Parts Header
        try: ws_p = sh.worksheet("Parts")
        except: ws_p = sh.add_worksheet("Parts", 1000, 10)
        header_p = ["ID", "TicketID", "NamaPart", "Supplier", "TarikhMasuk", "WarrantyBulan", "TarikhExpire", "HargaBeli"]
        if ws_p.row_values(1) != header_p: ws_p.update("A1:H1", [header_p])
        
        st.session_state.db_checked = True
    except: pass

def robust_api_call(func, *args, **kwargs):
    """Fungsi Sabar: Kalau error, tunggu dan cuba lagi"""
    for i in range(3): # Cuba 3 kali
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if "Quota exceeded" in str(e) or "APIError" in str(e):
                time.sleep(2) # Tunggu 2 saat
                continue
            else:
                return None
    return None

def load_data(tab_name):
    init_db() # Cuma check flag, tak panggil API kalau dah check
    client = get_client()
    try:
        sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
        data = robust_api_call(sheet.get_all_records)
        return pd.DataFrame(data) if data else pd.DataFrame()
    except: return pd.DataFrame()

def add_row(tab_name, row):
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
    robust_api_call(sheet.append_row, row)
    st.cache_data.clear()

# --- 3. PDF GENERATOR ---
def generate_pdf(t, type="SERVICE"):
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    
    p.setFont("Helvetica-Bold", 22); p.drawString(50, h-50, "DCK TECH SERVICES")
    p.setFont("Helvetica", 10); p.drawString(50, h-65, "Resit & Borang Penerimaan Servis")
    p.line(50, h-75, w-50, h-75)
    
    p.setFont("Helvetica", 10)
    p.drawString(50, h-100, f"TIKET ID: {t.get('ID', '-')}")
    p.drawString(300, h-100, f"Tarikh: {t.get('Tarikh', '-')}")
    p.drawString(50, h-115, f"Nama: {t.get('Customer', '-')}")
    p.drawString(300, h-115, f"No HP: {t.get('Phone', '-')}")
    p.drawString(50, h-130, f"Model: {t.get('Model', '-')}")
    p.drawString(300, h-130, f"Serial No: {t.get('SN', '-')}")
    p.setFont("Helvetica-Bold", 10)
    p.drawString(50, h-150, f"PASSWORD: {t.get('Password', 'Tiada')}")
    
    y = h-180
    p.line(50, y+10, w-50, y+10)
    p.drawString(50, y, "DIAGNOSIS:"); y-=15
    p.setFont("Helvetica", 10)
    p.drawString(50, y, f"Masalah: {t.get('Masalah', '-')}"); y-=12
    p.drawString(50, y, f"Fizikal: {t.get('Fizikal', '-')}"); y-=12
    p.drawString(50, y, f"Aksesori: {t.get('Aksesori', '-')}"); y-=30
    
    if type == "INVOICE":
        p.setFont("Helvetica-Bold", 14)
        p.drawString(50, y, f"TOTAL: RM {float(t.get('Harga_Jual', 0)):.2f}"); y-=30
    
    p.setFont("Helvetica-Bold", 10); p.drawString(50, y, "TERMA & SYARAT:"); y-=15
    tc = ["1. Data hilang bukan tanggungjawab kedai.", "2. Barang tak tuntut > 3 bulan jadi hak milik kedai.", "3. Warranty sparepart sahaja."]
    p.setFont("Helvetica", 8)
    for line in tc: p.drawString(50, y, line); y-=12
    
    y -= 30
    p.drawString(50, y, "Tandatangan Pelanggan:"); p.drawString(300, y, "Tandatangan Admin:")
    p.line(50, y-30, 200, y-30); p.line(300, y-30, 450, y-30)
    p.save(); buffer.seek(0)
    return buffer

# --- 4. NAVIGATION ---
if 'page' not in st.session_state: st.session_state.page = "📊 DASHBOARD"
if 'selected_id' not in st.session_state: st.session_state.selected_id = None

menu = st.sidebar.radio("MENU", ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"], 
                        index=["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY"].index(st.session_state.page))

# --- DASHBOARD ---
if menu == "📊 DASHBOARD":
    st.title("📊 DCK Tech Dashboard")
    df = load_data("Tickets")
    
    if not df.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.info(f"PENDING: {len(df[df['Status'] == 'Pending'])}")
        c2.warning(f"CHECKING: {len(df[df['Status'] == 'Checking'])}")
        c3.success(f"DONE: {len(df[df['Status'] == 'Done'])}")
        c4.error(f"COLLECTED: {len(df[df['Status'] == 'Collected'])}")
        
        df['Harga_Jual'] = pd.to_numeric(df['Harga_Jual'], errors='coerce').fillna(0)
        df['Kos_Part'] = pd.to_numeric(df['Kos_Part'], errors='coerce').fillna(0)
        profit = df['Harga_Jual'].sum() - df['Kos_Part'].sum()
        
        st.divider()
        k1, k2 = st.columns(2)
        k1.metric("Total Sales", f"RM {df['Harga_Jual'].sum():.2f}")
        k2.metric("Total Profit", f"RM {profit:.2f}")
        
        st.divider()
        search = st.text_input("🔍 Cari Ticket:")
        if search: df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        
        st.write("Senarai Job Terkini:")
        for _, row in df.iloc[::-1].iterrows():
            with st.expander(f"{row['ID']} - {row['Customer']} ({row['Status']})"):
                st.write(f"Model: {row['Model']} | Masalah: {row['Masalah']}")
                if st.button("🔧 Manage Job", key=f"btn_{row['ID']}"):
                    st.session_state.page = "🔧 UPDATE STATUS"
                    st.session_state.selected_id = row['ID']
                    st.rerun()

# --- DAFTAR TIKET ---
elif menu == "📝 DAFTAR TIKET":
    st.title("📝 Tiket Masuk Baru")
    with st.container(border=True):
        c1, c2 = st.columns(2)
        nama = c1.text_input("Nama"); phone = c2.text_input("Phone")
        c3, c4, c5 = st.columns(3)
        model = c3.text_input("Model"); sn = c4.text_input("Serial No"); pwd = c5.text_input("Password Device")
        
        st.subheader("Checklist")
        mslh = st.multiselect("Masalah", LIST_MASALAH)
        fiz = st.multiselect("Fizikal", LIST_FIZIKAL)
        acc = st.multiselect("Aksesori", LIST_AKSESORI)
        note = st.text_area("Nota Tambahan")
        img = st.camera_input("Gambar")
        tnc = st.checkbox("Setuju T&C")
        
        if st.button("SIMPAN REKOD", use_container_width=True):
            if nama and tnc:
                tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                u_img = cloudinary.uploader.upload(img)["secure_url"] if img else ""
                row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, model, sn, pwd, ", ".join(mslh), ", ".join(fiz), ", ".join(acc), "Pending", 0, 0, u_img, note]
                add_row("Tickets", row)
                st.success("Berjaya!")
                st.session_state.last_pdf = {"ID": tid, "Customer": nama, "Phone": phone, "Model": model, "SN": sn, "Password": pwd, "Masalah": ", ".join(mslh), "Fizikal": ", ".join(fiz), "Aksesori": ", ".join(acc), "Tarikh": row[1]}
                st.rerun()
                
    if 'last_pdf' in st.session_state:
        st.download_button("📥 Download PDF Tiket", generate_pdf(st.session_state.last_pdf, "SERVICE"), "Tiket.pdf")

# --- UPDATE STATUS ---
elif menu == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Technician")
    df = load_data("Tickets")
    if not df.empty:
        ids = df['ID'].astype(str).tolist()
        tgt = str(st.session_state.selected_id) if st.session_state.selected_id else ids[-1]
        pid = st.selectbox("Pilih Job:", ids, index=ids.index(tgt) if tgt in ids else 0)
        
        job = df[df['ID'].astype(str) == str(pid)].iloc[0]
        st.info(f"CUSTOMER: {job['Customer']} | MODEL: {job['Model']} | PWD: {job['Password']}")
        
        c1, c2 = st.columns([1, 2])
        with c1:
            if str(job['Image_Link']).startswith("http"): st.image(job['Image_Link'])
            else: st.caption("Tiada Gambar")
            st.download_button("Print Tiket Asal", generate_pdf(job, "SERVICE"), f"Tiket_{pid}.pdf")
            
        with c2:
            df_p = load_data("Parts")
            parts = df_p[df_p['TicketID'].astype(str) == str(pid)]
            total_kos = pd.to_numeric(parts['HargaBeli'], errors='coerce').sum() if not parts.empty else 0
            
            st.write(f"**Kos Modal Part: RM {total_kos:.2f}**")
            
            with st.form("upd"):
                stt = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Done", "Collected"], index=["Pending", "Checking", "Waiting Part", "Done", "Collected"].index(job['Status']) if job['Status'] in ["Pending", "Checking", "Waiting Part", "Done", "Collected"] else 0)
                nt = st.text_area("Nota Tech", value=job['Tech_Note'])
                hj = st.number_input("Harga Jual (Total Bill)", value=float(job['Harga_Jual']))
                
                if st.form_submit_button("UPDATE"):
                    sheet = get_client().open_by_key(SHEET_ID).worksheet("Tickets")
                    cl = robust_api_call(sheet.find, str(pid)) # Guna robust call
                    if cl:
                        robust_api_call(sheet.update_cell, cl.row, 11, stt)
                        robust_api_call(sheet.update_cell, cl.row, 12, total_kos)
                        robust_api_call(sheet.update_cell, cl.row, 13, hj)
                        robust_api_call(sheet.update_cell, cl.row, 15, nt)
                        st.cache_data.clear()
                        st.success("Updated!"); st.rerun()
            
            if stt == "Done" or stt == "Collected":
                st.download_button("🖨️ PRINT INVOICE", generate_pdf(job, "INVOICE"), "Resit.pdf")

        st.divider()
        st.write("Parts List:")
        if not parts.empty: st.table(parts[['NamaPart', 'Supplier', 'HargaBeli']])
        
        with st.expander("Tambah Part"):
            with st.form("add_p"):
                pn = st.text_input("Part"); ps = st.text_input("Supplier"); hb = st.number_input("Harga Beli", 0.0)
                if st.form_submit_button("Simpan"):
                    add_row("Parts", [f"P-{int(time.time())}", pid, pn, ps, str(datetime.now().date()), 0, "", hb])
                    st.rerun()

# --- INVENTORY ---
elif menu == "📦 INVENTORY":
    st.title("📦 Inventory")
    df_p = load_data("Parts")
    if not df_p.empty:
        for i, row in df_p.iterrows():
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                c1.write(f"**{row['NamaPart']}** (Supp: {row['Supplier']}) - Ticket: {row['TicketID']}")
                if c2.button("Go to Job", key=f"inv_{i}"):
                    st.session_state.page = "🔧 UPDATE STATUS"
                    st.session_state.selected_id = row['TicketID']
                    st.rerun()
