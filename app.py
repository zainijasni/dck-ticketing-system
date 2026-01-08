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

# --- 2. SESSION STATE SETUP (PENTING UNTUK TELEPORT) ---
if 'page' not in st.session_state: st.session_state.page = "📊 DASHBOARD"
if 'selected_id' not in st.session_state: st.session_state.selected_id = None

# --- LIST DATA ---
LIST_MASALAH = ["Slow", "Screen Pecah", "Hinge Rosak", "Keyboard Rosak", "Tiada Display", "Tiada Power", "Tak Boleh Charge", "Format", "Upgrade RAM/SSD", "Lain-lain"]
LIST_FIZIKAL = ["Calar Biasa", "Calar Teruk", "Skru Hilang", "Case Pecah", "I/O Port Rosak", "Sempurna"]
LIST_AKSESORI = ["Beg", "Charger", "Mouse", "Tiada"]

# --- 3. DATABASE ENGINE (ROBUST) ---
@st.cache_resource
def get_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(st.secrets["google_creds"], scopes=scope)
    return gspread.authorize(creds)

def robust_api_call(func, *args, **kwargs):
    """Cuba 3 kali kalau Google block traffic"""
    for i in range(3):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if "Quota" in str(e) or "APIError" in str(e):
                time.sleep(2)
                continue
            return None
    return None

def init_db():
    if 'db_checked' in st.session_state: return
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        
        # Tickets Setup
        try: ws = sh.worksheet("Tickets")
        except: ws = sh.add_worksheet("Tickets", 1000, 20)
        header_t = ["ID", "Tarikh", "Customer", "Phone", "Model", "SN", "Password", "Masalah", "Fizikal", "Aksesori", "Status", "Kos_Part", "Harga_Jual", "Image_Link", "Tech_Note"]
        if ws.row_values(1) != header_t: ws.update("A1:O1", [header_t])
            
        # Parts Setup
        try: ws_p = sh.worksheet("Parts")
        except: ws_p = sh.add_worksheet("Parts", 1000, 10)
        header_p = ["ID", "TicketID", "NamaPart", "Supplier", "TarikhMasuk", "WarrantyBulan", "TarikhExpire", "HargaBeli"]
        if ws_p.row_values(1) != header_p: ws_p.update("A1:H1", [header_p])
        
        st.session_state.db_checked = True
    except: pass

def load_data(tab_name):
    init_db()
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

def delete_part(part_id):
    """Fungsi delete part based on ID"""
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet("Parts")
    cell = robust_api_call(sheet.find, str(part_id))
    if cell:
        robust_api_call(sheet.delete_rows, cell.row)
        st.cache_data.clear()
        return True
    return False

# --- 4. PDF GENERATOR ---
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
    
    p.drawString(50, y-50, "Tandatangan Pelanggan: _________________"); p.drawString(300, y-50, "Tandatangan Admin: _________________")
    p.save(); buffer.seek(0)
    return buffer

# --- 5. SIDEBAR NAVIGATION (LOGIC FIXED) ---
NAV_OPTIONS = ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY", "📈 LAPORAN"]

# Cari index semasa untuk paksa Sidebar ikut arahan kita
try:
    nav_index = NAV_OPTIONS.index(st.session_state.page)
except:
    nav_index = 0

menu = st.sidebar.radio("NAVIGASI", NAV_OPTIONS, index=nav_index)

# Jika user klik radio button manual, update session state
if menu != st.session_state.page:
    st.session_state.page = menu
    # Jangan rerun sini, biar dia flow ke bawah

# --- DASHBOARD ---
if st.session_state.page == "📊 DASHBOARD":
    st.title("📊 DCK Tech Dashboard")
    df = load_data("Tickets")
    
    if not df.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.info(f"PENDING: {len(df[df['Status'] == 'Pending'])}")
        c2.warning(f"CHECKING: {len(df[df['Status'] == 'Checking'])}")
        c3.success(f"DONE: {len(df[df['Status'] == 'Done'])}")
        c4.error(f"COLLECTED: {len(df[df['Status'] == 'Collected'])}")
        
        st.divider()
        search = st.text_input("🔍 Cari Ticket:")
        if search: df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        
        st.write("### Senarai Job Terkini")
        for _, row in df.iloc[::-1].iterrows():
            with st.expander(f"{row['ID']} - {row['Customer']} ({row['Status']})"):
                st.write(f"Model: {row['Model']} | Masalah: {row['Masalah']}")
                # --- FIX TELEPORT ---
                if st.button("🔧 Manage Job", key=f"btn_{row['ID']}"):
                    st.session_state.selected_id = row['ID']
                    st.session_state.page = "🔧 UPDATE STATUS"
                    st.rerun() # Paksa refresh segera

# --- DAFTAR TIKET ---
elif st.session_state.page == "📝 DAFTAR TIKET":
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

# --- UPDATE STATUS & BILLING ---
elif st.session_state.page == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Technician & Billing")
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
            st.download_button("Print Tiket Asal", generate_pdf(job, "SERVICE"), f"Tiket_{pid}.pdf")
            
        with c2:
            df_p = load_data("Parts")
            parts = df_p[df_p['TicketID'].astype(str) == str(pid)]
            # --- AUTO CALC KOS PART ---
            total_kos = pd.to_numeric(parts['HargaBeli'], errors='coerce').sum() if not parts.empty else 0
            
            st.markdown(f"### 💰 KOS MODAL: RM {total_kos:.2f}")
            
            with st.form("upd"):
                stt = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Done", "Collected"], index=["Pending", "Checking", "Waiting Part", "Done", "Collected"].index(job['Status']) if job['Status'] in ["Pending", "Checking", "Waiting Part", "Done", "Collected"] else 0)
                nt = st.text_area("Nota Tech", value=job['Tech_Note'])
                hj = st.number_input("Harga Jual (Total Bill)", value=float(job['Harga_Jual']))
                
                # Kira Untung Live
                if hj > 0:
                    untung = hj - total_kos
                    st.caption(f"Anggaran Untung Bersih: RM {untung:.2f}")

                if st.form_submit_button("UPDATE STATUS & HARGA"):
                    sheet = get_client().open_by_key(SHEET_ID).worksheet("Tickets")
                    cl = robust_api_call(sheet.find, str(pid))
                    if cl:
                        robust_api_call(sheet.update_cell, cl.row, 11, stt)
                        robust_api_call(sheet.update_cell, cl.row, 12, total_kos)
                        robust_api_call(sheet.update_cell, cl.row, 13, hj)
                        robust_api_call(sheet.update_cell, cl.row, 15, nt)
                        st.cache_data.clear()
                        st.success("Updated!"); st.rerun()
            
            if stt == "Done" or stt == "Collected":
                st.download_button("🖨️ PRINT INVOICE", generate_pdf(job, "INVOICE"), "Resit.pdf", use_container_width=True)

        st.divider()
        st.subheader("🔩 Pengurusan Parts (Edit/Delete)")
        if not parts.empty: 
            # Papar table biasa dulu
            st.table(parts[['ID', 'NamaPart', 'HargaBeli']])
            
            # --- RUANGAN EDIT/DELETE ---
            c_del, c_edit = st.columns(2)
            with c_del:
                del_id = st.selectbox("Pilih Part untuk Delete:", parts['ID'].tolist())
                if st.button("Hapus Part"):
                    if delete_part(del_id):
                        st.warning("Part Dihapus!"); st.rerun()
                        
        with st.expander("➕ Tambah Part Baru"):
            with st.form("add_p"):
                pn = st.text_input("Part"); ps = st.text_input("Supplier"); hb = st.number_input("Harga Beli (Modal)", 0.0)
                if st.form_submit_button("Simpan Part"):
                    add_row("Parts", [f"P-{int(time.time())}", pid, pn, ps, str(datetime.now().date()), 0, "", hb])
                    st.rerun()

# --- INVENTORY ---
elif st.session_state.page == "📦 INVENTORY":
    st.title("📦 Inventory Log")
    df_p = load_data("Parts")
    if not df_p.empty:
        for i, row in df_p.iterrows():
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                c1.write(f"**{row['NamaPart']}** (RM {row['HargaBeli']}) - Ticket: {row['TicketID']}")
                if c2.button("Go to Job", key=f"inv_{i}"):
                    st.session_state.selected_id = row['TicketID']
                    st.session_state.page = "🔧 UPDATE STATUS"
                    st.rerun()

# --- LAPORAN (NEW) ---
elif st.session_state.page == "📈 LAPORAN":
    st.title("📈 Laporan Prestasi Bisnes")
    df = load_data("Tickets")
    if not df.empty:
        # Convert Data
        df['Harga_Jual'] = pd.to_numeric(df['Harga_Jual'], errors='coerce').fillna(0)
        df['Kos_Part'] = pd.to_numeric(df['Kos_Part'], errors='coerce').fillna(0)
        df['Untung'] = df['Harga_Jual'] - df['Kos_Part']
        df['Tarikh'] = pd.to_datetime(df['Tarikh'], errors='coerce')
        
        # 1. Total Metrics
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Sales", f"RM {df['Harga_Jual'].sum():.2f}")
        m2.metric("Total Kos", f"RM {df['Kos_Part'].sum():.2f}")
        m3.metric("Total Untung", f"RM {df['Untung'].sum():.2f}")
        
        st.divider()
        
        # 2. Chart Sales per Bulan
        st.subheader("📊 Jualan Bulanan")
        if not df['Tarikh'].isnull().all():
            sales_chart = df.groupby(df['Tarikh'].dt.strftime('%Y-%m'))['Harga_Jual'].sum()
            st.bar_chart(sales_chart)
            
        # 3. Chart Masalah
        st.subheader("🔧 Masalah Paling Kerap (Top 5)")
        if 'Masalah' in df.columns:
            # Pecahkan masalah kalau ada koma
            all_prob = ",".join(df['Masalah'].astype(str).tolist()).split(",")
            # Bersihkan whitespace
            clean_prob = [p.strip() for p in all_prob if p.strip() != ""]
            prob_series = pd.Series(clean_prob).value_counts().head(5)
            st.bar_chart(prob_series)
    else:
        st.info("Belum ada data secukupnya.")
