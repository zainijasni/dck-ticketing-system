import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import gspread
from google.oauth2.service_account import Credentials
import cloudinary
import cloudinary.uploader
import time
import urllib.parse
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="DCK One Stop System", layout="wide")

cloudinary.config( 
  cloud_name = "dmnk9vbyj", 
  api_key = "382361922884785", 
  api_secret = "qZBl9CL1886_Ohvv_5JWPxzUQz8",
  secure = True
)

SHEET_ID = "1ssuZ3BzAih5goP5m_XsgAPjCj1OeDX_CdE--S0h-xek"

# Initialize Session State
if 'page' not in st.session_state: st.session_state.page = "📊 DASHBOARD"
if 'selected_id' not in st.session_state: st.session_state.selected_id = None
if 'email_user' not in st.session_state: st.session_state.email_user = ""
if 'email_pass' not in st.session_state: st.session_state.email_pass = ""
# Default Config (Kalau Sheet Kosong)
if 'config' not in st.session_state:
    st.session_state.config = {
        "Company_Name": "DCK TECH SERVICES",
        "Address": "No 123, Jalan Gadget, 70000 Seremban",
        "Phone": "012-3456789",
        "Logo": "",
        "TNC_Ticket": "1. Data loss is customer's responsibility.\n2. Warranty on changed parts only.",
        "TNC_Invoice": "1. Goods sold are not returnable.\n2. Warranty 1 week for accessories."
    }

# --- 2. HELPER FUNCTIONS ---
def safe_float(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return 0.0
        clean = str(val).upper().replace("RM", "").replace(",", ".").strip()
        return float(clean)
    except: return 0.0

def safe_int(val):
    try: return int(float(val))
    except: return 0

def robust_api_call(func, *args, **kwargs):
    for i in range(3):
        try: return func(*args, **kwargs)
        except Exception: time.sleep(1); continue
    return None

def clean_phone_number_my(phone_input):
    p = str(phone_input).replace("-", "").replace(" ", "").replace("+", "").strip()
    if p.startswith("0"): return "6" + p
    elif p.startswith("1"): return "60" + p
    elif p.startswith("60"): return p
    else: return "6" + p

# --- 3. DATABASE ENGINE ---
@st.cache_resource
def get_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(st.secrets["google_creds"], scopes=scope)
    return gspread.authorize(creds)

def init_db():
    if 'db_checked' in st.session_state: return
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        
        # 1. Tickets
        try: ws = sh.worksheet("Tickets")
        except: ws = sh.add_worksheet("Tickets", 1000, 20)
        h_t = ["ID", "Tarikh", "Customer", "Phone", "Email", "Model", "SN", "Password", "Masalah", "Fizikal", "Aksesori", "Status", "Kos_Part", "Harga_Jual", "Image_Link", "Tech_Note"]
        if len(ws.row_values(1)) != len(h_t): ws.update("A1:P1", [h_t])
        
        # 2. Parts
        try: ws_p = sh.worksheet("Parts")
        except: ws_p = sh.add_worksheet("Parts", 1000, 10)
        h_p = ["ID", "TicketID", "NamaPart", "Supplier", "TarikhMasuk", "WarrantyBulan", "TarikhExpire", "HargaBeli"]
        if ws_p.row_values(1) != h_p: ws_p.update("A1:H1", [h_p])
        
        # 3. Sales (BARU: POS)
        try: ws_s = sh.worksheet("Sales")
        except: ws_s = sh.add_worksheet("Sales", 1000, 10)
        h_s = ["ID", "Tarikh", "Item", "Qty", "Harga_Unit", "Total", "Customer", "PaymentMethod"]
        if ws_s.row_values(1) != h_s: ws_s.update("A1:H1", [h_s])
        
        # 4. Config (BARU: SETTING)
        try: ws_c = sh.worksheet("Config")
        except: 
            ws_c = sh.add_worksheet("Config", 100, 2)
            # Default Data
            ws_c.append_row(["Key", "Value"])
            ws_c.append_row(["Company_Name", "DCK TECH SERVICES"])
            ws_c.append_row(["Address", "Alamat Kedai Anda"])
            ws_c.append_row(["Phone", "012-3456789"])
            ws_c.append_row(["TNC_Ticket", "1. Data customer tanggungjawab sendiri."])
            ws_c.append_row(["TNC_Invoice", "1. Barang dijual tidak boleh dikembalikan."])

        st.session_state.db_checked = True
    except: pass

def load_data(tab_name):
    init_db()
    client = get_client()
    try:
        sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
        data = robust_api_call(sheet.get_all_records)
        df = pd.DataFrame(data) if data else pd.DataFrame()
        if tab_name == "Tickets" and not df.empty and "Email" not in df.columns: df["Email"] = ""
        return df
    except: return pd.DataFrame()

def load_config():
    """Load settings dari Sheet Config"""
    try:
        df = load_data("Config")
        if not df.empty:
            config = dict(zip(df['Key'], df['Value']))
            # Update session state
            st.session_state.config.update(config)
    except: pass

def save_config_to_db(new_config):
    """Simpan setting ke Sheet"""
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet("Config")
    sheet.clear()
    sheet.append_row(["Key", "Value"])
    for k, v in new_config.items():
        sheet.append_row([k, v])
    st.session_state.config = new_config
    st.cache_data.clear()

def add_row(tab_name, row):
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
    robust_api_call(sheet.append_row, row)
    st.cache_data.clear()

def update_part_data(part_id, new_name, new_supp, new_price, new_warranty):
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet("Parts")
    cell = robust_api_call(sheet.find, str(part_id))
    if cell:
        robust_api_call(sheet.update_cell, cell.row, 3, new_name)
        robust_api_call(sheet.update_cell, cell.row, 4, new_supp)
        robust_api_call(sheet.update_cell, cell.row, 6, new_warranty)
        new_exp = (datetime.now() + pd.DateOffset(months=int(new_warranty))).strftime("%Y-%m-%d")
        robust_api_call(sheet.update_cell, cell.row, 7, new_exp)
        robust_api_call(sheet.update_cell, cell.row, 8, new_price)
        st.cache_data.clear()
        return True
    return False

def delete_part(part_id):
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet("Parts")
    cell = robust_api_call(sheet.find, str(part_id))
    if cell:
        robust_api_call(sheet.delete_rows, cell.row)
        st.cache_data.clear()
        return True
    return False

def update_customer_info_db(tid, nama, phone, email, model, sn, pwd, masalah):
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet("Tickets")
    cell = robust_api_call(sheet.find, str(tid))
    if cell:
        robust_api_call(sheet.update_cell, cell.row, 3, nama)
        robust_api_call(sheet.update_cell, cell.row, 4, phone)
        robust_api_call(sheet.update_cell, cell.row, 5, email)
        robust_api_call(sheet.update_cell, cell.row, 6, model)
        robust_api_call(sheet.update_cell, cell.row, 7, sn)
        robust_api_call(sheet.update_cell, cell.row, 8, pwd)
        robust_api_call(sheet.update_cell, cell.row, 9, masalah)
        st.cache_data.clear()
        return True
    return False

# --- 4. PDF GENERATOR (DYNAMIC CONFIG) ---
def generate_pdf(t, type="SERVICE"):
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    cfg = st.session_state.config
    
    # Header Dinamik (Dari Config)
    p.setFont("Helvetica-Bold", 20); p.drawString(50, h-50, cfg.get("Company_Name", "DCK TECH"))
    p.setFont("Helvetica", 10); p.drawString(50, h-65, cfg.get("Address", ""))
    p.drawString(50, h-78, f"Tel: {cfg.get('Phone', '')}")
    p.line(50, h-85, w-50, h-85)
    
    # Detail
    p.setFont("Helvetica", 10)
    p.drawString(50, h-110, f"REF ID: {t.get('ID', '-')}"); p.drawString(300, h-110, f"DATE: {t.get('Tarikh', '-')}")
    
    if 'Model' in t: # Kalau Service Ticket
        p.drawString(50, h-125, f"CUSTOMER: {t.get('Customer', '-')}")
        p.drawString(300, h-125, f"MODEL: {t.get('Model', '-')}")
        p.drawString(50, h-140, f"CONTACT: {t.get('Phone', '-')}")
        p.drawString(300, h-140, f"S/N: {t.get('SN', '-')}")
        
        y = h-170
        p.line(50, y+10, w-50, y+10)
        p.setFont("Helvetica-Bold", 10); p.drawString(50, y, "SERVICE DETAILS:"); y-=15
        p.setFont("Helvetica", 10)
        p.drawString(50, y, f"Problem: {t.get('Masalah', '-')}"); y-=15
        p.drawString(50, y, f"Condition: {t.get('Fizikal', '-')}"); y-=15
        p.drawString(50, y, f"Accessories: {t.get('Aksesori', '-')}"); y-=15
        p.drawString(50, y, f"Note: {t.get('Tech_Note', '-')}"); y-=30
        
        if type == "INVOICE":
            p.setFont("Helvetica-Bold", 14)
            p.drawString(50, y, f"TOTAL: RM {safe_float(t.get('Harga_Jual', 0)):.2f}"); y-=30
            tnc_text = cfg.get("TNC_Invoice", "")
        else:
            tnc_text = cfg.get("TNC_Ticket", "")

    else: # Kalau Sales Receipt (Barang Runcit)
        p.drawString(50, h-125, f"CUSTOMER: {t.get('Customer', 'Walk-in')}")
        y = h-160
        p.line(50, y+10, w-50, y+10)
        p.setFont("Helvetica-Bold", 10)
        p.drawString(50, y, "ITEM DESCRIPTION"); p.drawString(300, y, "QTY"); p.drawString(400, y, "PRICE"); y-=20
        p.setFont("Helvetica", 10)
        p.drawString(50, y, str(t.get('Item', '-'))); p.drawString(300, y, str(t.get('Qty', 0)))
        p.drawString(400, y, f"RM {safe_float(t.get('Harga_Unit', 0)):.2f}"); y-=30
        p.line(50, y+10, w-50, y+10)
        p.setFont("Helvetica-Bold", 14)
        p.drawString(300, y, "TOTAL:"); p.drawString(400, y, f"RM {safe_float(t.get('Total', 0)):.2f}"); y-=30
        tnc_text = cfg.get("TNC_Invoice", "")

    # T&C Dinamik
    p.setFont("Helvetica-Bold", 10); p.drawString(50, y, "TERMS & CONDITIONS:"); y-=15
    p.setFont("Helvetica", 8)
    # Split by enter
    for line in tnc_text.split('\n'):
        p.drawString(50, y, line.strip()); y-=12
    
    y -= 30
    p.drawString(50, y, "Customer Signature: _________________"); p.drawString(300, y, "Authorized Signature: _________________")
    p.save(); buffer.seek(0)
    return buffer

# --- 5. EMAIL & LINKS ---
def generate_message_content(data):
    d = {k: data.get(k, '-') for k in data}
    msg = f"Hai {d['Customer']},\n\nTerima kasih berurusan dengan {st.session_state.config.get('Company_Name')}."
    if d['Status'] == 'Pending': msg += "\nKami telah menerima peranti anda."
    elif d['Status'] in ['Done', 'Collected']: msg += "\n✅ Peranti SIAP."
    else: msg += f"\nStatus terkini: {d['Status']}"
    
    msg += f"\n\n--- BUTIRAN ---\nID: {d['ID']}\nModel: {d['Model']}\nMasalah: {d['Masalah']}\nNota: {d.get('Tech_Note', '-')}"
    if d['Status'] in ['Done', 'Collected']: msg += f"\n\n💰 TOTAL: RM {safe_float(d.get('Harga_Jual', 0)):.2f}"
    msg += "\n\nSekian,\nTeam DCK Tech"
    return msg

def generate_links(type, data):
    phone = clean_phone_number_my(data.get('Phone', ''))
    full_msg = generate_message_content(data)
    if type == "WA": return f"https://wa.me/{phone}?text={urllib.parse.quote(full_msg)}"
    elif type == "EMAIL": return f"mailto:{data.get('Email')}?subject=Status%20Tiket&body={urllib.parse.quote(full_msg)}"

def send_email_with_pdf(to_email, data, pdf_buffer, pdf_name):
    sender = st.session_state.email_user
    password = st.session_state.email_pass
    if not sender or not password: return False, "Sila set Email di Tetapan."
    msg = MIMEMultipart()
    msg['From'] = sender; msg['To'] = to_email; msg['Subject'] = f"Tiket: {data.get('ID')}"
    msg.attach(MIMEText(generate_message_content(data), 'plain'))
    part = MIMEApplication(pdf_buffer.getvalue(), Name=pdf_name)
    part['Content-Disposition'] = f'attachment; filename="{pdf_name}"'
    msg.attach(part)
    try:
        s = smtplib.SMTP('smtp.gmail.com', 587); s.starttls(); s.login(sender, password); s.send_message(msg); s.quit()
        return True, "Berjaya!"
    except Exception as e: return False, str(e)

# --- 6. NAVIGATION ---
# Load config dulu sebelum render page
load_config()

PAGES = ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🛒 JUALAN KEDAI", "🔧 UPDATE STATUS", "📦 INVENTORY", "📈 LAPORAN", "⚙️ TETAPAN"]
try: idx = PAGES.index(st.session_state.page)
except: idx = 0
sel = st.sidebar.radio("NAVIGASI", PAGES, index=idx)
if sel != st.session_state.page: st.session_state.page = sel; st.rerun()

# === PAGE: TETAPAN (CONFIG) ===
if st.session_state.page == "⚙️ TETAPAN":
    st.title("⚙️ Tetapan & Konfigurasi")
    
    # 1. Email Server
    with st.expander("📧 Tetapan Email Server", expanded=True):
        with st.form("set_email"):
            eu = st.text_input("Email Gmail", value=st.session_state.email_user)
            ep = st.text_input("App Password", value=st.session_state.email_pass, type="password")
            if st.form_submit_button("Simpan Email"):
                st.session_state.email_user = eu; st.session_state.email_pass = ep
                st.toast("Email Saved!", icon='✅')

    # 2. Company Info & TNC
    with st.expander("🏢 Maklumat Syarikat & T&C", expanded=True):
        with st.form("set_comp"):
            cn = st.text_input("Nama Syarikat", value=st.session_state.config.get("Company_Name"))
            ad = st.text_area("Alamat", value=st.session_state.config.get("Address"))
            ph = st.text_input("No. Telefon Kedai", value=st.session_state.config.get("Phone"))
            
            c1, c2 = st.columns(2)
            tnc_t = c1.text_area("T&C Tiket Masuk", value=st.session_state.config.get("TNC_Ticket"), height=150)
            tnc_i = c2.text_area("T&C Invois/Resit", value=st.session_state.config.get("TNC_Invoice"), height=150)
            
            if st.form_submit_button("💾 Simpan Konfigurasi"):
                new_conf = st.session_state.config.copy()
                new_conf.update({"Company_Name": cn, "Address": ad, "Phone": ph, "TNC_Ticket": tnc_t, "TNC_Invoice": tnc_i})
                save_config_to_db(new_conf)
                st.success("Konfigurasi dikemaskini! PDF akan guna data baru.")

# === PAGE: JUALAN KEDAI (POS) ===
elif st.session_state.page == "🛒 JUALAN KEDAI":
    st.title("🛒 Jualan Runcit (POS)")
    with st.container(border=True):
        c1, c2 = st.columns(2)
        item = c1.text_input("Nama Barang")
        cust = c2.text_input("Nama Pelanggan (Optional)", value="Walk-in")
        
        c3, c4, c5 = st.columns(3)
        qty = c3.number_input("Kuantiti", 1, 100, 1)
        price = c4.number_input("Harga Seunit (RM)", 0.0)
        total = qty * price
        c5.metric("TOTAL", f"RM {total:.2f}")
        
        pay = st.selectbox("Cara Bayaran", ["Cash", "QR/Transfer", "Card"])
        
        if st.button("✅ REKOD JUALAN", use_container_width=True):
            if item and total > 0:
                sid = f"SALE-{datetime.now().strftime('%d%H%M')}"
                # ID, Tarikh, Item, Qty, Unit, Total, Cust, Pay
                add_row("Sales", [sid, datetime.now().strftime("%Y-%m-%d"), item, qty, price, total, cust, pay])
                st.toast("Jualan Direkod!", icon='💰')
                # Auto generate receipt data
                st.session_state.last_sale = {"ID": sid, "Tarikh": datetime.now().strftime("%Y-%m-%d"), "Item": item, "Qty": qty, "Harga_Unit": price, "Total": total, "Customer": cust}
                time.sleep(1); st.rerun()
            else: st.error("Masukkan nama barang & harga.")

    if 'last_sale' in st.session_state:
        ls = st.session_state.last_sale
        st.divider(); st.success("Transaksi Berjaya!")
        st.download_button("🖨️ CETAK RESIT JUALAN", generate_pdf(ls, "SALES"), "Resit_Jualan.pdf", use_container_width=True)

# === PAGE: DASHBOARD ===
elif st.session_state.page == "📊 DASHBOARD":
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
            with st.expander(f"{row.get('ID', '-')} - {row.get('Customer', '-')} ({row.get('Status', '-')})"):
                st.write(f"Model: {row.get('Model', '-')} | Masalah: {row.get('Masalah', '-')}")
                if st.button("🔧 Manage Job", key=f"btn_{row.get('ID')}"):
                    st.session_state.selected_id = row.get('ID'); st.session_state.page = "🔧 UPDATE STATUS"; st.rerun()

# === PAGE: DAFTAR TIKET ===
elif st.session_state.page == "📝 DAFTAR TIKET":
    st.title("📝 Tiket Masuk Baru")
    with st.container(border=True):
        c1, c2, c3 = st.columns(3)
        nama = c1.text_input("Nama"); phone = c2.text_input("No HP"); email = c3.text_input("Email")
        c4, c5, c6 = st.columns(3)
        model = c4.text_input("Model"); sn = c5.text_input("S/N"); pwd = c6.text_input("Password")
        mslh = st.multiselect("Masalah", ["Slow", "Screen", "Battery", "Keyboard", "Format", "Other"])
        fiz = st.multiselect("Fizikal", ["Calar", "Pecah", "Sempurna"]); acc = st.multiselect("Aksesori", ["Bag", "Charger", "Mouse"])
        note = st.text_area("Nota"); img = st.file_uploader("Gambar", accept_multiple_files=True)
        tnc = st.checkbox("Setuju T&C")
        
        if st.button("SIMPAN", use_container_width=True):
            if nama and tnc:
                tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                img_str = "" # Simplified image handling for this snippet
                if img: 
                    urls = [cloudinary.uploader.upload(f)['secure_url'] for f in img]
                    img_str = ",".join(urls)
                
                row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, email, model, sn, pwd, ",".join(mslh), ",".join(fiz), ",".join(acc), "Pending", 0, 0, img_str, note]
                add_row("Tickets", row)
                st.toast("Tiket Dibuka!", icon='✅')
                st.session_state.last_data = {"ID": tid, "Customer": nama, "Phone": phone, "Email": email, "Model": model, "SN": sn, "Masalah": ",".join(mslh), "Fizikal": ",".join(fiz), "Aksesori": ",".join(acc), "Tech_Note": note, "Status": "Pending", "Images": img_str, "Image_Link": img_str, "Tarikh": row[1]}
                time.sleep(1); st.rerun()

    if 'last_data' in st.session_state:
        ld = st.session_state.last_data
        st.divider(); st.success(f"Tiket {ld['ID']} Dibuka!")
        c1, c2, c3 = st.columns(3)
        c1.download_button("📥 PDF", generate_pdf(ld, "SERVICE"), "Tiket.pdf", "application/pdf", use_container_width=True)
        c2.link_button("📱 WhatsApp", generate_links("WA", ld), use_container_width=True)
        if ld['Email']: 
            if c3.button("📧 Email (Auto)"):
                ok, m = send_email_with_pdf(ld['Email'], ld, generate_pdf(ld, "SERVICE"), "Tiket.pdf")
                if ok: st.toast(m)
                else: st.error(m)

# === PAGE: UPDATE STATUS ===
elif st.session_state.page == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Technician")
    df = load_data("Tickets")
    if not df.empty:
        ids = df['ID'].astype(str).tolist()
        tgt = str(st.session_state.selected_id) if st.session_state.selected_id else ids[-1]
        pid = st.selectbox("Pilih Job:", ids, index=ids.index(tgt) if tgt in ids else 0)
        job = df[df['ID'].astype(str) == str(pid)].iloc[0]
        
        with st.expander("ℹ️ MAKLUMAT & EDIT", expanded=True):
            edit = st.checkbox("Edit Info")
            if edit:
                with st.form("ed"):
                    nm = st.text_input("Nama", job.get('Customer')); ph = st.text_input("Phone", job.get('Phone'))
                    em = st.text_input("Email", job.get('Email')); md = st.text_input("Model", job.get('Model'))
                    sn = st.text_input("SN", job.get('SN')); pw = st.text_input("Pwd", job.get('Password'))
                    ms = st.text_area("Masalah", job.get('Masalah'))
                    if st.form_submit_button("Simpan"):
                        update_customer_info_db(pid, nm, ph, em, md, sn, pw, ms); st.rerun()
            else:
                st.write(f"**{job.get('Customer')}** | {job.get('Model')} | {job.get('Masalah')}")
                st.error(f"PWD: {job.get('Password')}")
                c1, c2 = st.columns(2)
                c1.link_button("WhatsApp", generate_links("WA", job), use_container_width=True)
                if job.get('Email'): 
                    if c2.button("Email Auto"): 
                        doc = "INVOICE" if job.get('Status') in ['Done', 'Collected'] else "SERVICE"
                        send_email_with_pdf(job.get('Email'), job, generate_pdf(job, doc), "Status.pdf")

        c_l, c_r = st.columns([1, 2])
        with c_l:
            st.download_button("Print Tiket", generate_pdf(job, "SERVICE"), "Tiket.pdf")
            im = str(job.get('Image_Link'))
            if im and len(im) > 5:
                for i in im.split(','): st.image(i)

        with c_r:
            df_p = load_data("Parts")
            parts = df_p[df_p['TicketID'].astype(str) == str(pid)]
            kos = sum([safe_float(x) for x in parts['HargaBeli']]) if not parts.empty else 0
            
            t1, t2 = st.tabs(["Status", "Parts"])
            with t1:
                st.write(f"**KOS:** RM {kos:.2f}")
                with st.form("up"):
                    s = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Done", "Collected"], index=["Pending", "Checking", "Waiting Part", "Done", "Collected"].index(job.get('Status','Pending')) if job.get('Status','Pending') in ["Pending", "Checking", "Waiting Part", "Done", "Collected"] else 0)
                    n = st.text_area("Nota", job.get('Tech_Note')); h = st.number_input("Harga Jual", safe_float(job.get('Harga_Jual')))
                    if st.form_submit_button("Update"):
                        sheet = get_client().open_by_key(SHEET_ID).worksheet("Tickets")
                        cl = robust_api_call(sheet.find, str(pid))
                        robust_api_call(sheet.update_cell, cl.row, 12, s); robust_api_call(sheet.update_cell, cl.row, 13, kos)
                        robust_api_call(sheet.update_cell, cl.row, 14, h); robust_api_call(sheet.update_cell, cl.row, 16, n)
                        st.cache_data.clear(); st.rerun()
                if s in ["Done", "Collected"]: st.download_button("Print Invois", generate_pdf(job, "INVOICE"), "Invois.pdf")

            with t2:
                with st.form("ap", clear_on_submit=True):
                    pn = st.text_input("Part"); ps = st.text_input("Supp"); w = st.number_input("Warr", 1); pr = st.number_input("Harga", 0.0)
                    if st.form_submit_button("Add"):
                        exp = (datetime.now() + pd.DateOffset(months=int(w))).strftime("%Y-%m-%d")
                        add_row("Parts", [f"P-{int(time.time())}", pid, pn, ps, str(datetime.now().date()), w, exp, pr]); st.rerun()
                if not parts.empty:
                    st.dataframe(parts[['ID', 'NamaPart', 'HargaBeli']])
                    d = st.selectbox("Del", ["-"]+parts['ID'].tolist())
                    if d != "-" and st.button("Delete"): delete_part(d); st.rerun()

# === PAGE: INVENTORY & LAPORAN ===
elif st.session_state.page == "📦 INVENTORY":
    st.title("📦 Inventory Log"); df = load_data("Parts"); st.dataframe(df)

elif st.session_state.page == "📈 LAPORAN":
    st.title("📈 Laporan Prestasi")
    df = load_data("Tickets"); df_s = load_data("Sales"); df_p = load_data("Parts")
    
    # KIRA TOTAL SALES (SERVICE + KEDAI)
    sales_serv = sum([safe_float(x) for x in df['Harga_Jual']]) if not df.empty else 0
    sales_shop = sum([safe_float(x) for x in df_s['Total']]) if not df_s.empty else 0
    kos_part = sum([safe_float(x) for x in df['Kos_Part']]) if not df.empty else 0
    total_rev = sales_serv + sales_shop
    total_prof = total_rev - kos_part
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Revenue", f"RM {total_rev:.2f}")
    m2.metric("Sales Kedai (POS)", f"RM {sales_shop:.2f}")
    m3.metric("Total Untung", f"RM {total_prof:.2f}")
    
    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("🔧 Top Masalah")
        if not df.empty and 'Masalah' in df.columns:
            txt = ",".join(df['Masalah'].astype(str)); itm = [x.strip() for x in txt.split(",") if x.strip()]
            if itm: st.bar_chart(pd.Series(itm).value_counts().head(5))
            
    with c2:
        st.subheader("🔩 Top Parts")
        if not df_p.empty: st.bar_chart(df_p['NamaPart'].value_counts().head(5))
        
    st.divider()
    if not df_s.empty:
        st.subheader("🛒 Rekod Jualan Kedai")
        st.dataframe(df_s)
