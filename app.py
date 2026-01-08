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
import urllib.parse
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="DCK Tech System", layout="wide")

cloudinary.config( 
  cloud_name = "dmnk9vbyj", 
  api_key = "382361922884785", 
  api_secret = "qZBl9CL1886_Ohvv_5JWPxzUQz8",
  secure = True
)

SHEET_ID = "1ssuZ3BzAih5goP5m_XsgAPjCj1OeDX_CdE--S0h-xek"

if 'page' not in st.session_state: st.session_state.page = "📊 DASHBOARD"
if 'selected_id' not in st.session_state: st.session_state.selected_id = None
if 'email_user' not in st.session_state: st.session_state.email_user = ""
if 'email_pass' not in st.session_state: st.session_state.email_pass = ""
if 'config' not in st.session_state:
    st.session_state.config = {
        "Company_Name": "DCK TECH SERVICES",
        "Address": "No 123, Jalan Gadget, 70000 Seremban",
        "Phone": "012-3456789",
        "TNC_Ticket": "1. Data customer tanggungjawab sendiri.\n2. Warranty part sahaja.",
        "TNC_Invoice": "1. Barang dijual tiada refund."
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
        try: ws = sh.worksheet("Tickets")
        except: ws = sh.add_worksheet("Tickets", 1000, 20)
        h_t = ["ID", "Tarikh", "Customer", "Phone", "Email", "Model", "SN", "Password", "Masalah", "Fizikal", "Aksesori", "Status", "Kos_Part", "Harga_Jual", "Image_Link", "Tech_Note"]
        if len(ws.row_values(1)) != len(h_t): ws.update("A1:P1", [h_t])
        
        try: ws_p = sh.worksheet("Parts")
        except: ws_p = sh.add_worksheet("Parts", 1000, 10)
        h_p = ["ID", "TicketID", "NamaPart", "Supplier", "TarikhMasuk", "WarrantyBulan", "TarikhExpire", "HargaBeli"]
        if ws_p.row_values(1) != h_p: ws_p.update("A1:H1", [h_p])
        
        try: ws_s = sh.worksheet("Sales")
        except: ws_s = sh.add_worksheet("Sales", 1000, 10)
        h_s = ["ID", "Tarikh", "Item", "Qty", "Harga_Unit", "Total", "Customer", "PaymentMethod"]
        if ws_s.row_values(1) != h_s: ws_s.update("A1:H1", [h_s])
        
        try: ws_c = sh.worksheet("Config")
        except: 
            ws_c = sh.add_worksheet("Config", 100, 2)
            ws_c.append_row(["Key", "Value"]); ws_c.append_row(["Company_Name", "DCK TECH"])
        st.session_state.db_checked = True
    except: pass

def load_data(tab_name):
    init_db(); client = get_client()
    try:
        sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
        data = robust_api_call(sheet.get_all_records)
        df = pd.DataFrame(data) if data else pd.DataFrame()
        if tab_name == "Tickets" and not df.empty and "Email" not in df.columns: df["Email"] = ""
        return df
    except: return pd.DataFrame()

def load_config():
    try:
        df = load_data("Config")
        if not df.empty: st.session_state.config.update(dict(zip(df['Key'], df['Value'])))
    except: pass

def save_config_to_db(new_config):
    client = get_client(); sheet = client.open_by_key(SHEET_ID).worksheet("Config")
    sheet.clear(); sheet.append_row(["Key", "Value"])
    for k, v in new_config.items(): sheet.append_row([k, v])
    st.session_state.config = new_config; st.cache_data.clear()

def add_row(tab_name, row):
    client = get_client(); sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
    robust_api_call(sheet.append_row, row); st.cache_data.clear()

def update_cell_data(tab_name, id_val, col_dict):
    client = get_client(); sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
    cell = robust_api_call(sheet.find, str(id_val))
    if cell:
        for col_idx, val in col_dict.items(): robust_api_call(sheet.update_cell, cell.row, col_idx, val)
        st.cache_data.clear(); return True
    return False

def delete_row_data(tab_name, id_val):
    client = get_client(); sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
    cell = robust_api_call(sheet.find, str(id_val))
    if cell: robust_api_call(sheet.delete_rows, cell.row); st.cache_data.clear(); return True
    return False

def update_customer_info_db(tid, nama, phone, email, model, sn, pwd, masalah):
    client = get_client(); sheet = client.open_by_key(SHEET_ID).worksheet("Tickets")
    cell = robust_api_call(sheet.find, str(tid))
    if cell:
        robust_api_call(sheet.update_cell, cell.row, 3, nama)
        robust_api_call(sheet.update_cell, cell.row, 4, phone)
        robust_api_call(sheet.update_cell, cell.row, 5, email)
        robust_api_call(sheet.update_cell, cell.row, 6, model)
        robust_api_call(sheet.update_cell, cell.row, 7, sn)
        robust_api_call(sheet.update_cell, cell.row, 8, pwd)
        robust_api_call(sheet.update_cell, cell.row, 9, masalah)
        st.cache_data.clear(); return True
    return False

def update_part_data(part_id, new_name, new_supp, new_price, new_warranty):
    client = get_client(); sheet = client.open_by_key(SHEET_ID).worksheet("Parts")
    cell = robust_api_call(sheet.find, str(part_id))
    if cell:
        robust_api_call(sheet.update_cell, cell.row, 3, new_name)
        robust_api_call(sheet.update_cell, cell.row, 4, new_supp)
        robust_api_call(sheet.update_cell, cell.row, 6, new_warranty)
        new_exp = (datetime.now() + pd.DateOffset(months=int(new_warranty))).strftime("%Y-%m-%d")
        robust_api_call(sheet.update_cell, cell.row, 7, new_exp)
        robust_api_call(sheet.update_cell, cell.row, 8, new_price)
        st.cache_data.clear(); return True
    return False

def delete_part(part_id):
    client = get_client(); sheet = client.open_by_key(SHEET_ID).worksheet("Parts")
    cell = robust_api_call(sheet.find, str(part_id))
    if cell: robust_api_call(sheet.delete_rows, cell.row); st.cache_data.clear(); return True
    return False

# --- 4. PDF GENERATOR ---
def generate_pdf(t, type="SERVICE"):
    buffer = BytesIO(); p = canvas.Canvas(buffer, pagesize=A4); w, h = A4
    cfg = st.session_state.config
    
    p.setFont("Helvetica-Bold", 20); p.drawString(50, h-50, cfg.get("Company_Name", "DCK TECH"))
    p.setFont("Helvetica", 10); p.drawString(50, h-65, cfg.get("Address", "")); p.drawString(50, h-78, f"Tel: {cfg.get('Phone', '')}")
    p.line(50, h-85, w-50, h-85)
    p.setFont("Helvetica", 10); p.drawString(50, h-110, f"REF ID: {t.get('ID', '-')}"); p.drawString(300, h-110, f"DATE: {t.get('Tarikh', '-')}")
    
    if 'Model' in t: # Service
        p.drawString(50, h-125, f"CUSTOMER: {t.get('Customer', '-')}"); p.drawString(300, h-125, f"MODEL: {t.get('Model', '-')}")
        p.drawString(50, h-140, f"CONTACT: {t.get('Phone', '-')}"); p.drawString(300, h-140, f"S/N: {t.get('SN', '-')}")
        y = h-170; p.line(50, y+10, w-50, y+10); p.setFont("Helvetica-Bold", 10); p.drawString(50, y, "SERVICE DETAILS:"); y-=15
        p.setFont("Helvetica", 10)
        p.drawString(50, y, f"Problem: {t.get('Masalah', '-')}"); y-=15
        p.drawString(50, y, f"Condition: {t.get('Fizikal', '-')}"); y-=15
        p.drawString(50, y, f"Accessories: {t.get('Aksesori', '-')}"); y-=15
        p.drawString(50, y, f"Note: {t.get('Tech_Note', '-')}"); y-=30
        tnc_text = cfg.get("TNC_Invoice", "") if type == "INVOICE" else cfg.get("TNC_Ticket", "")
        if type == "INVOICE":
            p.setFont("Helvetica-Bold", 14); p.drawString(50, y, f"TOTAL: RM {safe_float(t.get('Harga_Jual', 0)):.2f}"); y-=30
    else: # Sales
        p.drawString(50, h-125, f"CUSTOMER: {t.get('Customer', 'Walk-in')}")
        y = h-160; p.line(50, y+10, w-50, y+10); p.setFont("Helvetica-Bold", 10)
        p.drawString(50, y, "ITEM"); p.drawString(300, y, "QTY"); p.drawString(400, y, "PRICE"); y-=20
        p.setFont("Helvetica", 10)
        p.drawString(50, y, str(t.get('Item', '-'))); p.drawString(300, y, str(t.get('Qty', 0)))
        p.drawString(400, y, f"RM {safe_float(t.get('Harga_Unit', 0)):.2f}"); y-=30
        p.line(50, y+10, w-50, y+10); p.setFont("Helvetica-Bold", 14)
        p.drawString(300, y, "TOTAL:"); p.drawString(400, y, f"RM {safe_float(t.get('Total', 0)):.2f}"); y-=30
        tnc_text = cfg.get("TNC_Invoice", "")

    p.setFont("Helvetica-Bold", 10); p.drawString(50, y, "TERMS & CONDITIONS:"); y-=15
    p.setFont("Helvetica", 8)
    for line in tnc_text.split('\n'): p.drawString(50, y, line.strip()); y-=12
    y -= 30; p.drawString(50, y, "Customer Signature: _________________"); p.drawString(300, y, "Authorized Signature: _________________")
    p.save(); buffer.seek(0); return buffer

# --- 5. EMAIL & LINKS ---
def generate_message_content(data):
    cust = str(data.get('Customer', 'Pelanggan'))
    status = str(data.get('Status', 'Pending'))
    tid = str(data.get('ID', '-'))
    model = str(data.get('Model', '-'))
    sn = str(data.get('SN', '-'))
    masalah = str(data.get('Masalah', '-'))
    note = str(data.get('Tech_Note', '-'))
    price = safe_float(data.get('Harga_Jual', 0))
    company = st.session_state.config.get('Company_Name', 'DCK TECH')

    msg = f"Hai {cust},\n\nTerima kasih berurusan dengan {company}."
    if status == 'Pending': msg += "\nKami telah menerima peranti anda untuk pemeriksaan."
    elif status in ['Done', 'Collected']: msg += "\n✅ Berita Baik! Peranti anda telah SIAP dibaiki."
    else: msg += f"\nStatus terkini peranti anda: {status}"
    
    msg += f"\n\n--- BUTIRAN ---\nID: {tid}\nModel: {model}\nS/N: {sn}\nMasalah: {masalah}\nNota: {note}"
    if status in ['Done', 'Collected']: msg += f"\n\n💰 TOTAL: RM {price:.2f}"
    msg += "\n\nSekian,\nTeam DCK Tech"
    return msg

def generate_links(type, data):
    phone = clean_phone_number_my(data.get('Phone', ''))
    full_msg = generate_message_content(data)
    if type == "WA":
        wa_msg = full_msg.replace("BUTIRAN:", "*BUTIRAN:*")
        return f"https://wa.me/{phone}?text={urllib.parse.quote(wa_msg)}"
    elif type == "EMAIL":
        sub = f"Tiket DCK: {data.get('ID', '-')} - {data.get('Model', '-')}"
        return f"mailto:{data.get('Email', '')}?subject={urllib.parse.quote(sub)}&body={urllib.parse.quote(full_msg)}"

def send_email_with_pdf(to_email, data, pdf_buffer, pdf_name):
    sender = st.session_state.email_user; password = st.session_state.email_pass
    if not sender or not password: return False, "Sila set Email di Tetapan."
    msg = MIMEMultipart(); msg['From'] = sender; msg['To'] = to_email; msg['Subject'] = f"Tiket: {data.get('ID', '-')}"
    msg.attach(MIMEText(generate_message_content(data), 'plain'))
    part = MIMEApplication(pdf_buffer.getvalue(), Name=pdf_name)
    part['Content-Disposition'] = f'attachment; filename="{pdf_name}"'
    msg.attach(part)
    try:
        s = smtplib.SMTP('smtp.gmail.com', 587); s.starttls(); s.login(sender, password); s.send_message(msg); s.quit()
        return True, "Berjaya!"
    except Exception as e: return False, str(e)

# --- 6. NAVIGATION ---
load_config()
PAGES = ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🛒 JUALAN KEDAI", "🔧 UPDATE STATUS", "📦 INVENTORY", "📈 LAPORAN", "⚙️ TETAPAN"]
try: idx = PAGES.index(st.session_state.page)
except: idx = 0
sel = st.sidebar.radio("NAVIGASI", PAGES, index=idx)
if sel != st.session_state.page: st.session_state.page = sel; st.rerun()

# === PAGE: DASHBOARD ===
if st.session_state.page == "📊 DASHBOARD":
    st.title("📊 DCK Tech Dashboard")
    df = load_data("Tickets"); df_s = load_data("Sales")
    
    today_str = datetime.now().strftime("%Y-%m-%d")
    sales_today = 0.0
    if not df_s.empty: sales_today += df_s[df_s['Tarikh'] == today_str]['Total'].apply(safe_float).sum()

    if not df.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.info(f"PENDING: {len(df[df['Status'] == 'Pending'])}")
        c2.warning(f"CHECKING: {len(df[df['Status'] == 'Checking'])}")
        c3.success(f"DONE: {len(df[df['Status'] == 'Done'])}")
        c4.error(f"COLLECTED: {len(df[df['Status'] == 'Collected'])}")
    
    st.divider(); st.metric("💰 JUALAN KEDAI (HARI INI)", f"RM {sales_today:.2f}")
    
    st.write("### 🔍 Cari Ticket")
    search = st.text_input("Masukkan Nama / ID / Model:", placeholder="Contoh: DCK-12345")
    st.write("### Senarai Job Terkini")
    if not df.empty:
        if search: df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        for _, row in df.iloc[::-1].head(10).iterrows():
            with st.expander(f"{row.get('ID')} - {row.get('Customer')} ({row.get('Status')})"):
                st.write(f"Model: {row.get('Model')} | Masalah: {row.get('Masalah')}")
                if st.button("🔧 Manage Job", key=f"btn_{row.get('ID')}"):
                    st.session_state.selected_id = row.get('ID'); st.session_state.page = "🔧 UPDATE STATUS"; st.rerun()
    else: st.info("Tiada rekod tiket.")

# === PAGE: DAFTAR TIKET ===
elif st.session_state.page == "📝 DAFTAR TIKET":
    st.title("📝 Tiket Masuk Baru")
    with st.form("reg_ticket", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("### 👤 Pelanggan")
            nama = st.text_input("Nama"); phone = st.text_input("No HP"); email = st.text_input("Email")
        with c2:
            st.markdown("### 💻 Peranti")
            model = st.text_input("Model"); sn = st.text_input("Serial No"); pwd = st.text_input("Password")
        with c3:
            st.markdown("### 🔧 Diagnosis")
            mslh = st.multiselect("Masalah", ["Slow", "Screen", "Battery", "Keyboard", "Format", "Other"])
            fiz = st.multiselect("Fizikal", ["Calar", "Pecah", "Sempurna"]); acc = st.multiselect("Aksesori", ["Bag", "Charger", "Mouse"])
        
        note = st.text_area("Nota Tambahan")
        img = st.file_uploader("Gambar Peranti", accept_multiple_files=True)
        tnc = st.checkbox("Saya setuju dengan Terma & Syarat")
        
        submitted = st.form_submit_button("SIMPAN REKOD", use_container_width=True)
        
        if submitted:
            if nama and tnc:
                with st.spinner("Menyimpan..."):
                    tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                    img_str = ""
                    if img: 
                        try: img_str = ",".join([cloudinary.uploader.upload(f)['secure_url'] for f in img])
                        except: pass
                    
                    row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, email, model, sn, pwd, ",".join(mslh), ",".join(fiz), ",".join(acc), "Pending", 0, 0, img_str, note]
                    add_row("Tickets", row)
                    st.toast("Tiket Berjaya Dibuka!", icon='✅')
                    st.session_state.last_data = {"ID": tid, "Customer": nama, "Phone": phone, "Email": email, "Model": model, "SN": sn, "Masalah": ",".join(mslh), "Fizikal": ",".join(fiz), "Aksesori": ",".join(acc), "Tech_Note": note, "Status": "Pending", "Images": img_str, "Image_Link": img_str, "Tarikh": row[1]}
                    time.sleep(1); st.rerun()

    if 'last_data' in st.session_state:
        ld = st.session_state.last_data
        st.divider(); st.success(f"Tiket {ld['ID']} Telah Dibuka!")
        c1, c2, c3 = st.columns(3)
        c1.download_button("📥 PDF Tiket", generate_pdf(ld, "SERVICE"), "Tiket.pdf", use_container_width=True)
        c2.link_button("📱 WhatsApp", generate_links("WA", ld), use_container_width=True)
        if ld['Email']: 
            if c3.button("📧 Email Auto"):
                ok, m = send_email_with_pdf(ld['Email'], ld, generate_pdf(ld, "SERVICE"), "Tiket.pdf")
                if ok: st.toast(m, icon='✅')
                else: st.error(m)

# === PAGE: JUALAN KEDAI ===
elif st.session_state.page == "🛒 JUALAN KEDAI":
    st.title("🛒 Sistem Jualan (POS)")
    tab_pos, tab_manage = st.tabs(["🛒 Jualan Baru", "📋 Urus Jualan"])
    
    with tab_pos:
        with st.form("pos_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            item = c1.text_input("Nama Barang"); cust = c2.text_input("Nama Pelanggan", value="Walk-in")
            c3, c4, c5 = st.columns(3)
            qty = c3.number_input("Qty", 1, 100, 1); price = c4.number_input("Harga Unit (RM)", 0.0)
            pay = st.selectbox("Bayaran", ["Cash", "QR", "Transfer"])
            sub_pos = st.form_submit_button("✅ REKOD JUALAN")
            
            if sub_pos:
                if item and price > 0:
                    total = qty * price
                    sid = f"SALE-{datetime.now().strftime('%d%H%M')}"
                    add_row("Sales", [sid, datetime.now().strftime("%Y-%m-%d"), item, qty, price, total, cust, pay])
                    st.toast("Jualan Direkod!", icon='💰')
                    st.session_state.last_sale = {"ID": sid, "Tarikh": datetime.now().strftime("%Y-%m-%d"), "Item": item, "Qty": qty, "Harga_Unit": price, "Total": total, "Customer": cust}
                    time.sleep(1); st.rerun()
        if 'last_sale' in st.session_state:
            st.download_button("🖨️ Resit", generate_pdf(st.session_state.last_sale, "SALES"), "Resit.pdf", use_container_width=True)

    with tab_manage:
        df_s = load_data("Sales")
        if not df_s.empty:
            st.dataframe(df_s)
            sale_id = st.selectbox("Pilih ID Jualan:", ["-"] + df_s['ID'].tolist())
            if sale_id != "-":
                curr = df_s[df_s['ID'] == sale_id].iloc[0]
                with st.form("edit_sale"):
                    e_item = st.text_input("Item", curr['Item'])
                    e_qty = st.number_input("Qty", value=int(curr['Qty']))
                    e_price = st.number_input("Harga Unit", value=float(curr['Harga_Unit']))
                    c_del, c_upd = st.columns(2)
                    if c_del.form_submit_button("🗑️ Hapus"): delete_row_data("Sales", sale_id); st.rerun()
                    if c_upd.form_submit_button("💾 Update"):
                        update_cell_data("Sales", sale_id, {3: e_item, 4: e_qty, 5: e_price, 6: e_qty * e_price})
                        st.toast("Jualan Dikemaskini!", icon='✅'); time.sleep(1); st.rerun()

# === PAGE: UPDATE STATUS ===
elif st.session_state.page == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Technician")
    df = load_data("Tickets")
    if not df.empty:
        ids = df['ID'].astype(str).tolist()
        tgt = str(st.session_state.selected_id) if st.session_state.selected_id else ids[-1]
        pid = st.selectbox("Pilih Job:", ids, index=ids.index(tgt) if tgt in ids else 0)
        job = df[df['ID'].astype(str) == str(pid)].iloc[0]
        
        with st.expander("ℹ️ MAKLUMAT TIKET & KOMUNIKASI (KLIK EDIT)", expanded=True):
            edit_mode = st.checkbox("✏️ Tick Untuk Edit Info")
            if edit_mode:
                with st.form("edit_cust_form"):
                    ec_nama = st.text_input("Nama", value=job.get('Customer',''))
                    c_ec1, c_ec2 = st.columns(2)
                    ec_phone = c_ec1.text_input("Phone", value=job.get('Phone',''))
                    ec_email = c_ec2.text_input("Email", value=job.get('Email',''))
                    c_ec3, c_ec4, c_ec5 = st.columns(3)
                    ec_model = c_ec3.text_input("Model", value=job.get('Model',''))
                    ec_sn = c_ec4.text_input("Serial No", value=job.get('SN',''))
                    ec_pwd = c_ec5.text_input("Password", value=job.get('Password',''))
                    ec_mslh = st.text_area("Masalah", value=job.get('Masalah',''))
                    if st.form_submit_button("💾 SIMPAN PERUBAHAN"):
                        if update_customer_info_db(pid, ec_nama, ec_phone, ec_email, ec_model, ec_sn, ec_pwd, ec_mslh):
                            st.toast("Info Berjaya Diubah!", icon='✅'); time.sleep(1); st.rerun()
            else:
                c1, c2 = st.columns(2)
                c1.write(f"**Nama:** {job.get('Customer','-')}"); c1.write(f"**Model:** {job.get('Model','-')}")
                c1.write(f"**Phone:** {job.get('Phone','-')}"); c1.write(f"**Email:** {job.get('Email','-')}")
                c2.write(f"**Masalah:** {job.get('Masalah','-')}"); c2.error(f"🔐 PWD: {job.get('Password','-')}")

            st.divider()
            ca, cb = st.columns(2)
            wa_url = generate_links("WA", job)
            ca.link_button("📱 WhatsApp Status", wa_url, use_container_width=True)
            if job.get('Email'): 
                with cb:
                    if st.button("📧 Hantar Email + PDF (Auto Server)", key=f"em_{pid}"):
                        doc = "INVOICE" if job.get('Status') in ['Done', 'Collected'] else "SERVICE"
                        ok, m = send_email_with_pdf(job.get('Email'), job, generate_pdf(job, doc), "Status.pdf")
                        if ok: st.toast(m, icon='✅')
                        else: st.error(m)
            else: cb.caption("Tiada Email.")
        
        c_l, c_r = st.columns([1, 2])
        with c_l:
            st.download_button("Print Tiket", generate_pdf(job, "SERVICE"), "Tiket.pdf")
            im = str(job.get('Image_Link'))
            if im and len(im) > 5:
                for i in im.split(','): st.image(i)

        with c_r:
            df_p = load_data("Parts")
            parts = df_p[df_p['TicketID'].astype(str) == str(pid)]
            kos = sum([safe_float(x) for x in parts['HargaBeli'].tolist()]) if not parts.empty else 0
            
            t1, t2 = st.tabs(["Status", "Parts"])
            with t1:
                st.write(f"**KOS:** RM {kos:.2f}")
                with st.form("upd_status"):
                    stt = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Done", "Collected"], index=["Pending", "Checking", "Waiting Part", "Done", "Collected"].index(job.get('Status','Pending')) if job.get('Status','Pending') in ["Pending", "Checking", "Waiting Part", "Done", "Collected"] else 0)
                    nt = st.text_area("Nota Technician", value=job.get('Tech_Note',''))
                    hj = st.number_input("Harga Jual (Total Bill)", value=safe_float(job.get('Harga_Jual',0)))
                    if st.form_submit_button("UPDATE"):
                        sheet = get_client().open_by_key(SHEET_ID).worksheet("Tickets")
                        cl = robust_api_call(sheet.find, str(pid))
                        if cl:
                            robust_api_call(sheet.update_cell, cl.row, 12, stt)
                            robust_api_call(sheet.update_cell, cl.row, 13, kos)
                            robust_api_call(sheet.update_cell, cl.row, 14, hj)
                            robust_api_call(sheet.update_cell, cl.row, 16, nt)
                            st.cache_data.clear()
                            st.toast("Status Dikemaskini!", icon='🎉'); time.sleep(1); st.rerun()
                if stt in ["Done", "Collected"]: st.download_button("🖨️ CETAK RESIT", generate_pdf(job, "INVOICE"), "Resit.pdf", use_container_width=True)

            with t2:
                with st.form("ap", clear_on_submit=True):
                    pn = st.text_input("Part"); ps = st.text_input("Supp"); w = st.number_input("Warr", 1); pr = st.number_input("Harga", 0.0)
                    if st.form_submit_button("Add"):
                        exp = (datetime.now() + pd.DateOffset(months=int(w))).strftime("%Y-%m-%d")
                        add_row("Parts", [f"P-{int(time.time())}", pid, pn, ps, str(datetime.now().date()), w, exp, pr])
                        st.toast("Part ditambah!", icon='➕'); time.sleep(1); st.rerun()
                if not parts.empty:
                    st.dataframe(parts[['ID', 'NamaPart', 'HargaBeli']])
                    d = st.selectbox("Del", ["-"]+parts['ID'].tolist())
                    if d != "-" and st.button("Delete"): 
                        delete_part(d); st.toast("Part dihapus!", icon='🗑️'); time.sleep(1); st.rerun()

# === PAGE: INVENTORY ===
elif st.session_state.page == "📦 INVENTORY":
    st.title("📦 Inventory Log")
    df_p = load_data("Parts")
    search_p = st.text_input("🔍 Cari Part (Nama/Ticket ID):")
    
    if not df_p.empty:
        if search_p: df_p = df_p[df_p.apply(lambda r: r.astype(str).str.contains(search_p, case=False).any(), axis=1)]
        # FUNGSI LOOPING UNTUK CARD + BUTTON TELEPORT
        for i, row in df_p.iterrows():
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                c1.write(f"**{row.get('NamaPart')}** (RM {row.get('HargaBeli')}) | Ticket: {row.get('TicketID')}")
                if c2.button("Go to Job", key=f"inv_{row.get('ID')}"):
                    st.session_state.selected_id = row.get('TicketID'); st.session_state.page = "🔧 UPDATE STATUS"; st.rerun()
    else: st.info("Tiada barang.")

# === PAGE: LAPORAN ===
elif st.session_state.page == "📈 LAPORAN":
    st.title("📈 Laporan Prestasi")
    df = load_data("Tickets"); df_s = load_data("Sales"); df_p = load_data("Parts")
    
    if not df.empty:
        df['Tarikh'] = pd.to_datetime(df['Tarikh'], errors='coerce')
        df['Harga_Jual'] = df['Harga_Jual'].apply(safe_float)
        df['Kos_Part'] = df['Kos_Part'].apply(safe_float)
        df['Untung'] = df['Harga_Jual'] - df['Kos_Part']
        
        serv_sales = df['Harga_Jual'].sum()
        shop_sales = df_s['Total'].apply(safe_float).sum() if not df_s.empty else 0
        total_rev = serv_sales + shop_sales
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Revenue", f"RM {total_rev:.2f}")
        m2.metric("Untung Servis", f"RM {df['Untung'].sum():.2f}")
        m3.metric("Jualan Kedai", f"RM {shop_sales:.2f}")
        
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("🔧 Masalah Utama")
            if 'Masalah' in df.columns:
                txt = ",".join(df['Masalah'].astype(str)); itm = [x.strip() for x in txt.split(",") if x.strip()]
                if itm: st.bar_chart(pd.Series(itm).value_counts().head(5))
        with c2:
            st.subheader("🔩 Barang Laju (Parts)")
            if not df_p.empty: st.bar_chart(df_p['NamaPart'].value_counts().head(5))
            
        st.divider()
        st.subheader("📅 Prestasi Berkala")
        tab_h, tab_m, tab_b = st.tabs(["Harian", "Mingguan", "Bulanan"])
        
        with tab_h: st.line_chart(df.groupby(df['Tarikh'].dt.date)[['Harga_Jual', 'Untung']].sum().tail(30))
        with tab_m: st.bar_chart(df.groupby(df['Tarikh'].dt.to_period('W').astype(str))[['Harga_Jual', 'Untung']].sum())
        with tab_b: st.bar_chart(df.groupby(df['Tarikh'].dt.to_period('M').astype(str))[['Harga_Jual', 'Untung']].sum())
        
        st.download_button("📥 Download Laporan (CSV)", df.to_csv(index=False).encode('utf-8'), "Laporan.csv", "text/csv")

# === PAGE: TETAPAN ===
elif st.session_state.page == "⚙️ TETAPAN":
    st.title("⚙️ Tetapan")
    with st.form("conf"):
        cn = st.text_input("Nama Syarikat", st.session_state.config.get("Company_Name"))
        ad = st.text_area("Alamat", st.session_state.config.get("Address"))
        ph = st.text_input("No Tel", st.session_state.config.get("Phone"))
        tnc1 = st.text_area("T&C Tiket", st.session_state.config.get("TNC_Ticket"))
        tnc2 = st.text_area("T&C Invois", st.session_state.config.get("TNC_Invoice"))
        if st.form_submit_button("Simpan"):
            new_c = st.session_state.config.copy()
            new_c.update({"Company_Name":cn, "Address":ad, "Phone":ph, "TNC_Ticket":tnc1, "TNC_Invoice":tnc2})
            save_config_to_db(new_c); st.toast("Tetapan Disimpan!", icon='✅')
    
    with st.expander("📧 Email Server"):
        with st.form("em"):
            eu = st.text_input("Email", st.session_state.email_user)
            ep = st.text_input("App Password", st.session_state.email_pass, type="password")
            if st.form_submit_button("Set Email"):
                st.session_state.email_user = eu; st.session_state.email_pass = ep; st.toast("Email Disimpan!", icon='✅')
