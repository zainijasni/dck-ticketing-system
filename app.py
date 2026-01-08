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
import qrcode # Pastikan library ini ada dalam requirements.txt

# =============================================================================
# 1. CONFIGURATION & SETUP
# =============================================================================
st.set_page_config(page_title="DCK Tech System", layout="wide")

# Cloudinary Config (Untuk simpan gambar)
cloudinary.config( 
  cloud_name = "dmnk9vbyj", 
  api_key = "382361922884785", 
  api_secret = "qZBl9CL1886_Ohvv_5JWPxzUQz8",
  secure = True
)

# Google Sheet ID (Database Utama)
SHEET_ID = "1ssuZ3BzAih5goP5m_XsgAPjCj1OeDX_CdE--S0h-xek"

# --- SESSION STATE INITIALIZATION ---
if 'page' not in st.session_state:
    st.session_state.page = "📊 DASHBOARD"
if 'selected_id' not in st.session_state:
    st.session_state.selected_id = None
if 'email_user' not in st.session_state:
    st.session_state.email_user = ""
if 'email_pass' not in st.session_state:
    st.session_state.email_pass = ""
if 'pos_cart' not in st.session_state:
    st.session_state.pos_cart = []
# V74: New Session State untuk Restock Cart
if 'restock_cart' not in st.session_state:
    st.session_state.restock_cart = []

if 'config' not in st.session_state:
    st.session_state.config = {
        "Company_Name": "DCK TECH SERVICES",
        "Address": "No 123, Jalan Gadget, 70000 Seremban",
        "Phone": "012-3456789",
        "Logo": "",
        "TNC_Ticket": "1. Data customer tanggungjawab sendiri.\n2. Warranty part sahaja.",
        "TNC_Invoice": "1. Barang dijual tiada refund.",
        "Options_Masalah": "Slow, Screen, Battery, Keyboard, Format, Water Damage, Other",
        "Options_Fizikal": "Calar, Pecah, Kemek, Sempurna",
        "Options_Aksesori": "Bag, Charger, Mouse, Box"
    }

# =============================================================================
# 2. HELPER FUNCTIONS (Alat Bantuan)
# =============================================================================
def safe_float(val):
    """Tukar duit 'RM 50.00' jadi nombor 50.0"""
    try:
        if pd.isna(val) or str(val).strip() == "":
            return 0.0
        clean = str(val).upper().replace("RM", "").replace(",", ".").strip()
        return float(clean)
    except:
        return 0.0

def safe_int(val):
    """Tukar teks jadi nombor bulat"""
    try:
        return int(float(val))
    except:
        return 0

def robust_api_call(func, *args, **kwargs):
    """Kalau Google Sheet jam, cuba 3 kali sebelum error"""
    for i in range(3):
        try:
            return func(*args, **kwargs)
        except Exception:
            time.sleep(1)
            continue
    return None

def clean_phone_number_my(phone_input):
    """Format nombor telefon Malaysia untuk WhatsApp"""
    p = str(phone_input).replace("-", "").replace(" ", "").replace("+", "").strip()
    if p.startswith("0"):
        return "6" + p
    elif p.startswith("1"):
        return "60" + p
    elif p.startswith("60"):
        return p
    else:
        return "6" + p

# =============================================================================
# 3. DATABASE ENGINE (Google Sheets)
# =============================================================================
@st.cache_resource
def get_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(st.secrets["google_creds"], scopes=scope)
    return gspread.authorize(creds)

def init_db():
    """Cek database, kalau tak ada tab, buat baru"""
    if 'db_checked' in st.session_state:
        return
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        
        # 1. Tickets (Rekod Servis)
        try:
            ws = sh.worksheet("Tickets")
        except:
            ws = sh.add_worksheet("Tickets", 1000, 20)
        h_t = ["ID", "Tarikh", "Customer", "Phone", "Email", "Model", "SN", "Password", "Masalah", "Fizikal", "Aksesori", "Status", "Kos_Part", "Harga_Jual", "Image_Link", "Tech_Note"]
        if len(ws.row_values(1)) != len(h_t):
            ws.update("A1:P1", [h_t])
        
        # 2. Parts (Used Log - Barang Keluar untuk Repair)
        try:
            ws_p = sh.worksheet("Parts")
        except:
            ws_p = sh.add_worksheet("Parts", 1000, 10)
        h_p = ["ID", "TicketID", "NamaPart", "Supplier", "TarikhMasuk", "WarrantyBulan", "TarikhExpire", "HargaBeli"]
        if ws_p.row_values(1) != h_p:
            ws_p.update("A1:H1", [h_p])
        
        # 3. Sales (Rekod Jualan Kaunter - V73 Added Warranty Column)
        try:
            ws_s = sh.worksheet("Sales")
        except:
            ws_s = sh.add_worksheet("Sales", 1000, 10)
        h_s = ["ID", "Tarikh", "Item", "Qty", "Harga_Unit", "Total", "Customer", "PaymentMethod", "Warranty"]
        if len(ws_s.row_values(1)) != len(h_s):
            ws_s.update("A1:I1", [h_s])
        
        # 4. Config (Tetapan)
        try:
            ws_c = sh.worksheet("Config")
        except: 
            ws_c = sh.add_worksheet("Config", 100, 2)
            ws_c.append_row(["Key", "Value"])
            ws_c.append_row(["Company_Name", "DCK TECH"])
            
        # 5. Master Inventory (Gudang Utama - V72/V74 Simplified)
        try:
            ws_m = sh.worksheet("Master_Inventory")
        except:
            ws_m = sh.add_worksheet("Master_Inventory", 1000, 10)
        h_m = ["ItemCode", "ItemName", "CostPrice", "SellPrice", "CurrentStock", "LastUpdated"]
        if len(ws_m.row_values(1)) != len(h_m):
            ws_m.update("A1:F1", [h_m])

        # 6. Restock Log (Rekod Beli Barang - V72/V74 Simplified)
        try:
            ws_r = sh.worksheet("Restock_Log")
        except:
            ws_r = sh.add_worksheet("Restock_Log", 1000, 10)
        h_r = ["LogID", "Date", "InvoiceNo", "Supplier", "ItemName", "QtyAdded", "CostPrice", "TotalCost"]
        if len(ws_r.row_values(1)) != len(h_r):
            ws_r.update("A1:H1", [h_r])

        st.session_state.db_checked = True
    except:
        pass

def load_data(tab_name):
    init_db()
    client = get_client()
    try:
        sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
        data = robust_api_call(sheet.get_all_records)
        df = pd.DataFrame(data) if data else pd.DataFrame()
        
        # Patch Empty Dataframes to avoid errors
        if df.empty:
            if tab_name == "Tickets":
                df = pd.DataFrame(columns=["ID", "Tarikh", "Customer", "Phone", "Email", "Model", "SN", "Password", "Masalah", "Fizikal", "Aksesori", "Status", "Kos_Part", "Harga_Jual", "Image_Link", "Tech_Note"])
            elif tab_name == "Parts":
                 df = pd.DataFrame(columns=["ID", "TicketID", "NamaPart", "Supplier", "TarikhMasuk", "WarrantyBulan", "TarikhExpire", "HargaBeli"])
            elif tab_name == "Sales":
                 df = pd.DataFrame(columns=["ID", "Tarikh", "Item", "Qty", "Harga_Unit", "Total", "Customer", "PaymentMethod", "Warranty"])
            elif tab_name == "Master_Inventory":
                 df = pd.DataFrame(columns=["ItemCode", "ItemName", "CostPrice", "SellPrice", "CurrentStock", "LastUpdated"])

        if tab_name == "Tickets" and "Email" not in df.columns:
            df["Email"] = ""
        return df
    except:
        return pd.DataFrame()

def load_config():
    try:
        df = load_data("Config")
        if not df.empty:
            st.session_state.config.update(dict(zip(df['Key'], df['Value'])))
    except:
        pass

def save_config_to_db(new_config):
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

def update_cell_data(tab_name, id_val, col_dict):
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
    cell = robust_api_call(sheet.find, str(id_val))
    if cell:
        for col_idx, val in col_dict.items():
            robust_api_call(sheet.update_cell, cell.row, col_idx, val)
        st.cache_data.clear()
        return True
    return False

def delete_row_data(tab_name, id_val):
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet(tab_name)
    cell = robust_api_call(sheet.find, str(id_val))
    if cell:
        robust_api_call(sheet.delete_rows, cell.row)
        st.cache_data.clear()
        return True
    return False

# --- V72: INVENTORY LOGIC (Update Stock) ---
def update_stock(item_code, qty_change):
    # qty_change: Positif tambah stok, Negatif tolak stok
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet("Master_Inventory")
    try:
        cell = sheet.find(str(item_code))
        if cell:
            curr = safe_int(sheet.cell(cell.row, 5).value) # Col 5 is CurrentStock in V74
            new_qty = curr + int(qty_change)
            sheet.update_cell(cell.row, 5, new_qty)
            st.cache_data.clear()
            return True
    except:
        pass
    return False

def check_and_update_master(code, name, cost, sell, qty):
    # This function handles "If exists, update qty. If new, create row."
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet("Master_Inventory")
    try:
        cell = sheet.find(str(code))
        if cell:
            # Update existing
            curr_qty = safe_int(sheet.cell(cell.row, 5).value)
            sheet.update_cell(cell.row, 5, curr_qty + int(qty)) # Update Stock
            sheet.update_cell(cell.row, 3, cost) # Update latest cost
            sheet.update_cell(cell.row, 4, sell) # Update latest sell price
        else:
            # Create new
            tgl = datetime.now().strftime("%Y-%m-%d")
            sheet.append_row([code, name, cost, sell, qty, tgl])
        st.cache_data.clear()
        return True
    except:
        return False

def update_customer_info_db(tid, nama, phone, email, model, sn, pwd, masalah):
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet("Tickets")
    cell = robust_api_call(sheet.find, str(tid))
    if cell:
        sheet.update_cell(cell.row, 3, nama)
        sheet.update_cell(cell.row, 4, phone)
        sheet.update_cell(cell.row, 5, email)
        sheet.update_cell(cell.row, 6, model)
        sheet.update_cell(cell.row, 7, sn)
        sheet.update_cell(cell.row, 8, pwd)
        sheet.update_cell(cell.row, 9, masalah)
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

# =============================================================================
# 4. PDF GENERATOR
# =============================================================================
def draw_wrapped_text(c, text, x, y, max_width, font="Helvetica", size=10):
    c.setFont(font, size)
    text = str(text)
    words = text.split()
    lines = []
    current_line = []
    
    for word in words:
        test_line = ' '.join(current_line + [word])
        if c.stringWidth(test_line, font, size) < max_width:
            current_line.append(word)
        else:
            lines.append(' '.join(current_line))
            current_line = [word]
    lines.append(' '.join(current_line))
    
    curr_y = y
    for line in lines:
        c.drawString(x, curr_y, line)
        curr_y -= (size + 2)
    return curr_y

def generate_pdf(t, type="SERVICE"):
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    cfg = st.session_state.config
    
    logo_url = cfg.get("Logo", "")
    if logo_url and len(logo_url) > 10:
        try:
            logo = ImageReader(logo_url)
            p.drawImage(logo, 40, h-100, width=80, height=80, mask='auto', preserveAspectRatio=True)
        except:
            pass

    p.setFont("Helvetica-Bold", 20)
    p.drawString(140, h-50, cfg.get("Company_Name", "DCK TECH SERVICES"))
    p.setFont("Helvetica", 10)
    p.drawString(140, h-65, cfg.get("Address", "Alamat Kedai"))
    p.drawString(140, h-78, f"Tel: {cfg.get('Phone', '')}")
    
    p.setLineWidth(1.5)
    p.line(40, h-110, w-40, h-110)
    
    p.setFont("Helvetica-Bold", 14)
    doc_title = "SERVICE TICKET" if type == "SERVICE" else "OFFICIAL RECEIPT"
    if type == "INVOICE": doc_title = "OFFICIAL INVOICE"
    p.drawCentredString(w/2, h-130, doc_title)
    
    p.setLineWidth(0.5)
    p.rect(40, h-220, w-80, 80)
    p.setFont("Helvetica-Bold", 10)
    p.drawString(50, h-155, "CUSTOMER DETAILS:")
    p.setFont("Helvetica", 10)
    
    if 'Model' in t: 
        p.drawString(50, h-175, f"Name: {t.get('Customer', '-')}")
        p.drawString(300, h-175, f"Ticket ID: {t.get('ID', '-')}")
        p.drawString(50, h-190, f"Phone: {t.get('Phone', '-')}")
        p.drawString(300, h-190, f"Date: {t.get('Tarikh', '-')}")
        p.drawString(50, h-205, f"Email: {t.get('Email', '-')}")
        p.drawString(300, h-205, f"S/N: {t.get('SN', '-')}")
    else: 
        p.drawString(50, h-175, f"Name: {t.get('Customer', '-')}")
        p.drawString(300, h-175, f"Receipt ID: {t.get('ID', '-')}")
        p.drawString(300, h-190, f"Date: {t.get('Tarikh', '-')}")

    y = h-250
    p.setFont("Helvetica-Bold", 11)
    p.setFillColorRGB(0.9, 0.9, 0.9)
    p.rect(40, y, w-80, 20, fill=1)
    p.setFillColorRGB(0, 0, 0)
    
    if 'Model' in t: 
        p.drawString(50, y+6, "DEVICE / MODEL")
        p.drawString(250, y+6, "DIAGNOSIS / PROBLEM")
        p.drawString(450, y+6, "REMARKS")
        y -= 25
        y_model = draw_wrapped_text(p, t.get('Model', '-'), 50, y, 180)
        y_prob = draw_wrapped_text(p, t.get('Masalah', '-'), 250, y, 180)
        y_note = draw_wrapped_text(p, t.get('Tech_Note', '-'), 450, y, 100)
        y = min(y_model, y_prob, y_note) - 15
        p.line(40, y, w-40, y)
        y -= 20
        p.setFont("Helvetica-Bold", 10)
        p.drawString(50, y, "Condition & Accessories:")
        p.setFont("Helvetica", 10)
        p.drawString(200, y, f"{t.get('Fizikal', '-')} | {t.get('Aksesori', '-')}")
    else: 
        p.drawString(50, y+6, "ITEM DESCRIPTION")
        p.drawString(350, y+6, "QTY")
        p.drawString(450, y+6, "PRICE")
        y -= 25
        p.setFont("Helvetica", 10)
        items = t.get('Items_List', [])
        if not items:
            items = [{'Item': t.get('Item'), 'Qty': t.get('Qty'), 'Harga_Unit': t.get('Harga_Unit'), 'Warranty': t.get('Warranty', '-')}]
        
        for itm in items:
            p.drawString(350, y, str(itm.get('Qty', 1)))
            p.drawString(450, y, f"RM {safe_float(itm.get('Harga_Unit', 0)):.2f}")
            # Papar Warranty di Resit
            desc_text = f"{str(itm.get('Item', '-'))} (W: {itm.get('Warranty','-')})"
            y_item = draw_wrapped_text(p, desc_text, 50, y, 280)
            y = y_item - 10
        p.line(40, y, w-40, y)

    if type in ["INVOICE", "SALES"] or t.get('Status') in ['Done', 'Collected']:
        total = safe_float(t.get('Harga_Jual', t.get('Total', 0)))
        y -= 30
        p.setFont("Helvetica-Bold", 14)
        p.drawRightString(w-50, y, f"TOTAL: RM {total:.2f}")
    
    y_footer = 150
    p.line(40, y_footer, w-40, y_footer)
    p.setFont("Helvetica-Bold", 9)
    p.drawString(40, y_footer-15, "TERMS & CONDITIONS:")
    p.setFont("Helvetica", 8)
    tnc_text = cfg.get("TNC_Invoice", "") if type in ["INVOICE", "SALES"] else cfg.get("TNC_Ticket", "")
    curr_y = y_footer - 30
    for line in tnc_text.split('\n'):
        p.drawString(40, curr_y, line.strip())
        curr_y -= 12
    
    y_sig = 50
    p.line(50, y_sig, 200, y_sig)
    p.line(350, y_sig, 500, y_sig)
    p.drawString(50, y_sig-15, "Customer Signature")
    p.drawString(350, y_sig-15, "Technician / Manager")
    p.save()
    buffer.seek(0)
    return buffer

# =============================================================================
# 5. EMAIL & LINKS
# =============================================================================
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
    if status == 'Pending':
        msg += "\nKami telah menerima peranti anda untuk pemeriksaan."
    elif status in ['Done', 'Collected']:
        msg += "\n✅ Berita Baik! Peranti anda telah SIAP dibaiki."
    else:
        msg += f"\nStatus terkini peranti anda: {status}"
    
    msg += f"\n\n--- BUTIRAN ---\nID: {tid}\nModel: {model}\nS/N: {sn}\nMasalah: {masalah}\nSolution: {note}"
    
    if status in ['Done', 'Collected']:
        msg += f"\n\n💰 TOTAL: RM {price:.2f}"
    
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
    sender = st.session_state.email_user
    password = st.session_state.email_pass
    if not sender or not password:
        return False, "Sila set Email di Tetapan."
    
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = to_email
    msg['Subject'] = f"Tiket: {data.get('ID', '-')}"
    msg.attach(MIMEText(generate_message_content(data), 'plain'))
    
    part = MIMEApplication(pdf_buffer.getvalue(), Name=pdf_name)
    part['Content-Disposition'] = f'attachment; filename="{pdf_name}"'
    msg.attach(part)
    
    try:
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(sender, password)
        s.send_message(msg)
        s.quit()
        return True, "Berjaya!"
    except Exception as e:
        return False, str(e)

# =============================================================================
# 🚦 V70 GATEKEEPER LOGIC: DIGITAL HEALTH CARD (PROFILING MODE)
# =============================================================================
query_params = st.query_params 
sn_query = query_params.get("sn", None)

if sn_query:
    # --- MOD PUBLIC (DIGITAL HEALTH CARD) ---
    st.markdown("""
    <style>
        [data-testid="stSidebar"] {display: none;}
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

    load_config()
    cfg = st.session_state.config
    
    # Header: Digital Profile
    c_logo, c_title = st.columns([1, 4])
    with c_title:
        st.title(f"💻 {cfg.get('Company_Name', 'DCK TECH')} - Digital Profile")
        st.caption("Rekod Profil & Sejarah Peranti")
    st.divider()

    with st.spinner(f"🔍 Memuatkan profil S/N: {sn_query}..."):
        df = load_data("Tickets")
        
    found = False
    if not df.empty:
        # --- FIX: ROBUST S/N MATCHING ---
        history = df[df['SN'].astype(str).str.strip().str.upper() == str(sn_query).strip().upper()]
        
        if not history.empty:
            found = True
            latest = history.iloc[-1]
            
            # --- 1. DEVICE IDENTITY CARD (No Status) ---
            with st.container(border=True):
                c1, c2 = st.columns(2)
                c1.write(f"**Model:** {latest.get('Model')}")
                c2.write(f"**S/N:** {latest.get('SN')}")
            
            # --- 2. UNIFIED HISTORY TIMELINE (ALL RECORDS) ---
            st.write("### 📜 Sejarah Pembaikan")
            
            # Loop through ALL records (Newest first)
            for _, row in history.iloc[::-1].iterrows():
                # Clean Timeline UI: Date - Problem
                with st.expander(f"📅 {row['Tarikh']} : {row['Masalah']}"):
                    # Green Solution Box
                    st.markdown(f"""
                    <div style="background-color: #d4edda; padding: 10px; border-radius: 5px; border: 1px solid #c3e6cb; color: #155724;">
                        <strong>✅ Solution / Tindakan:</strong><br>
                        {row['Tech_Note']}
                    </div>
                    """, unsafe_allow_html=True)
                    
            st.divider()
            st.caption(f"Profil ini dijana automatik oleh sistem {cfg.get('Company_Name')}")

    if not found:
        st.error(f"❌ Maaf, tiada profil dijumpai untuk S/N: {sn_query}")
        st.warning("Sila pastikan Serial Number dimasukkan dengan betul.")

    st.stop()


# =============================================================================
# 6. NAVIGATION & PAGES
# =============================================================================
load_config()
PAGES = ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🛒 JUALAN KEDAI", "🔧 UPDATE STATUS", "📦 PENGURUSAN STOK", "🔎 HISTORY DEVICE", "📈 LAPORAN", "⚙️ TETAPAN"]
try: idx = PAGES.index(st.session_state.page)
except: idx = 0
sel = st.sidebar.radio("NAVIGASI", PAGES, index=idx)
if sel != st.session_state.page: st.session_state.page = sel; st.rerun()

# === PAGE: DASHBOARD ===
if st.session_state.page == "📊 DASHBOARD":
    st.title("📊 DCK Tech Dashboard")
    df = load_data("Tickets")
    df_s = load_data("Sales")
    
    today_str = datetime.now().strftime("%Y-%m-%d")
    sales_today = 0.0
    if not df_s.empty:
        sales_today += df_s[df_s['Tarikh'] == today_str]['Total'].apply(safe_float).sum()

    if not df.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.info(f"PENDING: {len(df[df['Status'] == 'Pending'])}")
        c2.warning(f"CHECKING: {len(df[df['Status'] == 'Checking'])}")
        c3.success(f"DONE: {len(df[df['Status'] == 'Done'])}")
        c4.error(f"COLLECTED: {len(df[df['Status'] == 'Collected'])}")
    
    st.divider()
    st.metric("💰 JUALAN KEDAI (HARI INI)", f"RM {sales_today:.2f}")
    
    st.write("### 🔍 Cari Ticket")
    search = st.text_input("Masukkan Nama / ID / Model:", placeholder="Contoh: DCK-12345")
    st.write("### Senarai Job Terkini")
    if not df.empty:
        if search: df = df[df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]
        for _, row in df.iloc[::-1].head(10).iterrows():
            with st.expander(f"{row.get('ID')} - {row.get('Customer')} ({row.get('Status')})"):
                st.write(f"Model: {row.get('Model')} | Masalah: {row.get('Masalah')}")
                if st.button("🔧 Manage Job", key=f"btn_{row.get('ID')}"):
                    st.session_state.selected_id = row.get('ID')
                    st.session_state.page = "🔧 UPDATE STATUS"
                    st.rerun()
    else:
        st.info("Tiada rekod tiket.")

# === PAGE: DAFTAR TIKET ===
elif st.session_state.page == "📝 DAFTAR TIKET":
    st.title("📝 Tiket Masuk Baru")
    
    # V69: Read options from Config (Dynamic Checkbox)
    opt_mslh = [x.strip() for x in st.session_state.config.get("Options_Masalah", "Slow, Screen, Battery").split(",")]
    opt_fiz = [x.strip() for x in st.session_state.config.get("Options_Fizikal", "Calar, Pecah").split(",")]
    opt_acc = [x.strip() for x in st.session_state.config.get("Options_Aksesori", "Bag, Charger").split(",")]

    with st.form("reg_ticket", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("### 👤 Pelanggan")
            nama = st.text_input("Nama")
            phone = st.text_input("No HP")
            email = st.text_input("Email")
        with c2:
            st.markdown("### 💻 Peranti")
            model = st.text_input("Model")
            sn = st.text_input("Serial No")
            pwd = st.text_input("Password")
        with c3:
            st.markdown("### 🔧 Diagnosis")
            mslh = st.multiselect("Masalah", opt_mslh)
            fiz = st.multiselect("Fizikal", opt_fiz)
            acc = st.multiselect("Aksesori", opt_acc)
        
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
                        try:
                            img_str = ",".join([cloudinary.uploader.upload(f)['secure_url'] for f in img])
                        except:
                            pass
                    
                    row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, email, model, sn, pwd, ",".join(mslh), ",".join(fiz), ",".join(acc), "Pending", 0, 0, img_str, note]
                    add_row("Tickets", row)
                    st.toast("Tiket Berjaya Dibuka!", icon='✅')
                    st.session_state.last_data = {"ID": tid, "Customer": nama, "Phone": phone, "Email": email, "Model": model, "SN": sn, "Masalah": ",".join(mslh), "Fizikal": ",".join(fiz), "Aksesori": ",".join(acc), "Tech_Note": note, "Status": "Pending", "Images": img_str, "Image_Link": img_str, "Tarikh": row[1]}
                    time.sleep(1)
                    st.rerun()

    if 'last_data' in st.session_state:
        ld = st.session_state.last_data
        st.divider()
        st.success(f"Tiket {ld['ID']} Telah Dibuka!")
        c1, c2, c3 = st.columns(3)
        c1.download_button("📥 PDF Tiket", generate_pdf(ld, "SERVICE"), "Tiket.pdf", use_container_width=True)
        c2.link_button("📱 WhatsApp", generate_links("WA", ld), use_container_width=True)
        if ld['Email']:
            if c3.button("📧 Email Auto"):
                ok, m = send_email_with_pdf(ld['Email'], ld, generate_pdf(ld, "SERVICE"), "Tiket.pdf")
                if ok:
                    st.toast(m, icon='✅')
                else:
                    st.error(m)

# === PAGE: JUALAN KEDAI (V74 - SIMPLE & BULK) ===
elif st.session_state.page == "🛒 JUALAN KEDAI":
    st.title("🛒 Sistem Jualan (POS)")
    df_m = load_data("Master_Inventory")
    
    stock_options = {}
    if not df_m.empty:
        for idx, row in df_m.iterrows():
            if safe_int(row['CurrentStock']) > 0:
                stock_options[f"{row['ItemName']} (Stok: {row['CurrentStock']}) - RM{row['SellPrice']}"] = row

    tab_pos, tab_manage = st.tabs(["🛒 Kaunter Bayaran", "📋 Rekod Jualan"])
    with tab_pos:
        with st.container(border=True):
            st.subheader("➕ Tambah Barang")
            with st.form("add_item_form", clear_on_submit=True):
                c1, c2, c3 = st.columns([3, 1, 1])
                selected_item_label = c1.selectbox("Pilih Barang (Dari Stok)", ["-"] + list(stock_options.keys()))
                qty = c2.number_input("Qty", 1, 100, 1)
                warr = c3.selectbox("Warranty Kedai", ["Tiada", "1 Bulan", "3 Bulan"])
                
                if st.form_submit_button("Masuk Bakul"):
                    if selected_item_label != "-":
                        item_data = stock_options[selected_item_label]
                        current_stock = safe_int(item_data['CurrentStock'])
                        
                        if qty <= current_stock:
                            price = safe_float(item_data['SellPrice'])
                            st.session_state.pos_cart.append({
                                "ItemCode": item_data['ItemCode'],
                                "Item": item_data['ItemName'],
                                "Qty": qty,
                                "Harga_Unit": price,
                                "Total": qty * price,
                                "Warranty": warr
                            })
                            st.toast(f"{item_data['ItemName']} ditambah!", icon='🛒')
                        else:
                            st.error(f"Stok tak cukup! Tinggal {current_stock} unit.")
                    else:
                        st.warning("Sila pilih barang.")

        if st.session_state.pos_cart:
            st.divider()
            st.subheader("🛍️ Bakul Jualan")
            st.markdown("---")
            for i, row in enumerate(st.session_state.pos_cart):
                c1, c2, c3, c4 = st.columns([3, 1, 1, 0.5])
                c1.write(f"**{row['Item']}** ({row['Warranty']})")
                c2.write(f"x {row['Qty']}")
                c3.write(f"RM {row['Total']:.2f}")
                if c4.button("❌", key=f"del_{i}"):
                    st.session_state.pos_cart.pop(i)
                    st.rerun()
            st.markdown("---")
            grand_total = sum([x['Total'] for x in st.session_state.pos_cart])
            st.metric("GRAND TOTAL", f"RM {grand_total:.2f}")
            
            with st.form("checkout_form"):
                c1, c2 = st.columns(2)
                cust_name = c1.text_input("Nama Pelanggan", "Walk-in")
                pay_method = c2.selectbox("Cara Bayaran", ["Cash", "QR DuitNow", "Online Transfer"])
                
                if st.form_submit_button("✅ Bayar & Cetak Resit", use_container_width=True):
                    sid = f"SALE-{datetime.now().strftime('%d%H%M')}"
                    tgl = datetime.now().strftime("%Y-%m-%d")
                    for row in st.session_state.pos_cart:
                        add_row("Sales", [sid, tgl, row['Item'], row['Qty'], row['Harga_Unit'], row['Total'], cust_name, pay_method, row['Warranty']])
                        if 'ItemCode' in row:
                            update_stock(row['ItemCode'], -row['Qty'])
                            
                    st.session_state.last_sale = {"ID": sid, "Tarikh": tgl, "Customer": cust_name, "Total": grand_total, "Items_List": st.session_state.pos_cart}
                    st.session_state.pos_cart = []
                    st.toast("Transaksi Berjaya! Stok Dikemaskini.", icon='💰')
                    time.sleep(1)
                    st.rerun()
            
            if st.button("❌ Kosongkan Semua Bakul"):
                st.session_state.pos_cart = []
                st.rerun()
            
        if 'last_sale' in st.session_state and st.session_state.last_sale:
            st.divider()
            st.success("Transaksi Selesai.")
            st.download_button("🖨️ CETAK RESIT", generate_pdf(st.session_state.last_sale, "SALES"), "Resit_Jualan.pdf", use_container_width=True)
            
    with tab_manage:
        df_s = load_data("Sales")
        st.dataframe(df_s, use_container_width=True)

# === PAGE: UPDATE STATUS ===
elif st.session_state.page == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Technician")
    df = load_data("Tickets")
    df_m = load_data("Master_Inventory")
    
    stock_options = {}
    if not df_m.empty:
        for idx, row in df_m.iterrows():
            if safe_int(row['CurrentStock']) > 0:
                stock_options[f"{row['ItemName']} (Stok: {row['CurrentStock']}) - Cost: RM{row['CostPrice']}"] = row

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
                            st.toast("Info Berjaya Diubah!", icon='✅')
                            time.sleep(1)
                            st.rerun()
            else:
                c1, c2 = st.columns(2)
                c1.write(f"**Nama:** {job.get('Customer','-')}")
                c1.write(f"**Model:** {job.get('Model','-')}")
                c1.write(f"**Phone:** {job.get('Phone','-')}")
                c1.write(f"**S/N:** {job.get('SN','-')}")
                c2.write(f"**Masalah:** {job.get('Masalah','-')}")
                c2.error(f"🔐 PWD: {job.get('Password','-')}")

            st.divider()
            
            with st.expander("🖨️ GENERATE QR STICKER (HEALTH CARD)"):
                st.info("Tampal ini di bawah laptop customer.")
                
                my_url = "https://dck-ticketing-system-r7vonv3ctwvxzfh2yqn4s5.streamlit.app"
                app_url = st.text_input("Link Sistem (Auto-Set)", value=my_url)
                
                if app_url and job.get('SN'):
                    if app_url.endswith("/"): app_url = app_url[:-1]
                    qr_data = f"{app_url}/?sn={job.get('SN')}"
                    
                    qr = qrcode.QRCode(box_size=10, border=2)
                    qr.add_data(qr_data)
                    qr.make(fit=True)
                    img_qr = qr.make_image(fill_color="black", back_color="white")
                    
                    buffer = BytesIO()
                    img_qr.save(buffer, format="PNG")
                    img_bytes = buffer.getvalue()
                    
                    c_qr1, c_qr2 = st.columns([1, 2])
                    c_qr1.image(img_bytes, caption=f"S/N: {job.get('SN')}", width=150)
                    c_qr2.write(f"**URL:** {qr_data}")
                    c_qr2.info("👉 Right-click gambar QR > 'Save Image' untuk print.")

            st.divider()
            ca, cb = st.columns(2)
            wa_url = generate_links("WA", job)
            ca.link_button("📱 WhatsApp Status", wa_url, use_container_width=True)
            if job.get('Email'): 
                with cb:
                    if st.button("📧 Hantar Email + PDF (Auto Server)", key=f"em_{pid}"):
                        doc = "INVOICE" if job.get('Status') in ['Done', 'Collected'] else "SERVICE"
                        ok, m = send_email_with_pdf(job.get('Email'), job, generate_pdf(job, doc), "Status.pdf")
                        if ok:
                            st.toast(m, icon='✅')
                        else:
                            st.error(m)
            else:
                cb.caption("Tiada Email.")
        
        c_l, c_r = st.columns([1, 2])
        with c_l:
            st.download_button("📄 Cetak Tiket", generate_pdf(job, "SERVICE"), "Tiket.pdf")
            img_str = str(job.get('Image_Link',''))
            if img_str and img_str != "nan":
                urls = img_str.split(",")
                for u in urls:
                    if u.startswith("http"):
                        st.image(u, use_container_width=True)

        with c_r:
            df_p = load_data("Parts")
            # --- FIX: CHECK IF DATAFRAME HAS DATA & COLUMNS BEFORE FILTERING ---
            if not df_p.empty and 'TicketID' in df_p.columns:
                 parts = df_p[df_p['TicketID'].astype(str) == str(pid)]
            else:
                 parts = pd.DataFrame()
            
            kos = sum([safe_float(x) for x in parts['HargaBeli'].tolist()]) if not parts.empty else 0
            
            t1, t2 = st.tabs(["Status", "Parts (Stok)"])
            with t1:
                st.write(f"**KOS:** RM {kos:.2f}")
                with st.form("upd_status"):
                    stt = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Done", "Collected"], index=["Pending", "Checking", "Waiting Part", "Done", "Collected"].index(job.get('Status','Pending')) if job.get('Status','Pending') in ["Pending", "Checking", "Waiting Part", "Done", "Collected"] else 0)
                    nt = st.text_area("Solution", value=job.get('Tech_Note',''))
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
                            st.toast("Status Dikemaskini!", icon='🎉')
                            time.sleep(1)
                            st.rerun()
                if stt in ["Done", "Collected"]:
                    st.download_button("🖨️ CETAK RESIT", generate_pdf(job, "INVOICE"), "Resit.pdf", use_container_width=True)

            with t2:
                if not parts.empty:
                    st.dataframe(parts[['NamaPart', 'WarrantyBulan', 'TarikhExpire', 'HargaBeli']], use_container_width=True)
                    d = st.selectbox("Pilih ID untuk Hapus (Guna jika salah masuk):", ["-"] + parts['ID'].tolist())
                    if d != "-" and st.button("Hapus Part"):
                        delete_part(d)
                        st.toast("Part dihapus!", icon='🗑️')
                        time.sleep(1)
                        st.rerun()
                else:
                    st.info("Tiada part digunakan.")
                
                st.markdown("---")
                st.write("### ➕ Guna Part Dari Stok")
                with st.form("use_part_form", clear_on_submit=True):
                    sel_part = st.selectbox("Pilih Part", ["-"] + list(stock_options.keys()))
                    warr_part = st.radio("Warranty Part (Untuk Client)", [1, 3], horizontal=True, format_func=lambda x: f"{x} Bulan")
                    
                    if st.form_submit_button("Guna Part Ini"):
                        if sel_part != "-":
                            item_data = stock_options[sel_part]
                            exp = (datetime.now() + pd.DateOffset(months=int(warr_part))).strftime("%Y-%m-%d")
                            add_row("Parts", [f"P-{int(time.time())}", pid, item_data['ItemName'], item_data['Supplier'], str(datetime.now().date()), warr_part, exp, item_data['CostPrice']])
                            update_stock(item_data['ItemCode'], -1)
                            st.toast(f"{item_data['ItemName']} ditambah ke Job (Warranty {warr_part} Bulan)!", icon='✅')
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.warning("Pilih part dulu.")

# === PAGE: PENGURUSAN STOK (V74: BAKUL RESTOCK SIMPLE) ===
elif st.session_state.page == "📦 PENGURUSAN STOK":
    st.title("📦 Pengurusan Stok (Inventory)")
    t_restock, t_master, t_log = st.tabs(["📥 Masuk Stok (Restock)", "📋 Master List", "📜 Log Pembelian"])
    
    with t_restock:
        st.subheader("1. Maklumat Invoice")
        c1, c2 = st.columns(2)
        inv_no = c1.text_input("No Invoice / Resit Supplier")
        supp = c2.text_input("Nama Supplier")
        
        st.markdown("---")
        st.subheader("2. Isi Barang & Tambah ke List")
        
        # Helper: Existing Items for Autocomplete
        df_m = load_data("Master_Inventory")
        existing_items = df_m['ItemName'].tolist() if not df_m.empty else []
        
        with st.form("add_to_list_form", clear_on_submit=True):
            c_name, c_qty = st.columns([3, 1])
            # Pilihan: Taip Baru atau Pilih Lama
            item_input = c_name.text_input("Nama Barang (Taip baru atau copy nama lama)")
            if not item_input and existing_items:
                c_name.caption(f"Contoh barang sedia ada: {', '.join(existing_items[:3])}...")
            
            qty_input = c_qty.number_input("Qty", 1, 1000, 1)
            
            c_cost, c_sell = st.columns(2)
            cost_input = c_cost.number_input("Harga Kos (RM)", 0.0)
            sell_input = c_sell.number_input("Harga Jual (RM)", 0.0)
            
            if st.form_submit_button("⬇️ Tambah Ke List Bawah"):
                if item_input and cost_input > 0:
                    st.session_state.restock_cart.append({
                        "ItemName": item_input,
                        "Qty": qty_input,
                        "Cost": cost_input,
                        "Sell": sell_input,
                        "Total": qty_input * cost_input
                    })
                    st.toast(f"{item_input} masuk list!", icon='⬇️')
                else:
                    st.warning("Nama barang & Harga Kos wajib isi.")

        # DISPLAY LIST
        if st.session_state.restock_cart:
            st.divider()
            st.subheader("3. Semak & Simpan")
            
            # Show Table
            cart_df = pd.DataFrame(st.session_state.restock_cart)
            st.dataframe(cart_df, use_container_width=True)
            
            # Delete Button Logic
            if st.button("❌ Kosongkan List (Reset)"):
                st.session_state.restock_cart = []
                st.rerun()
            
            st.markdown("---")
            if st.button("✅ SIMPAN SEMUA KE DATABASE", type="primary", use_container_width=True):
                if inv_no and supp:
                    tgl = datetime.now().strftime("%Y-%m-%d")
                    log_id = f"LOG-{int(time.time())}"
                    
                    progress_text = "Sedang menyimpan..."
                    my_bar = st.progress(0, text=progress_text)
                    
                    for i, item in enumerate(st.session_state.restock_cart):
                        # Generate Code (Simple Hash based on name to keep consistent or Create New)
                        # Check if exists in Master
                        found_code = None
                        if not df_m.empty:
                            match = df_m[df_m['ItemName'].str.lower() == item['ItemName'].lower()]
                            if not match.empty:
                                found_code = match.iloc[0]['ItemCode']
                        
                        if not found_code:
                            found_code = f"ITM-{int(time.time())}-{i}" # Unique ID
                        
                        # 1. Update/Add Master
                        check_and_update_master(found_code, item['ItemName'], item['Cost'], item['Sell'], item['Qty'])
                        
                        # 2. Add Log
                        add_row("Restock_Log", [log_id, tgl, inv_no, supp, item['ItemName'], item['Qty'], item['Cost'], item['Total']])
                        
                        my_bar.progress((i + 1) / len(st.session_state.restock_cart), text=f"Menyimpan {item['ItemName']}...")
                    
                    st.session_state.restock_cart = [] # Clear cart
                    my_bar.empty()
                    st.success("Selesai! Semua stok telah direkodkan.")
                    time.sleep(2)
                    st.rerun()
                else:
                    st.error("Sila isi No Invoice dan Nama Supplier di atas (Bahagian 1).")
        else:
            st.info("List kosong. Sila isi barang di atas.")

    with t_master:
        df_m = load_data("Master_Inventory")
        if not df_m.empty:
            st.dataframe(df_m, use_container_width=True)
        else:
            st.info("Belum ada stok.")

    with t_log:
        df_r = load_data("Restock_Log")
        if not df_r.empty:
            st.dataframe(df_r, use_container_width=True)
        else:
            st.info("Tiada rekod pembelian.")

# === PAGE: HISTORY DEVICE ===
elif st.session_state.page == "🔎 HISTORY DEVICE":
    st.title("🔎 Semakan Sejarah Peranti (SN)")
    st.info("Masukkan Serial Number (S/N) untuk melihat rekod servis lama.")
    search_sn = st.text_input("Scan / Taip Serial Number:", placeholder="Contoh: SN12345678")
    
    if search_sn:
        df = load_data("Tickets")
        if not df.empty:
            history = df[df['SN'].astype(str).str.strip().str.upper() == search_sn.strip().upper()]
            if not history.empty:
                st.success(f"Jumpa {len(history)} rekod untuk S/N: {search_sn}")
                st.dataframe(history[['Tarikh', 'ID', 'Customer', 'Model', 'Masalah', 'Status', 'Tech_Note']], use_container_width=True)
                st.write("---")
                st.write("### 📜 Butiran Terperinci")
                for _, row in history.iterrows():
                    with st.expander(f"{row['Tarikh']} - {row['Masalah']} ({row['Status']})"):
                        st.write(f"**ID Tiket:** {row['ID']}")
                        st.write(f"**Customer:** {row['Customer']}")
                        st.write(f"**Solution (Tech Note):** {row['Tech_Note']}")
                        if st.button("Buka Job Ini", key=f"hist_{row['ID']}"):
                            st.session_state.selected_id = row['ID']
                            st.session_state.page = "🔧 UPDATE STATUS"
                            st.rerun()
            else:
                st.warning("Tiada rekod dijumpai untuk S/N ini.")
        else:
            st.error("Database kosong.")

# === PAGE: LAPORAN ===
elif st.session_state.page == "📈 LAPORAN":
    st.title("📈 Laporan Prestasi")
    df = load_data("Tickets")
    df_s = load_data("Sales")
    df_p = load_data("Parts")
    
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
                txt = ",".join(df['Masalah'].astype(str))
                itm = [x.strip() for x in txt.split(",") if x.strip()]
                if itm:
                    st.bar_chart(pd.Series(itm).value_counts().head(5))
        with c2:
            st.subheader("🔩 Barang Laju (Parts)")
            if not df_p.empty:
                st.bar_chart(df_p['NamaPart'].value_counts().head(5))
        
        st.divider()
        st.subheader("📅 Prestasi Berkala")
        tab_h, tab_m, tab_b = st.tabs(["Harian", "Mingguan", "Bulanan"])
        
        with tab_h:
            st.line_chart(df.groupby(df['Tarikh'].dt.date)[['Harga_Jual', 'Untung']].sum().tail(30))
        with tab_m:
            st.bar_chart(df.groupby(df['Tarikh'].dt.to_period('W').astype(str))[['Harga_Jual', 'Untung']].sum())
        with tab_b:
            st.bar_chart(df.groupby(df['Tarikh'].dt.to_period('M').astype(str))[['Harga_Jual', 'Untung']].sum())
            
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
        
        st.divider()
        st.markdown("### 🔧 Pilihan Masalah & Fizikal (Asingkan dengan koma)")
        opt_m = st.text_area("Senarai Masalah", st.session_state.config.get("Options_Masalah", "Slow, Screen, Battery, Keyboard, Format, Water Damage, Other"))
        opt_f = st.text_area("Senarai Fizikal", st.session_state.config.get("Options_Fizikal", "Calar, Pecah, Kemek, Sempurna"))
        opt_a = st.text_area("Senarai Aksesori", st.session_state.config.get("Options_Aksesori", "Bag, Charger, Mouse, Box"))

        if st.form_submit_button("Simpan"):
            new_c = st.session_state.config.copy()
            new_c.update({
                "Company_Name":cn, "Address":ad, "Phone":ph, "TNC_Ticket":tnc1, "TNC_Invoice":tnc2,
                "Options_Masalah": opt_m, "Options_Fizikal": opt_f, "Options_Aksesori": opt_a
            })
            save_config_to_db(new_c)
            st.toast("Tetapan Disimpan!", icon='✅')
    
    with st.expander("📧 Email Server"):
        with st.form("em"):
            eu = st.text_input("Email", st.session_state.email_user)
            ep = st.text_input("App Password", st.session_state.email_pass, type="password")
            if st.form_submit_button("Set Email"):
                st.session_state.email_user = eu
                st.session_state.email_pass = ep
                st.toast("Email Disimpan!", icon='✅')
