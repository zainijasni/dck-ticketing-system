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
        try:
            return func(*args, **kwargs)
        except Exception as e:
            time.sleep(1)
            continue
    return None

def generate_whatsapp_link(phone, nama, tid, model, status, img_links):
    clean_phone = str(phone).replace("+", "").replace("-", "").replace(" ", "")
    if clean_phone.startswith("0"): clean_phone = "6" + clean_phone
    msg = f"*DCK TECH SERVICES*\n\nHai {nama},\nTiket ID: {tid}\nModel: {model}\nStatus: {status}\n\nLihat gambar peranti:\n{img_links}\n\nTerima kasih!"
    return f"https://wa.me/{clean_phone}?text={urllib.parse.quote(msg)}"

def generate_email_link(email, nama, tid, model, status, img_links):
    subject = urllib.parse.quote(f"Update Tiket DCK Tech: {tid}")
    body = urllib.parse.quote(f"Hai {nama},\n\nStatus Terkini:\nTiket ID: {tid}\nModel: {model}\nStatus: {status}\n\nGambar:\n{img_links}\n\nTerima Kasih.\nDCK Tech Team")
    return f"mailto:{email}?subject={subject}&body={body}"

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
        
        # Tickets Header (16 Kolum)
        try: ws = sh.worksheet("Tickets")
        except: ws = sh.add_worksheet("Tickets", 1000, 20)
        h_t = ["ID", "Tarikh", "Customer", "Phone", "Email", "Model", "SN", "Password", "Masalah", "Fizikal", "Aksesori", "Status", "Kos_Part", "Harga_Jual", "Image_Link", "Tech_Note"]
        if len(ws.row_values(1)) != len(h_t): ws.update("A1:P1", [h_t])
        
        # Parts Header
        try: ws_p = sh.worksheet("Parts")
        except: ws_p = sh.add_worksheet("Parts", 1000, 10)
        h_p = ["ID", "TicketID", "NamaPart", "Supplier", "TarikhMasuk", "WarrantyBulan", "TarikhExpire", "HargaBeli"]
        if ws_p.row_values(1) != h_p: ws_p.update("A1:H1", [h_p])
        
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

# --- FUNGSI PENTING: UPDATE CUSTOMER INFO ---
def update_customer_info_db(tid, nama, phone, email, model, sn, pwd, masalah):
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet("Tickets")
    cell = robust_api_call(sheet.find, str(tid))
    if cell:
        # Col 3=Cust, 4=Phone, 5=Email, 6=Model, 7=SN, 8=Pwd, 9=Masalah
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
    p.drawString(50, h-130, f"Email: {t.get('Email', '-')}")
    p.drawString(300, h-130, f"Model: {t.get('Model', '-')}")
    p.drawString(50, h-145, f"Serial No: {t.get('SN', '-')}")
    
    y = h-170
    p.line(50, y+10, w-50, y+10)
    p.setFont("Helvetica-Bold", 10); p.drawString(50, y, "DETAIL DIAGNOSIS:"); y-=15
    p.setFont("Helvetica", 10)
    p.drawString(50, y, f"Masalah: {t.get('Masalah', '-')}"); y-=15
    p.drawString(50, y, f"Fizikal: {t.get('Fizikal', '-')}"); y-=15
    p.drawString(50, y, f"Aksesori: {t.get('Aksesori', '-')}"); y-=15
    p.drawString(50, y, f"Nota Tambahan: {t.get('Tech_Note', '-')}"); y-=30
    
    if type == "INVOICE":
        p.setFont("Helvetica-Bold", 14)
        p.drawString(50, y, f"TOTAL BILL: RM {safe_float(t.get('Harga_Jual', 0)):.2f}"); y-=30
    
    p.setFont("Helvetica-Bold", 10); p.drawString(50, y, "TERMA & SYARAT:"); y-=15
    tc = ["1. Data hilang tanggungjawab sendiri.", "2. Barang tak tuntut > 3 bulan hak milik kedai.", "3. Warranty sparepart shj."]
    p.setFont("Helvetica", 8)
    for line in tc: p.drawString(50, y, line); y-=12
    
    y -= 40
    p.drawString(50, y, "Tandatangan Pelanggan: _________________"); p.drawString(300, y, "Tandatangan Admin: _________________")
    p.save(); buffer.seek(0)
    return buffer

# --- 5. NAVIGATION ---
PAGES = ["📊 DASHBOARD", "📝 DAFTAR TIKET", "🔧 UPDATE STATUS", "📦 INVENTORY", "📈 LAPORAN"]
try: current_index = PAGES.index(st.session_state.page)
except: current_index = 0
selected_page = st.sidebar.radio("NAVIGASI UTAMA", PAGES, index=current_index)
if selected_page != st.session_state.page:
    st.session_state.page = selected_page
    st.rerun()

# === PAGE: DASHBOARD ===
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
            with st.expander(f"{row.get('ID', '-')} - {row.get('Customer', '-')} ({row.get('Status', '-')})"):
                st.write(f"Model: {row.get('Model', '-')} | Masalah: {row.get('Masalah', '-')}")
                if st.button("🔧 Manage Job", key=f"btn_{row['ID']}"):
                    st.session_state.selected_id = row['ID']
                    st.session_state.page = "🔧 UPDATE STATUS"
                    st.rerun()

# === PAGE: DAFTAR TIKET ===
elif st.session_state.page == "📝 DAFTAR TIKET":
    st.title("📝 Tiket Masuk Baru")
    with st.container(border=True):
        st.markdown("### 1. Info Pelanggan")
        c1, c2, c3 = st.columns(3)
        nama = c1.text_input("Nama"); phone = c2.text_input("No HP"); email = c3.text_input("Email (Optional)")
        
        st.markdown("### 2. Info Peranti")
        c4, c5, c6 = st.columns(3)
        model = c4.text_input("Model"); sn = c5.text_input("Serial No"); pwd = c6.text_input("Password Device")
        
        st.markdown("### 3. Diagnosis")
        mslh = st.multiselect("Masalah", ["Slow", "Screen Pecah", "Hinge Rosak", "Keyboard Rosak", "Tiada Display", "Tiada Power", "Format", "Upgrade", "Lain-lain"])
        fiz = st.multiselect("Fizikal", ["Calar", "Pecah", "Skru Hilang", "Sempurna"])
        acc = st.multiselect("Aksesori", ["Bag", "Charger", "Mouse", "Tiada"])
        note = st.text_area("Nota Tambahan")
        
        st.markdown("### 4. Gambar Peranti")
        uploaded_files = st.file_uploader("Pilih Gambar", accept_multiple_files=True, type=['jpg','png','jpeg'])
        tnc = st.checkbox("Setuju T&C")
        
        if st.button("SIMPAN REKOD & SEND NOTIF", use_container_width=True):
            if nama and tnc:
                with st.spinner("Processing..."):
                    tid = f"DCK-{datetime.now().strftime('%d%H%M')}"
                    img_urls = []
                    if uploaded_files:
                        for f in uploaded_files:
                            res = cloudinary.uploader.upload(f)
                            img_urls.append(res['secure_url'])
                    img_str = ",".join(img_urls)
                    row = [tid, datetime.now().strftime("%Y-%m-%d"), nama, phone, email, model, sn, pwd, ", ".join(mslh), ", ".join(fiz), ", ".join(acc), "Pending", 0, 0, img_str, note]
                    add_row("Tickets", row)
                    st.success("Berjaya!")
                    st.session_state.last_data = {"ID": tid, "Customer": nama, "Phone": phone, "Email": email, "Model": model, "SN": sn, "Masalah": ", ".join(mslh), "Fizikal": ", ".join(fiz), "Aksesori": ", ".join(acc), "Tarikh": row[1], "Tech_Note": note, "Status": "Pending", "Images": img_str}
                    st.rerun()

    if 'last_data' in st.session_state:
        ld = st.session_state.last_data
        st.divider()
        st.success(f"Tiket {ld['ID']} Telah Dibuka!")
        col_pdf, col_wa, col_email = st.columns(3)
        with col_pdf: st.download_button("📥 1. Download Tiket", generate_pdf(ld, "SERVICE"), "Tiket.pdf", use_container_width=True)
        with col_wa: st.link_button("📱 2. WhatsApp", generate_whatsapp_link(ld['Phone'], ld['Customer'], ld['ID'], ld['Model'], ld['Status'], ld['Images']), use_container_width=True)
        with col_email: 
            if ld['Email']: st.link_button("📧 3. Email", generate_email_link(ld['Email'], ld['Customer'], ld['ID'], ld['Model'], ld['Status'], ld['Images']), use_container_width=True)

# === PAGE: UPDATE STATUS (THE EDIT FIX) ===
elif st.session_state.page == "🔧 UPDATE STATUS":
    st.title("🔧 Bilik Technician")
    df = load_data("Tickets")
    if not df.empty:
        ids = df['ID'].astype(str).tolist()
        tgt = str(st.session_state.selected_id) if st.session_state.selected_id else ids[-1]
        pid = st.selectbox("Pilih Job:", ids, index=ids.index(tgt) if tgt in ids else 0)
        
        job = df[df['ID'].astype(str) == str(pid)].iloc[0]
        
        # --- BOX MAKLUMAT UTAMA ---
        with st.expander("ℹ️ MAKLUMAT PENUH TIKET & KOMUNIKASI", expanded=True):
            
            # --- CHECKBOX UNTUK EDIT CUSTOMER INFO ---
            edit_mode = st.checkbox("✏️ Tick Untuk Edit Maklumat Pelanggan")
            
            if edit_mode:
                st.warning("Anda sedang dalam Mode Edit. Sila berhati-hati.")
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
                    
                    if st.form_submit_button("💾 SIMPAN PERUBAHAN MAKLUMAT"):
                        if update_customer_info_db(pid, ec_nama, ec_phone, ec_email, ec_model, ec_sn, ec_pwd, ec_mslh):
                            st.success("Info Berjaya Diubah!"); st.rerun()
            else:
                # VIEW MODE (Tak Boleh Edit)
                c1, c2 = st.columns(2)
                c1.write(f"**Nama:** {job.get('Customer','-')}"); c1.write(f"**Model:** {job.get('Model','-')}")
                c1.write(f"**Phone:** {job.get('Phone','-')}"); c1.write(f"**Email:** {job.get('Email','-')}")
                c2.write(f"**Masalah:** {job.get('Masalah','-')}"); c2.error(f"🔐 PWD: {job.get('Password','-')}")

            st.divider()
            ca, cb = st.columns(2)
            ca.link_button("📱 WhatsApp Status", generate_whatsapp_link(job.get('Phone',''), job.get('Customer',''), job.get('ID',''), job.get('Model',''), job.get('Status',''), job.get('Image_Link','')))
            if job.get('Email'): cb.link_button("📧 Email Status", generate_email_link(job.get('Email',''), job.get('Customer',''), job.get('ID',''), job.get('Model',''), job.get('Status',''), job.get('Image_Link','')))
        
        c_left, c_right = st.columns([1, 2])
        
        with c_left:
            img_str = str(job.get('Image_Link',''))
            if img_str and img_str != "nan":
                urls = img_str.split(",")
                for u in urls:
                    if u.startswith("http"): st.image(u, use_container_width=True)
            st.download_button("📄 Cetak Tiket", generate_pdf(job, "SERVICE"), f"Tiket_{pid}.pdf")

        with c_right:
            df_p = load_data("Parts")
            parts = df_p[df_p['TicketID'].astype(str) == str(pid)]
            total_kos = sum([safe_float(x) for x in parts['HargaBeli'].tolist()]) if not parts.empty else 0
            
            tab_utama, tab_parts = st.tabs(["📝 Status & Invoice", "🔩 Parts & Kos"])
            
            with tab_utama:
                st.write(f"**TOTAL KOS PARTS:** RM {total_kos:.2f}")
                with st.form("upd_status"):
                    stt = st.selectbox("Status", ["Pending", "Checking", "Waiting Part", "Done", "Collected"], index=["Pending", "Checking", "Waiting Part", "Done", "Collected"].index(job.get('Status','Pending')) if job.get('Status','Pending') in ["Pending", "Checking", "Waiting Part", "Done", "Collected"] else 0)
                    nt = st.text_area("Nota Technician", value=job.get('Tech_Note',''))
                    hj = st.number_input("Harga Jual (Total Bill)", value=safe_float(job.get('Harga_Jual',0)))
                    
                    if st.form_submit_button("UPDATE STATUS & HARGA"):
                        sheet = get_client().open_by_key(SHEET_ID).worksheet("Tickets")
                        cl = robust_api_call(sheet.find, str(pid))
                        if cl:
                            robust_api_call(sheet.update_cell, cl.row, 12, stt)
                            robust_api_call(sheet.update_cell, cl.row, 13, total_kos)
                            robust_api_call(sheet.update_cell, cl.row, 14, hj)
                            robust_api_call(sheet.update_cell, cl.row, 16, nt)
                            st.cache_data.clear()
                            st.success("Updated!"); st.rerun()
                
                if stt in ["Done", "Collected"]:
                    st.download_button("🖨️ CETAK RESIT", generate_pdf(job, "INVOICE"), "Resit.pdf", use_container_width=True)

            with tab_parts:
                t_list, t_edit, t_add = st.tabs(["📋 List", "✏️ Edit", "➕ Add"])
                with t_list:
                    if not parts.empty:
                        st.dataframe(parts[['NamaPart', 'WarrantyBulan', 'TarikhExpire', 'HargaBeli']], use_container_width=True)
                        dp = st.selectbox("Hapus Part (Pilih ID):", ["-"] + parts['ID'].tolist())
                        if dp != "-" and st.button("Hapus"): 
                            if delete_part(dp): st.rerun()
                with t_edit:
                    if not parts.empty:
                        eid = st.selectbox("Edit Part:", parts['ID'].tolist())
                        cp = parts[parts['ID'] == eid].iloc[0]
                        with st.form("ep"):
                            n = st.text_input("Nama", value=cp['NamaPart']); s = st.text_input("Supp", value=cp['Supplier'])
                            c_e1, c_e2 = st.columns(2)
                            w = c_e1.number_input("Warr (Bln)", value=safe_int(cp['WarrantyBulan']))
                            h = c_e2.number_input("Harga", value=safe_float(cp['HargaBeli']))
                            if st.form_submit_button("Simpan"):
                                update_part_data(eid, n, s, h, w); st.rerun()
                with t_add:
                    with st.form("ap"):
                        n = st.text_input("Part"); s = st.text_input("Supp")
                        c_a1, c_a2 = st.columns(2)
                        w = c_a1.number_input("Warr (Bln)", value=1, min_value=0)
                        h = c_a2.number_input("Harga", 0.0)
                        if st.form_submit_button("Tambah"):
                            exp_date = (datetime.now() + pd.DateOffset(months=int(w))).strftime("%Y-%m-%d")
                            add_row("Parts", [f"P-{int(time.time())}", pid, n, s, str(datetime.now().date()), w, exp_date, h])
                            st.rerun()

# === PAGE: INVENTORY ===
elif st.session_state.page == "📦 INVENTORY":
    st.title("📦 Inventory Log")
    df_p = load_data("Parts")
    if not df_p.empty:
        for i, row in df_p.iterrows():
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                c1.write(f"**{row['NamaPart']}** (RM {row['HargaBeli']}) | Exp: {row['TarikhExpire']}")
                if c2.button("Go to Job", key=f"inv_{i}"):
                    st.session_state.selected_id = row['TicketID']
                    st.session_state.page = "🔧 UPDATE STATUS"
                    st.rerun()

# === PAGE: LAPORAN ===
elif st.session_state.page == "📈 LAPORAN":
    st.title("📈 Laporan Prestasi")
    df = load_data("Tickets")
    if not df.empty:
        df['Tarikh'] = pd.to_datetime(df['Tarikh'], errors='coerce')
        df['Harga_Jual'] = df['Harga_Jual'].apply(safe_float)
        df['Kos_Part'] = df['Kos_Part'].apply(safe_float)
        df['Untung'] = df['Harga_Jual'] - df['Kos_Part']
        
        m1, m2 = st.columns(2)
        m1.metric("Total Sales", f"RM {df['Harga_Jual'].sum():.2f}")
        m2.metric("Total Untung", f"RM {df['Untung'].sum():.2f}")
        
        st.divider()
        st.subheader("🔧 Analisis Masalah")
        if 'Masalah' in df.columns:
            all_text = ",".join(df['Masalah'].astype(str).tolist())
            all_items = [x.strip() for x in all_text.split(",") if x.strip() != ""]
            if all_items: st.bar_chart(pd.Series(all_items).value_counts().head(10))
            
        st.divider()
        tab_h, tab_m, tab_b = st.tabs(["📅 Harian", "📆 Mingguan", "🗓️ Bulanan"])
        with tab_h: st.line_chart(df.groupby(df['Tarikh'].dt.date)[['Harga_Jual', 'Untung']].sum().tail(30))
        with tab_m: st.bar_chart(df.groupby(df['Tarikh'].dt.to_period('W').astype(str))[['Harga_Jual', 'Untung']].sum())
        with tab_b: st.bar_chart(df.groupby(df['Tarikh'].dt.to_period('M').astype(str))[['Harga_Jual', 'Untung']].sum())
        
        st.download_button("📥 Download CSV", df.to_csv(index=False).encode('utf-8'), "Laporan.csv", "text/csv")
