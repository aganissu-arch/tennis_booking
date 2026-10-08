import datetime
import os
import sqlite3
import time
from zoneinfo import ZoneInfo
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# ----------------- ตั้งค่าหน้าเว็บ -----------------
st.set_page_config(
    page_title="Tennis Appointment System",
    layout="centered",
    initial_sidebar_state="collapsed",
)

SLIP_DIR = "uploads/slips"
os.makedirs(SLIP_DIR, exist_ok=True)

# ----------------- ตรวจสอบโหมดฐานข้อมูล -----------------
def is_gsheet_enabled() -> bool:
    try:
        return (
            hasattr(st, "secrets")
            and "gcp_service_account" in st.secrets
            and "sheets" in st.secrets
        )
    except Exception:
        return False

# ----------------- 1. จัดการฝั่ง Google Sheets -----------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

@st.cache_resource
def get_gspread_client():
    if not is_gsheet_enabled():
        return None
    try:
        service_account_info = dict(st.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(
            service_account_info,
            scopes=SCOPES,
        )
        return gspread.authorize(credentials)
    except Exception:
        return None

@st.cache_resource
def get_spreadsheet():
    client = get_gspread_client()
    if not client:
        return None
    try:
        url = st.secrets["sheets"]["spreadsheet_url"]
        return client.open_by_url(url)
    except Exception:
        return None

@st.cache_resource
def ensure_worksheets():
    sh = get_spreadsheet()
    if not sh:
        return
    try:
        titles = [ws.title for ws in sh.worksheets()]
        if "Bookings" not in titles:
            ws_b = sh.add_worksheet(title="Bookings", rows="1000", cols="10")
            ws_b.append_row(["id", "play_date", "player_name", "is_paid", "slip_file", "created_at"])
        if "DailySettings" not in titles:
            ws_s = sh.add_worksheet(title="DailySettings", rows="500", cols="5")
            ws_s.append_row(["play_date", "total_cost"])
    except Exception:
        pass

ensure_worksheets()

# Caching Data เพื่อไม่ให้ส่ง Request เกินโควตา 60 ครั้ง/นาที ของ Google
@st.cache_data(ttl=15, show_spinner=False)
def fetch_bookings_gsheet_cached():
    sh = get_spreadsheet()
    if not sh:
        return []
    try:
        ws = sh.worksheet("Bookings")
        return ws.get_all_records()
    except Exception as e:
        if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e):
            st.warning("Google API Rate limit: กำลังรอพักคำขอ 1 นาที (กำลังดึงข้อมูลสำรอง)")
        return []

@st.cache_data(ttl=15, show_spinner=False)
def fetch_settings_gsheet_cached():
    sh = get_spreadsheet()
    if not sh:
        return []
    try:
        ws = sh.worksheet("DailySettings")
        return ws.get_all_records()
    except Exception:
        return []

def get_bookings_gsheet(selected_date: str):
    records = fetch_bookings_gsheet_cached()
    results = []
    for r in records:
        if str(r.get("play_date")) == selected_date:
            results.append((
                r.get("id"),
                r.get("player_name"),
                int(r.get("is_paid", 0)),
                str(r.get("slip_file", "")),
                str(r.get("created_at", ""))
            ))
    return results

def add_booking_gsheet(selected_date: str, player_name: str):
    sh = get_spreadsheet()
    if not sh:
        return False, "ไม่สามารถเชื่อมต่อฐานข้อมูลได้"
    try:
        ws = sh.worksheet("Bookings")
        records = fetch_bookings_gsheet_cached()
        date_records = [r for r in records if str(r.get("play_date")) == selected_date]
        if len(date_records) >= 6:
            return False, "จำนวนผู้ลงชื่อครบกำหนด 6 ท่านแล้ว"
        if any(str(r.get("player_name")).strip().lower() == player_name.strip().lower() for r in date_records):
            return False, f"ชื่อ '{player_name.strip()}' มีในระบบสำหรับวันนี้แล้ว"

        b_id = int(datetime.datetime.now().timestamp() * 1000)
        created_at = datetime.datetime.now(ZoneInfo("Asia/Bangkok")).strftime("%Y-%m-%d %H:%M:%S")
        ws.append_row([b_id, selected_date, player_name.strip(), 0, "", created_at])
        st.cache_data.clear()
        return True, "บันทึกการลงชื่อเรียบร้อยแล้ว"
    except Exception as e:
        if "429" in str(e):
            return False, "ระบบกำลังรอคิว Google Sheets (Rate limit) โปรดลองใหม่อีกครั้งใน 30 วินาที"
        return False, f"เกิดข้อผิดพลาดในการบันทึก: {e}"

def find_row_by_id(ws, booking_id):
    try:
        cell = ws.find(str(booking_id), in_column=1)
        if cell:
            return cell.row
    except Exception:
        pass
    try:
        col_vals = ws.col_values(1)
        for idx, val in enumerate(col_vals, start=1):
            if str(val).strip() == str(booking_id).strip():
                return idx
    except Exception:
        pass
    return None

def delete_booking_gsheet(booking_id):
    sh = get_spreadsheet()
    if not sh:
        return
    try:
        ws = sh.worksheet("Bookings")
        row_idx = find_row_by_id(ws, booking_id)
        if row_idx:
            ws.delete_rows(row_idx)
        st.cache_data.clear()
    except Exception:
        pass

def update_payment_status_gsheet(booking_id, is_paid: int):
    sh = get_spreadsheet()
    if not sh:
        return
    try:
        ws = sh.worksheet("Bookings")
        row_idx = find_row_by_id(ws, booking_id)
        if row_idx:
            ws.update_cell(row_idx, 4, is_paid)
        st.cache_data.clear()
    except Exception:
        pass

def update_slip_gsheet(booking_id, slip_filename: str):
    sh = get_spreadsheet()
    if not sh:
        return
    try:
        ws = sh.worksheet("Bookings")
        row_idx = find_row_by_id(ws, booking_id)
        if row_idx:
            ws.update_cell(row_idx, 4, 1) # จ่ายแล้ว (column 4)
            ws.update_cell(row_idx, 5, slip_filename) # slip_file (column 5)
        st.cache_data.clear()
    except Exception:
        pass

def get_daily_settings_gsheet(selected_date: str) -> dict:
    records = fetch_settings_gsheet_cached()
    for r in records:
        if str(r.get("play_date")) == selected_date:
            try:
                cost = float(r.get("total_cost", 0))
            except Exception:
                cost = 0.0
            status = str(r.get("court_status", "")).strip() or "รอยืนยัน"
            return {
                "total_cost": cost,
                "court_status": status,
                "court_number": str(r.get("court_number", "")).strip(),
                "court_time": str(r.get("court_time", "")).strip(),
                "court_note": str(r.get("court_note", "")).strip(),
            }
    return {
        "total_cost": 0.0,
        "court_status": "รอยืนยัน",
        "court_number": "",
        "court_time": "",
        "court_note": "",
    }

def set_daily_settings_gsheet(selected_date: str, total_cost: float, court_status: str, court_number: str, court_time: str, court_note: str):
    sh = get_spreadsheet()
    if not sh:
        return
    try:
        ws = sh.worksheet("DailySettings")
        cell = ws.find(selected_date, in_column=1)
        row_vals = [selected_date, total_cost, court_status, court_number, court_time, court_note]
        if cell:
            ws.update(f"A{cell.row}:F{cell.row}", [row_vals])
        else:
            ws.append_row(row_vals)
        st.cache_data.clear()
    except Exception:
        pass

# ----------------- 2. จัดการฝั่ง SQLite (Fallback) -----------------
DB_FILE = "tennis_bookings.db"

def init_sqlite():
    with sqlite3.connect(DB_FILE) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS bookings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                play_date TEXT NOT NULL,
                player_name TEXT NOT NULL,
                is_paid INTEGER DEFAULT 0,
                slip_file TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("PRAGMA table_info(bookings)")
        cols_b = [row[1] for row in cursor.fetchall()]
        if "is_paid" not in cols_b:
            cursor.execute("ALTER TABLE bookings ADD COLUMN is_paid INTEGER DEFAULT 0")
        if "slip_file" not in cols_b:
            cursor.execute("ALTER TABLE bookings ADD COLUMN slip_file TEXT")

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS daily_settings (
                play_date TEXT PRIMARY KEY,
                total_cost REAL DEFAULT 0,
                court_status TEXT DEFAULT 'รอยืนยัน',
                court_number TEXT DEFAULT '',
                court_time TEXT DEFAULT '',
                court_note TEXT DEFAULT ''
            )
        """)
        cursor.execute("PRAGMA table_info(daily_settings)")
        cols_s = [row[1] for row in cursor.fetchall()]
        if "court_status" not in cols_s:
            cursor.execute("ALTER TABLE daily_settings ADD COLUMN court_status TEXT DEFAULT 'รอยืนยัน'")
        if "court_number" not in cols_s:
            cursor.execute("ALTER TABLE daily_settings ADD COLUMN court_number TEXT DEFAULT ''")
        if "court_time" not in cols_s:
            cursor.execute("ALTER TABLE daily_settings ADD COLUMN court_time TEXT DEFAULT ''")
        if "court_note" not in cols_s:
            cursor.execute("ALTER TABLE daily_settings ADD COLUMN court_note TEXT DEFAULT ''")
        conn.commit()

def get_bookings_sqlite(selected_date: str):
    with sqlite3.connect(DB_FILE) as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id, player_name, is_paid, slip_file, created_at FROM bookings WHERE play_date = ? ORDER BY id ASC",
            (selected_date,),
        )
        return cursor.fetchall()

def add_booking_sqlite(selected_date: str, player_name: str):
    with sqlite3.connect(DB_FILE) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM bookings WHERE play_date = ?", (selected_date,))
        count = cursor.fetchone()[0]
        if count >= 6:
            return False, "จำนวนผู้ลงชื่อครบกำหนด 6 ท่านแล้ว"

        cursor.execute(
            "SELECT COUNT(*) FROM bookings WHERE play_date = ? AND LOWER(player_name) = LOWER(?)",
            (selected_date, player_name.strip()),
        )
        if cursor.fetchone()[0] > 0:
            return False, f"ชื่อ '{player_name.strip()}' มีในระบบสำหรับวันนี้แล้ว"

        cursor.execute(
            "INSERT INTO bookings (play_date, player_name, is_paid, slip_file) VALUES (?, ?, 0, '')",
            (selected_date, player_name.strip()),
        )
        conn.commit()
        return True, "บันทึกการลงชื่อเรียบร้อยแล้ว"

def delete_booking_sqlite(booking_id: int):
    with sqlite3.connect(DB_FILE) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM bookings WHERE id = ?", (booking_id,))
        conn.commit()

def update_payment_status_sqlite(booking_id: int, is_paid: int):
    with sqlite3.connect(DB_FILE) as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE bookings SET is_paid = ? WHERE id = ?", (is_paid, booking_id))
        conn.commit()

def update_slip_sqlite(booking_id: int, slip_filename: str):
    with sqlite3.connect(DB_FILE) as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE bookings SET is_paid = 1, slip_file = ? WHERE id = ?", (slip_filename, booking_id))
        conn.commit()

def get_daily_settings_sqlite(selected_date: str) -> dict:
    with sqlite3.connect(DB_FILE) as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT total_cost, court_status, court_number, court_time, court_note FROM daily_settings WHERE play_date = ?",
            (selected_date,),
        )
        row = cursor.fetchone()
        if row:
            return {
                "total_cost": float(row[0]) if row[0] is not None else 0.0,
                "court_status": row[1] if row[1] else "รอยืนยัน",
                "court_number": row[2] or "",
                "court_time": row[3] or "",
                "court_note": row[4] or "",
            }
        return {
            "total_cost": 0.0,
            "court_status": "รอยืนยัน",
            "court_number": "",
            "court_time": "",
            "court_note": "",
        }

def set_daily_settings_sqlite(selected_date: str, total_cost: float, court_status: str, court_number: str, court_time: str, court_note: str):
    with sqlite3.connect(DB_FILE) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO daily_settings (play_date, total_cost, court_status, court_number, court_time, court_note)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(play_date) DO UPDATE SET
                total_cost = excluded.total_cost,
                court_status = excluded.court_status,
                court_number = excluded.court_number,
                court_time = excluded.court_time,
                court_note = excluded.court_note
        """, (selected_date, total_cost, court_status, court_number, court_time, court_note))
        conn.commit()

init_sqlite()

# ----------------- 3. Unified Database Router -----------------
use_gsheet = is_gsheet_enabled()

def get_bookings(selected_date: str):
    if use_gsheet:
        return get_bookings_gsheet(selected_date)
    return get_bookings_sqlite(selected_date)

def add_booking(selected_date: str, player_name: str):
    if use_gsheet:
        return add_booking_gsheet(selected_date, player_name)
    return add_booking_sqlite(selected_date, player_name)

def delete_booking(booking_id):
    if use_gsheet:
        delete_booking_gsheet(booking_id)
    else:
        delete_booking_sqlite(booking_id)

def update_payment_status(booking_id, is_paid: int):
    if use_gsheet:
        update_payment_status_gsheet(booking_id, is_paid)
    else:
        update_payment_status_sqlite(booking_id, is_paid)

def update_slip(booking_id, slip_filename: str):
    if use_gsheet:
        update_slip_gsheet(booking_id, slip_filename)
    else:
        update_slip_sqlite(booking_id, slip_filename)

def get_daily_settings(selected_date: str) -> dict:
    if use_gsheet:
        return get_daily_settings_gsheet(selected_date)
    return get_daily_settings_sqlite(selected_date)

def set_daily_settings(selected_date: str, total_cost: float, court_status: str, court_number: str, court_time: str, court_note: str):
    if use_gsheet:
        set_daily_settings_gsheet(selected_date, total_cost, court_status, court_number, court_time, court_note)
    else:
        set_daily_settings_sqlite(selected_date, total_cost, court_status, court_number, court_time, court_note)

def get_daily_cost(selected_date: str) -> float:
    return get_daily_settings(selected_date)["total_cost"]

def set_daily_cost(selected_date: str, cost: float):
    curr = get_daily_settings(selected_date)
    set_daily_settings(selected_date, cost, curr["court_status"], curr["court_number"], curr["court_time"], curr["court_note"])

# ----------------- Dialogs (Popup ยืนยัน) -----------------
@st.dialog("ยืนยันการลงชื่อ")
def confirm_signup_dialog(date_str: str, date_obj: datetime.date, name: str):
    st.write(f"ผู้เล่น: **{name}**")
    st.write(f"วันที่: **{date_obj.strftime('%d/%m/%Y')}**")
    st.caption("ระบบจำกัดจำนวน 6 ท่านต่อวัน โปรดตรวจสอบความถูกต้องก่อนยืนยัน")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("ยืนยัน", key="dlg_confirm_signup", type="primary", use_container_width=True):
            success, msg = add_booking(date_str, name)
            if success:
                st.session_state["just_signed"] = msg
            else:
                st.session_state["sign_error"] = msg
            st.rerun()
    with col2:
        if st.button("ยกเลิก", key="dlg_cancel_signup", use_container_width=True):
            st.rerun()

@st.dialog("ยืนยันการยกเลิกการจอง")
def confirm_cancel_dialog(booking_id, player_name: str, date_obj: datetime.date):
    st.write(f"ต้องการยกเลิกการจองของ: **{player_name}**")
    st.write(f"วันที่: **{date_obj.strftime('%d/%m/%Y')}**")
    st.caption("เมื่อยกเลิก สิทธิ์ในคิวจะเปิดว่างให้ผู้เล่นอื่นทันที")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("ยืนยันการยกเลิก", key="dlg_confirm_cancel", type="primary", use_container_width=True):
            delete_booking(booking_id)
            st.session_state["just_cancelled"] = f"ยกเลิกชื่อ '{player_name}' เรียบร้อยแล้ว"
            st.rerun()
    with col2:
        if st.button("ย้อนกลับ", key="dlg_back_cancel", use_container_width=True):
            st.rerun()

@st.dialog("แนบสลิปการโอนเงิน")
def upload_slip_dialog(booking_id, player_name: str):
    st.write(f"ผู้เล่น: **{player_name}**")
    if os.path.exists("assets/promptpay_qr.jpg"):
        with st.expander("คลิกดู QR Code สำหรับชำระเงิน"):
            st.image("assets/promptpay_qr.jpg", use_container_width=True)
    st.caption("อัปโหลดรูปภาพสลิปเพื่อยืนยันการชำระเงิน ระบบจะเปลี่ยนสถานะเป็น 'จ่ายแล้ว' ให้อัตโนมัติ")
    
    uploaded_file = st.file_uploader("เลือกรูปภาพสลิป", type=["jpg", "jpeg", "png"], key=f"uploader_{booking_id}")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("บันทึกสลิป", type="primary", use_container_width=True):
            if uploaded_file is not None:
                ext = os.path.splitext(uploaded_file.name)[1]
                filename = f"slip_{booking_id}_{int(time.time())}{ext}"
                filepath = os.path.join(SLIP_DIR, filename)
                with open(filepath, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                update_slip(booking_id, filename)
                st.session_state["just_signed"] = f"แนบสลิปของ '{player_name}' เรียบร้อย สถานะเปลี่ยนเป็นจ่ายแล้ว"
                st.rerun()
            else:
                st.error("กรุณาเลือกไฟล์สลิปก่อนบันทึก")
    with col2:
        if st.button("ปิด", use_container_width=True):
            st.rerun()

@st.dialog("หลักฐานการโอนเงิน")
def view_slip_dialog(booking_id, player_name: str, slip_filename: str):
    st.write(f"ผู้เล่น: **{player_name}**")
    st.write("สถานะ: **จ่ายแล้ว**")
    filepath = os.path.join(SLIP_DIR, slip_filename)
    if os.path.exists(filepath):
        st.image(filepath, caption="สลิปการโอนเงิน", use_container_width=True)
    else:
        st.info("แนบหลักฐานในระบบเรียบร้อยแล้ว")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("แนบสลิปใหม่", use_container_width=True):
            upload_slip_dialog(booking_id, player_name)
    with col2:
        if st.button("ปิด", use_container_width=True):
            st.rerun()

# ----------------- Professional Mobile-First Styling -----------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap');
    
    html, body, p, div, span:not([data-testid="stIconMaterial"]), input, button, select, textarea, label {
        font-family: 'Prompt', -apple-system, BlinkMacSystemFont, sans-serif !important;
        -webkit-tap-highlight-color: transparent;
    }
    [data-testid="stIconMaterial"] {
        font-family: 'Material Symbols Rounded', 'Material Symbols Outlined', 'Material Icons' !important;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2.5rem !important;
        padding-left: 0.9rem !important;
        padding-right: 0.9rem !important;
        max-width: 680px !important;
    }

    .club-header {
        background-color: #153e28;
        border-radius: 12px;
        padding: 1.2rem 1.4rem;
        color: #ffffff;
        margin-bottom: 0.8rem;
        border: 1px solid #1c5236;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.06);
    }
    .club-header-title {
        font-size: 1.35rem;
        font-weight: 700;
        letter-spacing: -0.3px;
        margin: 0;
        color: #ffffff;
    }
    .club-header-sub {
        font-size: 0.82rem;
        color: #a7f3d0;
        margin-top: 0.25rem;
        font-weight: 400;
        line-height: 1.4;
    }

    .db-status-bar {
        font-size: 0.75rem;
        padding: 5px 10px;
        border-radius: 6px;
        margin-bottom: 1rem;
        font-weight: 500;
        display: inline-block;
    }
    .db-status-cloud {
        background-color: #f0fdf4;
        color: #166534;
        border: 1px solid #bbf7d0;
    }

    .summary-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 8px;
        margin-bottom: 1.2rem;
    }
    .summary-box {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 10px 8px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    .summary-label {
        font-size: 0.72rem;
        color: #64748b;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.3px;
        margin-bottom: 3px;
    }
    .summary-val {
        font-size: 1.1rem;
        font-weight: 700;
        color: #0f172a;
        line-height: 1.2;
    }
    .summary-sub {
        font-size: 0.68rem;
        color: #10b981;
        font-weight: 600;
        margin-top: 2px;
    }

    .slot-row {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 10px 12px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 0.45rem;
        min-height: 54px;
    }
    .slot-row-filled {
        border-left: 4px solid #153e28;
    }
    .slot-row-empty {
        border-left: 4px solid #cbd5e1;
        background: #fafafa;
    }

    .slot-badge {
        font-size: 0.78rem;
        font-weight: 700;
        width: 26px;
        height: 26px;
        border-radius: 6px;
        display: flex;
        align-items: center;
        justify-content: center;
        margin-right: 10px;
        flex-shrink: 0;
    }
    .slot-badge-filled {
        background-color: #153e28;
        color: #ffffff;
    }
    .slot-badge-empty {
        background-color: #e2e8f0;
        color: #64748b;
    }

    .slot-info {
        flex-grow: 1;
        min-width: 0;
    }
    .player-title {
        font-size: 0.95rem;
        font-weight: 600;
        color: #0f172a;
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 5px;
    }
    .empty-title {
        font-size: 0.88rem;
        font-weight: 400;
        color: #94a3b8;
    }

    .fee-tag {
        font-size: 0.72rem;
        font-weight: 600;
        color: #0369a1;
        background-color: #e0f2fe;
        padding: 2px 6px;
        border-radius: 4px;
        display: inline-block;
    }

    .status-badge-paid {
        font-size: 0.72rem;
        font-weight: 600;
        color: #166534;
        background-color: #dcfce7;
        border: 1px solid #bbf7d0;
        padding: 2px 7px;
        border-radius: 4px;
        display: inline-block;
    }
    .status-badge-unpaid {
        font-size: 0.72rem;
        font-weight: 500;
        color: #854d0e;
        background-color: #fef9c3;
        border: 1px solid #fef08a;
        padding: 2px 7px;
        border-radius: 4px;
        display: inline-block;
    }

    .pill-empty {
        font-size: 0.72rem;
        font-weight: 500;
        color: #64748b;
        background-color: #f1f5f9;
        padding: 3px 8px;
        border-radius: 6px;
    }

    .section-title {
        font-size: 0.95rem;
        font-weight: 700;
        color: #0f172a;
        margin-top: 1rem;
        margin-bottom: 0.6rem;
        letter-spacing: -0.2px;
    }

    div.stButton > button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 0.82rem !important;
        min-height: 38px !important;
        transition: all 0.15s ease !important;
    }

    button[data-testid="baseButton-primary"],
    div.stButton > button[kind="primary"] {
        background-color: #153e28 !important;
        color: #ffffff !important;
        border: 1px solid #153e28 !important;
    }
    button[data-testid="baseButton-primary"]:hover,
    div.stButton > button[kind="primary"]:hover {
        background-color: #1c5236 !important;
        border-color: #1c5236 !important;
        color: #ffffff !important;
    }

    button[data-testid="baseButton-secondary"],
    div.stButton > button[kind="secondary"] {
        background-color: #ffffff !important;
        color: #334155 !important;
        border: 1px solid #cbd5e1 !important;
    }
    button[data-testid="baseButton-secondary"]:hover,
    div.stButton > button[kind="secondary"]:hover {
        background-color: #f8fafc !important;
        border-color: #94a3b8 !important;
        color: #0f172a !important;
    }

    /* เด่นพิเศษ: แผงวันที่เล่น (Date Hero Card) */
    .date-hero-card {
        background: #ffffff;
        border: 2px solid #153e28;
        border-left: 6px solid #153e28;
        border-radius: 12px;
        padding: 12px 16px;
        margin-bottom: 0.65rem;
        box-shadow: 0 4px 12px rgba(21, 62, 40, 0.08);
    }
    .date-hero-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 3px;
    }
    .date-hero-tag {
        font-size: 0.78rem;
        font-weight: 700;
        color: #153e28;
        letter-spacing: 0.4px;
        text-transform: uppercase;
    }
    .date-hero-pill {
        font-size: 0.75rem;
        font-weight: 700;
        color: #ffffff;
        background-color: #153e28;
        padding: 2px 10px;
        border-radius: 999px;
    }
    .date-hero-title {
        font-size: 1.25rem;
        font-weight: 800;
        color: #0f172a;
        letter-spacing: -0.3px;
        line-height: 1.3;
    }

    /* ช่องวันที่เล่นเด่นๆ (High Contrast & Prominent) */
    div[data-testid="stDateInput"] {
        background: #f0fdf4 !important;
        border: 2px solid #153e28 !important;
        border-radius: 10px !important;
        padding: 6px 10px 8px 10px !important;
        box-shadow: 0 2px 8px rgba(21, 62, 40, 0.08) !important;
    }
    div[data-testid="stDateInput"] label {
        font-size: 0.95rem !important;
        font-weight: 800 !important;
        color: #153e28 !important;
        margin-bottom: 3px !important;
    }
    div[data-testid="stDateInput"] div[data-baseweb="input"] {
        background: #ffffff !important;
        border: 1.5px solid #153e28 !important;
        border-radius: 6px !important;
    }
    div[data-testid="stDateInput"] input,
    div[data-testid="stDateInput"] input:disabled,
    div[data-testid="stDateInput"] input[disabled] {
        font-size: 1.12rem !important;
        font-weight: 800 !important;
        color: #153e28 !important;
        -webkit-text-fill-color: #153e28 !important;
        opacity: 1 !important;
        cursor: default !important;
        letter-spacing: 0.3px !important;
    }

    /* ช่องค่าสนามรวมให้เข้าชุดกัน */
    div[data-testid="stNumberInput"] {
        background: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
        padding: 6px 10px 8px 10px !important;
    }
    div[data-testid="stNumberInput"] label {
        font-size: 0.88rem !important;
        font-weight: 600 !important;
        color: #334155 !important;
        margin-bottom: 3px !important;
    }
    div[data-testid="stNumberInput"] input {
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        color: #0f172a !important;
    }

    /* ประกาศสถานะสนามประจำวัน (Daily Announcement) */
    .announcement-box {
        border-radius: 12px;
        padding: 12px 14px;
        margin-bottom: 0.85rem;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
    }
    .announcement-box-confirmed {
        background-color: #f0fdf4;
        border: 1.5px solid #166534;
        border-left: 5px solid #166534;
    }
    .announcement-box-pending {
        background-color: #fefce8;
        border: 1.5px solid #ca8a04;
        border-left: 5px solid #ca8a04;
    }
    .announcement-box-cancelled {
        background-color: #fef2f2;
        border: 1.5px solid #dc2626;
        border-left: 5px solid #dc2626;
    }

    .announcement-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 6px;
    }
    .announcement-label {
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 0.4px;
        text-transform: uppercase;
    }
    .announcement-label-confirmed { color: #166534; }
    .announcement-label-pending { color: #854d0e; }
    .announcement-label-cancelled { color: #991b1b; }

    .announcement-badge {
        font-size: 0.75rem;
        font-weight: 700;
        padding: 2px 9px;
        border-radius: 4px;
    }
    .announcement-badge-confirmed {
        background-color: #166534;
        color: #ffffff;
    }
    .announcement-badge-pending {
        background-color: #ca8a04;
        color: #ffffff;
    }
    .announcement-badge-cancelled {
        background-color: #dc2626;
        color: #ffffff;
    }

    .announcement-title {
        font-size: 1.05rem;
        font-weight: 700;
        margin-bottom: 4px;
        letter-spacing: -0.2px;
    }
    .announcement-title-confirmed { color: #14532d; }
    .announcement-title-pending { color: #713f12; }
    .announcement-title-cancelled { color: #7f1d1d; }

    .announcement-meta {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-top: 6px;
    }
    .announcement-meta-item {
        background: rgba(255, 255, 255, 0.85);
        border: 1px solid rgba(0, 0, 0, 0.08);
        border-radius: 6px;
        padding: 3px 9px;
        font-size: 0.82rem;
        font-weight: 600;
        color: #0f172a;
    }
    .announcement-note {
        font-size: 0.82rem;
        margin-top: 6px;
        color: #475569;
        font-weight: 500;
        line-height: 1.4;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- แจ้งเตือนสถานะ -----------------
if "just_signed" in st.session_state:
    st.toast(st.session_state.pop("just_signed"))
if "just_cancelled" in st.session_state:
    st.toast(st.session_state.pop("just_cancelled"))
if "sign_error" in st.session_state:
    st.error(st.session_state.pop("sign_error"))

# ----------------- Header Section -----------------
st.markdown(
    '<div class="club-header">'
    '<div class="club-header-title">ระบบนัดหมายเล่นเทนนิส</div>'
    '<div class="club-header-sub">จำกัด 6 ท่านต่อวัน • ระบบหารค่าสนามและแนบสลิปชำระเงิน</div>'
    '</div>',
    unsafe_allow_html=True,
)

if use_gsheet:
    st.markdown('<div class="db-status-bar db-status-cloud">Cloud Database: Google Sheets ซิงค์เรียบร้อย</div>', unsafe_allow_html=True)

# ----------------- กำหนดวันที่ & สถานะประจำวัน -----------------
today_th = datetime.datetime.now(ZoneInfo("Asia/Bangkok")).date()
date_str = today_th.strftime("%Y-%m-%d")
selected_date = today_th

THAI_DAYS = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"]
THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]

def get_thai_date_label(d: datetime.date) -> str:
    day_name = THAI_DAYS[d.weekday()]
    month_name = THAI_MONTHS[d.month - 1]
    year_be = d.year + 543
    return f"วัน{day_name}ที่ {d.day} {month_name} {year_be}"

active_date_thai = get_thai_date_label(today_th)

# การ์ดแสดงวันที่เล่นเด่นชัด (Date Hero Card)
st.markdown(
    f'<div class="date-hero-card">'
    f'<div class="date-hero-top">'
    f'<span class="date-hero-tag">วันที่เล่น</span>'
    f'<span class="date-hero-pill">วันนี้</span>'
    f'</div>'
    f'<div class="date-hero-title">{active_date_thai}</div>'
    f'</div>',
    unsafe_allow_html=True,
)

daily_info = get_daily_settings(date_str)
c_status = daily_info.get("court_status", "รอยืนยัน")
c_num = daily_info.get("court_number", "")
c_time = daily_info.get("court_time", "")
c_note = daily_info.get("court_note", "")

# ----------------- กล่องประกาศสถานะสนามประจำวัน -----------------
if c_status == "ยืนยันแล้ว":
    meta_parts = []
    if c_num:
        meta_parts.append(f'<span class="announcement-meta-item">คอร์ท: {c_num}</span>')
    if c_time:
        meta_parts.append(f'<span class="announcement-meta-item">เวลา: {c_time}</span>')
    meta_html = f'<div class="announcement-meta">{"".join(meta_parts)}</div>' if meta_parts else ''
    note_html = f'<div class="announcement-note">{c_note}</div>' if c_note else ''
    st.markdown(
        f'<div class="announcement-box announcement-box-confirmed">'
        f'<div class="announcement-header">'
        f'<span class="announcement-label announcement-label-confirmed">ประกาศสถานะสนาม</span>'
        f'<span class="announcement-badge announcement-badge-confirmed">ยืนยันแล้ว</span>'
        f'</div>'
        f'<div class="announcement-title announcement-title-confirmed">สนามได้รับการยืนยันเรียบร้อยแล้ว</div>'
        f'{meta_html}'
        f'{note_html}'
        f'</div>',
        unsafe_allow_html=True,
    )
elif c_status == "งดเล่น":
    note_html = f'<div class="announcement-note">{c_note}</div>' if c_note else '<div class="announcement-note">เนื่องจากสนามไม่ว่างหรือสภาพอากาศไม่เอื้ออำนวย</div>'
    st.markdown(
        f'<div class="announcement-box announcement-box-cancelled">'
        f'<div class="announcement-header">'
        f'<span class="announcement-label announcement-label-cancelled">ประกาศสถานะสนาม</span>'
        f'<span class="announcement-badge announcement-badge-cancelled">งดเล่น</span>'
        f'</div>'
        f'<div class="announcement-title announcement-title-cancelled">วันนี้งดเล่นเทนนิส</div>'
        f'{note_html}'
        f'</div>',
        unsafe_allow_html=True,
    )
else: # รอยืนยัน
    time_part = f'<span class="announcement-meta-item">เวลาคาดการณ์: {c_time}</span>' if c_time else ''
    meta_html = f'<div class="announcement-meta">{time_part}</div>' if time_part else ''
    note_html = f'<div class="announcement-note">{c_note}</div>' if c_note else '<div class="announcement-note">อยู่ระหว่างรอผลการจองสนามและจัดสรรเวลา</div>'
    st.markdown(
        f'<div class="announcement-box announcement-box-pending">'
        f'<div class="announcement-header">'
        f'<span class="announcement-label announcement-label-pending">ประกาศสถานะสนาม</span>'
        f'<span class="announcement-badge announcement-badge-pending">รอยืนยัน</span>'
        f'</div>'
        f'<div class="announcement-title announcement-title-pending">กำลังรอยืนยันผลการจองสนาม</div>'
        f'{meta_html}'
        f'{note_html}'
        f'</div>',
        unsafe_allow_html=True,
    )

with st.expander("แก้ไขประกาศ / สถานะสนามประจำวัน", expanded=False):
    status_options = ["รอยืนยัน", "ยืนยันแล้ว", "งดเล่น"]
    status_idx = status_options.index(c_status) if c_status in status_options else 0
    new_status = st.selectbox("สถานะสนาม", status_options, index=status_idx, key=f"sel_stat_{date_str}")
    
    col_st_c, col_st_t = st.columns(2)
    with col_st_c:
        new_court = st.text_input("หมายเลขคอร์ท", value=c_num, placeholder="เช่น คอร์ท 1, 2", key=f"in_court_{date_str}", disabled=(new_status == "งดเล่น"))
    with col_st_t:
        new_time = st.text_input("ช่วงเวลาเล่น", value=c_time, placeholder="เช่น 18:00 - 20:00 น.", key=f"in_time_{date_str}", disabled=(new_status == "งดเล่น"))
    
    new_note = st.text_input("ข้อความประกาศเพิ่มเติม / หมายเหตุ", value=c_note, placeholder="เช่น วันนี้สภาพอากาศแจ่มใส หรือ นำลูกเทนนิสมาเอง", key=f"in_note_{date_str}")
    
    if st.button("บันทึกประกาศสถานะสนาม", type="primary", key=f"btn_save_ann_{date_str}"):
        set_daily_settings(date_str, daily_info.get("total_cost", 0.0), new_status, new_court.strip(), new_time.strip(), new_note.strip())
        st.toast("บันทึกประกาศสถานะสนามเรียบร้อยแล้ว")
        st.rerun()

col_date, col_cost = st.columns([1.1, 1.1])

with col_date:
    st.date_input(
        "วันที่เล่น",
        value=today_th,
        min_value=today_th,
        max_value=today_th,
        format="DD/MM/YYYY",
        disabled=True,
        help="ระบบเปิดให้ลงชื่อเฉพาะวันนี้เท่านั้น",
    )

saved_cost = daily_info.get("total_cost", 0.0)

with col_cost:
    total_cost = st.number_input(
        "ค่าสนามรวม (บาท)",
        min_value=0.0,
        value=float(saved_cost),
        step=50.0,
        format="%.0f",
        help="กรอกยอดรวมของวันนั้น ระบบจะบันทึกและหารเฉลี่ยให้ผู้เล่นทุกคนอัตโนมัติ",
    )
    if total_cost != saved_cost:
        set_daily_cost(date_str, total_cost)
        st.rerun()

bookings = get_bookings(date_str)
count = len(bookings)
remaining = 6 - count

split_per_person = (total_cost / count) if count > 0 else 0.0
# นับคนที่จ่ายแล้ว (เช็กจาก is_paid == 1 หรือมี slip_file)
paid_count = sum(1 for b in bookings if (b[2] == 1 or bool(b[3])))

# ----------------- สรุปข้อมูลภาพรวม -----------------
rem_text = f"ว่าง {remaining} ที่" if remaining > 0 else "เต็ม"
cost_display = f"{split_per_person:,.0f} บ." if total_cost > 0 and count > 0 else "-"

st.markdown(
    f'<div class="summary-grid">'
    f'<div class="summary-box"><div class="summary-label">ผู้เล่น</div><div class="summary-val">{count}/6</div><div class="summary-sub">{rem_text}</div></div>'
    f'<div class="summary-box"><div class="summary-label">หารคนละ</div><div class="summary-val">{cost_display}</div><div class="summary-sub">ยอด {total_cost:,.0f} บ.</div></div>'
    f'<div class="summary-box"><div class="summary-label">ชำระแล้ว</div><div class="summary-val">{paid_count}/{count}</div><div class="summary-sub">คน</div></div>'
    f'</div>',
    unsafe_allow_html=True,
)

st.progress(count / 6.0)

# ----------------- รายชื่อผู้เล่น 6 สล็อต -----------------
st.markdown(f'<div class="section-title">รายชื่อผู้เล่น ({selected_date.strftime("%d/%m/%Y")})</div>', unsafe_allow_html=True)

slots_data = list(bookings)

for slot_num in range(1, 7):
    if slot_num <= len(slots_data):
        b_id, p_name, is_paid, slip_file, created_at = slots_data[slot_num - 1]
        
        # ตรวจสอบสถานะการจ่ายเงิน
        has_paid = (is_paid == 1 or bool(slip_file))
        status_badge = '<span class="status-badge-paid">จ่ายแล้ว</span>' if has_paid else '<span class="status-badge-unpaid">รอชำระ</span>'
        fee_html = f'<span class="fee-tag">{split_per_person:,.0f} ฿</span>' if total_cost > 0 else ''

        c_left, c_slip, c_del = st.columns([2.5, 1.1, 0.9])
        
        with c_left:
            card_html = (
                f'<div class="slot-row slot-row-filled">'
                f'<div style="display:flex;align-items:center;min-width:0;">'
                f'<div class="slot-badge slot-badge-filled">{slot_num}</div>'
                f'<div class="slot-info">'
                f'<div class="player-title"><span>{p_name}</span>{fee_html}{status_badge}</div>'
                f'</div>'
                f'</div>'
                f'</div>'
            )
            st.markdown(card_html, unsafe_allow_html=True)
            
        with c_slip:
            st.markdown("<div style='height: 3px;'></div>", unsafe_allow_html=True)
            if slip_file:
                if st.button("ดูสลิป", key=f"view_slip_{b_id}", use_container_width=True):
                    view_slip_dialog(b_id, p_name, slip_file)
            else:
                if st.button("แนบสลิป", key=f"upload_slip_{b_id}", use_container_width=True):
                    upload_slip_dialog(b_id, p_name)

        with c_del:
            st.markdown("<div style='height: 3px;'></div>", unsafe_allow_html=True)
            if st.button("ยกเลิก", key=f"del_{b_id}", help="ยกเลิกการจองนี้", use_container_width=True):
                confirm_cancel_dialog(b_id, p_name, selected_date)
    else:
        empty_html = (
            f'<div class="slot-row slot-row-empty">'
            f'<div style="display:flex;align-items:center;">'
            f'<div class="slot-badge slot-badge-empty">{slot_num}</div>'
            f'<div class="empty-title">ที่ว่างสำหรับผู้เล่นคนที่ {slot_num}</div>'
            f'</div>'
            f'<span class="pill-empty">ว่าง</span>'
            f'</div>'
        )
        st.markdown(empty_html, unsafe_allow_html=True)

# ----------------- ฟอร์มลงชื่อ -----------------
st.markdown('<div class="section-title">ลงชื่อร่วมเล่น</div>', unsafe_allow_html=True)

if c_status == "งดเล่น":
    st.warning("วันนี้ระบบงดการลงชื่อเล่นเทนนิสชั่วคราวเนื่องจากมีประกาศงดเล่น")
elif count >= 6:
    st.info("จำนวนผู้ลงชื่อในวันนี้ครบ 6 ท่านแล้ว (หากมีผู้เล่นยกเลิก ช่องจะเปิดว่างทันที)")
else:
    player_name_input = st.text_input(
        "ชื่อ หรือ ชื่อเล่น",
        placeholder="กรอกชื่อเพื่อลงชื่อจองสล็อต",
        key="player_name_input_field",
        label_visibility="collapsed",
    )

    if st.button("ยืนยันการลงชื่อ", use_container_width=True, type="primary"):
        cleaned_name = player_name_input.strip()
        if not cleaned_name:
            st.error("โปรดกรอกชื่อก่อนกดยืนยัน")
        else:
            confirm_signup_dialog(date_str, selected_date, cleaned_name)

# ----------------- ช่องทางชำระเงิน (QR Code) -----------------
QR_PATH = "assets/promptpay_qr.jpg"

if os.path.exists(QR_PATH):
    st.markdown("---")
    st.markdown('<div class="section-title">ช่องทางการชำระเงิน</div>', unsafe_allow_html=True)
    
    pay_note = f"ยอดชำระต่อคน: <b>{split_per_person:,.0f} บาท</b>" if split_per_person > 0 else "โปรดระบุยอดเงินตามแถบสรุปด้านบน"
    
    st.markdown(
        f'<div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:12px; text-align:center; max-width:320px; margin:0 auto 12px auto;">'
        f'<div style="font-size:0.88rem; font-weight:700; color:#0f172a;">สแกน QR Code พร้อมเพย์</div>'
        f'<div style="font-size:0.78rem; color:#64748b; margin-top:2px;">{pay_note}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    
    _, col_qr, _ = st.columns([1, 2, 1])
    with col_qr:
        st.image(QR_PATH, caption="เมื่อชำระแล้ว โปรดกดปุ่ม 'แนบสลิป' ที่ชื่อของท่าน", use_container_width=True)

