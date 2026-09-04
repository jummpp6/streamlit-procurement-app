import datetime
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st

# 🔗 กำหนด Spreadsheet ID ของ Google Sheet
SPREADSHEET_ID = "1k_hSSdF50uYcRZVffPNh0NfpXvbJFjajlG26SRUpYMs"

try:
    import holidays

    thai_holidays = holidays.TH()
except ImportError:
    thai_holidays = []


# 🟢 0. ฟังก์ชันเชื่อมต่อ Google Sheets ผ่าน Service Account
def get_gsheet_worksheet(sheet_name="Shops"):
    """เชื่อมต่อ Google Sheets ด้วย Service Account ผ่าน gspread"""
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
    return sheet


def clean_text_val(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return "" if val_str.lower() in ["nan", "none"] else val_str


# 🟢 1. โหลดข้อมูลร้านค้าจาก Google Sheets
@st.cache_data(ttl=60)
def load_shops_data():
    """อ่านข้อมูลร้านค้าจาก Google Sheets ผ่าน gspread"""
    try:
        ws = get_gsheet_worksheet("Shops")
        data = ws.get_all_records()
        df = pd.DataFrame(data)

        df.columns = df.columns.str.strip()
        for col in ["shop_name", "address", "phone", "tax_id"]:
            if col not in df.columns:
                df[col] = ""

        df = df.dropna(subset=["shop_name"])
        df = df[df["shop_name"].astype(str).str.strip() != ""]

        for col in ["shop_name", "address", "phone", "tax_id"]:
            df[col] = df[col].apply(clean_text_val)

        return df
    except Exception as e:
        st.error(f"Error loading shops from Google Sheets: {e}")
        return pd.DataFrame(columns=["shop_name", "address", "phone", "tax_id"])


# 🟢 2. บันทึกข้อมูลร้านค้ากลับลง Google Sheets
def save_shops_data(df):
    """นำ DataFrame ที่แก้ไข/เพิ่มแล้ว เขียนกลับลง Google Sheets"""
    try:
        ws = get_gsheet_worksheet("Shops")
        df_clean = df[["shop_name", "address", "phone", "tax_id"]].fillna("")
        header = df_clean.columns.tolist()
        values = [header] + df_clean.values.tolist()

        ws.clear()
        ws.update(values)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error saving shops to Google Sheets: {e}")
        return False


# 🟢 3. โหลดข้อมูลครูจาก Google Sheets (รองรับรับพารามิเตอร์เผื่อไฟล์หลักส่งชื่อไฟล์มา)
@st.cache_data(ttl=60)
def load_teacher_data(sheet_name="Teachers"):
    """อ่านข้อมูลครูจาก Google Sheets"""
    person_dict = {}
    person_options = [""]
    try:
        # หากไฟล์หลักเผลอส่งชื่อไฟล์ .xlsx มา ให้แปลงเป็นชื่อชีท Teachers แทน
        if isinstance(sheet_name, str) and sheet_name.endswith(".xlsx"):
            sheet_name = "Teachers"

        ws = get_gsheet_worksheet(sheet_name)
        rows = ws.get_all_values()

        if len(rows) > 1:
            for r in rows[1:]:
                if len(r) >= 2:
                    full_name = str(r[1]).strip() if r[1] else ""
                    acad = str(r[2]).strip() if len(r) >= 3 and r[2] else ""

                    if (
                        full_name
                        and full_name not in ["ชื่อ - สกุล", "ลำดับที่"]
                        and full_name not in person_dict
                    ):
                        person_options.append(full_name)
                        person_dict[full_name] = acad
    except Exception as e:
        st.warning(f"Error loading teachers from Google Sheets: {e}")

    return person_options, person_dict


# 🟢 4. คำนวณวันทำการ
def add_business_days(start_date, num_days=5):
    """ฟังก์ชันเพิ่มวันทำการ"""
    current_date = start_date
    added_days = 0
    while added_days < num_days:
        current_date += datetime.timedelta(days=1)
        if current_date.weekday() < 5 and current_date not in thai_holidays:
            added_days += 1
    return current_date
