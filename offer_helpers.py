import datetime
import os
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st

# กำหนด Spreadsheet ID ตามที่คุณต้องการเชื่อมต่อ
SPREADSHEET_ID = "1k_hSSdF50uYcRZVffPNh0NfpXvbJFjajlG26SRUpYMs"

try:
    import holidays

    thai_holidays = holidays.TH()
except ImportError:
    thai_holidays = []


@st.cache_resource
def get_gspread_client():
    """เชื่อมต่อ Google Sheets ผ่าน Streamlit Secrets"""
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credentials_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(
        credentials_dict, scopes=scope
    )
    client = gspread.authorize(creds)
    return client


def clean_text_val(val):
    """ฟังก์ชันช่วยลบ .0 ท้ายข้อความที่เกิดจากการอ่านเลข"""
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str


@st.cache_data(ttl=600)
def load_shops_data():
    """ดึงและจัดระเบียบข้อมูลร้านค้าจาก Google Sheet"""
    try:
        client = get_gspread_client()
        # เปิดไฟล์ผ่าน Spreadsheet ID และเลือก Worksheet ชื่อ "shops" (สามารถเปลี่ยนชื่อชีทได้ตามจริง)
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet("shops")
        data = sheet.get_all_records()
        df = pd.DataFrame(data)

        if df.empty:
            return pd.DataFrame(
                columns=["shop_name", "address", "phone", "tax_id"]
            )

        if "ชื่อร้านค้า" in df.columns or "shop_name" not in df.columns:
            col_name = (
                "ชื่อร้านค้า" if "ชื่อร้านค้า" in df.columns else df.columns[0]
            )
            col_addr = "ที่อยู่" if "ที่อยู่" in df.columns else df.columns[1]

            df["shop_clean"] = df[col_name].ffill()
            shops = []
            for shop, group in df.groupby("shop_clean", sort=False):
                if pd.isna(shop) or clean_text_val(shop) in [
                    "nan",
                    "None",
                    "",
                ]:
                    continue

                address, phone, tax_id = "", "", ""
                for val in group[col_addr].dropna():
                    val_str = clean_text_val(val)
                    if val_str.startswith("เบอร์โทร"):
                        phone = clean_text_val(
                            val_str.replace("เบอร์โทร", "", 1)
                        )
                    elif val_str.startswith("เลขประจำตัวผู้เสียภาษี"):
                        tax_id = clean_text_val(
                            val_str.replace("เลขประจำตัวผู้เสียภาษี", "", 1)
                        )
                    elif val_str not in ["nan", "None", "ที่อยู่", ""]:
                        if not address:
                            address = val_str

                shops.append(
                    {
                        "shop_name": clean_text_val(shop),
                        "address": address,
                        "phone": phone,
                        "tax_id": tax_id,
                    }
                )

            return pd.DataFrame(shops)

        for col in ["shop_name", "address", "phone", "tax_id"]:
            if col not in df.columns:
                df[col] = ""

        df = df.dropna(subset=["shop_name"])
        df["shop_name"] = df["shop_name"].apply(clean_text_val)

        for col in ["address", "phone", "tax_id"]:
            df[col] = df[col].apply(clean_text_val)

        return df

    except Exception as e:
        st.error(f"⚠️ ไม่สามารถดึงข้อมูลร้านค้าจาก Google Sheet ได้: {e}")
        return pd.DataFrame(columns=["shop_name", "address", "phone", "tax_id"])


def save_shops_data(df):
    """(กรณีต้องการบันทึกกลับ Google Sheet หรือเก็บบน Session)"""
    # หากต้องการบันทึกกลับไปที่ Google Sheet สามารถเขียนเพิ่มตรงนี้ได้
    # แต่เบื้องต้นใช้การคืนค่าสถานะจำลองหรือบันทึกลง Sheet โดยตรง
    try:
        client = get_gspread_client()
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet("shops")
        sheet.clear()
        sheet.update([df.columns.values.tolist()] + df.values.tolist())
        return True
    except Exception as e:
        st.error(f"⚠️ ไม่สามารถบันทึกข้อมูลลง Google Sheet ได้: {e}")
        return False


def add_business_days(start_date, num_days=5):
    """ฟังก์ชันคำนวณวันบวกเพิ่ม N วันทำการ"""
    current_date = start_date
    added_days = 0
    while added_days < num_days:
        current_date += datetime.timedelta(days=1)
        if current_date.weekday() < 5 and current_date not in thai_holidays:
            added_days += 1
    return current_date


@st.cache_data(ttl=600)
def load_teacher_data():
    """อ่านข้อมูลรายชื่อครูจาก Google Sheet ทุกชีท (ยกเว้นชีทสรุป)"""
    person_dict = {}
    person_options = [""]

    try:
        client = get_gspread_client()
        spreadsheet = client.open_by_key(SPREADSHEET_ID)

        for sheet in spreadsheet.worksheets():
            sheet_title = sheet.title
            if sheet_title in ["สรุป", "สรุปคศ.", "Sheet1"]:
                continue

            rows = sheet.get_all_values()
            for row in rows:
                if len(row) >= 3:
                    fname = str(row[1]).strip() if row[1] else ""
                    lname = str(row[2]).strip() if row[2] else ""
                    acad = str(row[3]).strip() if len(row) >= 4 and row[3] else ""

                    if fname and lname and lname not in ["nan", "None", ""]:
                        if not any(
                            fname.startswith(p) for p in ["ข้อมูล", "ลำดับ", "ชื่อ"]
                        ):
                            full_name = f"{fname} {lname}".strip()
                            if acad in ["nan", "None", ""]:
                                acad = ""
                            if full_name not in person_dict:
                                person_options.append(full_name)
                                person_dict[full_name] = acad
    except Exception as e:
        st.warning(f"⚠️ ไม่สามารถอ่านข้อมูลครูจาก Google Sheet ได้: {e}")

    return person_options, person_dict
