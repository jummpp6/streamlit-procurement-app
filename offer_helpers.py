# import datetime
# import gspread
# from google.oauth2.service_account import Credentials
# import pandas as pd
# import streamlit as st

# # 🔗 กำหนด Spreadsheet ID ของ Google Sheet
# SPREADSHEET_ID = "1k_hSSdF50uYcRZVffPNh0NfpXvbJFjajlG26SRUpYMs"


# # 🟢 0. ฟังก์ชันเชื่อมต่อ Google Sheets ผ่าน Service Account
# def get_gsheet_worksheet(sheet_name="Shops"):
#     """เชื่อมต่อ Google Sheets ด้วย Service Account ผ่าน gspread"""
#     scopes = ["https://www.googleapis.com/auth/spreadsheets"]
#     # ดึงค่า JSON Credentials จาก Secrets ของ Streamlit
#     creds = Credentials.from_service_account_info(
#         st.secrets["gcp_service_account"], scopes=scopes
#     )
#     client = gspread.authorize(creds)
#     sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
#     return sheet


# def clean_text_val(val):
#     if pd.isna(val) or val is None:
#         return ""
#     val_str = str(val).strip()
#     if val_str.endswith(".0"):
#         val_str = val_str[:-2]
#     return "" if val_str.lower() in ["nan", "none"] else val_str


# # 🟢 1. โหลดข้อมูลร้านค้า
# @st.cache_data(ttl=60)  # ตั้ง Cache 1 นาที เพื่อไม่ให้ดึงข้อมูลถี่เกินไป
# def load_shops_data():
#     """อ่านข้อมูลร้านค้าจาก Google Sheets ผ่าน gspread"""
#     try:
#         ws = get_gsheet_worksheet("Shops")
#         data = ws.get_all_records()
#         df = pd.DataFrame(data)

#         # แปลงชื่อคอลัมน์และจัดการข้อมูลให้สะอาด
#         df.columns = df.columns.str.strip()
#         for col in ["shop_name", "address", "phone", "tax_id"]:
#             if col not in df.columns:
#                 df[col] = ""

#         df = df.dropna(subset=["shop_name"])
#         df = df[df["shop_name"].astype(str).str.strip() != ""]

#         for col in ["shop_name", "address", "phone", "tax_id"]:
#             df[col] = df[col].apply(clean_text_val)

#         return df
#     except Exception as e:
#         st.error(f"Error loading shops: {e}")
#         return pd.DataFrame(columns=["shop_name", "address", "phone", "tax_id"])


# # 🟢 2. บันทึกข้อมูลร้านค้ากลับลง Google Sheets
# def save_shops_data(df):
#     """นำ DataFrame ที่แก้ไข/เพิ่มแล้ว เขียนกลับลง Google Sheets"""
#     try:
#         ws = get_gsheet_worksheet("Shops")

#         # ทำความสะอาดข้อมูลก่อนส่งกลับ
#         df_clean = df[["shop_name", "address", "phone", "tax_id"]].fillna("")

#         # แปลง DataFrame เป็น List รวม Header
#         header = df_clean.columns.tolist()
#         values = [header] + df_clean.values.tolist()

#         # ล้าง Sheet เดิม แล้วเขียนข้อมูลใหม่ทั้งหมดลงไป
#         ws.clear()
#         ws.update(values)

#         # ล้าง Cache ของ Streamlit เพื่อให้หน้าเว็บโหลดข้อมูลใหม่ทันที
#         st.cache_data.clear()
#         return True
#     except Exception as e:
#         st.error(f"Error saving shops to Google Sheets: {e}")
#         return False


# # 🟢 3. โหลดข้อมูลครู (ปรับตามตารางล่าสุด: B=ชื่อ-นามสกุล, C=วิทยฐานะ)
# def load_teacher_data(sheet_name="Teachers"):
#     """อ่านข้อมูลครูจาก Google Sheets"""
#     person_dict = {}
#     person_options = [""]
#     try:
#         ws = get_gsheet_worksheet(sheet_name)
#         rows = ws.get_all_values()

#         # อ่านข้อมูลตั้งแต่แถวที่ 2 เป็นต้นไป (ข้าม Header แถวแรก)
#         if len(rows) > 1:
#             for r in rows[1:]:
#                 if len(r) >= 2:
#                     # คอลัมน์ B (Index 1) คือ ชื่อ-สกุล
#                     full_name = str(r[1]).strip() if r[1] else ""
#                     # คอลัมน์ C (Index 2) คือ วิทยฐานะ
#                     acad = str(r[2]).strip() if len(r) >= 3 and r[2] else ""

#                     # กรองเอาเฉพาะแถวที่มีข้อมูลชื่อ และไม่ตรงกับหัวข้อตาราง
#                     if (
#                         full_name
#                         and full_name not in ["ชื่อ - สกุล", "ลำดับที่"]
#                         and full_name not in person_dict
#                     ):
#                         person_options.append(full_name)
#                         person_dict[full_name] = acad

#     except Exception as e:
#         st.warning(f"Error loading teachers: {e}")

#     return person_options, person_dict


# # 🟢 4. คำนวณวันทำการ
# def add_business_days(start_date, num_days):
#     """ฟังก์ชันเพิ่มวันทำการ"""
#     current_date = start_date
#     added = 0
#     while added < num_days:
#         current_date += datetime.timedelta(days=1)
#         if current_date.weekday() < 5:
#             added += 1
#     return current_date

import datetime
import os
import pandas as pd
import streamlit as st

EXCEL_FILE = r"D:\Project_AutoDoc\data.xlsx"


def clean_text_val(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return "" if val_str.lower() in ["nan", "none"] else val_str


# 🟢 1. โหลดข้อมูลร้านค้าจาก Excel
@st.cache_data(ttl=5)
def load_shops_data():
    try:
        if not os.path.exists(EXCEL_FILE):
            return pd.DataFrame(columns=["shop_name", "address", "phone", "tax_id"])

        df = pd.read_excel(EXCEL_FILE, sheet_name="Shops", dtype=str)
        df.columns = df.columns.str.strip()
        for col in ["shop_name", "address", "phone", "tax_id"]:
            if col not in df.columns:
                df[col] = ""

        df = df.dropna(subset=["shop_name"])
        df = df[df["shop_name"].astype(str).str.strip() != ""]

        for col in ["shop_name", "address", "phone", "tax_id"]:
            df[col] = df[col].apply(clean_text_val)

        return df
    except Exception:
        return pd.DataFrame(columns=["shop_name", "address", "phone", "tax_id"])


# 🟢 2. บันทึกข้อมูลร้านค้าลง Excel (ต้องมีฟังก์ชันนี้)
def save_shops_data(df):
    try:
        df_clean = df[["shop_name", "address", "phone", "tax_id"]].fillna("")
        with pd.ExcelWriter(
            EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            df_clean.to_excel(writer, sheet_name="Shops", index=False)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error saving shops to Excel: {e}")
        return False


# 🟢 3. โหลดข้อมูลครูจาก Excel
def load_teacher_data(sheet_name="Teachers"):
    person_dict = {}
    person_options = [""]
    try:
        if not os.path.exists(EXCEL_FILE):
            return person_options, person_dict

        xls = pd.ExcelFile(EXCEL_FILE)
        if sheet_name not in xls.sheet_names:
            return person_options, person_dict

        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)

        for _, row in df.iterrows():
            if len(row) >= 2:
                full_name = clean_text_val(row.iloc[1])
                acad = clean_text_val(row.iloc[2]) if len(row) >= 3 else ""

                if (
                    full_name
                    and full_name not in ["ชื่อ - สกุล", "ลำดับที่"]
                    and full_name not in person_dict
                ):
                    person_options.append(full_name)
                    person_dict[full_name] = acad

    except Exception:
        pass

    return person_options, person_dict


# 🟢 4. คำนวณวันทำการ
def add_business_days(start_date, num_days):
    current_date = start_date
    added = 0
    while added < num_days:
        current_date += datetime.timedelta(days=1)
        if current_date.weekday() < 5:
            added += 1
    return current_date
