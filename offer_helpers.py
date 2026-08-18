import datetime
import pandas as pd
import streamlit as st

# 🔗 กำหนด Spreadsheet ID ของ Google Sheet
SPREADSHEET_ID = "1hPP2Jks_k9-hKSWzgnVyIJMRWVmo32Qa"

def get_sheet_url(sheet_name):
    """สร้าง URL สำหรับดึงข้อมูลเป็น CSV"""
    return f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name}"

def clean_text_val(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str

def load_shops_data():
    """อ่านข้อมูลร้านค้าผ่าน CSV Link"""
    try:
        url = get_sheet_url("Shops")
        df = pd.read_csv(url, dtype=str)
        df.columns = df.columns.str.strip()
        for col in ["shop_name", "address", "phone", "tax_id"]:
            if col not in df.columns:
                df[col] = ""
        df = df.dropna(subset=["shop_name"])
        df = df[df["shop_name"].str.strip() != ""]
        for col in ["shop_name", "address", "phone", "tax_id"]:
            df[col] = df[col].apply(clean_text_val)
        return df
    except Exception as e:
        st.error(f"Error loading shops: {e}")
        return pd.DataFrame(columns=["shop_name", "address", "phone", "tax_id"])

def save_shops_data(df):
    """กรณีอ่านผ่าน CSV จะบันทึกกลับไม่ได้โดยตรง"""
    st.warning("⚠️ การบันทึกข้อมูลต้องทำผ่านหน้า Google Sheets โดยตรงครับ")
    return False

def load_teacher_data(sheet_name="Teachers"):
    """อ่านข้อมูลครูผ่าน CSV Link"""
    person_dict = {}
    person_options = [""]
    try:
        url = get_sheet_url(sheet_name)
        df = pd.read_csv(url, dtype=str)
        for _, row in df.iterrows():
            r_list = row.tolist()
            if len(r_list) >= 3:
                fname = str(r_list[1]).strip() if pd.notna(r_list[1]) else ""
                lname = str(r_list[2]).strip() if pd.notna(r_list[2]) else ""
                acad = str(r_list[3]).strip() if len(r_list) >= 4 and pd.notna(r_list[3]) else ""
                if fname and lname and lname not in ["nan", "None"]:
                    full_name = f"{fname} {lname}".strip()
                    if full_name not in person_dict:
                        person_options.append(full_name)
                        person_dict[full_name] = acad
    except Exception as e:
        st.warning(f"Error loading teachers: {e}")
    return person_options, person_dict

def add_business_days(start_date, num_days):
    """ฟังก์ชันเพิ่มวันทำการ"""
    current_date = start_date
    added = 0
    while added < num_days:
        current_date += datetime.timedelta(days=1)
        if current_date.weekday() < 5:
            added += 1
    return current_date
