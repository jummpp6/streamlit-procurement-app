import datetime
import os
import pandas as pd
import streamlit as st

EXCEL_SHOPS_FILE = "shops.xlsx"

try:
    import holidays

    thai_holidays = holidays.TH()
except ImportError:
    thai_holidays = []


def clean_text_val(val):
    """ฟังก์ชันช่วยลบ .0 ท้ายข้อความที่เกิดจากการอ่านเลขใน Excel"""
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str


def load_shops_data():
    """อ่านและจัดระเบียบข้อมูลร้านค้า"""
    if not os.path.exists(EXCEL_SHOPS_FILE):
        df_init = pd.DataFrame(
            [
                {
                    "shop_name": "บริษัท มานิตวิทยา จำกัด",
                    "address": "76,77,78-79 ซอยศรีสุริโยทัย 1 ต.ทะเลชุบศร อ.เมือง จ.ลพบุรี 15000",
                    "phone": "0817000767",
                    "tax_id": "0165533000070",
                }
            ]
        )
        df_init.to_excel(EXCEL_SHOPS_FILE, index=False)
        return df_init

    # อ่านแบบ dtype=str เพื่อป้องกันไม่ให้ pandas แปลงเลขเป็น float ตั้งแต่แรก
    df = pd.read_excel(EXCEL_SHOPS_FILE, dtype=str)

    if "ชื่อร้านค้า" in df.columns or "shop_name" not in df.columns:
        col_name = "ชื่อร้านค้า" if "ชื่อร้านค้า" in df.columns else df.columns[0]
        col_addr = "ที่อยู่" if "ที่อยู่" in df.columns else df.columns[1]

        df["shop_clean"] = df[col_name].ffill()
        shops = []
        for shop, group in df.groupby("shop_clean", sort=False):
            if pd.isna(shop) or clean_text_val(shop) in ["nan", "None", ""]:
                continue

            address, phone, tax_id = "", "", ""
            for val in group[col_addr].dropna():
                val_str = clean_text_val(val)
                if val_str.startswith("เบอร์โทร"):
                    phone = clean_text_val(val_str.replace("เบอร์โทร", "", 1))
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

        new_df = pd.DataFrame(shops)
        try:
            new_df.to_excel(EXCEL_SHOPS_FILE, index=False)
        except Exception:
            pass
        return new_df

    for col in ["shop_name", "address", "phone", "tax_id"]:
        if col not in df.columns:
            df[col] = ""

    df = df.dropna(subset=["shop_name"])
    df["shop_name"] = df["shop_name"].apply(clean_text_val)

    for col in ["address", "phone", "tax_id"]:
        df[col] = df[col].apply(clean_text_val)

    return df


def save_shops_data(df):
    """บันทึกข้อมูลร้านค้าลงไฟล์ Excel"""
    try:
        df.to_excel(EXCEL_SHOPS_FILE, index=False)
        return True
    except PermissionError:
        st.error(
            "⚠️ ไม่สามารถบันทึกได้ เนื่องจากไฟล์ 'shops.xlsx' กำลังถูกเปิดอยู่ใน Excel!"
        )
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


def load_teacher_data(file_path="teachers.xlsx"):
    """อ่านข้อมูลรายชื่อครูจากไฟล์ Excel"""
    person_dict = {}
    person_options = [""]

    if os.path.exists(file_path):
        try:
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                if sheet in ["สรุป", "สรุปคศ.", "Sheet1"]:
                    continue

                df = pd.read_excel(file_path, sheet_name=sheet)
                for _, row in df.iterrows():
                    r_list = row.tolist()
                    if len(r_list) >= 3:
                        fname = str(r_list[1]).strip() if pd.notna(r_list[1]) else ""
                        lname = str(r_list[2]).strip() if pd.notna(r_list[2]) else ""
                        acad = (
                            str(r_list[3]).strip()
                            if len(r_list) >= 4 and pd.notna(r_list[3])
                            else ""
                        )

                        if fname and lname and lname not in ["nan", "None"]:
                            if not any(
                                fname.startswith(p) for p in ["ข้อมูล", "ลำดับ", "ชื่อ"]
                            ):
                                full_name = f"{fname} {lname}".strip()
                                if acad in ["nan", "None"]:
                                    acad = ""
                                if full_name not in person_dict:
                                    person_options.append(full_name)
                                    person_dict[full_name] = acad
        except Exception as e:
            st.warning(f"⚠️ ไม่สามารถอ่านไฟล์ {file_path} ได้: {e}")

    return person_options, person_dict
