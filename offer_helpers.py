import os
import pandas as pd
import streamlit as st


def clean_text_val(val):
  """ทำความสะอาดค่าข้อความ ป้องกันค่า NaN หรือ None"""
  if pd.isna(val) or str(val).lower() in ["nan", "none", ""]:
    return ""
  return str(val).strip()


def load_shops_data():
  """โหลดข้อมูลร้านค้าจาก Google Sheets (พร้อมระบบสำรองอ่านจากไฟล์ Excel/CSV)"""
  try:
    conn = st.connection("gsheets", type="gsheets")
    df = conn.read(ttl=0)  # โหลดข้อมูลล่าสุดแบบ Real-time
    if not df.empty:
      return df
  except Exception as e:
    print(f"Google Sheets load error: {e}")

  # กรณีเชื่อมต่อ Google Sheets ไม่สำเร็จ จะอ่านจากไฟล์สำรองในเครื่องแทน
  if os.path.exists("shops.xlsx"):
    return pd.read_excel("shops.xlsx")
  elif os.path.exists("shops.csv"):
    return pd.read_csv("shops.csv")
  elif os.path.exists("vendors.xlsx"):
    return pd.read_excel("vendors.xlsx")

  # หากไม่มีไฟล์ใดๆ เลย ให้สร้าง DataFrame เปล่าโครงสร้างเริ่มต้น
  return pd.DataFrame(columns=["shop_name", "address", "phone", "tax_id"])


def save_shops_data(df):
  """บันทึกข้อมูลร้านค้าลง Google Sheets และสำรองเก็บบันทึกลงไฟล์ภายในเครื่อง"""
  try:
    conn = st.connection("gsheets", type="gsheets")
    conn.update(data=df)

    # สำรองบันทึกไฟล์ Excel ไว้ในเครื่องเพื่อความปลอดภัย
    df.to_excel("shops.xlsx", index=False)
    df.to_excel("vendors.xlsx", index=False)
    return True
  except Exception as e:
    print(f"Google Sheets save error: {e}")
    try:
      df.to_excel("shops.xlsx", index=False)
      df.to_excel("vendors.xlsx", index=False)
      return True
    except Exception as local_err:
      print(f"Local save error: {local_err}")
      return False


# ==========================================
# 🔄 Alias Functions (ป้องกัน ImportError กรณีโค้ดเดิมเรียกชื่ออื่น)
# ==========================================
def load_vendors_data():
  """ฟังก์ชันสำรอง กรณีที่โค้ดเก่าเรียกใช้ load_vendors_data"""
  return load_shops_data()


def save_vendors_data(df):
  """ฟังก์ชันสำรอง กรณีที่โค้ดเก่าเรียกใช้ save_vendors_data"""
  return save_shops_data(df)
