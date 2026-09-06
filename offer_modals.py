import pandas as pd
from offer_helpers import clean_text_val, load_shops_data, save_shops_data
import streamlit as st

dialog_decorator = getattr(st, "dialog", getattr(st, "experimental_dialog", None))

def add_shop_modal_content():
    new_shop_name = st.text_input(
        "🏪 ชื่อร้านค้า",
        placeholder="ตัวอย่าง: บริษัท มานิตวิทยา จำกัด",
    )
    new_address = st.text_area(
        "📍 ที่อยู่",
        value="",
        height=70,
        placeholder="ตัวอย่าง: 76,77,78-79 ซอยศรีสุริโยทัย 1 ต.ทะเลชุบศร อ.เมือง จ.ลพบุรี 15000",
    )
    new_phone = st.text_input(
        "📞 เบอร์โทรศัพท์",
        value="",
        placeholder="ตัวอย่าง: 081-7000767",
    )
    new_tax_id = st.text_input(
        "🏢 เลขประจำตัวผู้เสียภาษี",
        value="",
        placeholder="ตัวอย่าง: 0165533000070",
    )

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("บันทึกเพิ่มร้านค้า", type="primary", use_container_width=True):
            clean_name = clean_text_val(new_shop_name)
            clean_phone_val = clean_text_val(new_phone)
            clean_tax_val = clean_text_val(new_tax_id)

            if not clean_name:
                st.error("⚠️ กรุณากรอกชื่อร้านค้า")
            else:
                df_shops = load_shops_data()
                name_col = next(
                    (
                        col
                        for col in df_shops.columns
                        if str(col).lower()
                        in ["shop_name", "name", "shopname", "ร้านค้า", "ชื่อร้าน"]
                    ),
                    df_shops.columns[0],
                )

                existing_names = df_shops[name_col].astype(str).str.strip().tolist()
                if clean_name in existing_names:
                    st.error(f"⚠️ ร้านค้า '{clean_name}' มีอยู่ในระบบแล้ว!")
                else:
                    new_row = pd.DataFrame(
                        [
                            {
                                name_col: clean_name,
                                "address": new_address,
                                "phone": clean_phone_val,
                                "tax_id": clean_tax_val,
                            }
                        ]
                    )
                    df_shops = pd.concat([df_shops, new_row], ignore_index=True)

                    # บันทึกลง Google Sheets ผ่าน gspread
                    if save_shops_data(df_shops):
                        st.toast(f"เพิ่มร้านค้า '{clean_name}' เรียบร้อย!", icon="✅")
                        st.rerun()
                    else:
                        st.error("⚠️ ไม่สามารถบันทึกข้อมูลลง Google Sheets ได้")

    with col2:
        if st.button("ยกเลิกเพิ่มร้าน", use_container_width=True):
            st.rerun()

if dialog_decorator:
    add_shop_modal = dialog_decorator("➕ เพิ่มร้านค้าใหม่")(add_shop_modal_content)
else:
    add_shop_modal = add_shop_modal_content
