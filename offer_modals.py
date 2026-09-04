import pandas as pd
from offer_helpers import clean_text_val, load_shops_data, save_shops_data
import streamlit as st

dialog_decorator = getattr(st, "dialog", getattr(st, "experimental_dialog", None))


# 🟢 1. Modal สำหรับแก้ไขข้อมูลร้านค้า
def edit_address_modal_content(shop_name, c_address, c_phone, c_tax_id):
    st.write(f"**ร้านค้า:** {shop_name}")

    new_address = st.text_area(
        "📍 ที่อยู่",
        value=clean_text_val(c_address),
        height=70,
        placeholder="ตัวอย่าง: 76,77,78-79 ซอยศรีสุริโยทัย 1 ต.ทะเลชุบศร อ.เมือง จ.ลพบุรี 15000",
    )
    new_phone = st.text_input(
        "📞 เบอร์โทรศัพท์",
        value=clean_text_val(c_phone),
        placeholder="ตัวอย่าง: 0817000767",
    )
    new_tax_id = st.text_input(
        "🏢 เลขประจำตัวผู้เสียภาษี",
        value=clean_text_val(c_tax_id),
        placeholder="ตัวอย่าง: 0165533000070",
    )

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("บันทึกการแก้ไข", type="primary", use_container_width=True):
            clean_phone_val = clean_text_val(new_phone)
            clean_tax_val = clean_text_val(new_tax_id)

            df_shops = load_shops_data()
            name_col = next((col for col in df_shops.columns if str(col).lower() in ["shop_name", "name", "shopname", "ร้านค้า", "ชื่อร้าน"]), df_shops.columns[0])
            idx = df_shops[df_shops[name_col].astype(str).str.strip() == str(shop_name).strip()].index

            if not idx.empty:
                df_shops.loc[idx[0], "address" if "address" in df_shops.columns else df_shops.columns[1]] = new_address
                df_shops.loc[idx[0], "phone" if "phone" in df_shops.columns else df_shops.columns[2]] = clean_phone_val
                df_shops.loc[idx[0], "tax_id" if "tax_id" in df_shops.columns else df_shops.columns[3]] = clean_tax_val
            else:
                new_row = pd.DataFrame(
                    [
                        {
                            name_col: shop_name,
                            "address": new_address,
                            "phone": clean_phone_val,
                            "tax_id": clean_tax_val,
                        }
                    ]
                )
                df_shops = pd.concat([df_shops, new_row], ignore_index=True)

            if save_shops_data(df_shops):
                cache_key = f"disp_vendor_{shop_name}"
                if cache_key in st.session_state:
                    del st.session_state[cache_key]

                st.toast(f"บันทึกข้อมูลของ '{shop_name}' เรียบร้อย!", icon="✅")
                st.rerun()
            else:
                st.error("⚠️ ไม่สามารถบันทึกข้อมูลลง Google Sheets ได้ กรุณาตรวจสอบการเชื่อมต่อ")

    with col2:
        if st.button("ยกเลิก", use_container_width=True):
            st.rerun()


if dialog_decorator:
    edit_address_modal = dialog_decorator("✏️ แก้ไขข้อมูลร้านค้า")(
        edit_address_modal_content
    )
else:
    edit_address_modal = edit_address_modal_content


# 🟢 2. Modal สำหรับเพิ่มร้านค้าใหม่
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
                name_col = next((col for col in df_shops.columns if str(col).lower() in ["shop_name", "name", "shopname", "ร้านค้า", "ชื่อร้าน"]), df_shops.columns[0])

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
                    
                    if save_shops_data(df_shops):
                        st.toast(f"เพิ่มร้านค้า '{clean_name}' เรียบร้อย!", icon="✅")
                        st.rerun()
                    else:
                        st.error("⚠️ ไม่สามารถบันทึกข้อมูลลง Google Sheets ได้")

    with col2:
        if st.button("ยกเลิกเพิ่มร้าน", use_container_width=True):
            st.rerun()

if dialog_decorator:
    add_shop_modal = dialog_decorator("➕ เพิ่มร้านค้าใหม่")(
        add_shop_modal_content
    )
else:
    add_shop_modal = add_shop_modal_content
