# ชื่อไฟล์: space_fourcolor.py
import io
import os
import pandas as pd
import streamlit as st

# นำเข้าฟังก์ชันจากไฟล์แยกทั้งสอง
from fourcolor import generate_fourcolor_excel
from space import generate_space_excel, get_only_teacher_names


@st.dialog("🎨 สร้างเอกสารสี่สี และ รายละเอียดคุณลักษณะ", width="large")
def render_space_fourcolor_dialog(
    default_receiver=None,
    default_receiver_sub=None,
    default_items=None,
    default_total_amount=0.0,
    project_name="",
    department="",
    budget_type="",
    default_parcel_no="",
    **kwargs,
):
    st.markdown(
        """
        <style>
            div[data-testid="stDialog"] div[data-baseweb="modal"] {
                max-width: 900px !important;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    receiver_options = get_only_teacher_names("teachers.xlsx")

    st.markdown("##### 👤 ข้อมูลผู้รับพัสดุ")
    idx_rec = (
        receiver_options.index(default_receiver)
        if default_receiver in receiver_options
        else 0
    )
    idx_sub = (
        receiver_options.index(default_receiver_sub)
        if default_receiver_sub in receiver_options
        else (1 if len(receiver_options) > 1 else 0)
    )

    col_rec1, col_rec2 = st.columns(2)
    with col_rec1:
        receiver = st.selectbox(
            "ผู้รับพัสดุ (หลัก)",
            options=receiver_options,
            index=idx_rec,
            key="sf_receiver",
        )
    with col_rec2:
        receiver_sub = st.selectbox(
            "ผู้รับพัสดุแทน",
            options=receiver_options,
            index=idx_sub,
            key="sf_receiver_sub",
        )

    st.markdown("---")
    st.write(
        "ตรวจสอบและแก้ไขรายการสินค้า จำนวน และราคา (ใช้สร้างทั้งเอกสารสี่สี และ รายละเอียดคุณลักษณะ พร้อมกัน)"
    )

    # 🟢 ฟังก์ชันแปลงข้อมูลให้เป็น DataFrame ที่ปลอดภัย 100%
    def ensure_dataframe(data):
        if isinstance(data, pd.DataFrame):
            return data.copy()
        if isinstance(data, list) and len(data) > 0:
            try:
                return pd.DataFrame(data)
            except Exception:
                pass
        elif isinstance(data, dict):
            try:
                return pd.DataFrame(data)
            except Exception:
                try:
                    return pd.DataFrame.from_dict(data, orient="index")
                except Exception:
                    pass
        return pd.DataFrame(
            [
                {
                    "name": "",
                    "quantity": 1.0,
                    "unit": "รายการ",
                    "price_per_unit": float(default_total_amount),
                }
            ]
        )

    # 🟢 ล็อกข้อมูลตั้งต้นไว้ครั้งเดียวตอนเปิดไดอะล็อก (ป้องกันตารางรีเซ็ตระหว่างพิมพ์)
    if "sf_data_editor_initialized" not in st.session_state:
        st.session_state.sf_data_editor_initialized = True
        st.session_state.sf_initial_df = ensure_dataframe(default_items)

    # 🟢 เรนเดอร์ data_editor โดยใช้ข้อมูลตั้งต้นนิ่งๆ และให้ key จัดการสถานะการพิมพ์อย่างต่อเนื่อง
    edited_df = st.data_editor(
        st.session_state.sf_initial_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="sf_data_editor",
        column_config={
            "name": st.column_config.Column("รายการ", width="large", required=True),
            "quantity": st.column_config.NumberColumn(
                "จำนวน", min_value=0.0, step=1.0, default=1.0, required=True
            ),
            "unit": st.column_config.TextColumn(
                "หน่วยนับ", default="รายการ", required=True
            ),
            "price_per_unit": st.column_config.NumberColumn(
                "ราคาต่อหน่วย (บาท)",
                min_value=0.0,
                step=1.0,
                format="%.2f",
                default=0.0,
                required=True,
            ),
        },
    )

    processed_df = ensure_dataframe(edited_df)
    processed_df["total_price"] = processed_df["quantity"].fillna(0) * processed_df[
        "price_per_unit"
    ].fillna(0)

    valid_items = processed_df[processed_df["name"].str.strip() != ""].copy()
    valid_items.reset_index(drop=True, inplace=True)

    total_amount = valid_items["total_price"].sum() if not valid_items.empty else 0.0

    st.markdown(f"##### 📊 สรุปรายการทั้งหมด ({len(valid_items)} รายการ)")

    preview_df = valid_items.copy()
    preview_df.insert(0, "ลำดับ", range(1, len(preview_df) + 1))
    preview_df = preview_df.rename(
        columns={
            "name": "รายการ",
            "quantity": "จำนวน",
            "unit": "หน่วยนับ",
            "price_per_unit": "ราคาต่อหน่วย",
            "total_price": "ราคาสินค้า",
        }
    )

    st.dataframe(
        preview_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "ราคาต่อหน่วย": st.column_config.NumberColumn(format="%.2f บาท"),
            "ราคาสินค้า": st.column_config.NumberColumn(format="%.2f บาท"),
        },
    )

    st.info(f"💰 **ราคารวมทั้งหมด:** {total_amount:,.2f} บาท")

    is_price_matched = abs(total_amount - default_total_amount) < 0.01

    if not is_price_matched and default_total_amount > 0:
        st.warning(
            f"⚠️ ราคารวมในตาราง ({total_amount:,.2f} บาท) **ยังไม่ตรงกับ** ราคารวมของร้านค้า ({default_total_amount:,.2f} บาท)"
        )

    cf_btn1, cf_btn2 = st.columns([1, 1])
    with cf_btn1:
        if st.button("❌ ปิดหน้าต่าง", use_container_width=True):
            for k in [
                "sf_data_editor",
                "sf_initial_df",
                "sf_data_editor_initialized",
            ]:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()

    with cf_btn2:
        space_template = (
            "space.xlsx"
            if os.path.exists("space.xlsx")
            else os.path.join("templates_4color", "space.xlsx")
        )
        fourcolor_template = (
            "fourcolor_template.xlsx"
            if os.path.exists("fourcolor_template.xlsx")
            else os.path.join("templates_4color", "fourcolor_template.xlsx")
        )

        if os.path.exists(space_template) and os.path.exists(fourcolor_template):
            # สร้างไฟล์ Excel ทั้งสองแบบเตรียมไว้ทันทีจากตารางเดียวกัน
            fourcolor_bytes = generate_fourcolor_excel(
                fourcolor_template,
                receiver,
                receiver_sub,
                valid_items,
                total_amount,
                project_name=project_name,
                department=department,
                budget_type=budget_type,
                percel_no=default_parcel_no,
            )
            space_bytes = generate_space_excel(
                space_template,
                receiver,
                receiver_sub,
                valid_items,
                total_amount,
                project_name=project_name,
                department=department,
                budget_type=budget_type,
                percel_no=default_parcel_no,
            )

            if default_total_amount > 0 and not is_price_matched:
                st.button(
                    "📥 ดาวน์โหลดเอกสาร",
                    disabled=True,
                    type="primary",
                    use_container_width=True,
                    help="ราคารวมในตารางต้องเท่ากับราคารวมของร้านค้าจึงจะดาวน์โหลดได้",
                )
            else:
                # แสดงปุ่มดาวน์โหลดทั้ง 2 ไฟล์เคียงข้างกัน กดแยกอันไหนก็ได้จากข้อมูลชุดเดียวกัน
                dl_col1, dl_col2 = st.columns(2)
                with dl_col1:
                    st.download_button(
                        label="เอกสารสี่สี",
                        data=fourcolor_bytes,
                        file_name="เอกสารคุณลักษณะสินค้า_สี่สี.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                with dl_col2:
                    st.download_button(
                        label="รายละเอียดคุณลักษณะ",
                        data=space_bytes,
                        file_name="เอกสาร Space.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
        else:
            st.error("⚠️ ไม่พบไฟล์ Template สำหรับ Fourcolor หรือ Space")
