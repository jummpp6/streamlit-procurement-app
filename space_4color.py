# ชื่อไฟล์: space_4color.py
import os
import pandas as pd
import streamlit as st

from fourcolor import generate_fourcolor_excel
from space import generate_space_excel, get_only_teacher_names


def render_space_4color_page():
    # 🟢 เพิ่มการตรวจสอบและกำหนดค่าเริ่มต้น (Initialization) ป้องกัน Error
    if "sf_initial_df" not in st.session_state:
        st.session_state.sf_initial_df = pd.DataFrame([
            {"name": "", "quantity": 1.0, "unit": "รายการ", "price_per_unit": 0.0}
        ])
    if "sf_data_editor_initialized" not in st.session_state:
        st.session_state.sf_data_editor_initialized = True
        
    if "sc_items_df" not in st.session_state:
        st.session_state.sc_items_df = pd.DataFrame([
            {"name": "", "quantity": 1.0, "unit": "รายการ", "price_per_unit": 0.0}
        ])

    if st.button("⬅️ กลับหน้าหลัก", use_container_width=False):
        st.session_state.page = "home"
        st.rerun()

    st.title("🎨 ระบบสร้างเอกสารสี่สี และ รายละเอียดคุณลักษณะ")
    st.write("กรอกข้อมูลโครงการ รายการสินค้า และผู้รับพัสดุเพื่อสร้างเอกสารสี่สีและ Space พร้อมกันได้ทันที")
    st.markdown("---")

    # โค้ดส่วนที่เหลือของฟังก์ชัน...

    # --- ส่วนที่ 1: ข้อมูลโครงการและพัสดุ ---
    st.subheader("📝 1. ข้อมูลโครงการและพัสดุ")
    col1, col2 = st.columns(2, gap="large")

    with col1:
        # ย้ายข้อมูลทั่วไป (ชื่อโครงการ, แผนก, เลขพัสดุ) มาไว้ Col 1 ทั้งหมด
        project_name = st.text_input(
            "ชื่อโครงการ",
            placeholder="ตัวอย่าง: จัดซื้อวัสดุฝึกปฏิบัติการช่างยนต์",
            key="sc_project_name",
        )
        department = st.text_input(
            "งาน หรือ แผนกวิชา",
            placeholder="ตัวอย่าง: แผนกวิชาช่างยนต์",
            key="sc_department",
        )
        parcel_no = st.text_input(
            "เลขพัสดุ",
            placeholder="ตัวอย่าง: 317-69",
            key="sc_parcel_no",
        )

    with col2:
        # ย้ายแหล่งเงินงบประมาณ (st.radio) และเงื่อนไขย่อยมาไว้ Col 2 ทั้งหมด
        budget_source_type = st.radio(
            "แหล่งเงินงบประมาณ",
            [
                "บกศ. (เงินรายได้สถานศึกษา)",
                "เงินอุดหนุน",
                "งปม. (เงินงบประมาณ)",
                "อื่นๆ (กำหนดเอง)",
            ],
            key="sc_budget_source",
        )

        budget_type = ""
        if budget_source_type == "บกศ. (เงินรายได้สถานศึกษา)":
            budget_type = "เงินรายได้ของสถานศึกษา (บกศ.)"
        elif budget_source_type == "เงินอุดหนุน":
            sub_detail = st.text_input(
                "ระบุรายละเอียดเงินอุดหนุน",
                value="กิจกรรมพัฒนาผู้เรียน",
                placeholder="ตัวอย่าง: ค่ากิจกรรมพัฒนาผู้เรียน / ค่าเรียนฟรี 15 ปี",
                key="sc_sub_grant",
            )
            budget_type = f"เงินอุดหนุน {sub_detail}"
        elif budget_source_type == "งปม. (เงินงบประมาณ)":
            sub_detail = st.text_input(
                "ระบุรายละเอียดเงินงบประมาณ",
                value="เงินอุดหนุน",
                placeholder="ตัวอย่าง: งบอุดหนุน / งบดำเนินงาน",
                key="sc_sub_budget",
            )
            budget_type = f"เงินงบประมาณ {sub_detail}"
        elif budget_source_type == "อื่นๆ (กำหนดเอง)":
            sub_detail = st.text_input(
                "ระบุแหล่งเงินอื่นๆ",
                placeholder="พิมพ์ระบุแหล่งเงิน...",
                key="sc_sub_other",
            )
            budget_type = sub_detail

    st.markdown("---")

    # --- ส่วนที่ 2: ข้อมูลผู้รับพัสดุ ---
    st.subheader("👤 2. ข้อมูลผู้รับพัสดุ")
    receiver_options = get_only_teacher_names("teachers.xlsx")
    if not receiver_options:
        receiver_options = ["-"]

    col_rec1, col_rec2 = st.columns(2, gap="large")
    with col_rec1:
        receiver = st.selectbox(
            "ผู้รับพัสดุ (หลัก)",
            options=receiver_options,
            key="sc_receiver",
        )
    with col_rec2:
        receiver_sub = st.selectbox(
            "ผู้รับพัสดุแทน",
            options=receiver_options,
            index=min(1, len(receiver_options) - 1),
            key="sc_receiver_sub",
        )

    st.markdown("---")

    # --- ส่วนที่ 3: รายการสินค้า ---
    st.subheader("📋 3. รายการสินค้า")
    st.write("กรอกรายการสินค้า จำนวน และราคาต่อหน่วยตามต้องการ")

    # กำหนดค่าเริ่มต้นตารางสินค้าว่างๆ สำหรับกรอกเอง
    if "sc_items_df" not in st.session_state:
        st.session_state.sc_items_df = pd.DataFrame(
            [{"name": "", "quantity": 1.0, "unit": "รายการ", "price_per_unit": 0.0}]
        )

    edited_df = st.data_editor(
        st.session_state.sf_initial_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="sf_data_editor",
        column_config={
            "name": st.column_config.Column("รายการ", width="large", required=True),
            "quantity": st.column_config.NumberColumn(
                "จำนวน",
                min_value=0.0,
                step=0.01,  # รองรับทศนิยม
                format="%g",
                default=1.0,
                required=True,
            ),
            "unit": st.column_config.TextColumn(
                "หน่วยนับ", default="รายการ", required=True
            ),
            "price_per_unit": st.column_config.NumberColumn(
                "ราคาต่อหน่วย (บาท)",
                min_value=0.0,
                step=0.01,  # รองรับทศนิยมสตางค์
                format="%.2f",
                default=0.0,
                required=True,
            ),
        },
    )

    processed_df = (
        pd.DataFrame(edited_df)
        if not isinstance(edited_df, pd.DataFrame)
        else edited_df.copy()
    )
    processed_df["total_price"] = processed_df["quantity"].fillna(0) * processed_df[
        "price_per_unit"
    ].fillna(0)

    valid_items = processed_df[processed_df["name"].str.strip() != ""].copy()
    valid_items.reset_index(drop=True, inplace=True)

    total_amount = valid_items["total_price"].sum() if not valid_items.empty else 0.0

    st.markdown(f"##### 📊 สรุปรายการทั้งหมด ({len(valid_items)} รายการ)")

    if not valid_items.empty:
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

    st.markdown("---")
    st.subheader("📥 4. ดาวน์โหลดเอกสาร")

    _template = (
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
        if valid_items.empty:
            st.warning("⚠️ กรุณากรอกรายการสินค้าอย่างน้อย 1 รายการเพื่อดาวน์โหลดเอกสาร")
        else:
            # สร้างไฟล์ Excel ทั้งสองแบบจากข้อมูลที่กรอกเข้ามาสดๆ
            fourcolor_bytes = generate_fourcolor_excel(
                fourcolor_template,
                receiver,
                receiver_sub,
                valid_items,
                total_amount,
                project_name=project_name,
                department=department,
                budget_type=budget_type,
                percel_no=parcel_no,
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
                percel_no=parcel_no,
            )

            dl_col1, dl_col2 = st.columns(2, gap="medium")
            with dl_col1:
                st.download_button(
                    label="📥 ดาวน์โหลดเอกสารสี่สี (.xlsx)",
                    data=fourcolor_bytes,
                    file_name=f"เอกสารคุณลักษณะสินค้า_สี่สี_{project_name if project_name else 'พัสดุ'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
            with dl_col2:
                st.download_button(
                    label="📥 ดาวน์โหลดเอกสารรายละเอียดคุณลักษณะ (.xlsx)",
                    data=space_bytes,
                    file_name=f"เอกสาร Space_{project_name if project_name else 'พัสดุ'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
    else:
        st.error("⚠️ ไม่พบไฟล์ Template สำหรับ Fourcolor หรือ Space ในระบบ")
