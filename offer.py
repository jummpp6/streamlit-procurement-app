import datetime
import io
import os
import re
import zipfile

import pandas as pd
from pythainlp.util import bahttext
import streamlit as st

from doc_processor import (
    format_budget_money,
    format_thai_date,
    process_docx,
    to_thai_num,
)
from offer_components import render_person_inputs
from offer_helpers import (
    add_business_days,
    load_shops_data,
    load_teacher_data,
)
from offer_modals import add_shop_modal, dialog_decorator, edit_address_modal


def reset_purchase_form():
    """ฟังก์ชันสำหรับล้างข้อมูลในฟอร์มจัดซื้อ (เรียกใช้ผ่าน on_click)"""
    keys_to_clear = [
        "purchase_parcel_no",  # 👈 ล้างเลขพัสดุ
        "purchase_project_name",
        "purchase_budget",
        "purchase_item_count",
        "purchase_dept",
        "purchase_doc_no",
        "purchase_doc_save",
        "purchase_budget_mid",
        "purchase_sub_grant",
        "purchase_sub_budget",
        "purchase_submit_no",
    ]
    for key in keys_to_clear:
        if key in st.session_state:
            st.session_state[key] = ""

    st.toast("ล้างข้อมูลเรียบร้อยแล้ว!", icon="🧹")


def render_purchase_page():
    st.markdown(
        """
        <style>
            @import url('https://fonts.googleapis.com/css2?family=TH+Sarabun+New:wght@400;600;700&display=swap');
            @import url('https://fonts.cdnfonts.com/css/th-sarabunpsk');

            header[data-testid="stHeader"] { display: none !important; }
            .main .block-container { padding-top: 1rem !important; padding-bottom: 2.5rem !important; max-width: 1100px !important; }
            html, body, [class*="css"], .stMarkdown, .stText, p, label, input, select, textarea, button, span, div { font-family: 'TH SarabunPSK', 'TH Sarabun New', 'Sarabun', sans-serif !important; font-size: 103.5% !important; }
            div[data-testid="stHorizontalBlock"] { align-items: flex-start !important; gap: 1.5rem !important; }
            .stTextInput label, .stSelectbox label, .stDateInput label, .stRadio label { font-weight: 600 !important; font-size: 105% !important; margin-bottom: 4px !important; color: #334155 !important; }
            div[data-baseweb="select"] *, div[data-baseweb="input"] input { font-family: 'TH SarabunPSK', 'TH Sarabun New', 'Sarabun', sans-serif !important; font-size: 105% !important; }
            div[data-testid="stDateInput"] input { font-family: 'TH SarabunPSK', 'TH Sarabun New', 'Sarabun', sans-serif !important; font-size: 105% !important; line-height: 44px !important; height: 44px !important; padding-top: 0px !important; padding-bottom: 0px !important; }
            div[data-baseweb="input"] > div, div[data-baseweb="select"] > div { min-height: 44px !important; height: 44px !important; align-items: center !important; border-radius: 6px !important; background-color: #F8FAFC !important; border: 1px solid #CBD5E1 !important; }
            h1 { font-size: 3.5rem !important; font-weight: 700 !important; margin-top: 0rem !important; margin-bottom: 0.5rem !important; color: #0F172A !important; }
            h3, .stSubheader { font-size: 2.6rem !important; font-weight: 700 !important; color: #0F172A !important; margin-top: 0.8rem !important; margin-bottom: 0.6rem !important; border-bottom: 2px solid #CBD5E1; padding-bottom: 4px; }
            h5 { font-size: 1.8rem !important; font-weight: 600 !important; color: #1E293B !important; margin-top: 0.2rem !important; margin-bottom: 0.5rem !important; background-color: #F1F5F9; padding: 4px 10px; border-left: 4px solid #2563EB; border-radius: 0 4px 4px 0; }
            .stButton > button { padding: 0.5rem 1.5rem !important; font-size: 100% !important; border-radius: 8px !important; font-weight: 600 !important; }
            .stElementContainer { margin-bottom: 0.4rem !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("🛒 ระบบสร้างเอกสารจัดซื้อ")

    # --- แถวบนสุด: ปุ่มกลับหน้าหลัก + ช่องกรอกเลขพัสดุ ---
    col_top1, col_top2 = st.columns([1, 2], gap="medium")
    with col_top1:
        if st.button("⬅️ กลับหน้าหลัก", use_container_width=True):
            st.session_state.page = "home"
            st.rerun()

    with col_top2:
        parcel_no = st.text_input(
            "",
            placeholder="ตัวอย่าง: เลขพัสดุ (เช่น 317-69 หรือ 111-69)",
            key="purchase_parcel_no",
            # ซ่อน Label เพื่อให้วางเรียงสวยงาม
        )

    TEMPLATE_DIR = "templates"
    if not os.path.exists(TEMPLATE_DIR):
        os.makedirs(TEMPLATE_DIR)

    template_files = [
        f
        for f in os.listdir(TEMPLATE_DIR)
        if f.endswith(".docx") and not f.startswith("~$")
    ]

    # --- ส่วนที่ 1: ข้อมูลการจัดซื้อหลัก ---
    st.subheader("📝 1. ข้อมูลการจัดซื้อ (หน้า ส.1)")
    col1, col2 = st.columns([1, 1], gap="large")

    with col1:
        project_name = st.text_input(
            "ชื่อโครงการ",
            placeholder="ตัวอย่าง: จัดซื้อวัสดุฝึกปฏิบัติการช่างยนต์",
            key="purchase_project_name",
        )
        budget = st.text_input(
            "จำนวนเงิน / วงเงิน (พิมพ์เฉพาะตัวเลข)",
            placeholder="ตัวอย่าง: 4996",
            key="purchase_budget",
        )
        clean_num = 0.0
        try:
            clean_num = float(
                budget.replace(",", "").replace(" ", "").replace("บาท", "")
            )
            budget_text = bahttext(clean_num)
        except ValueError:
            budget_text = budget

        st.info(f"💡 **แปลงเป็นตัวหนังสืออัตโนมัติ:** {budget_text}")

        item_count_raw = st.text_input(
            "จำนวนรายการ (เช่น 5 รายการ)",
            placeholder="ตัวอย่าง: 5 รายการ",
            key="purchase_item_count",
        )

        budget_source_type = st.radio(
            "แหล่งเงินงบประมาณ",
            [
                "บกศ. (เงินรายได้สถานศึกษา)",
                "เงินอุดหนุน",
                "งปม. (เงินงบประมาณ)",
            ],
            key="purchase_budget_source",
        )

        budget_type_text = ""
        if budget_source_type == "บกศ. (เงินรายได้สถานศึกษา)":
            budget_type_text = "เงินรายได้ของสถานศึกษา (บกศ.)"
        elif budget_source_type == "เงินอุดหนุน":
            sub_detail = st.text_input(
                "ระบุรายละเอียดเงินอุดหนุน",
                value="กิจกรรมพัฒนาผู้เรียน",
                placeholder="ตัวอย่าง: ค่ากิจกรรมพัฒนาผู้เรียน / ค่าเรียนฟรี 15 ปี",
                key="purchase_sub_grant",
            )
            budget_type_text = f"เงินอุดหนุน {sub_detail}"
        elif budget_source_type == "งปม. (เงินงบประมาณ)":
            sub_detail = st.text_input(
                "ระบุรายละเอียดเงินงบประมาณ",
                value="เงินอุดหนุน",
                placeholder="ตัวอย่าง: งบอุดหนุน / งบดำเนินงาน",
                key="purchase_sub_budget",
            )
            budget_type_text = f"เงินงบประมาณ {sub_detail}"

    with col2:
        department = st.text_input(
            "งาน หรือ แผนกวิชาที่จัดทำ",
            placeholder="ตัวอย่าง: แผนกวิชาช่างยนต์ / งานกิจกรรมนักเรียนนักศึกษา",
            key="purchase_dept",
        )
        selected_date = st.date_input(
            "วันที่เอกสาร (หน้า ส.1)",
            datetime.date.today(),
            key="purchase_doc_date",
        )
        doc_no_raw = st.text_input(
            "เลขที่คำสั่ง",
            placeholder="ตัวอย่าง: 960/2569",
            key="purchase_doc_no",
        )
        doc_no_save = st.text_input(
            "เลขที่บันทึก",
            placeholder="ตัวอย่าง: 960",
            key="purchase_doc_save",
        )

    use_thai_num = True
    use_thai_num2 = True

    # --- ส่วนที่ 2: คณะกรรมการ ---
    st.write("")
    st.subheader("👥 2. คำสั่งคณะกรรมการจัดซื้อ / ตรวจรับ")
    EXCEL_PATH = "teachers.xlsx"
    person_options, person_dict = load_teacher_data(EXCEL_PATH)

    col_count1, col_count2 = st.columns([1, 1], gap="large")
    with col_count1:
        buy_count = st.selectbox(
            "จำนวนกรรมการจัดซื้อ",
            options=[1, 2, 3],
            index=2,
            key="purchase_buy_count_select",
        )
    with col_count2:
        check_count = st.selectbox(
            "จำนวนกรรมการตรวจรับ",
            options=[1, 2, 3],
            index=2,
            key="purchase_check_count_select",
        )

    # 1. ฝั่งคณะกรรมการจัดซื้อ
    st.markdown("##### 🛒 คณะกรรมการจัดซื้อ")
    with st.container(border=True):
        buy_name1, buy_acad1, buy_pos1 = render_person_inputs(
            "จัดซื้อ 1",
            "purchase_buy1",
            "นายสมชาย ใจดี",
            "",
            "ประธานกรรมการฯ",
            person_options,
            person_dict,
        )
        if buy_count >= 2:
            buy_name2, buy_acad2, buy_pos2 = render_person_inputs(
                "จัดซื้อ 2",
                "purchase_buy2",
                "นางสาวสมหญิง ใจงาม",
                "",
                "กรรมการฯ",
                person_options,
                person_dict,
            )
        else:
            buy_name2, buy_acad2, buy_pos2 = "", "", ""

        if buy_count >= 3:
            buy_name3, buy_acad3, buy_pos3 = render_person_inputs(
                "จัดซื้อ 3",
                "purchase_buy3",
                "นายสมปอง ใจกล้า",
                "",
                "กรรมการฯ",
                person_options,
                person_dict,
            )
        else:
            buy_name3, buy_acad3, buy_pos3 = "", "", ""

    # 2. ฝั่งคณะกรรมการตรวจรับ
    st.markdown("##### 🔍 คณะกรรมการตรวจรับพัสดุ")
    with st.container(border=True):
        check_pos_options = (
            ["ผู้ตรวจรับพัสดุ"]
            if check_count == 1
            else ["ประธานกรรมการฯ", "กรรมการฯ", "กรรมการและเลขานุการฯ"]
        )

        default_check_pos1 = "ผู้ตรวจรับพัสดุ" if check_count == 1 else "ประธานกรรมการฯ"

        check_name1, check_acad1, check_pos1 = render_person_inputs(
            "ตรวจรับ 1",
            "purchase_check1",
            "นายสมชาย ใจดี",
            "",
            default_check_pos1,
            person_options,
            person_dict,
            pos_options_custom=check_pos_options,
        )
        if check_count >= 2:
            check_name2, check_acad2, check_pos2 = render_person_inputs(
                "ตรวจรับ 2",
                "purchase_check2",
                "นางสาวสมหญิง ใจงาม",
                "",
                "กรรมการฯ",
                person_options,
                person_dict,
            )
        else:
            check_name2, check_acad2, check_pos2 = "", "", ""

        if check_count >= 3:
            check_name3, check_acad3, check_pos3 = render_person_inputs(
                "ตรวจรับ 3",
                "purchase_check3",
                "นายสมปอง ใจกล้า",
                "",
                "กรรมการฯ",
                person_options,
                person_dict,
            )
        else:
            check_name3, check_acad3, check_pos3 = "", "", ""

    # --- ส่วนที่ 3: บันทึกรายงานผลการพิจารณา ---
    st.write("")
    st.subheader("📋 3. บันทึกรายงานผลการพิจารณา")

    col_report1, col_report2 = st.columns([1, 1], gap="large")

    with col_report1:
        selected_date2 = st.date_input(
            "วันที่รายงานผล",
            datetime.date.today(),
            key="purchase_date_report",
        )
        budgetmid = st.text_input(
            "วงเงินราคากลาง (พิมพ์เฉพาะตัวเลข)",
            placeholder="ตัวอย่าง: 4996",
            key="purchase_budget_mid",
        )

        clean_num_mid = 0.0
        try:
            clean_num_mid = float(
                budgetmid.replace(",", "").replace(" ", "").replace("บาท", "")
            )
            budget_text_mid = bahttext(clean_num_mid)
        except ValueError:
            budget_text_mid = budgetmid

        st.info(f"💡 **แปลงเป็นตัวหนังสืออัตโนมัติ:** {budget_text_mid}")

        if clean_num > 0 and clean_num_mid > clean_num:
            st.error("⚠️ **คำเตือน:** วงเงินราคากลางสูงกว่าวงเงินงบประมาณ!")

    with col_report2:
        df_shops = load_shops_data()
        shop_list = [
            s
            for s in df_shops["shop_name"].tolist()
            if s and str(s).strip() not in ["nan", "None", ""]
        ]

        vendor_name = st.selectbox(
            "ชื่อบริษัท / ร้านค้า",
            options=shop_list,
            index=0 if shop_list else None,
            key="purchase_selected_vendor_name",
        )

        c_address, c_phone, c_tax_id = "", "", ""
        display_text = "ไม่มีข้อมูล"

        if vendor_name:
            match = df_shops[df_shops["shop_name"] == vendor_name]
            if not match.empty:
                row = match.iloc[0]
                c_address = str(row["address"]) if pd.notna(row["address"]) else ""
                c_phone = str(row["phone"]) if pd.notna(row["phone"]) else ""
                c_tax_id = str(row["tax_id"]) if pd.notna(row["tax_id"]) else ""

                lines = []
                if c_address:
                    lines.append(c_address)
                if c_phone:
                    lines.append(f"เบอร์โทรศัพท์: {c_phone}")
                if c_tax_id:
                    lines.append(f"เลขประจำตัวผู้เสียภาษี: {c_tax_id}")

                if lines:
                    display_text = "\n".join(lines)

        st.markdown("**รายละเอียดร้านค้า**")
        st.text_area(
            "รายละเอียดร้านค้า",
            value=display_text,
            disabled=True,
            height=130,
            label_visibility="collapsed",
            key="purchase_vendor_display",
        )

        col_btn_add, col_btn_edit = st.columns([1, 1], gap="small")
        with col_btn_add:
            if st.button(
                "➕ เพิ่มร้านค้า",
                use_container_width=True,
                key="purchase_btn_add_shop",
            ):
                if dialog_decorator:
                    add_shop_modal()
        with col_btn_edit:
            if st.button(
                "✏️ แก้ไขข้อมูลร้าน",
                use_container_width=True,
                key="purchase_btn_edit_shop",
            ):
                if dialog_decorator:
                    edit_address_modal(vendor_name, c_address, c_phone, c_tax_id)

    # --- ส่วนที่ 4: ใบสั่งซื้อ ---
    st.write("")
    st.write("")
    st.subheader("📄 4. ใบสั่งซื้อ")
    col1, col2 = st.columns([1, 1], gap="large")
    with col1:
        submit_no = st.text_input(
            "เลขที่ข้อตกลง",
            placeholder="ตัวอย่าง: 960/2569",
            key="purchase_submit_no",
        )

    with col2:
        default_order_date = add_business_days(selected_date2, 5)

        selected_date5 = st.date_input(
            "วันที่ใบสั่งซื้อ",
            value=default_order_date,
            key="purchase_date_order_input",
        )

    # --- ส่วนที่ 5: ใบตรวจรับการจัดซื้อ ---
    st.write("")
    st.subheader("✅ 5. ใบตรวจรับการจัดซื้อ")
    selected_date4 = st.date_input(
        "วันที่ตรวจรับ",
        datetime.date.today(),
        key="purchase_check_date",
    )

    # --- ประมวลผลแปลงตัวเลข/วันที่ ---
    formatted_budget = format_budget_money(budget, use_thai=use_thai_num)
    formatted_budget_mid = format_budget_money(budgetmid, use_thai=use_thai_num2)

    formatted_date = format_thai_date(selected_date, use_thai=use_thai_num)
    formatted_date2 = format_thai_date(selected_date2, use_thai=use_thai_num)

    calc_date3 = add_business_days(selected_date5, 5)
    formatted_date3 = format_thai_date(calc_date3, use_thai=use_thai_num)

    formatted_date4 = format_thai_date(selected_date4, use_thai=use_thai_num)
    formatted_date5 = format_thai_date(selected_date5, use_thai=use_thai_num)

    formatted_doc_no = to_thai_num(doc_no_raw) if use_thai_num else str(doc_no_raw)
    formatted_doc_no_save = (
        to_thai_num(doc_no_save) if use_thai_num else str(doc_no_save)
    )
    formatted_submit_no = to_thai_num(submit_no) if use_thai_num else str(submit_no)
    formatted_item_count = (
        to_thai_num(item_count_raw) if use_thai_num else str(item_count_raw)
    )

    budget_with_text = f"{formatted_budget} บาท ({budget_text})"
    budget_with_text_mid = f"{formatted_budget_mid} บาท ({budget_text_mid})"

    replacements_data = {
        "{{PROJECT_NAME}}": project_name,
        "{{BUDGET}}": formatted_budget,
        "{{BUDGET_MID}}": formatted_budget_mid,
        "{{BUDGET_TEXT}}": budget_text,
        "{{BUDGET_TEXT_MID}}": budget_text_mid,
        "{{BUDGET_WITH_TEXT}}": budget_with_text,
        "{{BUDGET_WITH_TEXT_MID}}": budget_with_text_mid,
        "{{BUDGET_TYPE}}": budget_type_text,
        "{{ITEM_COUNT}}": formatted_item_count,
        "{{DEPARTMENT}}": department,
        "{{VENDOR_NAME}}": vendor_name,
        "{{VENDOR_ADDRESS}}": to_thai_num(c_address),
        "{{VENDOR_PHONE}}": to_thai_num(c_phone),
        "{{VENDOR_TAX_ID}}": to_thai_num(c_tax_id),
        "{{DOC_DATE}}": formatted_date,
        "{{DOC_DATE2}}": formatted_date2,
        "{{DOC_DATE3}}": formatted_date3,
        "{{DOC_DATE4}}": formatted_date4,
        "{{DOC_DATE5}}": formatted_date5,
        "{{DOC_NO}}": formatted_doc_no,
        "{{DOC_NO_SAVE}}": formatted_doc_no_save,
        "{{SUBMIT_NO}}": formatted_submit_no,
        "{{DIRECTOR_INSPECTOR}}": "คณะกรรมการ" if check_count > 1 else "ผู้",
        "{{DIRECTOR_NAME_BUY}}": f"๑. {buy_name1}" if buy_name1 else "",
        "{{DIRECTOR_ACAD_BUY}}": buy_acad1 if buy_name1 else "",
        "{{DIRECTOR_POS_BUY}}": buy_pos1 if buy_name1 else "",
        "{{DIRECTOR_NAME_BUY2}}": (
            f"๒. {buy_name2}" if buy_name2 and buy_count >= 2 else ""
        ),
        "{{DIRECTOR_ACAD_BUY2}}": buy_acad2 if buy_name2 and buy_count >= 2 else "",
        "{{DIRECTOR_POS_BUY2}}": buy_pos2 if buy_name2 and buy_count >= 2 else "",
        "{{DIRECTOR_NAME_BUY3}}": (
            f"๓. {buy_name3}" if buy_name3 and buy_count >= 3 else ""
        ),
        "{{DIRECTOR_ACAD_BUY3}}": buy_acad3 if buy_name3 and buy_count >= 3 else "",
        "{{DIRECTOR_POS_BUY3}}": buy_pos3 if buy_name3 and buy_count >= 3 else "",
        "{{DIRECTOR_NAME_BUY_PLAIN}}": buy_name1 if buy_name1 else "",
        "{{DIRECTOR_NAME_BUY2_PLAIN}}": (
            buy_name2 if buy_name2 and buy_count >= 2 else ""
        ),
        "{{DIRECTOR_NAME_BUY3_PLAIN}}": (
            buy_name3 if buy_name3 and buy_count >= 3 else ""
        ),
        "{{SIGN_LINE_BUY1}}": (
            "ลงชื่อ....................................ประธานกรรมการฯ"
            if buy_name1
            else ""
        ),
        "{{SIGN_NAME_BUY1}}": f"({buy_name1})" if buy_name1 else "",
        "{{SIGN_LINE_BUY2}}": (
            "ลงชื่อ....................................กรรมการฯ"
            if buy_name2 and buy_count >= 2
            else ""
        ),
        "{{SIGN_NAME_BUY2}}": (
            f"({buy_name2})" if buy_name2 and buy_count >= 2 else ""
        ),
        "{{SIGN_LINE_BUY3}}": (
            "ลงชื่อ....................................กรรมการฯ"
            if buy_name3 and buy_count >= 3
            else ""
        ),
        "{{SIGN_NAME_BUY3}}": (
            f"({buy_name3})" if buy_name3 and buy_count >= 3 else ""
        ),
        "{{CHECKITEM_NAME}}": f"๑. {check_name1}" if check_name1 else "",
        "{{CHECKITEM_ACAD}}": check_acad1 if check_name1 else "",
        "{{CHECKITEM_POS}}": (
            ("ผู้ตรวจรับพัสดุ" if check_count == 1 else check_pos1)
            if check_name1
            else ""
        ),
        "{{CHECKITEM_NAME2}}": (
            f"๒. {check_name2}" if check_name2 and check_count >= 2 else ""
        ),
        "{{CHECKITEM_ACAD2}}": (
            check_acad2 if check_name2 and check_count >= 2 else ""
        ),
        "{{CHECKITEM_POS2}}": check_pos2 if check_name2 and check_count >= 2 else "",
        "{{CHECKITEM_NAME3}}": (
            f"๓. {check_name3}" if check_name3 and check_count >= 3 else ""
        ),
        "{{CHECKITEM_ACAD3}}": (
            check_acad3 if check_name3 and check_count >= 3 else ""
        ),
        "{{CHECKITEM_POS3}}": check_pos3 if check_name3 and check_count >= 3 else "",
        "{{CHECKITEM_NAME_PLAIN}}": check_name1 if check_name1 else "",
        "{{CHECKITEM_NAME2_PLAIN}}": (
            check_name2 if check_name2 and check_count >= 2 else ""
        ),
        "{{CHECKITEM_NAME3_PLAIN}}": (
            check_name3 if check_name3 and check_count >= 3 else ""
        ),
        "{{SIGN_LINE_CHECK1}}": (
            "ลงชื่อ....................................ผู้ตรวจรับพัสดุ"
            if (check_name1 and check_count == 1)
            else (
                "ลงชื่อ....................................ประธานกรรมการฯ"
                if check_name1
                else ""
            )
        ),
        "{{SIGN_NAME_CHECK1}}": f"({check_name1})" if check_name1 else "",
        "{{SIGN_LINE_CHECK2}}": (
            "ลงชื่อ....................................กรรมการฯ"
            if check_name2 and check_count >= 2
            else ""
        ),
        "{{SIGN_NAME_CHECK2}}": (
            f"({check_name2})" if check_name2 and check_count >= 2 else ""
        ),
        "{{SIGN_LINE_CHECK3}}": (
            "ลงชื่อ....................................กรรมการฯ"
            if check_name3 and check_count >= 3
            else ""
        ),
        "{{SIGN_NAME_CHECK3}}": (
            f"({check_name3})" if check_name3 and check_count >= 3 else ""
        ),
    }

    # --- ส่วนที่ 6: ตรวจสอบข้อมูลสรุป (Preview), Validation & Actions ---
    st.write("")
    st.subheader("👁️ 6. ตรวจสอบข้อมูลสรุป (Preview)")

    with st.container(border=True):
        st.markdown("##### 📄 สรุปรายละเอียดเอกสารจัดซื้อ")

        col_prev1, col_prev2 = st.columns([1, 1], gap="medium")
        with col_prev1:
            st.markdown(
                f"**โครงการ:** {project_name if project_name else '⚠️ *(ยังไม่ได้กรอก)*'}"
            )
            st.markdown(
                f"**หน่วยงาน/แผนก:** {department if department else '⚠️ *(ยังไม่ได้กรอก)*'}"
            )
            st.markdown(
                f"**วงเงินงบประมาณ:** {budget_with_text if budget else '⚠️ *(ยังไม่ได้กรอก)*'}"
            )

            if clean_num > 0 and clean_num_mid > clean_num:
                st.markdown(
                    f"**ราคากลาง:** {budget_with_text_mid} 🚨 **(สูงกว่าวงเงินงบประมาณ)**"
                )
            else:
                st.markdown(
                    f"**ราคากลาง:** {budget_with_text_mid if budgetmid else '⚠️ *(ยังไม่ได้กรอก)*'}"
                )

            st.markdown(f"**แหล่งเงิน:** {budget_type_text}")
            st.markdown(
                f"**จำนวนรายการ:** {formatted_item_count if item_count_raw else '⚠️ *(ยังไม่ได้กรอก)*'}"
            )

        with col_prev2:
            st.markdown(
                f"**เลขพัสดุที่ใช้ระบุชื่อไฟล์:** `{parcel_no if parcel_no else 'ไม่ได้ระบุ (ใช้ตามชื่อต้นแบบ)'}`"
            )
            st.markdown(
                f"**เลขที่คำสั่ง / ข้อตกลง:** {formatted_doc_no if doc_no_raw else '-'} / {formatted_submit_no if submit_no else '-'}"
            )
            st.markdown(
                f"**ร้านค้า / ผู้เสนอราคา:** {vendor_name if vendor_name else '⚠️ *(ยังไม่ได้เลือก)*'}"
            )
            st.markdown(f"**วันที่เอกสาร (ส.1):** {formatted_date}")
            st.markdown(f"**วันที่รายงานผล:** {formatted_date2}")
            st.markdown(f"**วันที่ใบสั่งซื้อ:** {formatted_date5}")
            st.markdown(f"**วันที่ตรวจรับ:** {formatted_date4}")

        st.markdown("---")
        col_prev_buy, col_prev_check = st.columns([1, 1], gap="medium")

        with col_prev_buy:
            st.markdown("**🛒 รายชื่อคณะกรรมการจัดซื้อ:**")
            buy_list_preview = []
            if buy_name1:
                buy_list_preview.append(f"1. {buy_name1} ({buy_pos1})")
            if buy_name2 and buy_count >= 2:
                buy_list_preview.append(f"2. {buy_name2} ({buy_pos2})")
            if buy_name3 and buy_count >= 3:
                buy_list_preview.append(f"3. {buy_name3} ({buy_pos3})")

            if buy_list_preview:
                for item in buy_list_preview:
                    st.text(f"  • {item}")
            else:
                st.caption("⚠️ ยังไม่มีการเลือกกรรมการจัดซื้อ")

        with col_prev_check:
            st.markdown("**🔍 รายชื่อคณะกรรมการตรวจรับ:**")
            check_list_preview = []
            if check_name1:
                check_list_preview.append(
                    f"1. {check_name1} ({'ผู้ตรวจรับพัสดุ' if check_count == 1 else check_pos1})"
                )
            if check_name2 and check_count >= 2:
                check_list_preview.append(f"2. {check_name2} ({check_pos2})")
            if check_name3 and check_count >= 3:
                check_list_preview.append(f"3. {check_name3} ({check_pos3})")

            if check_list_preview:
                for item in check_list_preview:
                    st.text(f"  • {item}")
            else:
                st.caption("⚠️ ยังไม่มีการเลือกกรรมการตรวจรับ")

    st.write("")
    col_action1, col_action2 = st.columns([1, 4], gap="small")

    with col_action1:
        st.button(
            "🗑️ ล้างข้อมูล",
            use_container_width=True,
            help="ล้างข้อมูลการกรอกทั้งหมด",
            key="purchase_reset_btn",
            on_click=reset_purchase_form,
        )

    with col_action2:
        btn_generate = st.button(
            "🚀 สร้างเอกสารจัดซื้อทั้งหมด",
            type="primary",
            use_container_width=True,
            key="purchase_submit_btn",
        )

    if btn_generate:
        missing_fields = []
        if not project_name.strip():
            missing_fields.append("ชื่อโครงการ")
        if not budget.strip():
            missing_fields.append("จำนวนเงิน / วงเงิน")
        if not department.strip():
            missing_fields.append("งาน หรือ แผนกวิชา")
        if not vendor_name:
            missing_fields.append("ชื่อบริษัท / ร้านค้า")
        if not buy_name1:
            missing_fields.append("ประธาน/กรรมการจัดซื้อคนที่ 1")
        if not check_name1:
            missing_fields.append("ผู้ตรวจรับ / ประธานกรรมการตรวจรับคนที่ 1")

        if missing_fields:
            st.error(
                "❌ กรุณากรอกข้อมูลสำคัญให้ครบถ้วนก่อนสร้างเอกสาร:\n- "
                + "\n- ".join(missing_fields)
            )
        elif clean_num > 0 and clean_num_mid > clean_num:
            st.error(
                "❌ ไม่สามารถสร้างเอกสารได้เนื่องจาก **ราคากลางสูงกว่าวงเงินงบประมาณ** "
                f"({formatted_budget_mid} > {formatted_budget} บาท) กรุณาตรวจสอบและแก้ไขข้อมูล"
            )
        elif not template_files:
            st.error(
                "ไม่สามารถสร้างเอกสารได้"
                f" เนื่องจากไม่มีไฟล์ต้นแบบในโฟลเดอร์ '{TEMPLATE_DIR}'"
            )
        else:
            with st.spinner("กำลังประมวลผลเอกสารจัดซื้อ..."):
                zip_buffer = io.BytesIO()

                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for file_name in template_files:
                        file_path = os.path.join(TEMPLATE_DIR, file_name)
                        processed_stream = process_docx(file_path, replacements_data)

                        # --- ลอจิกเปลี่ยนเลขพัสดุในชื่อไฟล์ปลายทาง ---
                        new_filename = file_name
                        if parcel_no.strip():
                            # เปลี่ยนตัวเลข-ตัวเลข (เช่น 317-69 หรือเลขอื่น) ให้เป็น parcel_no ที่กรอกมา
                            new_filename = re.sub(
                                r"\d+-\d+", parcel_no.strip(), file_name
                            )

                        zip_file.writestr(
                            f"{new_filename}", processed_stream.getvalue()
                        )

                zip_buffer.seek(0)

            st.success("🎉 สร้างเอกสารจัดซื้อทั้งหมดสำเร็จแล้ว!")

            st.download_button(
                label="📦 ดาวน์โหลดชุดเอกสารจัดซื้อ (.zip)",
                data=zip_buffer,
                file_name=f"{project_name}_{parcel_no}.zip",
                mime="application/zip",
                use_container_width=True,
            )


if __name__ == "__main__":
    render_purchase_page()
