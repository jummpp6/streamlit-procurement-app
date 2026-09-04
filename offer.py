# ชื่อ offer.py
import datetime
import io
import os
import re
import zipfile

# เปลี่ยนจาก: from fourcolor import render_fourcolor_dialog
from space_fourcolor import render_space_fourcolor_dialog

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
from offer_helpers import add_business_days, load_shops_data, load_teacher_data
from offer_modals import add_shop_modal


def _extract_shop_info(row):
    """Helper ดึงข้อมูล ที่อยู่, เบอร์โทร, เลขภาษี จาก row ของ df_shops"""
    def find_val(candidates):
        for cand in candidates:
            for c in row.index:
                if str(c).lower().strip() == cand.lower():
                    val = row[c]
                    if pd.notna(val) and str(val).strip().lower() not in [
                        "nan",
                        "none",
                        "",
                    ]:
                        return str(val).strip()
        return ""

    return (
        find_val(["address", "shop_address", "ที่อยู่", "addr"]),
        find_val(["phone", "tel", "เบอร์โทร", "เบอร์โทรศัพท์", "telephone"]),
        find_val(["tax_id", "taxid", "เลขประจำตัวผู้เสียภาษี", "tax_no", "tax"]),
    )


def reset_purchase_form():
    """ฟังก์ชันสำหรับล้างข้อมูลในฟอร์มจัดซื้อ (เรียกใช้ผ่าน on_click)"""
    keys_to_clear = [
        "purchase_parcel_no",
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


def on_edit_vendor_change(df_shops):
    selected_vendor = st.session_state.get("dialog_edit_vendor_select")
    if selected_vendor and not df_shops.empty:
        match = df_shops[df_shops["clean_shop_name"] == str(selected_vendor).strip()]
        if not match.empty:
            addr, phone, tax = _extract_shop_info(match.iloc[0])
            st.session_state["dialog_edit_vendor_addr"] = addr
            st.session_state["dialog_edit_vendor_phone"] = phone
            st.session_state["dialog_edit_vendor_tax"] = tax


@st.dialog("✏️ แก้ไขข้อมูลร้านค้า")
def edit_shop_dialog(shop_list_options, df_shops):
    if not shop_list_options:
        st.warning("ยังไม่มีข้อมูลร้านค้าในระบบ")
        return

    if "dialog_edit_vendor_select" not in st.session_state:
        st.session_state["dialog_edit_vendor_select"] = shop_list_options[0]
        on_edit_vendor_change(df_shops)

    st.selectbox(
        "เลือกชื่อร้านค้า / บริษัท ที่ต้องการแก้ไข",
        options=shop_list_options,
        key="dialog_edit_vendor_select",
        on_change=on_edit_vendor_change,
        args=(df_shops,),
    )
    st.text_area("ที่อยู่", key="dialog_edit_vendor_addr")
    st.text_input("เบอร์โทรศัพท์", key="dialog_edit_vendor_phone")
    st.text_input("เลขประจำตัวผู้เสียภาษี", key="dialog_edit_vendor_tax")

    if st.button("💾 บันทึกการแก้ไข", type="primary", use_container_width=True):
        selected_vendor = st.session_state.get("dialog_edit_vendor_select")
        st.success(f"บันทึกการแก้ไขร้านค้า '{selected_vendor}' เรียบร้อยแล้ว")
        st.rerun()


def render_purchase_page():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=TH+Sarabun+New:wght@400;600;700&display=swap');
        @import url('https://fonts.cdnfonts.com/css/th-sarabunpsk');

        header[data-testid="stHeader"] { display: none !important; }
        .main .block-container { padding-top: 1rem !important; padding-bottom: 2.5rem !important; max-width: 1100px !important; }
        html, body, [class*="css"], .stMarkdown, .stText, p, label, input, select, textarea, button, span, div { font-family: 'TH SarabunPSK', 'TH Sarabun New', 'Sarabun', sans-serif !important; font-size: 103.5% !important; }
        div[data-testid="stHorizontalBlock"] { align-items: flex-start !important; gap: 1rem !important; }
        .stTextInput label, .stSelectbox label, .stDateInput label, .stRadio label { font-weight: 600 !important; font-size: 105% !important; margin-bottom: 4px !important; color: #334155 !important; }
        div[data-baseweb="select"] *, div[data-baseweb="input"] input { font-family: 'TH SarabunPSK', 'TH Sarabun New', 'Sarabun', sans-serif !important; font-size: 105% !important; }
        div[data-testid="stDateInput"] input { font-family: 'TH SarabunPSK', 'TH Sarabun New', 'Sarabun', sans-serif !important; font-size: 105% !important; line-height: 44px !important; height: 44px !important; padding-top: 0px !important; padding-bottom: 0px !important; }
        div[data-baseweb="input"] > div, div[data-baseweb="select"] > div { min-height: 44px !important; height: 44px !important; align-items: center !important; border-radius: 6px !important; background-color: #F8FAFC !important; border: 1px solid #CBD5E1 !important; }
        div[data-baseweb="textarea"] > div { min-height: 44px !important; border-radius: 6px !important; background-color: #F8FAFC !important; border: 1px solid #CBD5E1 !important; }
        div[data-baseweb="textarea"] textarea { font-family: 'TH SarabunPSK', 'TH Sarabun New', 'Sarabun', sans-serif !important; font-size: 95% !important; line-height: 1.2 !important; padding: 6px 8px !important; }
        h1 { font-size: 3.5rem !important; font-weight: 700 !important; margin-top: 0rem !important; margin-bottom: 0.5rem !important; color: #0F172A !important; }
        h3, .stSubheader { font-size: 2.6rem !important; font-weight: 700 !important; color: #0F172A !important; margin-top: 0.8rem !important; margin-bottom: 0.6rem !important; border-bottom: 2px solid #CBD5E1; padding-bottom: 4px; }
        h5 { font-size: 1.8rem !important; font-weight: 600 !important; color: #1E293B !important; margin-top: 0.2rem !important; margin-bottom: 0.5rem !important; background-color: #F1F5F9; padding: 4px 10px; border-left: 4px solid #2563EB; border-radius: 0 4px 4px 0; }
        .stButton > button { padding: 0.5rem 1.5rem !important; font-size: 100% !important; border-radius: 8px !important; font-weight: 600 !important; }
        .stElementContainer { margin-bottom: 0.4rem !important; }
        .header-btn-container { margin-top: 25px; }
        .header-btn-container button { height: 44px !important; min-height: 44px !important; width: 100% !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("🛒 ระบบสร้างเอกสารจัดซื้อ")

    col_top1, col_top2 = st.columns([1, 2], gap="medium")
    with col_top1:
        if st.button("⬅️ กลับหน้าหลัก", use_container_width=True):
            st.session_state.page = "home"
            st.rerun()

    with col_top2:
        if "purchase_parcel_no" not in st.session_state:
            st.session_state["purchase_parcel_no"] = ""
        else:
            val_init = st.session_state["purchase_parcel_no"]
            if "/" in val_init or "_" in val_init:
                st.session_state["purchase_parcel_no"] = val_init.replace("/", "-").replace("_", "-")

        def update_purchase_parcel():
            val = st.session_state["purchase_parcel_no"]
            if "/" in val or "_" in val:
                st.session_state["purchase_parcel_no"] = val.replace("/", "-").replace("_", "-")

        parcel_no = st.text_input(
            "",
            placeholder="ตัวอย่าง: เลขพัสดุ (เช่น 317-69 หรือ 111-69)",
            key="purchase_parcel_no",
            on_change=update_purchase_parcel,
        )

    TEMPLATE_DIR = "templates"
    os.makedirs(TEMPLATE_DIR, exist_ok=True)
    template_files = sorted(
        [
            f
            for f in os.listdir(TEMPLATE_DIR)
            if f.endswith(".docx") and not f.startswith("~$")
        ]
    )

    st.subheader("📝 1. ข้อมูลการจัดซื้อ (หน้า ส.1)")
    col1, col2 = st.columns([1, 1], gap="large")

    with col1:
        project_name = st.text_input(
            "ชื่อโครงการ (ไม่ต้องใส่โครงการนำหน้า)",
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
            "จำนวนรายการทั้งหมด (เช่น 5 รายการ)",
            placeholder="ตัวอย่าง: 5 รายการ",
            key="purchase_item_count",
        )

        clean_max_item_count = 0
        try:
            match_cnt = re.search(r"\d+", item_count_raw.replace(",", ""))
            if match_cnt:
                clean_max_item_count = int(match_cnt.group())
        except ValueError:
            clean_max_item_count = 0

        budget_source_type = st.radio(
            "แหล่งเงินงบประมาณ",
            ["บกศ. (เงินรายได้สถานศึกษา)", "เงินอุดหนุน", "งปม. (เงินงบประมาณ)"],
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
            "วันที่เอกสาร (หน้า ส.1)", datetime.date.today(), key="purchase_doc_date"
        )
        doc_no_raw = st.text_input(
            "เลขที่คำสั่ง", placeholder="ตัวอย่าง: 960/2569", key="purchase_doc_no"
        )
        doc_no_save = st.text_input(
            "เลขที่บันทึก", placeholder="ตัวอย่าง: 960", key="purchase_doc_save"
        )

    use_thai_num, use_thai_num2 = True, True

    st.write("")
    st.subheader("👥 2. คำสั่งคณะกรรมการจัดซื้อ / ตรวจรับ")
    person_options, person_dict = load_teacher_data("teachers.xlsx")

    col_count1, col_count2 = st.columns([1, 1], gap="large")
    buy_count = col_count1.selectbox(
        "จำนวนกรรมการจัดซื้อ",
        options=[1, 2, 3],
        index=2,
        key="purchase_buy_count_select",
    )
    check_count = col_count2.selectbox(
        "จำนวนกรรมการตรวจรับ",
        options=[1, 2, 3],
        index=2,
        key="purchase_check_count_select",
    )

    st.markdown("##### 🛒 คณะกรรมการจัดซื้อ")
    buy_persons = []
    defaults_buy = [
        ("จัดซื้อ 1", "นายสมชาย ใจดี", "ประธานกรรมการฯ"),
        ("จัดซื้อ 2", "นางสาวสมหญิง ใจงาม", "กรรมการฯ"),
        ("จัดซื้อ 3", "นายสมปอง ใจกล้า", "กรรมการฯ"),
    ]
    with st.container(border=True):
        for idx in range(3):
            label, def_name, def_pos = defaults_buy[idx]
            if idx < buy_count:
                p = render_person_inputs(
                    label,
                    f"purchase_buy{idx+1}",
                    def_name,
                    "",
                    def_pos,
                    person_options,
                    person_dict,
                )
            else:
                p = ("", "", "")
            buy_persons.append(p)

    st.markdown("##### 🔍 คณะกรรมการตรวจรับพัสดุ")
    check_persons = []
    check_pos_options = (
        ["ผู้ตรวจรับพัสดุ"]
        if check_count == 1
        else ["ประธานกรรมการฯ", "กรรมการฯ", "กรรมการและเลขานุการฯ"]
    )
    default_check_pos1 = "ผู้ตรวจรับพัสดุ" if check_count == 1 else "ประธานกรรมการฯ"
    defaults_check = [
        ("ตรวจรับ 1", "นายสมชาย ใจดี", default_check_pos1),
        ("ตรวจรับ 2", "นางสาวสมหญิง ใจงาม", "กรรมการฯ"),
        ("ตรวจรับ 3", "นายสมปอง ใจกล้า", "กรรมการฯ"),
    ]

    with st.container(border=True):
        for idx in range(3):
            label, def_name, def_pos = defaults_check[idx]
            if idx < check_count:
                kw = {"pos_options_custom": check_pos_options} if idx == 0 else {}
                p = render_person_inputs(
                    label,
                    f"purchase_check{idx+1}",
                    def_name,
                    "",
                    def_pos,
                    person_options,
                    person_dict,
                    **kw,
                )
            else:
                p = ("", "", "")
            check_persons.append(p)

    st.write("")
    st.subheader("📋 3. บันทึกรายงานผลการพิจารณา")

    df_shops = load_shops_data()
    shop_list_options = []
    if not df_shops.empty:
        name_col = next(
            (
                col
                for col in df_shops.columns
                if str(col).lower()
                in ["shop_name", "name", "shopname", "ร้านค้า", "ชื่อร้าน"]
            ),
            df_shops.columns[0],
        )
        df_shops["clean_shop_name"] = df_shops[name_col].astype(str).str.strip()
        shop_list_options = [
            s
            for s in df_shops["clean_shop_name"].unique()
            if s and s.lower() not in ["nan", "none", "nat", ""]
        ]

    top_col1, top_col2, top_col3, top_col4 = st.columns([1.2, 1, 1, 1], gap="small")
    selected_date2 = top_col1.date_input(
        "วันที่รายงานผล", datetime.date.today(), key="purchase_date_report"
    )
    shop_count = top_col2.selectbox(
        "จำนวนร้านค้า/บริษัท",
        options=[1, 2, 3, 4],
        index=0,
        key="purchase_shop_count_select",
    )

    with top_col3:
        st.markdown('<div class="header-btn-container">', unsafe_allow_html=True)
        if st.button(
            "➕ เพิ่มร้านค้า", use_container_width=True, key="purchase_btn_add_shop"
        ):
            add_shop_modal()
        st.markdown("</div>", unsafe_allow_html=True)

    with top_col4:
        st.markdown('<div class="header-btn-container">', unsafe_allow_html=True)
        if st.button(
            "✏️ แก้ไขร้าน", use_container_width=True, key="purchase_btn_edit_shop"
        ):
            edit_shop_dialog(shop_list_options, df_shops)
        st.markdown("</div>", unsafe_allow_html=True)

    st.write("")
    col_ratios = [3, 2, 1, 5]
    h1, h2, h3, h4 = st.columns(col_ratios, gap="small")
    h1.caption("**ชื่อร้านค้า**")
    h2.caption("**วงเงิน**")
    h3.caption("**จำนวน**")
    h4.caption("**รายละเอียดร้านค้า**")

    selected_vendors, budgetmid_list, vendor_item_counts, vendor_details_list = (
        [],
        [],
        [],
        [],
    )

    def on_vendor_change(index):
        detail_key = f"purchase_vendor_detail_{index}"
        if detail_key in st.session_state:
            del st.session_state[detail_key]

    for i in range(shop_count):
        c1, c2, c3, c4 = st.columns(col_ratios, gap="small")

        with c1:
            vendor = st.selectbox(
                f"ร้านค้า {i+1}",
                options=shop_list_options,
                index=0 if shop_list_options else None,
                key=f"purchase_selected_vendor_name_{i+1}",
                on_change=on_vendor_change,
                args=(i + 1,),
                label_visibility="collapsed",
            )
            selected_vendors.append(vendor)

        with c2:
            b_mid = st.text_input(
                f"วงเงิน {i+1}",
                placeholder="ตัวอย่าง: 4996",
                key=f"purchase_budget_mid_{i+1}",
                label_visibility="collapsed",
            )
            budgetmid_list.append(b_mid)

        with c3:
            v_item_count = st.text_input(
                f"จำนวน {i+1}",
                value=item_count_raw if (item_count_raw and shop_count == 1) else "",
                placeholder="",
                key=f"purchase_item_count_{i+1}",
                label_visibility="collapsed",
            )
            vendor_item_counts.append(v_item_count)

        with c4:
            detail_lines = []
            c_addr, c_phone, c_tax = "", "", ""
            if vendor and not df_shops.empty:
                match = df_shops[df_shops["clean_shop_name"] == str(vendor).strip()]
                if not match.empty:
                    c_addr, c_phone, c_tax = _extract_shop_info(match.iloc[0])
                    if c_addr:
                        detail_lines.append(f"ที่อยู่: {c_addr}")
                    if c_phone:
                        detail_lines.append(f"โทร: {c_phone}")
                    if c_tax:
                        detail_lines.append(f"เลขภาษี: {c_tax}")
                    detail_text = (
                        "\n".join(detail_lines)
                        if detail_lines
                        else "ไม่มีข้อมูลเพิ่มเติม"
                    )
                else:
                    detail_text = "ไม่พบข้อมูลในระบบ"
            else:
                detail_text = ""

            st.text_area(
                f"รายละเอียด {i+1}",
                value=detail_text,
                disabled=True,
                height=80,
                label_visibility="collapsed",
            )
            vendor_details_list.append((c_addr, c_phone, c_tax))

    st.write("")
    st.write("")
    st.subheader("📄 4. ใบสั่งซื้อ")
    col1, col2 = st.columns([1, 1], gap="large")
    submit_no = col1.text_input(
        "เลขที่ข้อตกลง", placeholder="ตัวอย่าง: 960/2569", key="purchase_submit_no"
    )
    default_order_date = add_business_days(selected_date2, 5)
    selected_date5 = col2.date_input(
        "วันที่ใบสั่งซื้อ", value=default_order_date, key="purchase_date_order_input"
    )

    st.write("")
    st.subheader("✅ 5. ใบตรวจรับการจัดซื้อ")
    selected_date4 = st.date_input(
        "วันที่ตรวจรับ", datetime.date.today(), key="purchase_check_date"
    )

    formatted_budget = format_budget_money(budget, use_thai=use_thai_num)
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

    def generate_submit_no_for_vendor(base_submit_no, index):
        if not base_submit_no:
            return ""
        if index == 0:
            return to_thai_num(base_submit_no) if use_thai_num else str(base_submit_no)
        new_sub_no = (
            f"{base_submit_no.split('/', 1)[0]}.{index}/{base_submit_no.split('/', 1)[1]}"
            if "/" in base_submit_no
            else f"{base_submit_no}.{index}"
        )
        return to_thai_num(new_sub_no) if use_thai_num else str(new_sub_no)

    sub_nos = [generate_submit_no_for_vendor(submit_no, i) for i in range(4)]
    budget_with_text = f"{formatted_budget} บาท ({budget_text})"

    def get_vendor_info(idx):
        v_name = selected_vendors[idx] if len(selected_vendors) > idx else ""
        v_b_mid = budgetmid_list[idx] if len(budgetmid_list) > idx else ""
        v_cnt = vendor_item_counts[idx] if len(vendor_item_counts) > idx else ""
        c_addr, c_phone, c_tax = (
            vendor_details_list[idx] if len(vendor_details_list) > idx else ("", "", "")
        )

        c_num_mid = 0.0
        try:
            c_num_mid = float(
                v_b_mid.replace(",", "").replace(" ", "").replace("บาท", "")
            )
            b_text_mid = bahttext(c_num_mid)
        except ValueError:
            b_text_mid = v_b_mid if v_b_mid else "-"

        c_cnt_num = 0
        try:
            match_item = re.search(r"\d+", str(v_cnt).replace(",", ""))
            if match_item:
                c_cnt_num = int(match_item.group())
        except ValueError:
            c_cnt_num = 0

        fmt_b_mid = (
            format_budget_money(v_b_mid, use_thai=use_thai_num2) if v_b_mid else ""
        )
        b_with_text_mid = f"{fmt_b_mid} บาท ({b_text_mid})" if fmt_b_mid else ""
        fmt_item_cnt = to_thai_num(v_cnt) if (use_thai_num and v_cnt) else str(v_cnt)

        return {
            "name": v_name,
            "addr": to_thai_num(c_addr) if use_thai_num else c_addr,
            "phone": to_thai_num(c_phone) if use_thai_num else c_phone,
            "tax": to_thai_num(c_tax) if use_thai_num else c_tax,
            "budget_mid": fmt_b_mid,
            "budget_text_mid": b_text_mid,
            "budget_with_text_mid": b_with_text_mid,
            "item_count": fmt_item_cnt,
            "clean_num_mid": c_num_mid,
            "clean_cnt_num": c_cnt_num,
        }

    v_list = [get_vendor_info(i) for i in range(4)]
    v1, v2, v3, v4 = v_list

    total_budget_mid_num = sum(v["clean_num_mid"] for v in v_list[:shop_count])
    total_item_count_num = sum(v["clean_cnt_num"] for v in v_list[:shop_count])

    formatted_total_budget_mid = format_budget_money(
        str(total_budget_mid_num), use_thai=use_thai_num2
    )
    total_budget_text_mid = (
        bahttext(total_budget_mid_num) if total_budget_mid_num > 0 else "-"
    )
    total_budget_with_text_mid = (
        f"{formatted_total_budget_mid} บาท ({total_budget_text_mid})"
        if total_budget_mid_num > 0
        else ""
    )
    formatted_total_item_count = (
        to_thai_num(str(total_item_count_num))
        if use_thai_num
        else str(total_item_count_num)
    )
    formatted_project_name = to_thai_num(project_name) if use_thai_num else project_name

    replacements_data = {
        "{{PROJECT_NAME}}": formatted_project_name,
        "{{BUDGET}}": formatted_budget,
        "{{BUDGET_TEXT}}": budget_text,
        "{{BUDGET_WITH_TEXT}}": budget_with_text,
        "{{BUDGET_TYPE}}": budget_type_text,
        "{{DEPARTMENT}}": department,
        "{{ITEM_COUNT}}": v1["item_count"],
        "{{BUDGET_MID}}": v1["budget_mid"],
        "{{BUDGET_TEXT_MID}}": v1["budget_text_mid"],
        "{{BUDGET_WITH_TEXT_MID}}": v1["budget_with_text_mid"],
        "{{TOTAL_ITEM_COUNT}}": formatted_total_item_count,
        "{{TOTAL_BUDGET_MID}}": formatted_total_budget_mid,
        "{{TOTAL_BUDGET_TEXT_MID}}": total_budget_text_mid,
        "{{TOTAL_BUDGET_WITH_TEXT_MID}}": total_budget_with_text_mid,
        "{{VENDOR_NAME}}": v1["name"],
        "{{VENDOR_ADDRESS}}": v1["addr"],
        "{{VENDOR_PHONE}}": v1["phone"],
        "{{VENDOR_TAX_ID}}": v1["tax"],
        "{{SUBMIT_NO}}": sub_nos[0],
        "{{PERCEL_NO}}": parcel_no,
        "{{DOC_DATE}}": formatted_date,
        "{{DOC_DATE2}}": formatted_date2,
        "{{DOC_DATE3}}": formatted_date3,
        "{{DOC_DATE4}}": formatted_date4,
        "{{DOC_DATE5}}": formatted_date5,
        "{{DOC_NO}}": formatted_doc_no,
        "{{DOC_NO_SAVE}}": formatted_doc_no_save,
        "{{DIRECTOR_INSPECTOR}}": "คณะกรรมการ" if check_count > 1 else "ผู้",
    }

    for idx, key_suffix in enumerate(["2", "3", "4"], start=1):
        v = v_list[idx]
        replacements_data.update(
            {
                f"{{{{VENDOR_NAME{key_suffix}}}}}": v["name"],
                f"{{{{VENDOR_ADDRESS{key_suffix}}}}}": v["addr"],
                f"{{{{VENDOR_PHONE{key_suffix}}}}}": v["phone"],
                f"{{{{VENDOR_TAX_ID{key_suffix}}}}}": v["tax"],
                f"{{{{BUDGET_MID{key_suffix}}}}}": v["budget_mid"],
                f"{{{{BUDGET_TEXT_MID{key_suffix}}}}}": v["budget_text_mid"],
                f"{{{{BUDGET_WITH_TEXT_MID{key_suffix}}}}}": v["budget_with_text_mid"],
                f"{{{{ITEM_COUNT{key_suffix}}}}}": v["item_count"],
                f"{{{{SUBMIT_NO{key_suffix}}}}}": sub_nos[idx],
            }
        )

    thai_nums = ["๑", "๒", "๓"]
    for idx, (name, acad, pos) in enumerate(buy_persons):
        suf = "" if idx == 0 else str(idx + 1)
        valid = name and (buy_count >= idx + 1)
        replacements_data.update(
            {
                f"{{{{DIRECTOR_NAME_BUY{suf}}}}}": (
                    f"{thai_nums[idx]}. {name}" if valid else ""
                ),
                f"{{{{DIRECTOR_ACAD_BUY{suf}}}}}": acad if valid else "",
                f"{{{{DIRECTOR_POS_BUY{suf}}}}}": pos if valid else "",
                f"{{{{DIRECTOR_NAME_BUY{suf}_PLAIN}}}}": name if valid else "",
                f"{{{{SIGN_LINE_BUY{idx+1}}}}}": (
                    f"ลงชื่อ....................................{pos if pos else 'กรรมการฯ'}"
                    if valid
                    else ""
                ),
                f"{{{{SIGN_NAME_BUY{idx+1}}}}}": f"({name})" if valid else "",
            }
        )

    for idx, (name, acad, pos) in enumerate(check_persons):
        suf = "" if idx == 0 else str(idx + 1)
        valid = name and (check_count >= idx + 1)
        actual_pos = (
            ("ผู้ตรวจรับพัสดุ" if check_count == 1 else pos) if idx == 0 else pos
        )

        sign_line = (
            "ลงชื่อ....................................ผู้ตรวจรับพัสดุ"
            if (idx == 0 and check_count == 1)
            else (
                "ลงชื่อ....................................ประธานกรรมการฯ"
                if idx == 0
                else "ลงชื่อ....................................กรรมการฯ"
            )
        )

        replacements_data.update(
            {
                f"{{{{CHECKITEM_NAME{suf}}}}}": (
                    f"{thai_nums[idx]}. {name}" if valid else ""
                ),
                f"{{{{CHECKITEM_ACAD{suf}}}}}": acad if valid else "",
                f"{{{{CHECKITEM_POS{suf}}}}}": actual_pos if valid else "",
                f"{{{{CHECKITEM_NAME{suf}_PLAIN}}}}": name if valid else "",
                f"{{{{SIGN_LINE_CHECK{idx+1}}}}}": sign_line if valid else "",
                f"{{{{SIGN_NAME_CHECK{idx+1}}}}}": f"({name})" if valid else "",
            }
        )

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

            if clean_num > 0 and total_budget_mid_num > clean_num:
                st.markdown(
                    f"**ราคากลางรวม ({shop_count} ร้าน):** {formatted_total_budget_mid} บาท 🚨 **(สูงกว่าวงเงินงบประมาณ)**"
                )
            elif total_budget_mid_num > 0:
                st.markdown(
                    f"**ราคากลางรวม ({shop_count} ร้าน):** {formatted_total_budget_mid} บาท"
                )

            st.markdown(f"**แหล่งเงิน:** {budget_type_text}")

            if clean_max_item_count > 0 and total_item_count_num > clean_max_item_count:
                st.markdown(
                    f"**จำนวนรายการรวม ({shop_count} ร้าน):** {formatted_total_item_count} รายการ 🚨 **(เกินจำนวนรายการทั้งหมด {to_thai_num(str(clean_max_item_count))} รายการ)**"
                )
            else:
                st.markdown(
                    f"**จำนวนรายการทั้งหมด:** {to_thai_num(item_count_raw) if item_count_raw else '⚠️ *(ยังไม่ได้กรอก)*'}"
                )

        with col_prev2:
            st.markdown(
                f"**เลขพัสดุที่ใช้ระบุชื่อไฟล์:** `{parcel_no if parcel_no else 'ไม่ได้ระบุ (ใช้ตามชื่อต้นแบบ)'}`"
            )
            st.markdown(
                f"**เลขที่คำสั่ง / ข้อตกลง:** {formatted_doc_no if doc_no_raw else '-'} / {sub_nos[0] if submit_no else '-'}"
            )
            st.markdown(
                f"**ร้านค้า / ผู้เสนอราคา:** {v1['name'] if v1['name'] else '⚠️ *(ยังไม่ได้เลือก)*'}"
            )
            st.markdown(f"**วันที่เอกสาร (ส.1):** {formatted_date}")
            st.markdown(f"**วันที่รายงานผล:** {formatted_date2}")
            st.markdown(f"**วันที่ใบสั่งซื้อ:** {formatted_date5}")
            st.markdown(f"**วันที่ตรวจรับ:** {formatted_date4}")

        st.markdown("---")
        col_prev_buy, col_prev_check = st.columns([1, 1], gap="medium")

        with col_prev_buy:
            st.markdown("**🛒 รายชื่อคณะกรรมการจัดซื้อ:**")
            buy_list_preview = [
                f"{i+1}. {p[0]} ({p[2]})"
                for i, p in enumerate(buy_persons[:buy_count])
                if p[0]
            ]
            if buy_list_preview:
                for item in buy_list_preview:
                    st.text(f" • {item}")
            else:
                st.caption("⚠️ ยังไม่มีการเลือกกรรมการจัดซื้อ")

        with col_prev_check:
            st.markdown("**🔍 รายชื่อคณะกรรมการตรวจรับ:**")
            check_list_preview = [
                f"{i+1}. {p[0]} ({'ผู้ตรวจรับพัสดุ' if check_count == 1 and i == 0 else p[2]})"
                for i, p in enumerate(check_persons[:check_count])
                if p[0]
            ]
            if check_list_preview:
                for item in check_list_preview:
                    st.text(f" • {item}")
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
            if not v1["name"]:
                missing_fields.append("ชื่อบริษัท / ร้านค้า")
            if not buy_persons[0][0]:
                missing_fields.append("ประธาน/กรรมการจัดซื้อคนที่ 1")
            if not check_persons[0][0]:
                missing_fields.append("ผู้ตรวจรับ / ประธานกรรมการตรวจรับคนที่ 1")

            if missing_fields:
                st.error(
                    "❌ กรุณากรอกข้อมูลสำคัญให้ครบถ้วนก่อนสร้างเอกสาร:\n- "
                    + "\n- ".join(missing_fields)
                )
            elif clean_num > 0 and total_budget_mid_num > clean_num:
                st.error(
                    f"❌ ไม่สามารถสร้างเอกสารได้เนื่องจาก **ผลรวมราคากลางของทุกร้านค้าสูงกว่าวงเงินงบประมาณ** ({formatted_total_budget_mid} บาท > {formatted_budget} บาท) กรุณาตรวจสอบและแก้ไขข้อมูล"
                )
            elif clean_max_item_count > 0 and total_item_count_num > clean_max_item_count:
                st.error(
                    f"❌ ไม่สามารถสร้างเอกสารได้เนื่องจาก **ผลรวมจำนวนรายการของทุกร้านค้า เกินกว่าจำนวนรายการทั้งหมดที่ระบุไว้** ({formatted_total_item_count} รายการ > {to_thai_num(str(clean_max_item_count))} รายการ) กรุณาตรวจสอบและแก้ไขข้อมูล"
                )
            elif not template_files:
                st.error(
                    f"ไม่สามารถสร้างเอกสารได้ เนื่องจากไม่มีไฟล์ต้นแบบในโฟลเดอร์ '{TEMPLATE_DIR}'"
                )
            else:
                with st.spinner("กำลังประมวลผลเอกสารจัดซื้อ..."):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                        for file_name in template_files:
                            file_path = os.path.join(TEMPLATE_DIR, file_name)
                            processed_stream = process_docx(
                                file_path,
                                replacements_data,
                                shop_count=shop_count,
                                buy_count=buy_count,
                                check_count=check_count,
                            )
                            new_filename = (
                                re.sub(r"\d+-\d+", parcel_no.strip(), file_name)
                                if parcel_no.strip()
                                else file_name
                            )
                            zip_file.writestr(
                                f"{new_filename}", processed_stream.getvalue()
                            )

                    zip_buffer.seek(0)

                    st.success("🎉 สร้างเอกสารจัดซื้อทั้งหมดสำเร็จแล้ว!")
                    clean_filename_doc_no = (
                        parcel_no.strip()
                        if parcel_no.strip()
                        else (doc_no_raw.replace("/", "-") if doc_no_raw else "ไม่ระบุเลข")
                    )
                    st.download_button(
                        label="📦 ดาวน์โหลดชุดเอกสารจัดซื้อ (.zip)",
                        data=zip_buffer,
                        file_name=f"เอกสารจัดซื้อ_{project_name}_{clean_filename_doc_no}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

    st.write("")
    st.subheader("🖼️ 7. เอกสารสี่สี / คุณลักษณะ")

    with st.container(border=True):
        col_fc1, col_fc2 = st.columns([2, 1], gap="medium")
        with col_fc1:
            st.markdown(
                "**จัดการข้อมูลรายการสินค้าสำหรับออกเอกสารแนบ Space และ Fourcolor**"
            )

        with col_fc2:
            if st.button(
                "🎨 เปิดตารางกรอกรายการ",
                type="primary",
                use_container_width=True,
                key="btn_open_space_fourcolor",
            ):
                rec_1 = check_persons[0][0] if check_persons[0][0] else None
                rec_2 = (
                    check_persons[1][0]
                    if len(check_persons) > 1 and check_persons[1][0]
                    else None
                )

                render_space_fourcolor_dialog(
                    default_receiver=rec_1,
                    default_receiver_sub=rec_2,
                    default_total_amount=clean_num,
                    project_name=project_name,
                    default_parcel_no=parcel_no,
                    department=department,
                    budget_type=budget_type_text,
                )


if __name__ == "__main__":
    render_purchase_page()
