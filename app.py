import streamlit as st

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

from hiring import render_hiring_page  # 👈 Import หน้าจัดจ้างเข้ามา
from offer import render_purchase_page

st.set_page_config(
    page_title="ระบบสร้างเอกสารพัสดุอัตโนมัติโดยวิธีเฉพาะเจาะจง",
    page_icon="📄",
    layout="wide",
)

if "page" not in st.session_state:
    st.session_state.page = "home"

# --- หน้าเลือกประเภท ---
if st.session_state.page == "home":
    st.title("📄 ระบบสร้างเอกสารพัสดุอัตโนมัติโดยวิธีเฉพาะเจาะจง")
    st.write("กรุณาเลือกประเภทการดำเนินงานที่ต้องการ:")
    st.write("")

    col1, col2 = st.columns(2)

    with col1:
        if st.button(
            "🛒 จัดซื้อ (ซื้อวัสดุ/ครุภัณฑ์)",
            use_container_width=True,
            type="primary",
        ):
            st.session_state.page = "purchase"
            st.rerun()

    with col2:
        if st.button("🛠️ จัดจ้าง (จ้างทำของ/จ้างซ่อม)", use_container_width=True):
            st.session_state.page = "hiring"
            st.rerun()

# --- หน้าจัดซื้อ ---
elif st.session_state.page == "purchase":
    render_purchase_page()

# --- หน้าจัดจ้าง ---
elif st.session_state.page == "hiring":
    render_hiring_page()  # 👈 เรียกใช้งานระบบจัดจ้าง
