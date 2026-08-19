# ==========================================
# ไฟล์: doc_processor.py
# ==========================================
import io
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


def to_thai_num(text):
    if not text:
        return ""
    arabic_digits = "0123456789"
    thai_digits = "๐๑๒๓๔๕๖๗๘๙"
    return str(text).translate(str.maketrans(arabic_digits, thai_digits))


def format_thai_date(date_obj, use_thai=True):
    if not date_obj:
        return ""

    THAI_MONTHS = [
        "",
        "มกราคม",
        "กุมภาพันธ์",
        "มีนาคม",
        "เมษายน",
        "พฤษภาคม",
        "มิถุนายน",
        "กรกฎาคม",
        "สิงหาคม",
        "กันยายน",
        "ตุลาคม",
        "พฤศจิกายน",
        "ธันวาคม",
    ]

    day = date_obj.day
    month = THAI_MONTHS[date_obj.month]
    year = date_obj.year + 543  # แปลง ค.ศ. เป็น พ.ศ.

    if use_thai:
        day_str = to_thai_num(day)
        year_str = to_thai_num(year)
    else:
        day_str = str(day)
        year_str = str(year)

    # เว้นวรรค 2 สเปซบาร์ระหว่าง [วัน]  [เดือน]  [พ.ศ. ปี]
    return f"{day_str}  {month}  พ.ศ. {year_str}"


def format_budget_money(text, use_thai=True):
    if not text:
        return ""
    clean_str = str(text).replace(",", "").replace(" ", "").replace("บาท", "").strip()
    try:
        val = float(clean_str)
        formatted = f"{int(val):,}" if val.is_integer() else f"{val:,.2f}"
    except ValueError:
        formatted = text
    return to_thai_num(formatted) if use_thai else formatted


def set_font_exact_16(run, font_name="TH SarabunPSK", font_size_pt=16, is_bold=None):
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rPr.append(rFonts)
    if is_bold is not None:
        run.bold = is_bold


def replace_text_in_paragraph(paragraph, replacements, default_font="TH SarabunPSK"):
    full_text = "".join([run.text for run in paragraph.runs])
    has_target_key = any(key in full_text for key in replacements.keys())
    if not has_target_key:
        return

    first_run = paragraph.runs[0] if paragraph.runs else None
    is_bold = first_run.bold if first_run else False
    
    # 1. ดึงขนาดฟอนต์เดิมจาก run แรกมาเก็บไว้ (ถ้าไม่ได้ตั้งไว้ ให้ใช้ค่าเริ่มต้นเป็น 16)
    original_size = 16
    if first_run and first_run.font.size:
        original_size = first_run.font.size.pt

    for key, value in replacements.items():
        if key in full_text:
            full_text = full_text.replace(key, str(value))

    for run in paragraph.runs:
        run.text = ""

    if paragraph.runs:
        paragraph.runs[0].text = full_text
        # 2. ส่งขนาด original_size ที่ดึงมาได้เข้าไปแทนเลข 16
        set_font_exact_16(
            paragraph.runs[0], default_font, font_size_pt=original_size, is_bold=is_bold
        )


def process_table(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                replace_text_in_paragraph(p, replacements)
            for nested_table in cell.tables:
                process_table(nested_table, replacements)


def process_docx(file_path, replacements_processed):
    doc = Document(file_path)

    # แทนที่ข้อความในย่อหน้าปกติ
    for p in doc.paragraphs:
        replace_text_in_paragraph(p, replacements_processed)

    # แทนที่ข้อความในตาราง
    for table in doc.tables:
        process_table(table, replacements_processed)

    # ❌ ลบส่วนลูป section.header และ section.footer ออกทิ้งทั้งหมด
    # เพื่อไม่ให้ python-docx เข้าไปยุ่งหรือสร้าง XML Structure ในส่วนหัวกระดาษใหม่

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output
