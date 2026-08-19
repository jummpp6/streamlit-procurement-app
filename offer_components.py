import streamlit as st
from doc_processor import to_thai_num


def render_person_inputs(
    label_title,
    key_prefix,
    default_name,
    default_acad,
    default_pos,
    person_options,
    person_dict,
    pos_options_custom=None,
):
    """คอมโพเนนต์สำหรับสร้างช่องกรอกข้อมูล ชื่อ-วิทยฐานะ-ตำแหน่ง ของกรรมการแต่ละคน"""
    c1, c2, c3 = st.columns([1, 1, 1], gap="medium")
    with c1:
        if person_dict:
            try:
                default_idx = person_options.index(default_name)
            except ValueError:
                default_idx = 0

            selected_name = st.selectbox(
                f"ชื่อ-นามสกุล ({label_title})",
                options=person_options,
                index=default_idx,
                key=f"{key_prefix}_sel",
            )
            name = selected_name
        else:
            name = st.text_input(
                f"ชื่อ-นามสกุล ({label_title})",
                value=default_name,
                key=f"{key_prefix}_name",
            )

    acad_auto = person_dict.get(name, "") if (person_dict and name) else default_acad
    # 🟢 [จุดที่ 1] แปลงวิทยฐานะให้เป็นเลขไทยทันที (เช่น "ครู คศ.1" -> "ครู คศ.๑")
    acad_auto_thai = to_thai_num(acad_auto)

    with c2:
        academic = st.text_input(
            f"วิทยฐานะ ({label_title})",
            value=acad_auto_thai,  # 👈 เปลี่ยนมาใช้ค่าที่แปลงเป็นเลขไทยแล้ว
            key=f"{key_prefix}_acad_{name}",
        )

    with c3:
        if pos_options_custom is not None:
            pos_options = pos_options_custom
        else:
            pos_options = [
                "ประธานกรรมการฯ",
                "กรรมการฯ",
                "กรรมการและเลขานุการฯ",
            ]

        # 🟢 [จุดที่ 2] แปลงตัวเลือกตำแหน่งทั้งหมดให้เป็นเลขไทย
        pos_options = [to_thai_num(opt) for opt in pos_options]
        default_pos_thai = to_thai_num(default_pos)

        pos_index = pos_options.index(default_pos_thai) if default_pos_thai in pos_options else 0

        position = st.selectbox(
            f"ตำแหน่ง ({label_title})",
            options=pos_options,
            index=pos_index,
            key=f"{key_prefix}_pos",
        )

    # แปลงผลลัพธ์วิทยฐานะเป็นเลขไทยอีกชั้นก่อน return (เผื่อผู้ใช้อนุญาตให้พิมพ์แก้เองในช่อง text_input)
    return name, to_thai_num(academic), position
