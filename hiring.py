
🔨 ระบบสร้างเอกสารจัดจ้าง
ตัวอย่าง: เลขพัสดุ (เช่น 317-69 หรือ 111-69)
📝 1. ข้อมูลการจัดจ้าง (หน้า ส.1)
ชื่อโครงการ / ชื่องานจ้าง

ตัวอย่าง: จ้างซ่อมแซมเครื่องปรับอากาศ
จำนวนเงิน / วงเงินที่จะจ้าง (พิมพ์เฉพาะตัวเลข)

ตัวอย่าง: 15000
💡
แปลงเป็นตัวหนังสืออัตโนมัติ:

จำนวนรายการ / งานทั้งหมด (เช่น 1 งาน)

ตัวอย่าง: 1 งาน
แหล่งเงินงบประมาณ




งาน หรือ แผนกวิชาที่จัดทำ

ตัวอย่าง: แผนกวิชาช่างยนต์ / งานกิจกรรมนักเรียนนักศึกษา
วันที่เอกสาร (หน้า ส.1)

2026
/
09
/
04

09/04/2026
เลขที่คำสั่ง

ตัวอย่าง: 961/2569
เลขที่บันทึก

ตัวอย่าง: 961
👥 2. คำสั่งคณะกรรมการจัดจ้าง / ตรวจรับ
จำนวนกรรมการจัดจ้าง

3

จำนวนกรรมการตรวจรับ

3

🛠️ คณะกรรมการจัดจ้าง
ชื่อ-นามสกุล (จัดจ้าง 1)

Choose an option

วิทยฐานะ (จัดจ้าง 1)

ตำแหน่ง (จัดจ้าง 1)

ประธานกรรมการฯ

ชื่อ-นามสกุล (จัดจ้าง 2)

Choose an option

วิทยฐานะ (จัดจ้าง 2)

ตำแหน่ง (จัดจ้าง 2)

กรรมการฯ

ชื่อ-นามสกุล (จัดจ้าง 3)

Choose an option

วิทยฐานะ (จัดจ้าง 3)

ตำแหน่ง (จัดจ้าง 3)

กรรมการฯ

🔍 คณะกรรมการตรวจรับงานจ้าง
ชื่อ-นามสกุล (ตรวจรับ 1)

Choose an option

วิทยฐานะ (ตรวจรับ 1)

ตำแหน่ง (ตรวจรับ 1)

ประธานกรรมการฯ

ชื่อ-นามสกุล (ตรวจรับ 2)

Choose an option

วิทยฐานะ (ตรวจรับ 2)

ตำแหน่ง (ตรวจรับ 2)

กรรมการฯ

ชื่อ-นามสกุล (ตรวจรับ 3)

Choose an option

วิทยฐานะ (ตรวจรับ 3)

ตำแหน่ง (ตรวจรับ 3)

กรรมการฯ

📋 3. บันทึกรายงานผลการพิจารณา
วันที่รายงานผล

2026
/
09
/
04

09/04/2026
จำนวนผู้รับจ้าง/บริษัท

1

ชื่อผู้รับจ้าง/ร้านค้า

วงเงิน

จำนวน

รายละเอียดผู้รับจ้าง

บริษัท มานิตวิทยา จำกัด

ตัวอย่าง: 15000
ที่อยู่: 76,77,78-79 ซอยศรีสุริโยทัย 1 ต.ทะเลชุบศร อ.เมือง จ.ลพบุรี 15000
โทร: 081-7000767
เลขภาษี: 165533000070
📄 4. ใบสั่งจ้าง / ใบข้อตกลงจ้าง
เลขที่ข้อตกลง

ตัวอย่าง: 961/2569
วันที่ใบสั่งจ้าง

2026
/
09
/
09

09/09/2026
✅ 5. ใบตรวจรับการจ้าง
วันที่ตรวจรับ

2026
/
09
/
04

09/04/2026
📋 6. รายละเอียดรายการพัสดุและสเปค (Item Specifications)
กำหนดรายการพัสดุหรือขอบเขตงานจ้าง ระบบจะนำไปสร้างไฟล์ Excel ข้อกำหนดการจ้าง (Space) ให้อัตโนมัติเมื่อกดสร้างเอกสาร

AttributeError: This app has encountered an error. The original error message is redacted to prevent data leaks. Full error details have been recorded in the logs (if you're on Streamlit Cloud, click on 'Manage app' in the lower right of your app).
Traceback:
File "/mount/src/streamlit-procurement-app/app.py", line 67, in <module>
    render_hiring_page()  # 👈 เรียกใช้งานระบบจัดจ้าง
    ~~~~~~~~~~~~~~~~~~^^
File "/mount/src/streamlit-procurement-app/hiring.py", line 744, in render_hiring_page
    edited_items_df = st.data_editor(
        st.session_state["hiring_items_editor"],  # 👈 เติมตัวแปรข้อมูลตรงนี้ครับ
    ...<9 lines>...
        },
    )
File "/home/adminuser/venv/lib/python3.14/site-packages/streamlit/runtime/metrics_util.py", line 725, in wrapped_func
    result = non_optional_func(*args, **kwargs)
File "/home/adminuser/venv/lib/python3.14/site-packages/streamlit/elements/widgets/data_editor.py", line 1221, in data_editor
    data_df = dataframe_util.convert_anything_to_pandas_df(data, ensure_copy=True)
File "/home/adminuser/venv/lib/python3.14/site-packages/streamlit/dataframe_util.py", line 807, in convert_anything_to_pandas_df
    return _dict_to_pandas_df(data)
File "/home/adminuser/venv/lib/python3.14/site-packages/streamlit/dataframe_util.py", line 550, in _dict_to_pandas_df
    return _fix_column_naming(pd.DataFrame.from_dict(data, orient="index"))
                              ~~~~~~~~~~~~~~~~~~~~~~^^^^^^^^^^^^^^^^^^^^^^
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/core/frame.py", line 1988, in from_dict
    data = _from_nested_dict(data)
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/core/frame.py", line 16697, in _from_nested_dict
    for col, v in s.items():
                  ^^^^^^^
