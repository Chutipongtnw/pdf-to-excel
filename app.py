import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันดึงเฉพาะตัวเลขและตัวอักษรไทย/อังกฤษพื้นฐาน
def final_clean(text):
    if not text: return ""
    # กำจัดตัวอักษรแปลกปลอมที่มองไม่เห็น แต่เก็บ ก-ฮ, สระ, A-Z, 0-9
    return "".join(char for char in str(text) if char.isprintable()).strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v7", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน v7 (Fixed Positions)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครูแบบเจาะจงพิกัดหัวกระดาษ (Top area)
            # เราจะค้นหาตัวเลขในวงเล็บจากข้อความทั้งหมดในหน้า
            text_content = page.extract_text() or ""
            teacher_match = re.search(r'\((\d+)\)', text_content)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # 2. ดึงตารางแบบ "Raw Objects" เพื่อเลี่ยงปัญหาสี่เหลี่ยม
            # ใช้พิกัดในการช่วยแยกคอลัมน์
            table = page.extract_table({
                "vertical_strategy": "text", 
                "horizontal_strategy": "text",
                "intersection_tolerance": 15
            })
            
            if table:
                for row in table:
                    if row and len(row) >= 8:
                        # คอลัมน์ที่ 2 (index 1) คือ เลขประจำตัว
                        student_id = str(row[1]).replace('\n', '').replace(' ', '')
                        
                        # กรองเฉพาะแถวที่เป็นข้อมูลนักเรียนจริง (เลข 5 หลัก)
                        if student_id.isdigit() and len(student_id) == 5:
                            # ดึงรหัสวิชา (Index 3) 
                            # หากยังเป็นสี่เหลี่ยม เราจะใช้การดึงแบบ 'char' เพื่อบังคับอ่านค่า
                            subject_id = final_clean(row[3])
                            
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": subject_id,
                                "ระดับชั้น": final_clean(row[4]),
                                "เกรดปกติ": final_clean(row[7]),
                                "รหัสครู": teacher_id
                            })
            
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ {len(df)} รายการ")
        
        # แสดงผลในตาราง Streamlit
        st.dataframe(df, use_container_width=True)
        
        # ส่งออกเป็น Excel (ใช้ Engine xlsxwriter)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Data')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="student_grades_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
