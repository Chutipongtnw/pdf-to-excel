import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_val(text):
    if text is None: return ""
    # ลบอักขระขยะแต่รักษาไทย อังกฤษ เลข
    return "".join(char for char in str(text) if char.isprintable()).strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v8", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน v8 (พิกัดแม่นยำ)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครูจากข้อความดิบ (หาตัวเลขในวงเล็บ)
            full_text = page.extract_text() or ""
            teacher_match = re.search(r'\((\d+)\)', full_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # 2. ใช้พิกัดในการตัดแบ่งตาราง (คำนวณจากโครงสร้างโรงเรียนธาตุนารายณ์)
            # เราจะกำหนดจุดตัดคอลัมน์เองเพื่อไม่ให้ชื่อ/นามสกุลไหลไปปนรหัสวิชา
            table_settings = {
                "vertical_strategy": "explicit",
                "horizontal_strategy": "lines",
                "explicit_vertical_lines": [40, 75, 140, 360, 430, 490, 530, 580, 630, 750], 
                "snap_tolerance": 5,
            }
            
            table = page.extract_table(table_settings)
            
            if table:
                for row in table:
                    if len(row) >= 8:
                        student_id = clean_val(row[1]).replace(" ", "")
                        
                        # กรองแถวที่มีเลขประจำตัว 5 หลัก
                        if student_id.isdigit() and len(student_id) == 5:
                            # ดึงรหัสวิชา (Index 4) และ เกรดปกติ (Index 8) 
                            # หมายเหตุ: การใช้พิกัดอาจทำให้ index ขยับ ตรวจสอบจากผลลัพธ์อีกที
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": clean_val(row[4]),
                                "ระดับชั้น": clean_val(row[5]),
                                "เกรดปกติ": clean_text(row[8]) if len(row) > 8 else clean_val(row[7]),
                                "รหัสครู": teacher_id
                            })
            
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Data')
        
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel", data=output.getvalue(), 
                           file_name="student_grades_v8.xlsx")
