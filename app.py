import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_text(text):
    if not text: return ""
    # ลบอักขระขยะแต่เก็บไทย อังกฤษ เลข และช่องว่าง [cite: 4]
    return "".join(char for char in str(text) if char.isprintable())

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF (ธาตุนารายณ์วิทยา)", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียนบกพร่อง v5")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    # ใช้การตั้งค่าความละเอียดสูงเพื่อป้องกันตัวเลขหาย 
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครู (Teacher ID) โดยหาจากคำว่า "ครู" [cite: 2, 10, 17]
            words = page.extract_words()
            teacher_id = "N/A"
            for idx, w in enumerate(words):
                if "ครู" in w['text']:
                    # ค้นหาเลขในวงเล็บใน 15 คำถัดไป [cite: 2, 11, 17]
                    for next_w in words[idx+1 : idx+15]:
                        match = re.search(r'\((\d+)\)', next_w['text'])
                        if match:
                            teacher_id = match.group(1) [cite: 2, 11, 17]
                            break
                    if teacher_id != "N/A": break

            # 2. ดึงตารางโดยใช้ค่าพื้นฐานที่เสถียรที่สุด
            # แต่เพิ่มการจัดการข้อความในช่องให้รวมกัน (join_lines) 
            table_settings = {
                "vertical_strategy": "lines", 
                "horizontal_strategy": "lines",
                "snap_tolerance": 3,
                "join_tolerated_lines": 3
            }
            
            table = page.extract_table(table_settings)
            
            if table:
                for row in table:
                    # ตรวจสอบเลขประจำตัวนักเรียน 5 หลัก (มักอยู่คอลัมน์ index 1) [cite: 4, 12, 19]
                    if row and len(row) > 7:
                        student_id = str(row[1]).replace('\n', '').strip()
                        if student_id.isdigit() and len(student_id) == 5: [cite: 4, 12]
                            
                            # ดึงรหัสวิชา (Index 3) และรวมข้อความที่แยกบรรทัด [cite: 4, 12, 19]
                            raw_subject = str(row[3]).replace('\n', '').strip()
                            # ดึงเกรดปกติ (Index 7) 
                            raw_grade = str(row[7]).replace('\n', '').strip()
                            
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": raw_subject,
                                "ระดับชั้น": clean_text(row[4]),
                                "เกรดปกติ": raw_grade,
                                "รหัสครู": teacher_id
                            })
            
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Result')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="student_grades_v5.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
