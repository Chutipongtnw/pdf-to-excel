import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_text(text):
    if not text: return ""
    # ลบอักขระควบคุมที่ทำให้ Excel พัง แต่เก็บไทย อังกฤษ เลข และช่องว่าง
    return "".join(char for char in str(text) if char.isprintable())

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF (ธาตุนารายณ์วิทยา)", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียนบกพร่อง v4")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครู (Teacher ID) โดยหาคำว่า "ครู" แล้วดูคำถัดไปที่มีวงเล็บ
            words = page.extract_words()
            teacher_id = "N/A"
            for idx, w in enumerate(words):
                if "ครู" in w['text']:
                    # ค้นหาใน 10 คำถัดไปเพื่อหาเลขในวงเล็บ
                    for next_w in words[idx+1 : idx+10]:
                        match = re.search(r'\((\d+)\)', next_w['text'])
                        if match:
                            teacher_id = match.group(1)
                            break
                    if teacher_id != "N/A": break

            # 2. ดึงตารางโดยเน้นความละเอียดสูง
            table = page.extract_table({
                "vertical_strategy": "lines_relative",
                "horizontal_strategy": "lines_relative",
            })
            
            if table:
                for row in table:
                    # ตรวจสอบว่าแถวนี้มีเลขประจำตัว 5 หลักหรือไม่
                    if len(row) > 1:
                        student_id = str(row[1]).replace('\n', '').strip()
                        if student_id.isdigit() and len(student_id) == 5:
                            
                            # ดึงรหัสวิชา (Column Index 3) 
                            # แก้ปัญหาตัวเลขหายด้วยการดึงซ้ำและลบแค่ขึ้นบรรทัดใหม่
                            raw_subject = str(row[3]).replace('\n', '').strip()
                            
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": raw_subject,
                                "ระดับชั้น": clean_text(row[4]),
                                "เกรดปกติ": clean_text(row[7]),
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
            file_name="student_grades_v4.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
