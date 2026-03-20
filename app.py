import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_text(text):
    if not text: return ""
    return "".join(char for char in str(text) if char.isprintable())

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF (ธาตุนารายณ์วิทยา)", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียนบกพร่อง v6")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครู (Teacher ID)
            words = page.extract_words()
            teacher_id = "N/A"
            for idx, w in enumerate(words):
                if "ครู" in w['text']:
                    for next_w in words[idx+1 : idx+15]:
                        match = re.search(r'\((\d+)\)', next_w['text'])
                        if match:
                            teacher_id = match.group(1)
                            break
                    if teacher_id != "N/A": break

            # 2. ตั้งค่าการดึงตาราง (Table Settings)
            # ปรับเพิ่มความแม่นยำในการตรวจจับข้อความที่แยกบรรทัดกันในช่องเดียว
            table_settings = {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance": 3,
                "text_x_tolerance": 2,
                "text_y_tolerance": 2
            }
            
            table = page.extract_table(table_settings)
            
            if table:
                for row in table:
                    # ตรวจสอบจำนวนคอลัมน์และเลขประจำตัว 5 หลัก
                    if row and len(row) > 7:
                        # ลบช่องว่างและตัวอักษรแปลกๆ ออกก่อนเช็คตัวเลข
                        student_id = str(row[1]).replace('\n', '').replace(' ', '').strip()
                        
                        if student_id.isdigit() and len(student_id) == 5:
                            # ดึงรหัสวิชา (Index 3) - ใช้การล้างข้อมูลแบบถนอมตัวเลข
                            raw_subject = str(row[3]).replace('\n', '').strip()
                            
                            # ดึงระดับชั้น (Index 4) และ เกรดปกติ (Index 7)
                            raw_level = str(row[4]).replace('\n', '').strip()
                            raw_grade = str(row[7]).replace('\n', '').strip()
                            
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": raw_subject,
                                "ระดับชั้น": clean_text(raw_level),
                                "เกรดปกติ": clean_text(raw_grade),
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
            file_name="student_grades_v6.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
