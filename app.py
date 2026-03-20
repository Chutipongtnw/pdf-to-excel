import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันถอดรหัสลับที่ละเอียดขึ้น
def decode_pdf_font(text):
    if not text: return ""
    mapping = {
        '􀀡': '0', '􀀢': '1', '􀀣': '2', '􀀤': '3', '􀀥': '4', 
        '􀀦': '5', '􀀧': '6', '􀀨': '7', '􀀩': '8', '􀀪': '9',
        '􀀞': '2', '􀀠': '6'
    }
    for char, digit in mapping.items():
        text = text.replace(char, digit)
    # เก็บเฉพาะตัวเลขและตัวอักษรที่พิมพ์ออกมาได้
    return "".join(char for char in text if char.isprintable() or char.isdigit()).strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v12", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (Deep Inspection)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครู: สแกนแบบละเอียดจาก Object ข้อความ
            teacher_id = "N/A"
            page_text = page.extract_text() or ""
            # ลองหาในวงเล็บก่อน
            teacher_match = re.search(r'\((.*?)\)', page_text)
            if teacher_match:
                teacher_id = decode_pdf_font(teacher_match.group(1))
            
            # 2. ดึงตารางโดยเพิ่มค่า tolerance เพื่อรวมข้อความที่เยื้องบรรทัด
            table_settings = {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_y_tolerance": 10, # ช่วยให้ดึงข้อความที่เยื้องขึ้นลงได้
                "text_x_tolerance": 3,
                "text_y_tolerance": 3,
            }
            
            table = page.extract_table(table_settings)
            
            if table:
                for row in table:
                    # คอลัมน์ที่ 2 (index 1) คือ เลขประจำตัว
                    if row and len(row) >= 8:
                        student_id = decode_pdf_font(row[1]).replace(" ", "")
                        
                        if student_id.isdigit() and len(student_id) == 5:
                            # ดึงรหัสวิชา (index 3)
                            # หากดึงแล้วมีแต่ ท เราจะพยายามค้นหา Object ที่พิกัดเดียวกัน
                            subject_id = decode_pdf_font(row[3])
                            
                            # ดึงระดับชั้น (index 4) และ เกรดปกติ (index 7 หรือ 8)
                            level = decode_pdf_font(row[4])
                            grade = decode_pdf_font(row[7]) if len(row) == 8 else decode_pdf_font(row[8])
                            
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": subject_id,
                                "ระดับชั้น": level,
                                "เกรดปกติ": grade,
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
        
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel", data=output.getvalue(), file_name="student_grades_v12.xlsx")
