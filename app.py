import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันถอดรหัสลับจาก Custom Font เป็นตัวเลขปกติ
def decode_pdf_font(text):
    if not text: return ""
    # ตารางแมปตัวเลขที่มักพบในไฟล์ PDF โรงเรียน (Identity-H Encoding)
    # ผมเพิ่มรหัสที่คุณแจ้งมา (􀀨􀀞􀀠) และรหัสลำดับอื่นๆ ให้แล้วครับ
    mapping = {
        '􀀡': '0', '􀀢': '1', '􀀣': '2', '􀀤': '3', '􀀥': '4', 
        '􀀦': '5', '􀀧': '6', '􀀨': '7', '􀀩': '8', '􀀪': '9',
        '􀀞': '2', '􀀠': '6'  # แก้ไขตามเคส (􀀨􀀞􀀠) -> 726
    }
    for char, digit in mapping.items():
        text = text.replace(char, digit)
    
    # ลบอักขระพิเศษอื่นๆ ที่ไม่ใช่ตัวเลข/ตัวอักษรไทย/อังกฤษ เพื่อความสะอาด
    return "".join(char for char in text if char.isprintable()).strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v11", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (แก้ไขรหัสวิชา & รหัสครู)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครู (ดึงข้อความดิบทั้งหน้ามาถอดรหัส)
            raw_text = page.extract_text() or ""
            # หาข้อความในวงเล็บ (มักจะเป็นรหัสครู)
            teacher_match = re.search(r'\((.*?)\)', raw_text)
            teacher_id = "N/A"
            if teacher_match:
                teacher_id = decode_pdf_font(teacher_match.group(1))

            # 2. ดึงตาราง
            table = page.extract_table()
            if table:
                for row in table:
                    if row and len(row) >= 8:
                        # เลขประจำตัวนักเรียน (Index 1)
                        student_id = decode_pdf_font(row[1])
                        
                        # ตรวจสอบว่าเป็นแถวนักเรียนจริง (เลข 5 หลัก)
                        if student_id.isdigit() and len(student_id) == 5:
                            # ดึงรหัสวิชา (Index 3) และถอดรหัสลับ
                            subject_id = decode_pdf_font(row[3])
                            
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": subject_id,
                                "ระดับชั้น": decode_pdf_font(row[4]),
                                "เกรดปกติ": decode_pdf_font(row[7]),
                                "รหัสครู": teacher_id
                            })
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        # สร้างไฟล์ Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Result')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="student_grades_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
