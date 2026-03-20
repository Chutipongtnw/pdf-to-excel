import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_val(text):
    if text is None: return ""
    # ลบอักขระขยะแต่รักษาไทย อังกฤษ เลข และช่องว่าง
    return "".join(char for char in str(text) if char.isprintable()).strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v10", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน v10 (Deep Extraction)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครู: สแกนข้อความดิบหาตัวเลขในวงเล็บ (126)
            raw_text = page.extract_text() or ""
            teacher_id = "N/A"
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            if teacher_match:
                teacher_id = teacher_match.group(1)

            # 2. ดึงตารางโดยใช้โหมด 'text' เพื่อบังคับให้อ่านตัวเลขที่ฟอนต์มีปัญหา
            # และปรับพิกัดคอลัมน์ให้ตรงกับไฟล์ของโรงเรียนคุณ
            table = page.extract_table({
                "vertical_strategy": "text",
                "horizontal_strategy": "lines",
                "snap_tolerance": 4,
                "text_x_tolerance": 2
            })
            
            if table:
                for row in table:
                    # ตรวจสอบว่าแถวนี้มีข้อมูลนักเรียนหรือไม่ (เช็คคอลัมน์เลขประจำตัว)
                    if row and len(row) >= 8:
                        # พยายามดึงเลขประจำตัวจาก index 1
                        student_id = str(row[1]).replace('\n', '').replace(' ', '').strip()
                        
                        # ถ้าเจอเลขประจำตัว 5 หลัก ให้เริ่มเก็บข้อมูล
                        if student_id.isdigit() and len(student_id) == 5:
                            
                            # รหัสวิชา: รวมข้อความทุกบรรทัดในช่อง เพื่อให้ ท + 22101 มาครบ
                            subject_id = clean_val(row[3]).replace(' ', '')
                            
                            # ระดับชั้น: index 4
                            level = clean_val(row[4])
                            
                            # เกรดปกติ: ในไฟล์ของคุณมักอยู่ index 7 หรือ 8 (ลองดึงจากตัวท้ายๆ)
                            # เราจะเช็คค่าจากคอลัมน์ที่คาดว่าเป็นเกรด
                            grade = clean_val(row[7]) if len(row) == 8 else clean_val(row[8])

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
        st.success(f"ดึงข้อมูลสำเร็จ! พบทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        # ใช้ xlsxwriter เพื่อรองรับอักขระพิเศษ
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='StudentData')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="student_grades_v10.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("ไม่พบข้อมูล กรุณาลองอัปโหลดไฟล์ใหม่อีกครั้ง")
