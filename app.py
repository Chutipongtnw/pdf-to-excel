import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันสำหรับล้างตัวอักษรที่ Excel ไม่รองรับ (ป้องกัน IllegalCharacterError)
def clean_string(value):
    if isinstance(value, str):
        # ลบตัวอักษรที่ไม่ใช่ printable characters (ASCII Control Characters)
        return re.sub(r'[\x00-\x1f\x7f-\x9f]', '', value)
    return value

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF ผลการเรียน", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียนบกพร่อง (PDF to Excel)")
st.subheader("โรงเรียนธาตุนารายณ์วิทยา")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        num_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            
            # --- แก้ไขรหัสครู N/A: ปรับ Regex ให้ตรวจจับได้กว้างขึ้น ---
            # ค้นหาตัวเลขในวงเล็บที่อยู่ใกล้คำว่า "ครู" หรืออยู่เดี่ยวๆ ในบรรทัดที่คาดว่าเป็นชื่อครู
            teacher_match = re.search(r'\((\d{3})\)', text) 
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            table = page.extract_table()
            if table:
                for row in table[1:]:
                    # ตรวจสอบแถวข้อมูล (ต้องมีเลขประจำตัว 5 หลัก)
                    if row and len(row) > 1 and str(row[1]).strip().isdigit():
                        # ล้างค่าในแต่ละคอลัมน์ด้วยฟังก์ชัน clean_string
                        student_id = clean_string(row[1].replace('\n', '').strip())
                        subject_id = clean_string(row[3].replace('\n', '').strip())
                        level = clean_string(row[4].replace('\n', '').strip())
                        grade = clean_string(row[7].replace('\n', '').strip()) if row[7] else ""
                        
                        all_data.append({
                            "เลขประจำตัวนักเรียน": student_id,
                            "รหัสวิชา": subject_id,
                            "ระดับชั้น": level,
                            "เกรดปกติ": grade,
                            "รหัสครู": teacher_id
                        })
            
            progress_bar.progress((i + 1) / num_pages)

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(df)} รายการ")
        
        # แสดงตัวอย่างข้อมูล
        st.dataframe(df, use_container_width=True)
        
        # ปุ่มดาวน์โหลด Excel
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Student_Grades')
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
                data=output.getvalue(),
                file_name="student_grades_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการสร้างไฟล์ Excel: {e}")
