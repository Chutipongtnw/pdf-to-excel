import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v17", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (คอลัมน์ใหม่)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF (ไฟล์ใหม่)", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        total_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            # 1. ดึงข้อความดิบเพื่อหา "รหัสครู" และ "ชื่อวิชา"
            raw_text = page.extract_text() or ""
            
            # --- หารหัสครู (วงเล็บข้างบน) ---
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # --- หาชื่อวิชา (บรรทัดล่าง: ชื่อวิชา ... จำนวน) ---
            subject_name = "N/A"
            subject_name_match = re.search(r'ชื่อวิชา\s+(.*?)\s+จำนวน', raw_text)
            if subject_name_match:
                subject_name = subject_name_match.group(1).strip()

            # 2. ดึงตาราง
            table = page.extract_table()
            if table:
                for row in table:
                    # ตรวจสอบว่ามีคอลัมน์ครบและเป็นแถวข้อมูลนักเรียน (เลข 5 หลัก)
                    if row and len(row) >= 8:
                        student_id = str(row[1]).replace('\n', '').strip()
                        
                        if student_id.isdigit() and len(student_id) == 5:
                            # จัดเก็บข้อมูลเรียงตามคอลัมน์ที่ต้องการ
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": str(row[3]).replace('\n', '').strip(),
                                "ชื่อวิชา": subject_name,
                                "ระดับชั้น": str(row[4]).replace('\n', '').strip(),
                                "เกรดปกติ": str(row[7]).replace('\n', '').strip() if row[7] else "",
                                "รหัสครู": teacher_id
                            })
            
            progress_bar.progress((i + 1) / total_pages)

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(df)} รายการ")
        
        # แสดงตัวอย่างตาราง
        st.dataframe(df, use_container_width=True)
        
        # ปุ่มดาวน์โหลด Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Student_Grades')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
            data=output.getvalue(),
            file_name="student_grades_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("ไม่พบข้อมูลที่ตรงตามเงื่อนไขในไฟล์ PDF นี้")
