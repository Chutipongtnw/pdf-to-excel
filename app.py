import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF ผลการเรียน", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียนบกพร่อง (PDF to Excel)")
st.subheader("โรงเรียนธาตุนารายณ์วิทยา")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF (รองรับไฟล์หลายหน้า)", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        num_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            
            # 1. ดึงรหัสครู (หาตัวเลขในวงเล็บท้ายชื่อครู) 
            teacher_match = re.search(r'ครู\s+.*\((\d+)\)', text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # 2. ดึงตาราง
            table = page.extract_table()
            if table:
                for row in table[1:]: # ข้ามหัวตาราง 
                    # เช็คว่าเป็นแถวที่มีเลขประจำตัวนักเรียนจริงไหม (ต้องเป็นตัวเลข 5 หลัก) 
                    if row and len(row) > 1 and str(row[1]).strip().isdigit():
                        all_data.append({
                            "เลขประจำตัวนักเรียน": row[1].replace('\n', '').strip(),
                            "รหัสวิชา": row[3].replace('\n', '').strip(),
                            "ระดับชั้น": row[4].replace('\n', '').strip(),
                            "เกรดปกติ": row[7].replace('\n', '').strip() if row[7] else "",
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
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Student_Grades')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
            data=output.getvalue(),
            file_name="student_grades_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
