import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v18", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (แก้ไขการดึงชื่อวิชา)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        total_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            # ดึงข้อความดิบทั้งหมดในหน้านั้น
            raw_text = page.extract_text() or ""
            
            # 1. หารหัสครู (ตัวเลขในวงเล็บ)
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # 2. หาชื่อวิชา (ปรับปรุง Regex ให้ดึงข้อมูลได้ดีขึ้น)
            subject_name = "N/A"
            # ค้นหาข้อความระหว่าง "ชื่อวิชา" และ "จำนวน" หรือ "หน่วยการเรียน"
            subject_name_match = re.search(r'ชื่อวิชา\s+(.*?)\s+(จำนวน|หน่วยการเรียน)', raw_text)
            if subject_name_match:
                subject_name = subject_name_match.group(1).strip()
            # กรณีพิเศษ: ถ้ายังไม่เจอ ให้ลองหาบรรทัดที่มีคำว่า "ชื่อวิชา"
            elif "ชื่อวิชา" in raw_text:
                lines = raw_text.split('\n')
                for line in lines:
                    if "ชื่อวิชา" in line:
                        # ตัดเอาข้อความหลังคำว่า ชื่อวิชา
                        parts = line.split("ชื่อวิชา")
                        if len(parts) > 1:
                            subject_name = parts[1].split("จำนวน")[0].strip()
                            break

            # 3. ดึงข้อมูลจากตาราง
            table = page.extract_table()
            if table:
                for row in table:
                    if row and len(row) >= 8:
                        student_id = str(row[1]).replace('\n', '').strip()
                        
                        # ตรวจสอบว่าเป็นแถวข้อมูลนักเรียน (เลขประจำตัว 5 หลัก)
                        if student_id.isdigit() and len(student_id) == 5:
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
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Student_Grades')
            # ปรับความกว้างคอลัมน์อัตโนมัติ (เพิ่มความสวยงาม)
            worksheet = writer.sheets['Student_Grades']
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, max_len)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
            data=output.getvalue(),
            file_name="student_grades_report_v18.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
