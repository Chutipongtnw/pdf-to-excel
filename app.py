import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_subject_name(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน" ออกก่อน
    text = re.split(r'จ[ำา]นวน', text)[0]
    
    # 2. แก้คำผิดที่เกิดจากการจัดเรียงสระใน PDF ให้ถูกต้อง
    text = text.replace('ศลิป', 'ศิลป์').replace('หนวย', 'หน่วย')
    text = text.replace('คณิตศาสตร', 'คณิตศาสตร์').replace('พิ่มเติม', 'เพิ่มเติม')
    
    # 3. เก็บเฉพาะตัวอักษรไทย อังกฤษ และตัวเลข (ลบสี่เหลี่ยมรอยต่อทิ้ง)
    allowed_chars = re.findall(r'[ก-๙0-9a-zA-Z]', text)
    clean_text = "".join(allowed_chars)
    
    # 4. ตรวจสอบการสะกด "คณิตศาสตร์" อีกครั้ง
    if "คณิตศาสตร" in clean_text and not clean_text.endswith('์'):
        clean_text = clean_text.replace('คณิตศาสตร', 'คณิตศาสตร์')
        
    return clean_text.strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v25", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (แก้ไข Syntax Error)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        total_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # 1. หารหัสครู (เลขในวงเล็บ)
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # 2. หาชื่อวิชาจากบรรทัดท้ายหน้า
            subject_name = "N/A"
            lines = raw_text.split('\n')
            for line in lines:
                if "ชื่อวิชา" in line:
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = clean_subject_name(parts[-1])
                    break

            # 3. ดึงข้อมูลจากตาราง
            table = page.extract_table()
            if table:
                for row in table:
                    if row and len(row) >= 8:
                        student_id = str(row[1]).replace('\n', '').strip()
                        
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
            df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="student_grades_final.xlsx"
        )
