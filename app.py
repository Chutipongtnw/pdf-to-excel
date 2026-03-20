import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_subject_name(text):
    if not text: return "N/A"
    
    # ลบ "จำนวน" และ "หน่วยการเรียน" ออกให้หมดแบบเด็ดขาด
    # ใช้ Regex ตัดทุกอย่างตั้งแต่คำว่า 'จำนวน' ไปจนจบสตริง
    text = re.split(r'\s*จ[ำา]นวน.*', text)[0]
    
    # ลบ "หน่วยการเรียน" หรือ "หน่วยกิต" ที่อาจหลงเหลือ (กรณีไม่มีคำว่าจำนวน)
    text = re.split(r'\s*\d+\.\d+\s*หน[ว่]ย.*', text)[0]
    
    return text.strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v27", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (ลบหน่วยกิตเด็ดขาด)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF (Native PDF)", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # 1. หารหัสครู (ดึงจากตัวเลขในวงเล็บ)
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # 2. หาชื่อวิชาจากบรรทัดท้ายหน้า (รหัสวิชา ... ชื่อวิชา ...)
            subject_name = "N/A"
            for line in raw_text.split('\n'):
                if "ชื่อวิชา" in line:
                    # ตัดเอาข้อความหลังคำว่า "ชื่อวิชา"
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = clean_subject_name(parts[-1])
                    break

            # 3. ดึงข้อมูลจากตาราง
            table = page.extract_table()
            if table:
                for row in table:
                    # [1]=เลขประจำตัวนักเรียน, [3]=รหัสวิชา, [4]=ชั้น, [7]=เกรดปกติ
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
            
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel", data=output.getvalue(), file_name="student_grades_final.xlsx")
