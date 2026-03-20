import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_subject_name(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน" และข้อความที่ตามมาทั้งหมดออกทันที
    # ใช้ Regex ที่รองรับสระอำทุกรูปแบบ เพื่อให้ส่วนหน่วยกิตหายไปแน่นอน
    text = re.split(r'\s*จ[ำา]นวน.*', text)[0]
    
    # 2. ลบเฉพาะ "ช่องว่างส่วนเกิน" หัว-ท้าย
    # ไม่ลบสระ ไม่ลบวรรณยุกต์ และไม่ลบช่องว่างระหว่างชื่อกับตัวเลข (เช่น ทัศนศิลป์ 1)
    return text.strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v26", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (Fixed Subject Name)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # ดึงรหัสครูจากวงเล็บ [cite: 2, 8, 14]
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # ดึงชื่อวิชาจากบรรทัดล่าง [cite: 5, 12, 17]
            subject_name = "N/A"
            for line in raw_text.split('\n'):
                if "ชื่อวิชา" in line:
                    # แยกข้อความหลังคำว่า "ชื่อวิชา" มาใช้ตรงๆ ไม่กรองตัวอักษรทิ้ง
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = clean_subject_name(parts[-1])
                    break

            # ดึงข้อมูลจากตาราง [cite: 4, 10, 21]
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
            
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel", data=output.getvalue(), file_name="student_grades.xlsx")
