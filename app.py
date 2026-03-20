import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันล้างชื่อวิชาให้สะอาด 100%
def clean_subject_name(text):
    if not text: return "N/A"
    
    # 1. ลบอักขระขยะที่ทำให้เกิดช่องว่าง (เช่น เ หรืออักขระ Unicode แปลกๆ)
    # ลบ \uf02d และอักขระในช่วง Unicode ที่มักเป็นตัวปัญหาใน PDF ภาษาไทย
    text = re.sub(r'[\uf000-\uf0ff]', '', text) 
    text = text.replace('เ', '์').replace('สตรเ', 'สตร์เ')
    
    # 2. ตัดส่วนที่ขึ้นต้นด้วย "จำนวน" หรือ "จํานวน" ออกให้หมด
    text = re.split(r'จ[ำา]นวน', text)[0]
    
    # 3. ลบช่องว่างที่ซ้ำซ้อน
    text = " ".join(text.split())
    
    return text.strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v20", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (ล้างชื่อวิชาสมบูรณ์แบบ)")

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
            
            # 2. หาชื่อวิชาจากบรรทัดสรุปท้ายหน้า
            subject_name = "N/A"
            lines = raw_text.split('\n')
            for line in lines:
                if "ชื่อวิชา" in line:
                    # แยกเอาเฉพาะส่วนหลังคำว่า "ชื่อวิชา"
                    raw_subject = line.split("ชื่อวิชา")[-1]
                    subject_name = clean_subject_name(raw_subject)
                    break

            # 3. ดึงข้อมูลจากตาราง
            table = page.extract_table()
            if table:
                for row in table:
                    # ตรวจสอบโครงสร้างแถว [1]=เลขประจำตัว, [3]=รหัสวิชา, [4]=ชั้น, [7]=เกรดปกติ
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
            df.to_excel(writer, index=False, sheet_name='Grades')
            worksheet = writer.sheets['Grades']
            # ปรับความกว้างคอลัมน์อัตโนมัติ
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, max_len)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="student_grades_v20.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
