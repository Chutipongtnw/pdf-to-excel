import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_subject_name(text):
    if not text: return "N/A"
    
    # 1. ลบอักขระขยะ Unicode ที่มักกลายเป็นสี่เหลี่ยมหรือช่องว่าง (Private Use Area)
    # และตัวการันต์ที่แยกส่วน (เ )
    text = re.sub(r'[\uf000-\uf0ff]', '', text)
    text = text.replace('เ ', '์').replace('เ', '์').replace('ร ', 'ร์')
    
    # 2. แก้ไขคำเฉพาะที่มักผิดพลาดในไฟล์โรงเรียน
    text = text.replace('คณิตศาสตร ', 'คณิตศาสตร์')
    text = text.replace('คณิตศาสตร์พิ่มเติม', 'คณิตศาสตร์เพิ่มเติม')
    
    # 3. ตัด "จำนวน...หน่วยการเรียน" ออกให้หมด 
    # ใช้ Regex ค้นหาคำว่า 'จำนวน' หรือ 'จํานวน' และตัดทุกอย่างที่ตามมา
    text = re.split(r'จ[ำา]นวน', text)[0]
    
    # 4. จัดการช่องว่าง: ลบช่องว่างซ้ำซ้อนให้เหลือ 1 เคาะมาตรฐาน
    # แต่ถ้าเป็นชื่อวิชาแล้วตามด้วยตัวเลข (เช่น ภาษาไทย 2) ให้คงไว้ 1 เคาะ
    text = re.sub(r'\s+', ' ', text)
    
    return text.strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v21", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (ล้างข้อมูลชื่อวิชาแบบเด็ดขาด)")

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
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = clean_subject_name(parts[-1])
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
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 5
                worksheet.set_column(idx, idx, max_len)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="student_grades_v21.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
