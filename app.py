import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_subject_name(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน...หน่วยการเรียน" ออกให้ขาด
    text = re.split(r'จ[ำา]นวน', text)[0]
    
    # 2. แก้ไขลำดับสระและวรรณยุกต์ที่มักเพี้ยนจาก PDF (Glyph Correction)
    # เช่น 'ศลิป' -> 'ศิลป์', 'หนวย' -> 'หน่วย'
    corrections = {
        'ศลิป': 'ศิลป์',
        'หนวย': 'หน่วย',
        'คณิตศาสตร': 'คณิตศาสตร์',
        'พิ่มเติม': 'เพิ่มเติม',
        'ทัศนศลิป': 'ทัศนศิลป์',
        'ฟิสิกส': 'ฟิสิกส์',
        'ชีววิทยา': 'ชีววิทยา',
        'ภาษาไทย': 'ภาษาไทย'
    }
    
    # 3. ลบอักขระที่เป็นสี่เหลี่ยมหรือขยะ Unicode ออก (แต่เก็บสระไทยปกติไว้)
    text = re.sub(r'[\uf000-\uf0ff]', '', text)
    
    # 4. ลบช่องว่างทั้งหมดตามที่คุณต้องการเพื่อให้ข้อความติดกัน
    text = re.sub(r'\s+', '', text)
    
    # สั่งเปลี่ยนคำตามตาราง Correction ด้านบน
    for wrong, right in corrections.items():
        text = text.replace(wrong, right)
        
    return text.strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v23", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (แก้ไขสระเพี้ยนและวรรณยุกต์หาย)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        total_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # 1. หารหัสครู (วงเล็บข้างบน) [cite: 2, 8, 14]
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # 2. หาชื่อวิชาจากบรรทัดท้ายหน้า [cite: 5, 12, 18]
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
                    # [1]=เลขประจำตัวนักเรียน 
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
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="student_grades_v23.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
