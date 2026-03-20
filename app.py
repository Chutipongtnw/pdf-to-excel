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
    
    # 3. ใช้ Regex เก็บเฉพาะตัวอักษรไทย (รวมสระ/วรรณยุกต์) ตัวอังกฤษ และตัวเลข [cite: 5, 12, 18]
    # ลบทุกอย่างที่ไม่ใช่ ก-ฮ, ะ-ู, เ-์, a-z, A-Z, 0-9 และช่องว่าง
    # [ \u0E01-\u0E3A\u0E40-\u0E4E0-9a-zA-Z] คือกลุ่มตัวอักษรที่อนุญาต
    allowed_chars = re.findall(r'[ก-๙0-9a-zA-Z]', text)
    clean_text = "".join(allowed_chars)
    
    # 4. กรณีพิเศษ: ถ้าตัวสุดท้ายเป็น 'ร' (จาก ศาสตร์) ให้เติม '์' ถ้ามันหายไป
    if clean_text.endswith('คณิตศาสตร'):
        clean_text += '์'
        
    return clean_text.strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v24", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (ลบช่องว่างสี่เหลี่ยมรอยต่อตัวเลข)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        total_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or "" [cite: 4, 10, 16]
            
            # 1. หารหัสครู (วงเล็บข้างบน) [cite: 2, 8, 14]
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # 2. หาชื่อวิชาจากบรรทัดท้ายหน้า [cite: 5, 12, 18]
            subject_name = "N/A"
            lines = raw_text.split('\n')
            for line in lines:
                if "ชื่อวิชา" in line: [cite: 12, 18]
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = clean_subject_name(parts[-1])
                    break

            # 3. ดึงข้อมูลจากตาราง [cite: 4, 10, 16, 21]
            table = page.extract_table()
            if table:
                for row in table:
                    # [1]=เลขประจำตัวนักเรียน [cite: 4, 10, 16]
                    if row and len(row) >= 8:
                        student_id = str(row[1]).replace('\n', '').strip()
                        
                        if student_id.isdigit() and len(student_id) == 5: [cite: 4, 10, 16]
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id, [cite: 4]
                                "รหัสวิชา": str(row[3]).replace('\n', '').strip(), [cite: 4]
                                "ชื่อวิชา": subject_name,
                                "ระดับชั้น": str(row[4]).replace('\n', '').strip(), [cite: 4]
                                "เกรดปกติ": str(row[7]).replace('\n', '').strip() if row[7] else "", [cite: 4]
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
            file_name="student_grades_final_v24.xlsx"
        )
