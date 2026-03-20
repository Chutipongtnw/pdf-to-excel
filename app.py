import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_subject_name(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน...หน่วยการเรียน" ออกให้เด็ดขาดก่อนเป็นอันดับแรก
    text = re.split(r'จ[ำา]นวน', text)[0]
    
    # 2. ลบอักขระขยะทุกชนิดที่ไม่ใช่ ภาษาไทย ภาษาอังกฤษ และตัวเลข
    # วิธีนี้จะกำจัดสี่เหลี่ยมและช่องว่างประหลาดออกไปทั้งหมด
    text = re.sub(r'[^ก-๙a-zA-Z0-9]', '', text)
    
    # 3. ซ่อมคำเฉพาะที่มักจะเพี้ยนหลังจากการลบอักขระ
    # เช่น "คณิตศาสตร" (ที่การันต์หายไป) ให้คงสภาพที่อ่านออก
    text = text.replace('คณิตศาสตร', 'คณิตศาสตร์')
    text = text.replace('พิ่มเติม', 'เพิ่มเติม')
    
    return text.strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v22", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (ลบช่องว่างและสี่เหลี่ยมเด็ดขาด)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        total_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # 1. หารหัสครู (เลขในวงเล็บจากส่วนหัวหน้า) [cite: 2, 8, 14, 20, 26, 31, 37, 43, 49, 58, 64, 71, 77, 84, 91, 97, 104, 110, 116, 122, 128, 134, 140, 146, 152, 158, 164, 170, 176, 182, 188, 194, 200, 206, 212, 218, 224, 231, 237, 243, 249, 255, 262, 269, 276, 282, 288, 294, 300, 306, 312, 318, 326]
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # 2. หาชื่อวิชาจากบรรทัดสรุป (บรรทัดที่ขึ้นต้นด้วย "รหัสวิชา ... ชื่อวิชา") 
            subject_name = "N/A"
            lines = raw_text.split('\n')
            for line in lines:
                if "ชื่อวิชา" in line:
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        # ส่งเข้าฟังก์ชันลบช่องว่างและสิ่งแปลกปลอม
                        subject_name = clean_subject_name(parts[-1])
                    break

            # 3. ดึงข้อมูลจากตาราง [cite: 4, 10, 16, 21, 28, 33, 38, 45, 50, 55, 67, 73, 79, 87, 93, 99, 106, 112, 118, 124, 130, 136, 142, 148, 154, 160, 166, 172, 178, 184, 189, 196, 202, 208, 214, 220, 226, 233, 239, 245, 251, 257, 264, 271, 278, 284, 290, 296, 302, 308, 314, 320, 328]
            table = page.extract_table()
            if table:
                for row in table:
                    # [1]=เลขประจำตัวนักเรียน [cite: 4, 10, 16, 21, 28, 33, 38, 45, 50, 55, 67, 73, 79, 87, 93, 99, 106, 112, 118, 124, 130, 136, 142, 148, 154, 160, 166, 172, 178, 184, 189, 196, 202, 208, 214, 220, 226, 233, 239, 245, 251, 257, 264, 271, 278, 284, 290, 296, 302, 308, 314, 320, 328]
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
            file_name="student_grades_v22.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
