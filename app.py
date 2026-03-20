import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# เปลี่ยนชื่อฟังก์ชันให้ตรงกันทั้งหมด
def clean_val(text):
    if text is None: return ""
    # ลบอักขระควบคุมแต่รักษาไทย อังกฤษ เลข และช่องว่าง
    return "".join(char for char in str(text) if char.isprintable()).strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v9", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน v9 (ธาตุนารายณ์วิทยา)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        num_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครูจากข้อความดิบ (หาตัวเลขในวงเล็บ)
            full_text = page.extract_text() or ""
            teacher_match = re.search(r'\((\d+)\)', full_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # 2. ตั้งค่าการดึงตารางแบบเน้นเส้น (เหมาะกับตารางที่มีเส้นขอบชัดเจน)
            table_settings = {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance": 3,
            }
            
            table = page.extract_table(table_settings)
            
            if table:
                for row in table:
                    # ตารางโรงเรียนคุณมีประมาณ 9 คอลัมน์
                    if row and len(row) >= 8:
                        # ลบช่องว่างและขึ้นบรรทัดใหม่ในเลขประจำตัว
                        student_id = str(row[1]).replace('\n', '').replace(' ', '').strip()
                        
                        # กรองเฉพาะแถวที่มีเลขประจำตัว 5 หลัก
                        if student_id.isdigit() and len(student_id) == 5:
                            # ดึงข้อมูลตามลำดับคอลัมน์จริง
                            # [1]=เลขประจำตัว, [3]=รหัสวิชา, [4]=ระดับชั้น, [8]=เกรดปกติ (หรือ [7])
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": clean_val(row[3]),
                                "ระดับชั้น": clean_val(row[4]),
                                "เกรดปกติ": clean_val(row[8]) if len(row) > 8 else clean_val(row[7]),
                                "รหัสครู": teacher_id
                            })
            
            progress_bar.progress((i + 1) / num_pages)

    # แสดงผลถ้ามีข้อมูล
    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ! พบทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        # เตรียมไฟล์ Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Data')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="student_grades_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("ไม่พบข้อมูลในตาราง กรุณาตรวจสอบว่าไฟล์ PDF มีตารางและเป็น Native PDF")
