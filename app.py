import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันล้างอักขระขยะ แต่เก็บ "ภาษาไทย + ภาษาอังกฤษ + ตัวเลข" ไว้
def total_clean(value):
    if value is None:
        return ""
    s = str(value).replace('\n', ' ').strip()
    # ปรับ Regex: ให้เก็บ ก-ฮ, สระ, วรรณยุกต์ (ไทย), A-Z, a-z, 0-9 และเครื่องหมาย -
    s = re.sub(r'[^\u0e00-\u0e7fA-Za-z0-9\- ]', '', s)
    return s

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF ผลการเรียน", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียนบกพร่อง (PDF to Excel)")
st.subheader("โรงเรียนธาตุนารายณ์วิทยา")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        num_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            # ดึงข้อความดิบเพื่อหารหัสครู
            raw_text = page.extract_text() or ""
            
            # --- ปรับวิธีหารหัสครู (Teacher ID) ให้ยืดหยุ่นขึ้น ---
            # มองหาตัวเลข 3 หลักที่อยู่ในวงเล็บ โดยอนุญาตให้มีช่องว่างได้ เช่น ( 124 ) หรือ (124)
            teacher_id = "N/A"
            teacher_match = re.search(r'\(\s*(\d{3})\s*\)', raw_text)
            if teacher_match:
                teacher_id = teacher_match.group(1)
            
            # ดึงตาราง
            table = page.extract_table()
            if table:
                for row in table:
                    # ตรวจสอบเลขประจำตัวนักเรียน 5 หลัก (มักอยู่คอลัมน์ index 1)
                    if row and len(row) > 1:
                        student_id_raw = str(row[1]).replace('\n', '').strip()
                        if student_id_raw.isdigit() and len(student_id_raw) == 5:
                            
                            all_data.append({
                                "เลขประจำตัวนักเรียน": total_clean(row[1]),
                                "รหัสวิชา": total_clean(row[3]), # เก็บตัวเลขได้แล้ว
                                "ระดับชั้น": total_clean(row[4]),
                                "เกรดปกติ": total_clean(row[7]),
                                "รหัสครู": teacher_id # ใส่รหัสที่ดึงได้จากหน้าเว็บ
                            })
            
            progress_bar.progress((i + 1) / num_pages)

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        # ใช้ Engine xlsxwriter เพื่อความเสถียร
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Data')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
            data=output.getvalue(),
            file_name="student_report_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
