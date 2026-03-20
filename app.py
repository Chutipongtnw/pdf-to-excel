import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันล้างอักขระที่ Excel ไม่รองรับแบบเด็ดขาด
def total_clean(value):
    if value is None:
        return ""
    s = str(value).replace('\n', ' ').strip()
    # เก็บเฉพาะตัวอักษรไทย (ก-ฮ, สระ, วรรณยุกต์), อังกฤษ, ตัวเลข, และเครื่องหมายพื้นฐาน
    # ลบพวก Control characters และอักขระแปลกปลอมที่ PDF มักใส่มาออกให้หมด
    s = re.sub(r'[^\u0e01-\u0e5b\u0020-\u007e]', '', s)
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
            
            # --- ปรับวิธีหารหัสครูใหม่ ---
            # สแกนหาตัวเลข 3 หลักในวงเล็บที่มักจะอยู่บรรทัดบนๆ ของหน้า
            teacher_id = "N/A"
            teacher_matches = re.findall(r'\((\d{3})\)', raw_text)
            if teacher_matches:
                teacher_id = teacher_matches[0] # เอารหัสตัวแรกที่เจอในหน้านั้น
            
            # ดึงตาราง
            table = page.extract_table()
            if table:
                for row in table:
                    # ตรวจสอบว่าแถวนี้มี "เลขประจำตัวนักเรียน" หรือไม่ (หลักการ: ต้องเป็นตัวเลข 5 หลัก)
                    # โดยปกติเลขประจำตัวจะอยู่คอลัมน์ที่ 2 (index 1)
                    if row and len(row) > 1:
                        student_id_raw = str(row[1]).replace('\n', '').strip()
                        if student_id_raw.isdigit() and len(student_id_raw) == 5:
                            
                            # ดึงข้อมูลตามลำดับคอลัมน์ [1]=เลขประจำตัว, [3]=รหัสวิชา, [4]=ชั้น, [7]=เกรดปกติ
                            all_data.append({
                                "เลขประจำตัวนักเรียน": total_clean(row[1]),
                                "รหัสวิชา": total_clean(row[3]),
                                "ระดับชั้น": total_clean(row[4]),
                                "เกรดปกติ": total_clean(row[7]),
                                "รหัสครู": total_clean(teacher_id)
                            })
            
            progress_bar.progress((i + 1) / num_pages)

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        try:
            # ใช้ xlsxwriter แทนเพื่อให้จัดการอักขระได้ดีขึ้น
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Grades')
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
                data=output.getvalue(),
                file_name="student_grades_clean.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Excel Error: {e}")
