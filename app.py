import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันล้างอักขระขยะแบบถนอมข้อมูล (เก็บทุกอย่างที่เป็นตัวอักษรและตัวเลข)
def super_clean(value):
    if value is None:
        return ""
    # แทนที่การขึ้นบรรทัดใหม่ด้วยช่องว่าง
    s = str(value).replace('\n', ' ').strip()
    # ลบเฉพาะพวก Control Characters (อักขระสั่งการระบบ) ที่ Excel ไม่ชอบ
    # แต่จะเก็บ ภาษาไทย ภาษาอังกฤษ และ ตัวเลข ไว้ทั้งหมด
    s = "".join(char for char in s if char.isprintable())
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
            # ดึงข้อความดิบทั้งหมดในหน้า
            raw_text = page.extract_text() or ""
            
            # --- วิธีหารหัสครูแบบใหม่ (ค้นหาตัวเลข 3 หลักที่อยู่ต้นๆ หน้า) ---
            teacher_id = "N/A"
            # หาตัวเลข 3 หลัก (รหัสครู) เช่น 124, 203, 102
            # เราจะหาตัวเลข 3 หลักที่อาจจะอยู่ในวงเล็บหรือไม่ก็ได้
            teacher_matches = re.findall(r'\(?\s*(\d{3})\s*\)?', raw_text)
            if teacher_matches:
                # โดยปกติรหัสครูจะปรากฏเป็นลำดับต้นๆ ของข้อความในหน้า 
                teacher_id = teacher_matches[0]
            
            # ดึงตาราง
            table = page.extract_table()
            if table:
                for row in table:
                    # ตรวจสอบแถวข้อมูลนักเรียน (เช็คจากเลขประจำตัว 5 หลัก) 
                    if row and len(row) > 1:
                        # ลบช่องว่างและขึ้นบรรทัดใหม่ก่อนเช็คว่าเป็นตัวเลขไหม
                        student_id_check = str(row[1]).replace('\n', '').strip()
                        
                        if student_id_check.isdigit() and len(student_id_check) == 5:
                            all_data.append({
                                "เลขประจำตัวนักเรียน": super_clean(row[1]),
                                "รหัสวิชา": super_clean(row[3]), # ดึงค่าดิบ ไม่กรองตัวเลขออกแล้ว
                                "ระดับชั้น": super_clean(row[4]),
                                "เกรดปกติ": super_clean(row[7]),
                                "รหัสครู": teacher_id
                            })
            
            progress_bar.progress((i + 1) / num_pages)

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(df)} รายการ")
        
        # แสดงผลในหน้าเว็บ
        st.dataframe(df, use_container_width=True)
        
        # ส่วนการดาวน์โหลด
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Data')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
            data=output.getvalue(),
            file_name="student_report_v3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
