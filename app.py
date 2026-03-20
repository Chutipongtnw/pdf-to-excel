import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันล้างอักขระขยะและช่องว่างที่ผิดปกติในชื่อวิชา
def clean_subject_name(text):
    if not text: return "N/A"
    # ลบอักขระพิเศษ (เช่น \uf02d) ที่ทำให้เกิดช่องว่างกลางคำว่า "เพิ่มเติม"
    text = text.replace('\uf02d', '').replace('\u200b', '')
    # ตัดเอาเฉพาะส่วนก่อนคำว่า "จำนวน"
    if "จำนวน" in text:
        text = text.split("จำนวน")[0]
    return text.strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v19", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (แก้ไขชื่อวิชาและตัดหน่วยกิต)")

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
                    # ตัดคำว่า "ชื่อวิชา" ออก แล้วส่งเข้าฟังก์ชันล้างข้อมูล
                    raw_subject = line.split("ชื่อวิชา")[-1]
                    subject_name = clean_subject_name(raw_subject)
                    break

            # 3. ดึงข้อมูลจากตาราง
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
            
            progress_bar.progress((i + 1) / total_pages)

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Student_Grades')
            worksheet = writer.sheets['Student_Grades']
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, max_len)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
            data=output.getvalue(),
            file_name="student_grades_clean_v19.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
