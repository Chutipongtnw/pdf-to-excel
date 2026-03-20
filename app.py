import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันถอดรหัสสี่เหลี่ยม (Mapping Table)
# เนื่องจากคุณบอกว่าก๊อปไปวาง Word แล้วเพี้ยน เราจะดึงค่า Unicode ดิบมาแปลง
def decode_custom_font(text):
    if not text: return ""
    # ตารางแปลงค่า (ตัวอย่างการแมปรหัสที่มักพบใน PDF โรงเรียน)
    # หากรันแล้วเลขยังไม่ตรง เราจะปรับตารางนี้ตามค่าจริงที่ดึงได้
    mapping = {
        '􀀢': '1', '􀀣': '2', '􀀤': '3', '􀀥': '4', '􀀦': '5',
        '􀀧': '6', '􀀨': '7', '􀀩': '8', '􀀪': '9', '􀀡': '0',
        '􀀞': '2', '􀀠': '6' # เพิ่มตามที่คุณแจ้ง (􀀨􀀞􀀠) -> (726)
    }
    for char, digit in mapping.items():
        text = text.replace(char, digit)
    # ลบตัวอักษรพิเศษอื่นๆ ที่เหลือ
    return "".join(char for char in text if char.isprintable() or char in '0123456789')

st.title("📂 ระบบดึงข้อมูลผลการเรียน (แก้ไขตัวเลขสี่เหลี่ยม)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. ดึงรหัสครู (ดึงแบบ Raw เพื่อให้ติดรหัสสี่เหลี่ยมมาแปลง)
            raw_text = page.extract_text()
            teacher_id = "N/A"
            # หาข้อความในวงเล็บ
            found_id = re.search(r'\((.*?)\)', raw_text)
            if found_id:
                teacher_id = decode_custom_font(found_id.group(1))

            # 2. ดึงตาราง
            table = page.extract_table()
            if table:
                for row in table:
                    if row and len(row) > 7:
                        student_id = str(row[1]).replace('\n', '').strip()
                        # เลขประจำตัวถ้าเป็นสี่เหลี่ยมต้องแปลงก่อนเช็ค isdigit
                        student_id = decode_custom_font(student_id)
                        
                        if student_id.isdigit() and len(student_id) == 5:
                            # แปลงรหัสวิชาที่มีสี่เหลี่ยม
                            subject_raw = str(row[3]).replace('\n', '').strip()
                            subject_id = decode_custom_font(subject_raw)
                            
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": subject_id,
                                "ระดับชั้น": row[4].replace('\n', ' '),
                                "เกรดปกติ": row[7].replace('\n', ' '),
                                "รหัสครู": teacher_id
                            })
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data)
        st.dataframe(df)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลด Excel", output.getvalue(), "report.xlsx")
