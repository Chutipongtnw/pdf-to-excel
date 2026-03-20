import streamlit as st
import pdfplumber
import pandas as pd
import io

def decode_font(text):
    if not text: return ""
    # ตารางถอดรหัสลับ Identity-H ตามข้อมูลจริงจากไฟล์คุณ
    mapping = {
        '􀀞': '2', '􀀠': '6', '􀀨': '1', '􀀩': '4', 
        '􀀢': '2', '􀀣': '1', '􀀤': '0', '􀀡': '0',
        '􀀥': '3', '􀀦': '4', '􀀧': '5', '􀀪': '9'
    }
    for char, digit in mapping.items():
        text = text.replace(char, digit)
    return "".join(char for char in text if char.isprintable()).strip()

st.title("📂 ระบบดึงข้อมูลเวอร์ชันเสถียร (Fixed Position)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            # 1. ดึงรหัสครู (ระบุพิกัดส่วนหัวกระดาษ)
            header_area = page.crop((0, 0, page.width, 150))
            header_text = header_area.extract_text() or ""
            teacher_id = "N/A"
            import re
            t_match = re.search(r'\((\s*[\d􀀡-􀀪]+\s*)\)', header_text)
            if t_match:
                teacher_id = decode_font(t_match.group(1))

            # 2. ดึงตาราง (กำหนดตำแหน่งคอลัมน์ตายตัว ไม่ให้ชื่อไหลไปทับรหัสวิชา)
            # พิกัด x0 ของแต่ละช่อง: เลขประจำตัว(85), รหัสวิชา(380), ระดับชั้น(450), เกรด(615)
            table = page.extract_table({
                "vertical_strategy": "explicit",
                "explicit_vertical_lines": [80, 135, 375, 445, 500, 610, 680],
                "horizontal_strategy": "lines"
            })
            
            if table:
                for row in table:
                    if row and len(row) >= 5:
                        s_id = decode_font(row[0]) # เลขประจำตัว
                        if s_id.isdigit() and len(s_id) == 5:
                            all_data.append({
                                "เลขประจำตัวนักเรียน": s_id,
                                "รหัสวิชา": decode_font(row[1]), # ช่องรหัสวิชา
                                "ระดับชั้น": decode_font(row[2]), # ช่องระดับชั้น
                                "เกรดปกติ": decode_font(row[4]), # ช่องเกรดปกติ
                                "รหัสครู": teacher_id
                            })

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        st.dataframe(df)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลด Excel", output.getvalue(), "final_report.xlsx")
