import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ตารางถอดรหัสลับ (Mapping) จากไฟล์โรงเรียนธาตุนารายณ์
# หากเลขยังไม่ตรง เราจะปรับเลขในค่า 'digit' ตามผลลัพธ์ที่ได้
def decode_special_font(text):
    if not text: return ""
    mapping = {
        '􀀡': '0', '􀀢': '1', '􀀣': '2', '􀀤': '3', '􀀥': '4', 
        '􀀦': '5', '􀀧': '6', '􀀨': '7', '􀀩': '8', '􀀪': '9',
        '􀀞': '2', '􀀠': '6', '􀀚': '5', '􀀝': '8'
    }
    for char, digit in mapping.items():
        text = text.replace(char, digit)
    return "".join(char for char in text if char.isprintable() or char.isdigit()).strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v13", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (แก้ไขตัวเลขหาย & เกรดหาย)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. ดึงรหัสครู (หาจากวงเล็บและถอดรหัส)
            # ใช้พิกัดด้านบนของหน้าเพื่อความแม่นยำ
            header_text = page.within_bbox((0, 0, page.width, 150)).extract_text() or ""
            teacher_id = "N/A"
            teacher_match = re.search(r'\((\s*.*?\s*)\)', header_text)
            if teacher_match:
                teacher_id = decode_special_font(teacher_match.group(1))

            # 2. ดึงตารางแบบกำหนดคอลัมน์เอง (Fixed Columns) 
            # เพื่อป้องกันเกรดหายเนื่องจากเส้นตารางจาง
            table_settings = {
                "vertical_strategy": "text",
                "horizontal_strategy": "lines",
                "text_x_tolerance": 2,
            }
            
            table = page.extract_table(table_settings)
            
            if table:
                for row in table:
                    # ตารางปกติต้องมีอย่างน้อย 8 คอลัมน์ (ตามรูป m3.pdf)
                    if row and len(row) >= 8:
                        # [1] = เลขประจำตัว
                        student_id = decode_special_font(row[1]).replace(" ", "")
                        
                        if student_id.isdigit() and len(student_id) == 5:
                            # [3] = รหัสวิชา (รวม ท + เลข)
                            subject_id = decode_special_font(row[3])
                            # [4] = ระดับชั้น
                            level = decode_special_font(row[4])
                            # [7] = คะแนน (ถ้ามี)
                            # [8] = เกรดปกติ (ในบางหน้าอาจเป็น index 7 ถ้าตารางเบียด)
                            grade_idx = 8 if len(row) > 8 else 7
                            grade = decode_special_font(row[grade_idx])

                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": subject_id,
                                "ระดับชั้น": level,
                                "เกรดปกติ": grade,
                                "รหัสครู": teacher_id
                            })
            
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Final_Report')
        
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel", data=output.getvalue(), file_name="student_grades_v13.xlsx")
