import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def decode_pdf_font(text):
    if not text: return ""
    mapping = {
        '􀀡': '0', '􀀢': '1', '􀀣': '2', '􀀤': '3', '􀀥': '4', 
        '􀀦': '5', '􀀧': '6', '􀀨': '7', '􀀩': '8', '􀀪': '9',
        '􀀞': '2', '􀀠': '6'
    }
    for char, digit in mapping.items():
        text = text.replace(char, digit)
    return "".join(char for char in text if char.isprintable() or char.isdigit()).strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v13", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (Word Mapping Version)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครูจากข้อความดิบทั้งหน้า
            raw_text = page.extract_text() or ""
            teacher_id = "N/A"
            # ค้นหารูปแบบ (เลข 3 หลัก) หรือ (􀀨􀀞􀀠)
            teacher_match = re.search(r'\(([\d􀀡-􀀪\s]+)\)', raw_text)
            if teacher_match:
                teacher_id = decode_pdf_font(teacher_match.group(1))

            # 2. ดึงข้อมูลแบบสแกนคำ (Word-by-Word) เพื่อแก้ปัญหานามสกุลเบียดคอลัมน์
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            
            # จัดกลุ่มคำให้อยู่ในบรรทัดเดียวกัน
            lines = {}
            for w in words:
                y_pos = round(w['top'], 0)
                if y_pos not in lines: lines[y_pos] = []
                lines[y_pos].append(w)
            
            # วนลูปอ่านทีละบรรทัด
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda x: x['x0'])
                line_text = "".join([w['text'] for w in line_words])
                
                # เงื่อนไข: ถ้าบรรทัดนี้มี "เลขประจำตัวนักเรียน" (ตัวเลข 5 หลัก หรือรหัสสี่เหลี่ยมที่แปลงได้ 5 หลัก)
                # เราจะแกะข้อมูลจากพิกัด X (x0) ของแต่ละคำ
                for idx, w in enumerate(line_words):
                    clean_w = decode_pdf_font(w['text'])
                    if clean_w.isdigit() and len(clean_w) == 5 and 50 < w['x0'] < 100:
                        
                        student_id = clean_w
                        subject_id = ""
                        level = ""
                        grade = ""
                        
                        # วิ่งหาคำอื่นๆ ในบรรทัดเดียวกันตามพิกัด X
                        for other_w in line_words:
                            x_center = (other_w['x0'] + other_w['x1']) / 2
                            txt = decode_pdf_font(other_w['text'])
                            
                            # ช่วงพิกัดโดยประมาณ (ปรับตามโครงสร้างโรงเรียนธาตุนารายณ์)
                            if 380 < x_center < 480: subject_id += txt # คอลัมน์รหัสวิชา
                            if 480 < x_center < 540: level = txt      # คอลัมน์ระดับชั้น
                            if 630 < x_center < 700: grade = txt      # คอลัมน์เกรดปกติ
                        
                        if student_id:
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": subject_id,
                                "ระดับชั้น": level,
                                "เกรดปกติ": grade,
                                "รหัสครู": teacher_id
                            })

            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        st.success(f"ดึงข้อมูลสำเร็จ {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลด Excel", output.getvalue(), "student_grades.xlsx")
