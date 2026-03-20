import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันถอดรหัสลับ (อ้างอิงจากที่คุณก๊อปปี้มาล่าสุด)
def decode_pdf_font(text):
    if not text: return ""
    # เพิ่มตาราง Mapping ตามที่คุณให้ข้อมูลมา
    mapping = {
        '􀀢': '2', '􀀣': '1', '􀀤': '0', '􀀩': '4', 
        '􀀨': '1', '􀀞': '2', '􀀠': '6',
        '􀀡': '0', '􀀥': '3', '􀀦': '4', '􀀧': '5', '􀀪': '9'
    }
    for char, digit in mapping.items():
        text = text.replace(char, digit)
    # เก็บเฉพาะตัวเลขและตัวอักษรที่พิมพ์ออกมาได้
    return "".join(char for char in text if char.isprintable() or char.isdigit()).strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v14", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (แก้ไขพิกัดคอลัมน์)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครู (สแกนบรรทัดบนสุด)
            words = page.extract_words()
            teacher_id = "N/A"
            for w in words:
                # พิกัดครูมักอยู่ช่วงบน (top < 150)
                if w['top'] < 150 and '(' in w['text']:
                    match = re.search(r'\((.*?)\)', w['text'])
                    if match:
                        teacher_id = decode_pdf_font(match.group(1))
                        break

            # 2. จัดกลุ่มคำตามบรรทัดเพื่อดึงข้อมูลตาราง
            lines = {}
            for w in words:
                # กรองคำที่อยู่ในช่วงตาราง (ปกติ top > 200)
                if w['top'] > 200:
                    y_pos = round(w['top'], 0)
                    if y_pos not in lines: lines[y_pos] = []
                    lines[y_pos].append(w)
            
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda x: x['x0'])
                
                for w in line_words:
                    clean_w = decode_pdf_font(w['text'])
                    # เงื่อนไข: หาเลขประจำตัวนักเรียน (พิกัด x0 อยู่ในช่วง 80-110)
                    if clean_w.isdigit() and len(clean_w) == 5 and 80 < w['x0'] < 110:
                        
                        student_id = clean_w
                        subject_id = ""
                        level = ""
                        grade = ""
                        
                        # วนหาข้อมูลอื่นในบรรทัดเดียวกันโดยใช้พิกัด X ที่แน่นอน
                        for other_w in line_words:
                            x_mid = (other_w['x0'] + other_w['x1']) / 2
                            txt = decode_pdf_font(other_w['text'])
                            
                            # ปรับพิกัดตามหน้ากระดาษ PDF โรงเรียน
                            if 350 < x_mid < 420: subject_id = txt # ช่องรหัสวิชา
                            if 420 < x_mid < 480: level = txt      # ช่องระดับชั้น
                            if 580 < x_mid < 650: grade = txt      # ช่องเกรดปกติ
                        
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
        st.download_button("📥 ดาวน์โหลด Excel", output.getvalue(), "student_grades_v14.xlsx")
