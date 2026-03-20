import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันถอดรหัสลับ (อัปเดตชุดรหัสตามที่ตรวจสอบจากไฟล์โรงเรียน)
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

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF v15", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (Auto-Coordinate Version)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. หารหัสครูจากข้อความในวงเล็บ
            raw_text = page.extract_text() or ""
            teacher_id = "N/A"
            teacher_match = re.search(r'\((\s*[\d􀀡-􀀪]+\s*)\)', raw_text)
            if teacher_match:
                teacher_id = decode_pdf_font(teacher_match.group(1))

            # 2. ดึงคำทั้งหมดพร้อมพิกัด
            words = page.extract_words()
            
            # ค้นหาพิกัด X ของหัวตารางเพื่อใช้เป็นจุดอ้างอิง
            col_map = {"student_id": 90, "subject_id": 400, "level": 450, "grade": 610}
            for w in words:
                txt = decode_pdf_font(w['text'])
                if "เลขประจำตัว" in txt: col_map["student_id"] = w['x0']
                if "รหัสวิชา" in txt: col_map["subject_id"] = w['x0']
                if "ระดับชั้น" in txt: col_map["level"] = w['x0']
                if "เกรดปกติ" in txt: col_map["grade"] = w['x0']

            # 3. จัดกลุ่มคำตามบรรทัด (Y)
            lines = {}
            for w in words:
                if w['top'] > 180: # กรองเฉพาะส่วนที่เป็นตาราง
                    y_pos = round(w['top'], 1)
                    found_line = False
                    for existing_y in lines.keys():
                        if abs(y_pos - existing_y) < 5: # บรรทัดเดียวกันถ้าระยะห่างไม่เกิน 5
                            lines[existing_y].append(w)
                            found_line = True
                            break
                    if not found_line: lines[y_pos] = [w]
            
            # 4. ประมวลผลแต่ละบรรทัด
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda x: x['x0'])
                
                # หาเลขประจำตัวนักเรียนก่อน
                student_id = ""
                for w in line_words:
                    clean_w = decode_pdf_font(w['text'])
                    # เช็คเลขประจำตัว (5 หลัก)
                    if clean_w.isdigit() and len(clean_w) == 5 and abs(w['x0'] - col_map["student_id"]) < 30:
                        student_id = clean_w
                        break
                
                if student_id:
                    subject_id = ""
                    level = ""
                    grade = ""
                    # กวาดหาข้อมูลอื่นในบรรทัดนั้นตามระยะคอลัมน์
                    for w in line_words:
                        clean_w = decode_pdf_font(w['text'])
                        # ถ้าระยะ X ตรงกับคอลัมน์ที่หาไว้
                        if abs(w['x0'] - col_map["subject_id"]) < 40: subject_id += clean_w
                        if abs(w['x0'] - col_map["level"]) < 30: level = clean_w
                        if abs(w['x0'] - col_map["grade"]) < 30: grade = clean_w
                    
                    # คัดกรองรหัสวิชา (ต้องมี ท หรือตัวเลข)
                    if student_id:
                        all_data.append({
                            "เลขประจำตัวนักเรียน": student_id,
                            "รหัสวิชา": subject_id if "ท" in subject_id or any(d.isdigit() for d in subject_id) else "",
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
        st.download_button("📥 ดาวน์โหลด Excel", output.getvalue(), "student_grades_v15.xlsx")
    else:
        st.error("ไม่พบข้อมูลในพิกัดที่กำหนด กรุณาตรวจสอบไฟล์ PDF อีกครั้ง")
