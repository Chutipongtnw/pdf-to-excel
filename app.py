import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ฟังก์ชันถอดรหัสลับ (Mapping) ที่แม่นยำที่สุดจากข้อมูลที่คุณให้มา
def decode_font(text):
    if not text: return ""
    mapping = {
        '􀀢': '2', '􀀣': '1', '􀀤': '0', '􀀩': '4', 
        '􀀨': '1', '􀀞': '2', '􀀠': '6', '􀀡': '0',
        '􀀥': '3', '􀀦': '4', '􀀧': '5', '􀀪': '9'
    }
    for char, digit in mapping.items():
        text = text.replace(char, digit)
    return "".join(char for char in text if char.isprintable()).strip()

st.set_page_config(page_title="ระบบกู้ข้อมูล PDF", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (Rescue Version)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # 1. ดึงคำทั้งหมดออกมาพร้อมพิกัด (X, Y)
            words = page.extract_words()
            
            # 2. หารหัสครู (หาเลข 3 หลักในวงเล็บ)
            teacher_id = "N/A"
            full_text = page.extract_text() or ""
            t_match = re.search(r'\((\s*[\d􀀡-􀀪]+\s*)\)', full_text)
            if t_match:
                teacher_id = decode_font(t_match.group(1))

            # 3. จัดกลุ่มคำตามบรรทัด (Y) - ยอมรับความคลาดเคลื่อนได้ 3 พิกเซล
            lines = {}
            for w in words:
                y = round(w['top'])
                found = False
                for existing_y in lines.keys():
                    if abs(y - existing_y) <= 3:
                        lines[existing_y].append(w)
                        found = True
                        break
                if not found: lines[y] = [w]

            # 4. วิเคราะห์ทีละบรรทัด
            for y in sorted(lines.keys()):
                line = sorted(lines[y], key=lambda x: x['x0'])
                # รวมคำในบรรทัดเพื่อเช็คเบื้องต้น
                line_str = "".join([w['text'] for w in line])
                
                # มองหา "เลขประจำตัวนักเรียน" (ต้องเป็นตัวเลข 5 หลักหลังถอดรหัส)
                for idx, word in enumerate(line):
                    decoded_w = decode_font(word['text'])
                    if decoded_w.isdigit() and len(decoded_w) == 5:
                        # พบแถวนักเรียนแล้ว! ต่อไปจะดึงข้อมูลตามตำแหน่ง X
                        student_id = decoded_w
                        subject_id = ""
                        level = ""
                        grade = ""
                        
                        # กวาดหาข้อมูลในบรรทัดเดียวกันตามพิกัด X (มาตราส่วน PDF ทั่วไป)
                        for other_w in line:
                            x = other_w['x0']
                            txt = decode_font(other_w['text'])
                            
                            # ช่วงพิกัด X (ปรับให้กว้างขึ้นเพื่อไม่ให้ข้อมูลหาย)
                            if 360 < x < 430: subject_id += txt # รหัสวิชา
                            if 440 < x < 500: level = txt       # ระดับชั้น
                            if 600 < x < 710: grade = txt       # เกรดปกติ
                        
                        # คัดกรองรหัสวิชา (ต้องมี ท หรือตัวเลข)
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
        st.success(f"ดึงข้อมูลสำเร็จ! พบทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลด Excel", output.getvalue(), "student_grades.xlsx")
    else:
        st.warning("ไม่พบข้อมูลนักเรียนในไฟล์นี้ กรุณาตรวจสอบว่าเลขประจำตัวใน PDF เป็นข้อความที่ดึงได้")
