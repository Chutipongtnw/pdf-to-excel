import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_subject_name(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน/จํานวน...หน่วยการเรียน" ออกให้เด็ดขาด
    # ใช้ Regex ที่ดักจับทั้ง 'ำ' และ 'า' + 'นฤคหิต' รวมถึงช่องว่างรอบๆ
    text = re.split(r'\s*จ[ำา]นวน.*', text)[0]
    text = re.split(r'\s*จ\s*[ำา]\s*นวน.*', text)[0] # เผื่อช่องว่างกลางคำ
    
    # 2. แก้ไขรหัส Unicode พิเศษ (สระ/วรรณยุกต์ไทยที่มักกลายเป็นสี่เหลี่ยม)
    # ครอบคลุมชุดรหัสจาก PDF ระบบโรงเรียนที่ทำให้ "ฟิสิกส์" และ "นาฎศิลป์" เพี้ยน
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf707': '๊', '\uf708': '๋',
        '\uf70a': '่', '\uf70b': '้', '\uf70c': '๊', '\uf70d': '๋',
        '\uf70e': '์', '\uf710': '่', '\uf711': '้', '\uf712': '๊',
        '\uf713': '๋', '\uf714': '์', '\uf715': 'ิ', '\uf716': 'ุ',
        '\uf717': 'ู', '\uf718': 'ั', '\uf719': '็', '\uf71a': '์',
        '\uf71b': '์', '\uf71c': '์', 'เ ์': '์', 'เ': '์', 'ร ': 'ร์'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 3. ซ่อมคำเฉพาะที่คุณแจ้งว่าเพี้ยน (เช่น สระย้ายที่/ไม้เอกหาย)
    text = text.replace('ศลิป', 'ศิลป์').replace('หนวย', 'หน่วย').replace('ฟิสกิ', 'ฟิสิก')
    text = text.replace('ฟสิกส', 'ฟิสิกส์').replace('ศาสตร', 'ศาสตร์').replace('ศาสตร ', 'ศาสตร์')
    text = text.replace('นาฎศิลป', 'นาฏศิลป์').replace('ผลิตภัณฑ', 'ผลิตภัณฑ์')
    text = text.replace('สังคมศกึษา', 'สังคมศึกษา')
    
    # 4. ลบอักขระขยะ Unicode ที่ยังเหลือ (ตัวต้นเหตุสี่เหลี่ยม)
    text = re.sub(r'[\uf000-\uf0ff]', '', text)
    
    # 5. ลบช่องว่างที่ซ้ำซ้อนออก
    text = " ".join(text.split())
    
    return text.strip()

st.set_page_config(page_title="ระบบดึงข้อมูลสมบูรณ์ v29", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (ล้างชื่อวิชา & ตัดหน่วยกิต 100%)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # ดึงรหัสครู (เลขในวงเล็บ) [cite: 2, 8]
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # ดึงชื่อวิชาจากบรรทัดท้ายหน้า 
            subject_name = "N/A"
            for line in raw_text.split('\n'):
                if "ชื่อวิชา" in line:
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = clean_subject_name(parts[-1])
                    break

            # ดึงข้อมูลจากตาราง 
            table = page.extract_table()
            if table:
                for row in table:
                    if row and len(row) >= 8:
                        student_id = str(row[1]).replace('\n', '').strip()
                        # กรองเฉพาะเลขประจำตัว 5 หลัก 
                        if student_id.isdigit() and len(student_id) == 5:
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": str(row[3]).replace('\n', '').strip(),
                                "ชื่อวิชา": subject_name,
                                "ระดับชั้น": str(row[4]).replace('\n', '').strip(),
                                "เกรดปกติ": str(row[7]).replace('\n', '').strip() if row[7] else "",
                                "รหัสครู": teacher_id
                            })
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        st.success(f"ดึงข้อมูลสำเร็จ! พบ {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลด Excel", output.getvalue(), "report_v29.xlsx")
