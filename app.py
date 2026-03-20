import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_subject_name(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน 'จำนวน' และ 'หน่วยการเรียน' ออกทันที (ใช้ Regex ครอบคลุมทุกสระอำ)
    text = re.split(r'\s*จ[ำา]นวน.*', text)[0]
    
    # 2. แก้ไขรหัส Unicode พิเศษที่ PDF มักใช้แทนสระและวรรณยุกต์ไทย (ป้องกันสี่เหลี่ยม)
    # ชุดรหัสเหล่านี้คือต้นเหตุที่ทำให้ "ฟสกิ ส" หรือ "ศาสตร" อ่านไม่ได้
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf707': '๊', '\uf708': '๋',
        '\uf70a': '่', '\uf70b': '้', '\uf70c': '๊', '\uf70d': '๋',
        '\uf70e': '์', '\uf710': '่', '\uf711': '้', '\uf712': '๊',
        '\uf713': '๋', '\uf714': '์', '\uf715': 'ิ', '\uf716': 'ุ',
        '\uf717': 'ู', '\uf718': 'ั', '\uf719': '็', '\uf71a': '์'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)
    
    # 3. ลบอักขระขยะ Unicode อื่นๆ ที่ยังเหลือ (สี่เหลี่ยมล่องหน)
    text = re.sub(r'[\uf000-\uf0ff]', '', text)
    
    # 4. ซ่อมคำเฉพาะที่มักแตกออกจากกัน
    text = text.replace('ศาสตร ', 'ศาสตร์').replace('ศาสตร', 'ศาสตร์')
    text = text.replace('ฟสกิ ส', 'ฟิสิกส์').replace('ฟสิกส', 'ฟิสิกส์')
    text = text.replace('ผลิตภัณฑ', 'ผลิตภัณฑ์').replace('สิกส', 'สิกส์')
    
    # 5. ลบช่องว่างส่วนเกินเพื่อให้ข้อความชิดกันตามที่ต้องการ
    text = re.sub(r'\s+', ' ', text) # ให้เหลือ 1 เคาะมาตรฐานก่อน
    
    return text.strip()

st.set_page_config(page_title="ระบบดึงข้อมูลสมบูรณ์ v28", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (แก้ไขชื่อวิชาและลบหน่วยกิต)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # ดึงรหัสครู [cite: 2, 8, 14, 20]
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            teacher_id = teacher_match.group(1) if teacher_match else "N/A"
            
            # ดึงชื่อวิชา [cite: 5, 12, 18, 24]
            subject_name = "N/A"
            for line in raw_text.split('\n'):
                if "ชื่อวิชา" in line:
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = clean_subject_name(parts[-1])
                    break

            # ดึงตารางข้อมูล [cite: 4, 10, 16, 21]
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
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"ดึงข้อมูลสำเร็จ! พบ {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลด Excel", output.getvalue(), "report_v28.xlsx")
