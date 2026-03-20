import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def universal_thai_cleaner(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน" ออกก่อน (ตามความต้องการเดิม)
    for divider in ["จำนวน", "จํานวน", "หน่วย"]:
        if divider in text:
            text = text.split(divider)[0]

    # 2. แปลงรหัส Unicode พิเศษที่มักเป็นสี่เหลี่ยม (รหัสกลุ่ม \uf7xx)
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์', '\uf71b': '์', '\uf71c': '์'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 3. ลบอักขระขยะ Unicode ที่ไม่ใช่ตัวอักษรไทย/อังกฤษ/เลข (ป้องกันสี่เหลี่ยมที่เหลือ)
    text = re.sub(r'[\uf000-\uf0ff]', '', text)

    # 4. ลบช่องว่างที่แทรกระหว่างตัวอักษรไทย (เช่น สังคมศึก ษา -> สังคมศึกษา)
    # แต่จะเว้นช่องว่างไว้ 1 เคาะ หากเป็นจุดที่เชื่อมกับตัวเลข (เช่น ฟิสิกส์ 5)
    text = re.sub(r'(?<=[ก-๙])\s+(?=[ก-๙])', '', text)
    text = re.sub(r'\s+', ' ', text).strip()

    # 5. แก้ไขตัวอักษรซ้ำ (Duplicate Characters) เช่น ์์ หรือ เเ
    # ใช้ Regex ตรวจจับตัวที่ซ้ำติดกันตั้งแต่ 2 ตัวขึ้นไปให้เหลือตัวเดียว
    text = re.sub(r'(.)\1+', r'\1', text)
    
    # 6. ซ่อมแซมคำเฉพาะที่มีปัญหาเรื่องการวางตำแหน่งวรรณยุกต์ผิดที่
    text = text.replace('ศลิป', 'ศิลป์').replace('พิ่ม', 'เพิ่ม').replace('ศกึ', 'ศึก')
    text = text.replace('นาฎ', 'นาฏ').replace('ผลติ', 'ผลิต')

    return text.strip()

st.set_page_config(page_title="ระบบดึงข้อมูลอัจฉริยะ v32", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (Universal Cleaner)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF เพื่อเริ่มการดึงข้อมูล", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        pages = pdf.pages
        for i, page in enumerate(pages):
            raw_text = page.extract_text() or ""
            
            # ดึงรหัสครู (วงเล็บด้านบน)
            teacher_id = "N/A"
            t_match = re.search(r'\((\d+)\)', raw_text)
            if t_match: teacher_id = t_match.group(1)

            # ดึงชื่อวิชาจากบรรทัด "ชื่อวิชา"
            subject_name = "N/A"
            for line in raw_text.split('\n'):
                if "ชื่อวิชา" in line:
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = universal_thai_cleaner(parts[-1])
                    break

            # ดึงข้อมูลจากตาราง
            table = page.extract_table()
            if table:
                for row in table:
                    # ตรวจสอบว่ามีข้อมูลเลขประจำตัว (คอลัมน์ที่ 2)
                    if row and len(row) >= 8:
                        s_id = str(row[1]).replace('\n', '').strip()
                        if s_id.isdigit() and len(s_id) == 5:
                            all_data.append({
                                "เลขประจำตัวนักเรียน": s_id,
                                "รหัสวิชา": str(row[3]).replace('\n', '').strip(),
                                "ชื่อวิชา": subject_name,
                                "ระดับชั้น": str(row[4]).replace('\n', '').strip(),
                                "เกรดปกติ": str(row[7]).replace('\n', '').strip() if row[7] else "",
                                "รหัสครู": teacher_id
                            })
            progress_bar.progress((i + 1) / len(pages))

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        st.success(f"ดึงข้อมูลสำเร็จ! พบทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", output.getvalue(), "student_grades_report.xlsx")
