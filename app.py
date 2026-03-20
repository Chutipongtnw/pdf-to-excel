import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def universal_thai_cleaner(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน" และ "หน่วยการเรียน" ออกทันที 
    for divider in ["จำนวน", "จํานวน", "หน่วย"]:
        if divider in text:
            text = text.split(divider)[0]

    # 2. แปลงรหัส Unicode วรรณยุกต์ที่มักเป็นสี่เหลี่ยม (กลุ่ม \uf7xx) 
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์', '\uf71b': '์', 
        '\uf71c': '์', '\uf721': '์', '\uf72d': '์'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 3. ลบอักขระขยะ Unicode อื่นๆ ที่ไม่ใช่ตัวอักษรไทย/อังกฤษ/เลข [cite: 93]
    text = re.sub(r'[\uf000-\uf0ff]', '', text)

    # 4. กำจัดสระ "เ" ที่เบิ้ลมา (เเ) และกรณีที่ "เ" ผสมกับรหัสอื่นที่รูปร่างเหมือนกัน
    # ลบช่องว่างล่องหนก่อนแล้วทำการแทนที่ เเ ด้วย เ ตัวเดียว
    text = re.sub(r'\s+', '', text)
    text = text.replace('เเ', 'เ').replace('แแ', 'แ').replace('เ์', '์')
    
    # 5. แก้ไขคำเฉพาะที่มักสลับตำแหน่งหรือวรรณยุกต์หาย [cite: 33, 93, 160]
    corrections = {
        'คณิตศาสตร': 'คณิตศาสตร์',
        'พิ่มเติม': 'เพิ่มเติม',
        'ฟิสกิส์': 'ฟิสิกส์',
        'ทัศนศลิป': 'ทัศนศิลป์',
        'นาฎศิลป': 'นาฏศิลป์',
        'ศิลป์์': 'ศิลป์',
        'ศาสตร์์': 'ศาสตร์'
    }
    for wrong, right in corrections.items():
        text = text.replace(wrong, right)

    # 6. ขั้นตอนเด็ดขาด: ยุบวรรณยุกต์ที่ซ้ำกัน (เช่น ์์ -> ์)
    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)
    
    # 7. คืนค่าช่องว่าง 1 เคาะ หน้าตัวเลขท้ายชื่อวิชาเพื่อความสวยงาม [cite: 160]
    text = re.sub(r'(\d+)$', r' \1', text)

    return text.strip()

st.set_page_config(page_title="ระบบดึงข้อมูลอัจฉริยะ v35", layout="wide")
st.title("📂 ระบบดึงข้อมูล (กำจัดสระเบิ้ล เเ และตัว )")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # ดึงรหัสครูจากวงเล็บ [cite: 2, 8, 14]
            teacher_id = "N/A"
            t_match = re.search(r'\((\d+)\)', raw_text)
            if t_match: teacher_id = t_match.group(1)

            # ดึงชื่อวิชาจากบรรทัดท้ายหน้า [cite: 5, 12, 18]
            subject_name = "N/A"
            for line in raw_text.split('\n'):
                if "ชื่อวิชา" in line:
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = universal_thai_cleaner(parts[-1])
                    break

            # ดึงข้อมูลจากตาราง [cite: 4, 10, 16]
            table = page.extract_table()
            if table:
                for row in table:
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
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        st.success(f"ดึงข้อมูลสำเร็จ! พบ {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", output.getvalue(), "student_grades_v35.xlsx")
