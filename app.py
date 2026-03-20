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

    # 2. แปลงรหัส Unicode พิเศษที่มักเป็นสี่เหลี่ยม (กลุ่ม \uf7xx) [cite: 45, 93, 160]
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์', '\uf71b': '์', '\uf71c': '์'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 3. ลบอักขระขยะ Unicode อื่นๆ ที่ไม่ใช่ตัวอักษรไทย/อังกฤษ/เลข [cite: 93, 160, 290]
    text = re.sub(r'[\uf000-\uf0ff]', '', text)

    # 4. จัดการสระซ้ำที่พบบ่อย (เช่น เเ กลายเป็นสระแอ หรือ เ สองตัว) 
    # ลบช่องว่างล่องหนระหว่างสระก่อน แล้วค่อยยุบตัวซ้ำ
    text = text.replace('เเ', 'เ').replace('เ เ', 'เ')
    text = text.replace('แแ', 'แ').replace('ะะ', 'ะ').replace('รร', 'ร')

    # 5. ยุบวรรณยุกต์ซ้ำ (เช่น ์์ หรือ ่่) 
    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)

    # 6. ลบช่องว่างที่แทรกระหว่างตัวอักษรไทย (เช่น สังคมศึก ษา -> สังคมศึกษา) [cite: 112, 124, 284]
    text = re.sub(r'(?<=[ก-๙])\s+(?=[ก-๙])', '', text)
    
    # 7. ปรับช่องว่างให้เหลือ 1 เคาะ (คงไว้เฉพาะหน้าตัวเลข) [cite: 4, 10, 33, 208]
    text = re.sub(r'\s+', ' ', text).strip()

    # 8. ซ่อมคำเฉพาะที่มักเพี้ยนจากการวางตำแหน่งใน PDF [cite: 93, 160, 166, 308]
    text = text.replace('ศลิป', 'ศิลป์').replace('พิ่ม', 'เพิ่ม').replace('ศกึ', 'ศึก')
    text = text.replace('นาฎ', 'นาฏ').replace('ผลติ', 'ผลิต').replace('คณิตศาสตร', 'คณิตศาสตร์')
    
    # ตรวจสอบตัวการันต์ซ้ำซ้อนสุดท้าย [cite: 10, 28, 45, 160]
    text = text.replace('ร์์', 'ร์').replace('ล์์', 'ล์').replace('ต์์', 'ต์')

    return text.strip()

st.set_page_config(page_title="ระบบดึงข้อมูลอัจฉริยะ v33", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (แก้ไขสระเบิ้ล เเ)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        pages = pdf.pages
        for i, page in enumerate(pages):
            raw_text = page.extract_text() or ""
            
            # ดึงรหัสครู (วงเล็บด้านบน) [cite: 2, 8, 14, 20, 26, 31, 37, 43, 49, 58, 64, 71, 77, 84, 91, 97, 104, 110, 116, 122, 128, 134, 140, 146, 152, 158, 164, 170, 176, 182, 188, 194, 200, 206, 212, 218, 224, 231, 237, 243, 249, 255, 262, 269, 276, 282, 288, 294, 300, 306, 312, 318, 326]
            teacher_id = "N/A"
            t_match = re.search(r'\((\d+)\)', raw_text)
            if t_match: teacher_id = t_match.group(1)

            # ดึงชื่อวิชา [cite: 5, 12, 18, 24, 30, 35, 41, 47, 53, 61, 69, 75, 81, 89, 95, 101, 108, 114, 120, 126, 132, 138, 144, 149, 156, 162, 168, 174, 180, 186, 192, 198, 204, 210, 216, 222, 228, 235, 241, 247, 253, 259, 266, 273, 279, 286, 292, 298, 304, 310, 316, 322, 330]
            subject_name = "N/A"
            for line in raw_text.split('\n'):
                if "ชื่อวิชา" in line:
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = universal_thai_cleaner(parts[-1])
                    break

            # ดึงตาราง [cite: 4, 10, 16, 21, 28, 33, 38, 45, 50, 55, 67, 73, 79, 87, 93, 99, 106, 112, 118, 124, 130, 136, 142, 148, 154, 160, 166, 172, 178, 184, 189, 196, 202, 208, 214, 220, 226, 233, 239, 245, 251, 257, 264, 271, 278, 284, 290, 296, 302, 308, 314, 320, 328]
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
            progress_bar.progress((i + 1) / len(pages))

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        st.success(f"ดึงข้อมูลสำเร็จ! พบทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", output.getvalue(), "student_grades_v33.xlsx")
