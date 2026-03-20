import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def universal_thai_cleaner(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน" ออกทันที
    for divider in ["จำนวน", "จํานวน", "หน่วย"]:
        if divider in text:
            text = text.split(divider)[0]

    # 2. แมปปิ้งรหัส Unicode พิเศษที่ครอบคลุม 'ไม้เอก' และ 'ไม้โท' แบบต่างๆ (, )
    # รวมถึงรหัสที่มักกลายเป็นสี่เหลี่ยมในไฟล์ใหม่ของคุณ
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์', '\uf709': '',
        '\uf71b': '์', '\uf71c': '์', '\uf721': '์', '\uf72d': '์',
        'อ': 'อ่', 'ข': 'ข้อ', 'ค': 'ค้น', 'ต': 'ต่อ', 'นํ': 'นำ', 'ผ': 'แผ่น'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 3. กำจัด "เ" ที่ซ้ำซ้อน (เเ) โดยลบช่องว่างล่องหนออกก่อน
    text = re.sub(r'\s+', '', text)
    text = text.replace('เเ', 'เ').replace('เเ', 'เ') # ยุบ สระเอ สองตัว
    text = text.replace('แแ', 'แ').replace('เ์', '์')
    
    # 4. ลบอักขระขยะ Unicode ในช่วง Private Use Area (ตัวต้นเหตุสี่เหลี่ยม)
    text = re.sub(r'[\uf000-\uf0ff]', '', text)

    # 5. ซ่อมคำเฉพาะที่แจ้งว่าสะกดผิด (กลุ่ม ศึกษา, เพิ่มเติม, นาฏศิลป์)
    corrections = {
        'ศกึ': 'ศึก',
        'วิทยาศาตร์': 'วิทยาศาสตร์',
        'พิ่มเติม': 'เพิ่มเติม',
        'นาฏศิลป1': 'นาฏศิลป์ 1',
        'นาฏศิลป2': 'นาฏศิลป์ 2',
        'ทัศนศิลป': 'ทัศนศิลป์',
        'ฟิสิกส': 'ฟิสิกส์',
        'คณิตศาสตร': 'คณิตศาสตร์',
        'ผลติ': 'ผลิต',
        'ข้อมููล': 'ข้อมูล'
    }
    for wrong, right in corrections.items():
        text = text.replace(wrong, right)

    # 6. ยุบวรรณยุกต์ที่ซ้ำกัน (์์, ่่, ้้)
    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)
    
    # 7. คืนค่าช่องว่าง 1 เคาะ หน้าตัวเลขท้ายชื่อวิชา
    text = re.sub(r'(\d+)$', r' \1', text)

    return text.strip()

st.set_page_config(page_title="ระบบดึงข้อมูลอัจฉริยะ v37", layout="wide")
st.title("📂 ระบบดึงข้อมูล (ซ่อมวรรณยุกต์หาย & สระเบิ้ล)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF (ไฟล์ใหม่ที่มีปัญหา)", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        pages = pdf.pages
        for i, page in enumerate(pages):
            raw_text = page.extract_text() or ""
            
            # ดึงรหัสครู
            teacher_id = "N/A"
            t_match = re.search(r'\((\d+)\)', raw_text)
            if t_match: teacher_id = t_match.group(1)

            # ดึงชื่อวิชา
            subject_name = "N/A"
            for line in raw_text.split('\n'):
                if "ชื่อวิชา" in line:
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = universal_thai_cleaner(parts[-1])
                    break

            # ดึงตาราง
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
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", output.getvalue(), "student_report_v37.xlsx")
