import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import unicodedata

def universal_thai_cleaner(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน" ออกทันที
    for divider in ["จำนวน", "จํานวน", "หน่วย"]:
        if divider in text:
            text = text.split(divider)[0]

    # 2. จัดระเบียบ Unicode (Normalization) เพื่อให้รหัสตัวอักษรที่หน้าตาเหมือนกันกลายเป็นรหัสเดียวกัน
    text = unicodedata.normalize('NFKC', text)

    # 3. ล้างรหัส Unicode กลุ่มพิเศษที่มักแฝงมาใน PDF ระบบโรงเรียน
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์', '\uf709': '',
        '\uf712': 'เ', '\uf713': 'เ',
        'อ': 'อ่าน', 'ข': 'ข้อ', 'ค': 'ค้น', 'ต': 'ต่อ', 'นํ': 'นำ', 'ผ': 'แผ่น'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 4. กำจัดอักขระขยะ Unicode ล่องหน
    text = re.sub(r'[\uf000-\uf0ff]', '', text)

    # 5. ลบช่องว่างทั้งหมดเพื่อให้ตัวอักษรมาชนกันสนิท
    text = re.sub(r'\s+', '', text)

    # 6. แก้ไขปัญหา "เเ" (สระเอสองตัว) แบบเจาะจงลำดับ
    # ไล่เช็คทุกรูปแบบที่อาจเกิดขึ้นได้
    text = text.replace('เเ', 'เ').replace('เเ', 'เ') # แทนที่ซ้ำๆ
    text = text.replace('เ้', '้').replace('เ์', '์') # ลบสระเอที่ติดมากับวรรณยุกต์ใน 'เพิ่มเติม' หรือ 'ศาสตร์'
    
    # 7. ใช้ Regex ยุบสระที่เบิ้ล (เ+ แ+ ะ+)
    text = re.sub(r'เ+', 'เ', text)
    text = re.sub(r'แ+', 'แ', text)

    # 8. ซ่อมคำเฉพาะที่สะกดผิด (กลุ่ม ศึกษา, วิทยาศาสตร์, เพิ่มเติม)
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
        'ข้อมููล': 'ข้อมูล',
        'คาสตร์': 'ศาสตร์'
    }
    for wrong, right in corrections.items():
        text = text.replace(wrong, right)

    # 9. ยุบวรรณยุกต์ซ้ำครั้งสุดท้าย
    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)
    
    # 10. เว้นช่องว่าง 1 เคาะ หน้าตัวเลขท้ายชื่อวิชา
    text = re.sub(r'(\d+)$', r' \1', text)

    return text.strip()

st.set_page_config(page_title="ระบบดึงข้อมูลอัจฉริยะ v41", layout="wide")
st.title("📂 ระบบดึงข้อมูล (ถอนรากถอนโคน สระ เเ เบิ้ล)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        total_p = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
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
            progress_bar.progress((i + 1) / total_p)

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        st.success(f"ดึงข้อมูลสำเร็จ! พบทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", output.getvalue(), "student_report_v41.xlsx")
