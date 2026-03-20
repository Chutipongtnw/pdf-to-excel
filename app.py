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

    # 2. Normalization และแปลงรหัส Unicode พิเศษ
    text = unicodedata.normalize('NFKC', text)
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์', '\uf709': '',
        '\uf712': 'เ', '\uf713': 'เ',
        'อ': 'อ่าน', 'ข': 'ข้อ', 'ค': 'ค้น', 'ต': 'ต่อ', 'นํ': 'นำ', 'ผ': 'แผ่น'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 3. ลบอักขระขยะ Unicode และช่องว่างล่องหนทั้งหมด
    text = re.sub(r'[\u0000-\u001f\u007f-\u009f\uf000-\uf0ff\u200b\u00a0]', '', text)
    text = "".join(text.split())

    # 4. แก้ปัญหาคำเพี้ยนเฉพาะหน้า (ฟิสกิ -> ฟิสิก, ฟ่ง -> ฟัง, อ่านาน -> อ่าน)
    text = text.replace('ฟิสกิ', 'ฟิสิก')
    text = text.replace('ฟ่ง', 'ฟัง')
    text = text.replace('อ่านาน', 'อ่าน')
    text = text.replace('นำา', 'นำ')
    
    # 5. ยุบสระที่เบิ้ล (เเ, แแ, าา)
    for _ in range(2):
        text = text.replace('เเ', 'เ')
        text = text.replace('แแ', 'แ')
        text = text.replace('าา', 'า')
    
    # 6. ซ่อมคำกลุ่ม "ศึกษา" "เพิ่มเติม" "วิทยาศาสตร์"
    if 'พิ่มเติม' in text and 'เพิ่ม' not in text:
        text = text.replace('พิ่มเติม', 'เพิ่มเติม')

    corrections = {
        'ศกึ': 'ศึก',
        'วิทยาศาตร์': 'วิทยาศาสตร์',
        'นาฏศิลป1': 'นาฏศิลป์ 1',
        'นาฏศิลป2': 'นาฏศิลป์ 2',
        'ทัศนศิลป': 'ทัศนศิลป์',
        'ฟิสิกส': 'ฟิสิกส์',
        'คณิตศาสตร': 'คณิตศาสตร์',
        'ผลติ': 'ผลิต',
        'คาสตร์': 'ศาสตร์',
        'เ์': '์'
    }
    for wrong, right in corrections.items():
        text = text.replace(wrong, right)

    # 7. ยุบวรรณยุกต์ซ้ำ (์์, ่่, ้้)
    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)
    
    # 8. คืนค่าช่องว่าง 1 เคาะ หน้าตัวเลขท้ายชื่อวิชา
    text = re.sub(r'(\d+)$', r' \1', text)

    return text.strip()

st.set_config = st.set_page_config(page_title="ระบบดึงข้อมูลอัจฉริยะ v48", layout="wide")
st.title("📂 ระบบดึงข้อมูล (ซ่อมฟิสิกส์ และสระเบิ้ล)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF เพื่อรัน v48", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        total_p = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            teacher_id = "N/A"
            t_match = re.search(r'\((\d+)\)', raw_text)
            if t_match: teacher_id = t_match.group(1)

            subject_name = "N/A"
            for line in raw_text.split('\n'):
                if "ชื่อวิชา" in line:
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = universal_thai_cleaner(parts[-1])
                    break

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
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", output.getvalue(), "report_v48.xlsx")
