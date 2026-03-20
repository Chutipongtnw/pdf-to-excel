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

    # 2. Normalization ขั้นสูงสุด
    text = unicodedata.normalize('NFKC', text)
    
    # 3. Unicode Mapping และซ่อมคำเฉพาะหน้า (ค้นคว้า, ด้วย)
    text = text.replace('ค', 'ค้น').replace('คว', 'คว้า')
    text = text.replace('ด', 'ด้')
    
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์', '\uf709': '',
        '\uf712': 'เ', '\uf713': 'เ',
        'อ': 'อ่าน', 'ข': 'ข้อ', 'ค': 'ค้น', 'ต': 'ต่อ', 'นํ': 'นำ', 'ผ': 'แผ่น'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 4. ล้างรหัส Unicode ขยะและช่องว่างทั้งหมด
    text = re.sub(r'[\u0000-\u001f\u007f-\u009f\uf000-\uf0ff\u200b\u00a0]', '', text)
    text = "".join(text.split())

    # 5. แก้ไขพยัญชนะเบิ้ลและสระกระโดด (ศลิป -> ศิลป, ฟิสกิ -> ฟิสิก)
    text = text.replace('ศลิป', 'ศิลป์') 
    text = text.replace('ฟิสกิ', 'ฟิสิก') 
    text = text.replace('ต่ออ', 'ต่อ')
    text = text.replace('ค้นน', 'ค้น')
    text = text.replace('ข้ออ', 'ข้อ')
    text = text.replace('แผ่นน', 'แผ่น')
    text = text.replace('นำา', 'นำ')
    text = text.replace('ฟ่ง', 'ฟัง')
    
    # --- จุดที่แก้ไขเพิ่มเติมสำหรับวิชาภาษาอังกฤษ ---
    text = text.replace('อ่านาน', 'อ่าน') 
    # ------------------------------------------
    
    # 6. ยุบสระที่เบิ้ล (เเ, แแ, าา)
    for _ in range(2):
        text = text.replace('เเ', 'เ')
        text = text.replace('แแ', 'แ')
        text = text.replace('าา', 'า')
    
    # 7. คลังซ่อมคำมาตรฐาน
    if 'พิ่มเติม' in text and 'เพิ่ม' not in text:
        text = text.replace('พิ่มเติม', 'เพิ่มเติม')

    corrections = {
        'ศกึ': 'ศึก',
        'วิทยาศาตร์': 'วิทยาศาสตร์',
        'นาฏศิลป1': 'นาฏศิลป์ 1',
        'นาฏศิลป2': 'นาฏศิลป์ 2',
        'ทัศนศิลป1': 'ทัศนศิลป์ 1',
        'ทัศนศิลป2': 'ทัศนศิลป์ 2',
        'ฟิสิกส': 'ฟิสิกส์',
        'คณิตศาสตร': 'คณิตศาสตร์',
        'ผลติ': 'ผลิต',
        'คาสตร์': 'ศาสตร์',
        'วดีโอ': 'วิดีโอ',
        'เ์': '์'
    }
    for wrong, right in corrections.items():
        text = text.replace(wrong, right)

    # 8. ยุบวรรณยุกต์ซ้ำ
    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)
    
    # 9. คืนค่าช่องว่าง 1 เคาะ หน้าตัวเลขท้ายชื่อวิชา
    text = re.sub(r'(\d+)$', r' \1', text)

    return text.strip()

# ส่วนที่เหลือของ Streamlit เหมือนเดิมทั้งหมด
st.set_page_config(page_title="ระบบแปลงไฟล์ PDF to Excel เฉพาะไฟล์ผลการเรียนบกพร่อง รายครู ธนว เท่านั้น", layout="wide")
st.title("📂 ระบบแปลงไฟล์ PDF to Excel เฉพาะไฟล์ผลการเรียนบกพร่อง รายครู ธนว เท่านั้น")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF เพื่อรัน v53", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
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
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        st.success(f"ดึงข้อมูลสำเร็จ! พบ {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", output.getvalue(), "student_report_v53.xlsx")
