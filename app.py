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

    # 2. Normalization ขั้นสูงสุด (NFKC)
    text = unicodedata.normalize('NFKC', text)
    
    # 3. แปลงรหัส Unicode พิเศษเป็นมาตรฐาน
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์', '\uf709': '',
        '\uf712': 'เ', '\uf713': 'เ',
        'อ': 'อ่าน', 'ข': 'ข้อ', 'ค': 'ค้น', 'ต': 'ต่อ', 'นํ': 'นำ', 'ผ': 'แผ่น'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 4. ลบอักขระขยะ Unicode ทั้งหมด (รวมตัวที่มองไม่เห็น \u200b, \u00a0 ฯลฯ)
    text = re.sub(r'[\u0000-\u001f\u007f-\u009f\uf000-\uf0ff\u200b\u00a0]', '', text)
    
    # 5. ลบช่องว่าง "ทุกชนิด" ออกเพื่อให้ตัวอักษรชนกันสนิทที่สุด
    text = "".join(text.split())

    # 6. ขั้นตอนยุบตัวซ้ำ (เเ, แแ, าา, ์์) - ทำซ้ำหลายรอบเพื่อความชัวร์
    for _ in range(3):
        text = text.replace('เเ', 'เ')
        text = text.replace('แแ', 'แ')
        text = text.replace('าา', 'า')
        text = text.replace('นำา', 'นำ')
        text = text.replace('อ่านาน', 'อ่าน')
    
    # 7. ใช้ Regex ยุบตัวซ้ำอีกชั้น (กรณีมามากกว่า 2 ตัว)
    text = re.sub(r'เ{2,}', 'เ', text)
    text = re.sub(r'า{2,}', 'า', text)

    # 8. ซ่อมคำเฉพาะและตำแหน่งวรรณยุกต์
    text = text.replace('เ์', '์') # ลบสระเอที่ติดมากับการันต์
    
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
        'คาสตร์': 'ศาสตร์'
    }
    for wrong, right in corrections.items():
        text = text.replace(wrong, right)

    # 9. ยุบวรรณยุกต์ซ้ำ (์์, ่่, ้้)
    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)
    
    # 10. คืนค่าช่องว่าง 1 เคาะ หน้าตัวเลขท้ายชื่อวิชา
    text = re.sub(r'(\d+)$', r' \1', text)

    return text.strip()

st.set_page_config(page_title="ระบบดึงข้อมูลอัจฉริยะ v46", layout="wide")
st.title("📂 ระบบดึงข้อมูล (จบปัญหาสระเบิ้ล เเ)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF เพื่อทดสอบ v46", type="pdf")

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
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", output.getvalue(), "student_report_v46.xlsx")
