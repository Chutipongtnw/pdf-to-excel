import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def universal_thai_cleaner(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน" และ "หน่วยการเรียน" ออกให้ขาด 100% 
    for divider in ["จำนวน", "จํานวน", "หน่วย"]:
        if divider in text:
            text = text.split(divider)[0]

    # 2. แก้ไขรหัส Unicode ที่เป็นสระ/วรรณยุกต์ซ้อน (สาเหตุของสี่เหลี่ยมและตัวเบิ้ล) [cite: 93, 112]
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 3. กำจัด "เ เ" และสระซ้ำทุกรูปแบบ (รวมกรณีมีช่องว่างคั่นกลาง) 
    # ลบช่องว่างทุกชนิดทิ้งก่อน เพื่อให้สระที่เบิ้ลมาชนกัน แล้วค่อยยุบ
    text = re.sub(r'\s+', '', text) 
    text = text.replace('เเ', 'เ').replace('แแ', 'แ').replace('ะะ', 'ะ')
    
    # 4. ยุบวรรณยุกต์และตัวการันต์ที่ซ้ำกัน (เช่น ์์ -> ์) [cite: 4, 10, 28]
    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)

    # 5. ซ่อมคำเฉพาะที่มักจะสลับตำแหน่งสระ (เช่น ฟิสกิส์ -> ฟิสิกส์) 
    # ลบอักขระขยะ Unicode ออกให้หมด
    text = re.sub(r'[\uf000-\uf0ff]', '', text)
    
    corrections = {
        'คณิตศาสตร': 'คณิตศาสตร์',
        'พิ่มเติม': 'เพิ่มเติม',
        'ฟิสกิส์': 'ฟิสิกส์',
        'ฟิสิกส': 'ฟิสิกส์',
        'ทัศนศลิป': 'ทัศนศิลป์',
        'นาฎศิลป': 'นาฏศิลป์',
        'ศกึษา': 'ศึกษา'
    }
    for wrong, right in corrections.items():
        text = text.replace(wrong, right)

    # 6. ขั้นตอนสุดท้าย: ยุบตัวการันต์ที่อาจหลงเหลือจากการ replace [cite: 28, 45]
    text = text.replace('์์', '์')
    
    # เติมช่องว่าง 1 เคาะ หน้าตัวเลขตัวสุดท้าย (ถ้ามี) เพื่อความสวยงาม [cite: 4, 10]
    text = re.sub(r'(\d+)$', r' \1', text)

    return text.strip()

st.set_page_config(page_title="ระบบดึงข้อมูลอัจฉริยะ v34", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (Fixed เ เ & ฟิสิกส์)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # ดึงรหัสครู [cite: 2, 8, 14]
            teacher_id = "N/A"
            t_match = re.search(r'\((\d+)\)', raw_text)
            if t_match: teacher_id = t_match.group(1)

            # ดึงชื่อวิชา [cite: 5, 12, 18]
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
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        st.success(f"ดึงข้อมูลสำเร็จ! พบทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", output.getvalue(), "report_v34.xlsx")
