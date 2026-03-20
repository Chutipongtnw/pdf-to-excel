import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_subject_name(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน" และทุกอย่างที่ตามมาทิ้งทันที
    # ใช้การ Split แบบพื้นฐานเพื่อให้ตัดทิ้งได้ 100% ไม่ว่าสระอำจะเป็นแบบไหน
    for divider in ["จำนวน", "จํานวน", "หน่วย"]:
        if divider in text:
            text = text.split(divider)[0]
    
    # 2. แก้ไขสระและวรรณยุกต์ที่มักเป็นสี่เหลี่ยมหรือวางผิดที่
    unicode_fixes = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์',
        'ศลิป': 'ศิลป์', 'หนวย': 'หน่วย', 'ฟิสกิ': 'ฟิสิก',
        'สิกส': 'สิกส์', 'ศาสตร': 'ศาสตร์', 'ผลติ': 'ผลิต',
        'ศกึ': 'ศึก', 'นาฎ': 'นาฏ', 'พิ่ม': 'เพิ่ม'
    }
    for wrong, right in unicode_fixes.items():
        text = text.replace(wrong, right)

    # 3. ลบอักขระขยะ Unicode ที่ทำให้เกิดสี่เหลี่ยมออกให้หมด
    text = re.sub(r'[\uf000-\uf0ff]', '', text)
    
    # 4. ลบช่องว่างให้เหลือเพียงช่องเดียว (ถ้าคุณต้องการให้ติดกันหมด ให้เปลี่ยนเป็น "")
    text = " ".join(text.split())
    
    # 5. ตรวจสอบการสะกดคำว่า "ศาสตร์" อีกครั้ง
    if text.endswith('คณิตศาสตร'): text += '์'
    
    return text.strip()

st.set_page_config(page_title="ระบบดึงข้อมูลสมบูรณ์ v31", layout="wide")
st.title("📂 ระบบดึงข้อมูลผลการเรียน (ตัดหน่วยกิต & ล้างสี่เหลี่ยม)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF (ดไ.pdf)", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        total_p = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # ดึงรหัสครู (จากวงเล็บด้านบน)
            teacher_id = "N/A"
            teacher_match = re.search(r'\((\d+)\)', raw_text)
            if teacher_match:
                teacher_id = teacher_match.group(1)

            # ดึงชื่อวิชา (จากบรรทัดท้ายหน้า)
            subject_name = "N/A"
            for line in raw_text.split('\n'):
                if "ชื่อวิชา" in line:
                    parts = line.split("ชื่อวิชา")
                    if len(parts) > 1:
                        subject_name = clean_subject_name(parts[-1])
                    break

            # ดึงตารางข้อมูลนักเรียน
            table = page.extract_table()
            if table:
                for row in table:
                    if row and len(row) >= 8:
                        student_id = str(row[1]).replace('\n', '').strip()
                        # เก็บเฉพาะแถวที่มีเลขประจำตัว 5 หลัก
                        if student_id.isdigit() and len(student_id) == 5:
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": str(row[3]).replace('\n', '').strip(),
                                "ชื่อวิชา": subject_name,
                                "ระดับชั้น": str(row[4]).replace('\n', '').strip(),
                                "เกรดปกติ": str(row[7]).replace('\n', '').strip() if row[7] else "",
                                "รหัสครู": teacher_id
                            })
            progress_bar.progress((i + 1) / total_p)

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        st.success(f"ดึงข้อมูลสำเร็จ! พบ {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลด Excel", output.getvalue(), "student_report.xlsx")
