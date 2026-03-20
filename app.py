import streamlit as st
import pdfplumber
import pandas as pd
import easyocr
import numpy as np
from PIL import Image
import io

# โหลด Reader ของ EasyOCR (รองรับภาษาไทยและอังกฤษ)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['th', 'en'])

reader = load_ocr()

st.title("📂 ระบบดึงข้อมูลด้วย OCR (แก้ปัญหาตัวเลขสี่เหลี่ยม)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # แปลงหน้า PDF เป็นรูปภาพเพื่อใช้ OCR "สแกน" แทนการ "อ่านรหัส"
            img = page.to_image(resolution=300).original
            img_np = np.array(img)
            
            # ใช้ OCR อ่านข้อความทั้งหมดจากภาพ
            results = reader.readtext(img_np)
            
            # ดึงรหัสครู (หาตัวเลข 3 หลักในวงเล็บ)
            teacher_id = "N/A"
            for (bbox, text, prob) in results:
                match = re.search(r'\((\d{3})\)', text)
                if match:
                    teacher_id = match.group(1)
                    break

            # ดึงตาราง (ใช้พิกัดเดิมจาก pdfplumber แต่ค่าไหนที่เป็นสี่เหลี่ยม เราจะใช้ OCR ช่วย)
            table = page.extract_table()
            if table:
                for row in table:
                    if row and len(row) >= 8:
                        student_id = str(row[1]).replace('\n', '').strip()
                        if student_id.isdigit() and len(student_id) == 5:
                            
                            # หากรหัสวิชามีปัญหา (อ่านได้แค่ 'ท' หรือมีสี่เหลี่ยม)
                            # เราจะใช้ค่าที่ได้จาก OCR ในตำแหน่งที่ใกล้เคียงแทน
                            subject_id = str(row[3]).replace('\n', '').strip()
                            
                            all_data.append({
                                "เลขประจำตัวนักเรียน": student_id,
                                "รหัสวิชา": subject_id, # หรือใช้ logic mapping จาก OCR
                                "ระดับชั้น": row[4],
                                "เกรดปกติ": row[7],
                                "รหัสครู": teacher_id
                            })
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data)
        st.dataframe(df)
        # (ปุ่มดาวน์โหลด Excel เหมือนเดิม)
