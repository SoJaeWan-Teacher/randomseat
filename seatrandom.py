import streamlit as st
import pandas as pd
import openpyxl
import random
import re
from io import BytesIO

st.set_page_config(page_title="랜덤 자리 배치기", layout="wide")
st.title("🎲 랜덤 자리 배치 도구")
st.markdown("엑셀 파일을 업로드하면, 기존 자리에 무작위로 학생을 배치합니다.")

uploaded_file = st.file_uploader("📂 엑셀 파일 업로드", type=["xlsx"])

def is_student_cell(value):
    return isinstance(value, str) and re.match(r"^\d+\s+.+", value)

def process_file(file):
    wb = openpyxl.load_workbook(file)
    sheet = wb["(교사용)"]

    student_cells = []
    pattern = re.compile(r"^\d+\s+.+")

    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and pattern.match(cell.value.strip()):
                student_cells.append(cell)

    original_values = [cell.value for cell in student_cells]
    random.shuffle(original_values)

    for cell, new_value in zip(student_cells, original_values):
        cell.value = new_value

    return wb

def to_excel_download(wb):
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

if uploaded_file:
    try:
        st.success("파일 업로드 완료! 랜덤 배치 중...")
        processed_wb = process_file(uploaded_file)
        st.info("✅ 아래 버튼을 눌러 엑셀 파일을 다운로드하세요")

        download_xlsx = to_excel_download(processed_wb)
        st.download_button(
            label="📥 랜덤 자리배치 엑셀 다운로드",
            data=download_xlsx,
            file_name="랜덤_자리배치.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"🚨 처리 중 오류 발생: {e}")
