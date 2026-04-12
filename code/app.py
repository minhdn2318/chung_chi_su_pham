import streamlit as st
import os
import io
import re
import time

from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ===== LOAD ENV =====
load_dotenv()

# ===== CLEAN TEXT =====
def clean_text(text):
    text = re.sub(r'\*\*', '', text)
    text = re.sub(r'\*', '', text)
    text = re.sub(r'#+', '', text)
    text = re.sub(r'- ', '', text)
    text = re.sub(r'\n+', '\n', text)
    return text.strip()

# ===== PAGE NUMBER =====
def add_page_number(run):
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)

    instrText = OxmlElement('w:instrText')
    instrText.text = "PAGE"
    run._r.append(instrText)

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar2)

# ===== STREAMLIT UI =====
st.set_page_config(page_title="AI Report Generator PRO", layout="wide")
st.title("🎓 AI Report Generator (Gemini / Groq)")

# ===== SIDEBAR =====
with st.sidebar:
    st.header("⚙️ Cấu hình")

    provider = st.selectbox("Chọn AI Provider", ["Gemini", "Groq"])

    api_key_input = st.text_input("API Key (optional)", type="password")

    st.subheader("📄 Số trang")
    num_pages = st.number_input("Số trang", 5, 30, 12)

    st.subheader("Font")
    font_name = st.text_input("Font", "Times New Roman")
    font_size = st.number_input("Size", 12, 15, 14)
    line_spacing = st.number_input("Line spacing", 1.0, 2.5, 1.5)

# ===== INPUT =====
col1, col2 = st.columns(2)
with col1:
    ten_hoc_phan = st.text_input("Học phần", "Giáo dục học")
    ten_hoc_vien = st.text_input("Sinh viên", "Nguyễn Văn A")
with col2:
    de_bai = st.text_area("Đề tài", "Phân tích giáo dục")

# ===== LOAD API KEY =====
def get_api_key(provider):
    if api_key_input:
        return api_key_input

    if provider == "Gemini":
        try:
            return st.secrets["GEMINI_API_KEY"]
        except:
            return os.getenv("GEMINI_API_KEY")

    if provider == "Groq":
        return os.getenv("GROQ_API_KEY")

# ===== GENERATE TEXT =====
def generate_text(provider, api_key, prompt):
    if provider == "Gemini":
        import google.generativeai as genai
        genai.configure(api_key=api_key)

        model = genai.GenerativeModel("gemini-2.5-flash")

        return model.generate_content(prompt).text

    elif provider == "Groq":
        from groq import Groq
        client = Groq(api_key=api_key)

        response = client.chat.completions.create(
            model="llama3-70b-8192",
            messages=[{"role": "user", "content": prompt}],
        )

        return response.choices[0].message.content

# ===== MAIN =====
if st.button("🚀 Generate Report"):
    api_key = get_api_key(provider)

    if not api_key:
        st.error("❌ Thiếu API key")
        st.stop()

    try:
        progress = st.progress(0)
        status = st.empty()

        # ===== WORD CALC =====
        words_per_page = 400
        total_words = num_pages * words_per_page

        # ===== PROMPT (1 CALL ONLY) =====
        prompt = f"""
Viết báo cáo học thuật hoàn chỉnh

Học phần: {ten_hoc_phan}
Sinh viên: {ten_hoc_vien}
Đề tài: {de_bai}

Yêu cầu:
- Tổng độ dài: {total_words} từ (~{num_pages} trang)
- Cấu trúc:
1. Mở đầu
2. Cơ sở lý luận
3. Thực trạng Việt Nam
4. Giải pháp
5. Kết luận

- Viết tiếng Việt học thuật
- Không markdown, không ký tự *, #
- Có mục nhỏ 1.1, 1.2...
"""

        status.text("🤖 Đang gọi AI...")
        raw_text = generate_text(provider, api_key, prompt)

        progress.progress(40)

        text = clean_text(raw_text)

        # ===== LIMIT HARD =====
        words = text.split()
        if len(words) > total_words:
            text = " ".join(words[:total_words])

        # ===== CREATE DOC =====
        doc = Document()

        # margin
        section = doc.sections[0]
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(3)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)

        status.text("📝 Đang format Word...")

        for para_text in text.split("\n"):
            if para_text.strip():
                para = doc.add_paragraph(para_text)
                para.paragraph_format.line_spacing = line_spacing
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                run = para.runs[0]
                run.font.name = font_name
                run.font.size = Pt(font_size)
                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

        # page number
        footer = doc.sections[0].footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_page_number(p.add_run())

        # save
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        progress.progress(100)
        status.text("✅ Done")

        st.download_button(
            "📥 Tải file Word",
            buffer,
            file_name="report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(str(e))
