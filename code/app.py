import streamlit as st
import os
import io
import re
from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ===== KHỞI TẠO =====
load_dotenv()

def add_page_number(paragraph, position):
    """Thêm số trang vào Header/Footer theo vị trí"""
    if position == "TOP_CENTER":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif position == "BOTTOM_CENTER":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif position == "BOTTOM_RIGHT":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)

    instrText = OxmlElement('w:instrText')
    instrText.text = "PAGE"
    run._r.append(instrText)

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar2)

# ===== GIAO DIỆN STREAMLIT =====
st.set_page_config(page_title="AI Report Generator PRO", layout="wide")
st.title("🎓 Hệ thống Tạo BTL chuẩn Quy cách HPU2")

# ===== SIDEBAR: TRUNG TÂM CẤU HÌNH =====
with st.sidebar:
    st.header("⚙️ Cấu hình Hệ thống")
    provider = st.selectbox("AI Provider", ["Gemini", "Groq"])
    api_key_input = st.text_input("API Key (Ghi đè .env)", type="password")

    st.divider()
    st.subheader("📚 Quy cách trình bày")
    
    # Logic mặc định theo học phần
    hoc_phan_selection = st.selectbox("Học phần mục tiêu", [
        "Giáo dục học đại cương", 
        "Sử dụng phương tiện KH", 
        "Lý Luận dạy học đại học",
        "Khác"
    ])

    # Thiết lập giá trị mặc định dựa trên học phần
    def_margins = {"top": 2.5, "bottom": 3.0, "left": 3.0, "right": 2.0}
    def_spacing = 1.5
    def_page_pos = "TOP_CENTER"

    if "Sử dụng phương tiện" in hoc_phan_selection:
        def_page_pos = "BOTTOM_CENTER"
    elif "Lý Luận dạy học" in hoc_phan_selection:
        def_margins = {"top": 2.0, "bottom": 2.0, "left": 2.0, "right": 2.0}
        def_spacing = 1.3
        def_page_pos = "BOTTOM_RIGHT"
    elif "Khác" in hoc_phan_selection:
        def_page_pos = "BOTTOM_RIGHT"

    # Cho phép người dùng SỬA LẠI cấu hình bên Sidebar
    col_l, col_r = st.columns(2)
    with col_l:
        m_top = st.number_input("Lề trên (cm)", 1.0, 5.0, def_margins["top"])
        m_left = st.number_input("Lề trái (cm)", 1.0, 5.0, def_margins["left"])
    with col_r:
        m_bottom = st.number_input("Lề dưới (cm)", 1.0, 5.0, def_margins["bottom"])
        m_right = st.number_input("Lề phải (cm)", 1.0, 5.0, def_margins["right"])

    line_sp = st.number_input("Cách dòng (Line spacing)", 1.0, 3.0, def_spacing)
    font_sz = st.number_input("Cỡ chữ (Font size)", 10, 16, 14)
    page_pos = st.selectbox("Vị trí đánh số trang", 
                            ["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"], 
                            index=["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"].index(def_page_pos))

# ===== VÙNG NHẬP LIỆU CHÍNH =====
col_input1, col_input2 = st.columns([1, 2])
with col_input1:
    ten_hoc_vien = st.text_input("Học viên", "Đặng Nhật Minh")
    ten_mon = st.text_input("Tên môn (Hiển thị file)", hoc_phan_selection)
with col_input2:
    default_de_bai = "Bằng lý luận và thực tiễn, anh (chị) hãy phân tích các chức năng xã hội của giáo dục. Từ đó, anh (chị) hãy liên hệ với việc thực hiện các chức năng này ở Việt Nam."
    de_bai = st.text_area("Đề tài chi tiết", default_de_bai, height=120)

# ===== XỬ LÝ AI =====
def generate_content(api_key, provider, prompt):
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-1.5-flash")
            return model.generate_content(prompt).text
        else:
            from groq import Groq
            client = Groq(api_key=api_key)
            resp = client.chat.completions.create(model="GPT OSS 20B 128k", messages=[{"role":"user","content":prompt}])
            return resp.choices[0].message.content
    except Exception as e: return f"Lỗi: {e}"

# ===== THỰC THI =====
if st.button("🚀 Bắt đầu tạo bài tập lớn"):
    api_key = api_key_input or os.getenv("GEMINI_API_KEY") or os.getenv("GROQ_API_KEY")
    if not api_key:
        st.error("Vui lòng nhập API Key!")
        st.stop()

    status = st.empty()
    progress = st.progress(0)

    # 1. Gen nội dung (Ví dụ rút gọn, bạn có thể thêm vòng lặp gen từng chương ở đây)
    status.info("🤖 AI đang viết nội dung học thuật...")
    content = generate_content(api_key, provider, f"Viết bài luận 12 trang về: {de_bai}")
    progress.progress(50)

    # 2. Xử lý File Word
    status.info("📝 Đang khởi tạo định dạng chuẩn...")
    template_path = "De Thi/template.docx"
    doc = Document(template_path) if os.path.exists(template_path) else Document()

    # Áp dụng lề từ Sidebar
    section = doc.sections[0]
    section.top_margin = Cm(m_top)
    section.bottom_margin = Cm(m_bottom)
    section.left_margin = Cm(m_left)
    section.right_margin = Cm(m_right)

    # Đánh số trang theo Sidebar
    if page_pos == "TOP_CENTER":
        add_page_number(section.header.paragraphs[0], "TOP_CENTER")
    else:
        add_page_number(section.footer.paragraphs[0], page_pos)

    # Thêm nội dung và format
    for para_text in content.split('\n'):
        if para_text.strip():
            p = doc.add_paragraph(para_text)
            p.paragraph_format.line_spacing = line_sp
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = p.runs[0] if p.runs else p.add_run(para_text)
            run.font.name = "Times New Roman"
            run.font.size = Pt(font_sz)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "Times New Roman")

    # 3. Kết quả
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    progress.progress(100)
    status.success(f"✅ Đã xong! Lề: {m_left}-{m_right}-{m_top}-{m_bottom}. Số trang: {page_pos}")

    st.download_button(
        "📥 Tải file Word Hoàn thiện",
        buffer,
        file_name=f"BTL_{ten_hoc_vien}_{ten_mon}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )