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
    """Chèn số trang vào vị trí mong muốn theo quy định của từng học phần"""
    if "CENTER" in position:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif "RIGHT" in position:
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

def replace_info(doc, placeholders):
    """Thay thế thông tin trên trang bìa của Template"""
    for p in doc.paragraphs:
        for key, value in placeholders.items():
            if key in p.text:
                p.text = p.text.replace(key, value)

# ===== GIAO DIỆN STREAMLIT =====
st.set_page_config(page_title="HPU2 AI Report PRO", layout="wide")
st.title("🎓 Hệ thống Tạo BTL chuẩn Quy cách (Gemini 2.5 Ready)")

# ===== SIDEBAR: TRUNG TÂM CẤU HÌNH =====
with st.sidebar:
    st.header("⚙️ Cấu hình Mô hình")
    provider = st.selectbox("AI Provider", ["Gemini", "Groq"])
    
    # Cập nhật các Model mới nhất theo thông tin bạn cung cấp
    if provider == "Gemini":
        model_choice = st.selectbox("Chọn Model", ["gemini-2.5-pro", "gemini-2.5-flash"])
    else:
        model_choice = st.text_input("Groq Model ID", "GPT OSS 20B 128k")

    api_key_input = st.text_input("API Key (Ghi đè .env)", type="password")

    st.divider()
    st.subheader("📚 Quy cách trình bày")
    
    hoc_phan_selection = st.selectbox("Học phần mục tiêu", [
        "Giáo dục học đại cương", 
        "Sử dụng phương tiện KH", 
        "Lý Luận dạy học đại học",
        "Tùy chỉnh khác"
    ])

    # Thiết lập giá trị mặc định theo yêu cầu của bạn
    def_margins = {"top": 2.5, "bottom": 3.0, "left": 3.0, "right": 2.0}
    def_spacing = 1.5
    def_page_pos = "TOP_CENTER"

    if "Sử dụng phương tiện" in hoc_phan_selection:
        def_page_pos = "BOTTOM_CENTER"
    elif "Lý Luận dạy học" in hoc_phan_selection:
        def_margins = {"top": 2.0, "bottom": 2.0, "left": 2.0, "right": 2.0}
        def_spacing = 1.3
        def_page_pos = "BOTTOM_RIGHT"
    elif "Khác" in hoc_phan_selection or "Tùy chỉnh" in hoc_phan_selection:
        def_page_pos = "BOTTOM_RIGHT"

    # Cho phép người dùng SỬA LẠI cấu hình trực tiếp trên Sidebar
    col_l, col_r = st.columns(2)
    with col_l:
        m_top = st.number_input("Lề trên (cm)", 0.0, 5.0, def_margins["top"])
        m_left = st.number_input("Lề trái (cm)", 0.0, 5.0, def_margins["left"])
    with col_r:
        m_bottom = st.number_input("Lề dưới (cm)", 0.0, 5.0, def_margins["bottom"])
        m_right = st.number_input("Lề phải (cm)", 0.0, 5.0, def_margins["right"])

    line_sp = st.number_input("Cách dòng", 1.0, 3.0, def_spacing)
    font_sz = st.number_input("Cỡ chữ", 10, 16, 14)
    page_pos = st.selectbox("Vị trí số trang", 
                            ["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"], 
                            index=["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"].index(def_page_pos))

# ===== VÙNG NHẬP LIỆU CHÍNH =====
col_in1, col_in2 = st.columns([1, 2])
with col_in1:
    ten_hoc_vien = st.text_input("Học viên", "Đặng Nhật Minh")
    ten_mon = st.text_input("Tên môn hiển thị", hoc_phan_selection)
with col_input2:
    default_de_bai = "Bằng lý luận và thực tiễn, anh (chị) hãy phân tích các chức năng xã hội của giáo dục. Từ đó, anh (chị) hãy liên hệ với việc thực hiện các chức năng này ở Việt Nam."
    de_bai = st.text_area("Đề tài chi tiết", default_de_bai, height=120)

# ===== XỬ LÝ AI =====
def call_ai(api_key, provider, prompt, model_name):
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            # Sử dụng model ID mới nhất từ danh sách bạn cung cấp
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            return response.text
        else:
            from groq import Groq
            client = Groq(api_key=api_key)
            response = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": prompt}]
            )
            return response.choices[0].message.content
    except Exception as e:
        return f"ERROR: {str(e)}"

# ===== THỰC THI =====
if st.button("🚀 Bắt đầu tạo bài tập lớn"):
    api_key = api_key_input or os.getenv("GEMINI_API_KEY") or os.getenv("GROQ_API_KEY")
    if not api_key:
        st.error("❌ Vui lòng cung cấp API Key!")
        st.stop()

    status = st.empty()
    progress = st.progress(0)
    
    # 1. Lập dàn ý
    status.info("📝 Bước 1: Đang lập dàn ý chi tiết...")
    outline = call_ai(api_key, provider, f"Lập dàn ý chi tiết bài tập lớn đại học 12 trang: {de_bai}. Chỉ trả về các đầu mục.", model_choice)
    sections = [s for s in outline.split('\n') if len(s.strip()) > 5]
    progress.progress(20)

    # 2. Viết từng mục (Vòng lặp để đảm bảo độ dài 12 trang)
    full_report = ""
    for idx, section in enumerate(sections):
        status.write(f"⏳ Đang viết chương: {section}")
        chapter = call_ai(api_key, provider, f"Viết nội dung học thuật sâu sắc (800-1000 từ) cho mục '{section}' của đề tài '{de_bai}'. Không dùng markdown.", model_choice)
        full_report += f"\n\n{section}\n\n{chapter}"
        progress.progress(20 + int((idx+1)/len(sections)*60))

    # 3. Xuất file Word
    status.info("📄 Bước 3: Đang áp dụng quy cách trình bày...")
    template_path = "De Thi/template.docx"
    
    doc = Document(template_path) if os.path.exists(template_path) else Document()
    
    # Thay thế thông tin trên bìa
    replace_info(doc, {"Đặng Nhật Minh": ten_hoc_vien, "TÊN CHỦ ĐỀ": de_bai.upper()})

    # Cấu hình lề từ Sidebar
    section_word = doc.sections[0]
    section_word.top_margin = Cm(m_top)
    section_word.bottom_margin = Cm(m_bottom)
    section_word.left_margin = Cm(m_left)
    section_word.right_margin = Cm(m_right)

    # Đánh số trang theo cấu hình Sidebar
    if page_pos == "TOP_CENTER":
        add_page_number(section_word.header.paragraphs[0], "TOP_CENTER")
    else:
        # Đảm bảo footer có ít nhất 1 paragraph
        footer_para = section_word.footer.paragraphs[0] if section_word.footer.paragraphs else section_word.footer.add_paragraph()
        add_page_number(footer_para, page_pos)

    # Thêm văn bản và định dạng
    for line in full_report.split('\n'):
        if line.strip():
            p = doc.add_paragraph(line)
            p.paragraph_format.line_spacing = line_sp
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = p.runs[0] if p.runs else p.add_run(line)
            run.font.name = "Times New Roman"
            run.font.size = Pt(font_sz)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "Times New Roman")

    # Lưu và Tải về
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    progress.progress(100)
    status.success(f"✅ Hoàn thành! Model: {model_choice} | Lề: {m_left}-{m_right}cm")

    st.download_button(
        "📥 Tải Bài Tập Lớn Hoàn Thiện",
        buffer,
        file_name=f"BTL_{ten_hoc_vien}_{ten_mon}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )