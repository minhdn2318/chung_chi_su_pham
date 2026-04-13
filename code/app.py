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

# ===== KHỞI TẠO =====
load_dotenv()

def add_page_number(paragraph, position):
    """Chèn số trang tự động vào Header/Footer theo quy chuẩn HPU2"""
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
    """Thay thế thông tin học viên trên trang bìa của Template"""
    for p in doc.paragraphs:
        for key, value in placeholders.items():
            if key in p.text:
                p.text = p.text.replace(key, value)

# ===== GIAO DIỆN STREAMLIT =====
st.set_page_config(page_title="HPU2 AI Report PRO", layout="wide")
st.title("🎓 Hệ thống Tạo BTL chuẩn Quy cách (Fix Quota 429)")

# ===== SIDEBAR: TRUNG TÂM CẤU HÌNH =====
with st.sidebar:
    st.header("⚙️ Cấu hình Mô hình")
    provider = st.selectbox("AI Provider", ["Gemini", "Groq"])
    
    if provider == "Gemini":
        model_choice = st.selectbox("Chọn Model", ["gemini-2.5-pro", "gemini-2.5-flash"])
        st.caption("⚠️ Lưu ý: Free Tier giới hạn request. Hệ thống sẽ tự động chờ nghỉ.")
    else:
        # Thêm các option model mạnh nhất của Groq/Qwen hiện nay
        model_choice = st.selectbox("Chọn Groq Model", [
            "llama-3.3-70b-versatile", 
            "qwen-2.5-32b", 
            "llama3-70b-8192",
            "GPT OSS 20B 128k"
        ])

    api_key_input = st.text_input("API Key (Ghi đè .env)", type="password")

    st.divider()
    st.subheader("📚 Quy cách trình bày")
    
    hoc_phan_selection = st.selectbox("Học phần mục tiêu", [
        "Giáo dục học đại cương", 
        "Sử dụng phương tiện KH", 
        "Lý Luận dạy học đại học",
        "Tùy chỉnh khác"
    ])

    # Thiết lập giá trị mặc định dựa trên quy định HPU2
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

    col_l, col_r = st.columns(2)
    with col_l:
        m_top = st.number_input("Lề trên (cm)", 0.0, 5.0, def_margins["top"])
        m_left = st.number_input("Lề trái (cm)", 0.0, 5.0, def_margins["left"])
    with col_r:
        m_bottom = st.number_input("Lề dưới (cm)", 0.0, 5.0, def_margins["bottom"])
        m_right = st.number_input("Lề phải (cm)", 0.0, 5.0, def_margins["right"])

    line_sp = st.number_input("Cách dòng", 1.0, 2.5, def_spacing)
    font_sz = st.number_input("Cỡ chữ", 12, 16, 14)
    page_pos = st.selectbox("Vị trí số trang", 
                            ["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"], 
                            index=["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"].index(def_page_pos))

# ===== VÙNG NHẬP LIỆU CHÍNH =====
col_in1, col_in2 = st.columns([1, 2])
with col_in1:
    ten_hoc_vien = st.text_input("Học viên", "Đặng Nhật Minh")
    ten_mon = st.text_input("Tên môn hiển thị", hoc_phan_selection)
with col_in2:
    default_de_bai = "Bằng lý luận và thực tiễn, anh (chị) hãy phân tích các chức năng xã hội của giáo dục. Từ đó, anh (chị) hãy liên hệ với việc thực hiện các chức năng này ở Việt Nam."
    de_bai = st.text_area("Đề tài chi tiết", default_de_bai, height=120)

# ===== XỬ LÝ AI =====
def call_ai(api_key, provider, prompt, model_name):
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            # Thêm thời gian nghỉ 4s sau mỗi lần gọi Gemini để tránh 429
            time.sleep(4) 
            return response.text
        else:
            from groq import Groq
            client = Groq(api_key=api_key)
            response = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=4000
            )
            time.sleep(1) # Groq nhanh hơn nên chỉ cần nghỉ 1s
            return response.choices[0].message.content
    except Exception as e:
        if "429" in str(e):
            st.warning("⚠️ API đang bận (429), hệ thống sẽ tự thử lại sau 10s...")
            time.sleep(10)
            return call_ai(api_key, provider, prompt, model_name)
        return f"ERROR: {str(e)}"

# ===== THỰC THI CHÍNH =====
if st.button("🚀 Bắt đầu tạo bài tập lớn (Đảm bảo 12 trang)"):
    api_key = api_key_input or os.getenv("GEMINI_API_KEY") or os.getenv("GROQ_API_KEY")
    if not api_key:
        st.error("❌ Vui lòng cung cấp API Key!")
        st.stop()

    status = st.empty()
    progress = st.progress(0)
    
    # BƯỚC 1: LẬP DÀN Ý
    status.info("📝 Bước 1: Đang lập dàn ý chuyên sâu...")
    outline_prompt = f"Lập dàn ý báo cáo học thuật 12 trang cho đề tài: {de_bai}. Chỉ trả về các tiêu đề mục 1, 2, 3..."
    outline = call_ai(api_key, provider, outline_prompt, model_choice)
    sections = [s for s in outline.split('\n') if len(s.strip()) > 5]
    progress.progress(15)

    # BƯỚC 2: VIẾT NỘI DUNG TỪNG MỤC
    full_report = ""
    for idx, section in enumerate(sections):
        status.write(f"⏳ Đang viết chương {idx+1}/{len(sections)}: {section}")
        chapter_prompt = f"Viết bài luận chuyên sâu (khoảng 1000 từ) cho phần '{section}' của đề tài '{de_bai}'. Phong cách học thuật đại học, không dùng ký tự markdown."
        chapter = call_ai(api_key, provider, chapter_prompt, model_choice)
        full_report += f"\n\n{section}\n\n{chapter}"
        
        # Cập nhật tiến độ
        current_p = 15 + int((idx+1)/len(sections)*75)
        progress.progress(current_p)

    # BƯỚC 3: TẠO FILE WORD
    status.info("📄 Bước 3: Đang áp dụng quy cách trình bày & Template...")
    template_path = "DeThi/Template.docx"
    
    if os.path.exists(template_path):
        doc = Document(template_path)
        replace_info(doc, {"Đặng Nhật Minh": ten_hoc_vien, "TÊN CHỦ ĐỀ": de_bai.upper()})
    else:
        doc = Document()
        st.warning("⚠️ Không tìm thấy Template.docx, hệ thống tạo file mới.")

    sec = doc.sections[0]
    sec.top_margin, sec.bottom_margin = Cm(m_top), Cm(m_bottom)
    sec.left_margin, sec.right_margin = Cm(m_left), Cm(m_right)

    if page_pos == "TOP_CENTER":
        add_page_number(sec.header.paragraphs[0], "TOP_CENTER")
    else:
        footer_p = sec.footer.paragraphs[0] if sec.footer.paragraphs else sec.footer.add_paragraph()
        add_page_number(footer_p, page_pos)

    for line in full_report.split('\n'):
        if line.strip():
            p = doc.add_paragraph(line)
            p.paragraph_format.line_spacing = line_sp
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = p.runs[0] if p.runs else p.add_run(line)
            run.font.name = "Times New Roman"
            run.font.size = Pt(font_sz)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "Times New Roman")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    progress.progress(100)
    status.success(f"🎉 Hoàn thành! File đã được tối ưu hóa độ dài 12 trang.")

    st.download_button(
        "📥 Tải Bài Tập Lớn Hoàn Thiện",
        buffer,
        file_name=f"BTL_{ten_hoc_vien}_{ten_mon}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )