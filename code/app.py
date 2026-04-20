import streamlit as st
import os
import io
import time
import re
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ===== 1. TIỆN ÍCH HỆ THỐNG =====
def clean_text(text):
    """Loại bỏ triệt để ký tự Markdown và làm sạch văn bản"""
    # Xóa các dấu sao, thăng, gạch đầu dòng thường thấy trong Markdown
    text = re.sub(r'[*#_~-]', '', text)
    # Xóa các khoảng trắng thừa
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def add_toc(paragraph):
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar'); fldChar.set(qn('w:fldCharType'), 'begin'); run._r.append(fldChar)
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'TOC \\o "1-3" \\h \\z \\u' 
    run._r.append(instrText)
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'separate'); run._r.append(fldChar2)
    fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'end'); run._r.append(fldChar3)

def add_page_number(paragraph, position):
    if "CENTER" in position: paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif "RIGHT" in position: paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar'); fldChar.set(qn('w:fldCharType'), 'begin'); run._r.append(fldChar)
    instr = OxmlElement('w:instrText'); instr.text = "PAGE"; run._r.append(instr)
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end'); run._r.append(fldChar2)

def replace_placeholders(doc, data_dict):
    for p in doc.paragraphs:
        for key, value in data_dict.items():
            if key in p.text:
                new_text = p.text.replace(key, str(value))
                p.text = ""
                run = p.add_run(new_text)
                run.font.name = "Times New Roman"
                run.font.size = Pt(14)
                if "{{TEN_CHU_DE}}" in key:
                    run.bold = True
                    run.font.size = Pt(16)

# ===== 2. GIAO DIỆN SIDEBAR =====
st.set_page_config(page_title="AI Report Generator PRO", layout="wide", page_icon="🎓")
st.title("🎓 Hệ thống Tạo BTL chuẩn Quy cách HPU2")

with st.sidebar:
    st.header("⚙️ Cấu hình AI")
    provider = st.selectbox("AI Provider", ["Gemini", "Groq"], index=1)
    
    if provider == "Groq":
        model_map = {"Llama 3.3 70B": "llama-3.3-70b-versatile", "Qwen 2.5 32B": "qwen-2.5-32b"}
        model_choice = model_map[st.selectbox("Chọn Model", list(model_map.keys()))]
        # Ô nhập Groq Key riêng biệt
        groq_key = st.text_input("Groq API Key (Optional)", type="password")
    else:
        model_choice = st.selectbox("Chọn Model", ["gemini-2.0-pro-exp", "gemini-2.0-flash"])
        gemini_key = st.text_input("Gemini API Key (Optional)", type="password")

    st.divider()
    st.subheader("📏 Cấu hình Lề & Trang")
    hoc_phan_ui = st.selectbox("Học phần", ["Giáo dục học đại cương", "Sử dụng phương tiện KH", "Lý Luận dạy học đại học", "Khác"])
    
    presets = {
        "Giáo dục học đại cương": {"m": (2.0, 2.0, 3.0, 2.0), "sp": 1.5},
        "Khác": {"m": (2.0, 2.0, 3.0, 2.0), "sp": 1.5}
    }
    cp = presets.get(hoc_phan_ui, presets["Khác"])

    col_l, col_r = st.columns(2)
    with col_l:
        m_top, m_left = st.number_input("Trên (cm)", 0.0, 5.0, cp["m"][0]), st.number_input("Trái (cm)", 0.0, 5.0, cp["m"][2])
    with col_r:
        m_bottom, m_right = st.number_input("Dưới (cm)", 0.0, 5.0, cp["m"][1]), st.number_input("Phải (cm)", 0.0, 5.0, cp["m"][3])

    line_sp, font_sz = st.number_input("Cách dòng", 1.0, 2.5, cp["sp"]), st.number_input("Cỡ chữ", 12, 16, 14)
    page_pos = st.selectbox("Số trang", ["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"], index=2)

# ===== 3. VÙNG NHẬP LIỆU CHÍNH =====
col_in1, col_in2 = st.columns([1, 2])
with col_in1:
    st.subheader("👤 Thông tin Bìa")
    ten_hoc_vien = st.text_input("Họ và tên học viên", "Đặng Nhật Minh")
    so_bao_danh = st.text_input("Số báo danh", "39")
    ten_mon_bia = st.text_input("Chuyên đề", hoc_phan_ui)
    ten_chu_de_bia = st.text_input("Tên đề tài (Bìa)", "PHÂN TÍCH CÁC CHỨC NĂNG XÃ HỘI CỦA GIÁO DỤC")

with col_in2:
    st.subheader("🤖 Yêu cầu Nội dung (AI)")
    chi_tiet_ai = st.text_area("Mô tả đề bài", "Phân tích các chức năng xã hội của giáo dục và liên hệ thực tiễn Việt Nam.", height=270)

# ===== 4. LOGIC AI (VỚI BƯỚC REVIEW) =====
def call_ai(key, provider, prompt, model_name):
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=key)
            model = genai.GenerativeModel(model_name)
            res = model.generate_content(prompt); return res.text
        else:
            from groq import Groq
            client = Groq(api_key=key)
            res = client.chat.completions.create(model=model_name, messages=[{"role": "user", "content": prompt}])
            return res.choices[0].message.content
    except Exception as e: return f"ERROR: {e}"

if st.button("🚀 BẮT ĐẦU TẠO TIỂU LUẬN (QUY TRÌNH 3 BƯỚC)"):
    # Xác định API Key
    final_key = (groq_key if provider == "Groq" else gemini_key) or st.secrets.get(f"{provider.upper()}_API_KEY")
    if not final_key: st.error("❌ Thiếu API Key!"); st.stop()

    status = st.empty(); prog = st.progress(0)
    
    # BƯỚC 1: LẬP DÀN Ý
    status.info("📝 Bước 1: Lập dàn ý khoa học...")
    outline_prompt = f"Hãy đóng vai Tiến sĩ, lập dàn ý 6 mục chính cho đề tài: {chi_tiet_ai}. Chỉ trả về các tiêu đề mục, không kèm lời dẫn."
    outline = call_ai(final_key, provider, outline_prompt, model_choice)
    sections = [s.strip() for s in outline.split('\n') if len(s.strip()) > 10][:6]
    
    # BƯỚC 2 & 3: VIẾT & REVIEW
    full_content_list = []
    for i, sec in enumerate(sections):
        status.write(f"⏳ Đang xử lý mục {i+1}: {sec}")
        
        # Viết thô
        draft_prompt = f"Viết 700 từ chuyên sâu cho mục '{sec}' của đề tài '{chi_tiet_ai}'. Văn phong học thuật, không dùng ký tự đặc biệt, không dùng dấu * hay #."
        draft_content = call_ai(final_key, provider, draft_prompt, model_choice)
        
        # Review & Humanize
        status.write(f"🔍 Đang biên tập lại nội dung mục {i+1} để tránh 'vết' AI...")
        review_prompt = f"""Bạn là biên tập viên tạp chí khoa học. Hãy chỉnh sửa đoạn văn sau:
        1. Loại bỏ mọi ký tự lạ như *, #, -.
        2. Chỉnh sửa câu văn để tự nhiên như người viết, tránh các từ lặp máy móc.
        3. Đảm bảo tính liên kết với các phần trước đó.
        NỘI DUNG: {draft_content}"""
        
        polished_content = call_ai(final_key, provider, review_prompt, model_choice)
        # Làm sạch kỹ thuật lần cuối bằng Regex
        final_text = clean_text(polished_content)
        
        full_content_list.append((sec, final_text))
        prog.progress(int((i+1)/len(sections)*100))

    # ===== 5. XUẤT FILE WORD =====
    doc = Document()
    # Cấu hình lề
    section = doc.sections[0]
    section.top_margin, section.bottom_margin = Cm(m_top), Cm(m_bottom)
    section.left_margin, section.right_margin = Cm(m_left), Cm(m_right)
    
    # Header/Footer số trang
    add_page_number(section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph(), page_pos)

    # Đổ nội dung
    for title, content in full_content_list:
        h = doc.add_paragraph()
        h_run = h.add_run(clean_text(title).upper())
        h_run.bold = True
        h_run.font.size = Pt(14)
        
        for line in content.split('\n'):
            if line.strip():
                p = doc.add_paragraph(line)
                p.paragraph_format.line_spacing = line_sp
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                run = p.runs[0] if p.runs else p.add_run()
                run.font.name = "Times New Roman"
                run.font.size = Pt(font_sz)

    buffer = io.BytesIO()
    doc.save(buffer); buffer.seek(0)
    status.success("🎉 Đã hoàn thành và 'nhân bản' văn phong người viết thành công!")
    st.download_button("📥 Tải BTL đã qua kiểm duyệt", buffer, file_name=f"BTL_{ten_mon_bia}_{ten_hoc_vien}.docx")