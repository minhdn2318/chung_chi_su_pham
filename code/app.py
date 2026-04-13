import streamlit as st
import os
import io
import re
import time
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ===== 1. TIỆN ÍCH HỆ THỐNG =====
def add_toc(paragraph):
    """Chèn mã Field TOC để Word tự động tạo mục lục"""
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar)
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u' 
    run._r.append(instrText)
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    run._r.append(fldChar2)
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar3)

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
                p.text = p.text.replace(key, str(value))

# ===== 2. GIAO DIỆN SIDEBAR =====
st.set_page_config(page_title="HPU2 Report Generator PRO", layout="wide", page_icon="🎓")
st.title("🎓 Hệ thống Tạo BTL chuẩn Quy cách HPU2")

with st.sidebar:
    st.header("⚙️ Cấu hình AI")
    provider = st.selectbox("Chọn AI Provider", ["Gemini", "Groq"], index=1)
    
    if provider == "Groq":
        model_map = {
            "Llama 3.3 70B Versatile": "llama-3.3-70b-versatile",
            "Qwen 2.5 32B Coder": "qwen-2.5-32b"
        }
        model_choice = model_map[st.selectbox("Chọn Model", list(model_map.keys()))]
    else:
        model_choice = st.selectbox("Chọn Model", ["gemini-2.5-pro", "gemini-2.5-flash"])

    input_key = st.text_input(f"{provider} Key (Optional)", type="password", help="Hệ thống sẽ dùng Key mặc định nếu để trống")

    st.divider()
    st.subheader("📏 Quy cách trình bày")
    hoc_phan_ui = st.selectbox("Học phần", ["Giáo dục học đại cương", "Sử dụng phương tiện KH", "Lý Luận dạy học đại học", "Khác"])
    
    presets = {
        "Giáo dục học đại cương": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5},
        "Sử dụng phương tiện KH": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5},
        "Lý Luận dạy học đại học": {"m": (2.0, 2.0, 2.0, 2.0), "sp": 1.3},
        "Khác": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5}
    }
    cp = presets.get(hoc_phan_ui, presets["Khác"])

    col_l, col_r = st.columns(2)
    with col_l:
        m_top, m_left = st.number_input("Trên", 0.0, 5.0, cp["m"][0]), st.number_input("Trái", 0.0, 5.0, cp["m"][2])
    with col_r:
        m_bottom, m_right = st.number_input("Dưới", 0.0, 5.0, cp["m"][1]), st.number_input("Phải", 0.0, 5.0, cp["m"][3])

    line_sp, font_sz = st.number_input("Cách dòng", 1.0, 2.5, cp["sp"]), st.number_input("Cỡ chữ", 12, 16, 14)
    page_pos = st.selectbox("Vị trí số trang", ["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"], index=2)

# ===== 3. NHẬP LIỆU CHÍNH =====
col_in1, col_in2 = st.columns([1, 2])
with col_in1:
    st.subheader("👤 Thông tin học viên")
    ten_hoc_vien = st.text_input("Họ và tên", "Đặng Nhật Minh")
    so_bao_danh = st.text_input("Số báo danh", "SBD-2026-001")
    ten_mon = st.text_input("Chuyên đề (Bìa)", hoc_phan_ui)
    
    st.divider()
    st.subheader("📑 Thông tin Đề tài")
    ten_chu_de_bia = st.text_input("Tên đề tài (Bìa)", "PHÂN TÍCH CÁC CHỨC NĂNG XÃ HỘI CỦA GIÁO DỤC")

with col_in2:
    st.subheader("🤖 Yêu cầu chi tiết cho AI")
    default_prompt = "Bằng lý luận và thực tiễn, anh (chị) hãy phân tích các chức năng xã hội của giáo dục. Từ đó, anh (chị) hãy liên hệ với việc thực hiện các chức năng này ở Việt Nam."
    chi_tiet_ai = st.text_area("Yêu cầu nội dung (Dùng để gen bài)", default_prompt, height=265)

# ===== 4. LOGIC XỬ LÝ AI =====
def call_ai(key, provider, prompt, model_name):
    try:
        api_key = key or st.secrets.get(f"{provider.upper()}_API_KEY")
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            res = model.generate_content(prompt); time.sleep(8); return res.text
        else:
            from groq import Groq
            client = Groq(api_key=api_key)
            res = client.chat.completions.create(model=model_name, messages=[{"role": "user", "content": prompt}])
            time.sleep(2); return res.choices[0].message.content
    except Exception as e: return f"ERROR: {e}"

# ===== 5. THỰC THI =====
if st.button("🚀 BẮT ĐẦU TẠO BÀI TẬP LỚN"):
    api_key = input_key or st.secrets.get(f"{provider.upper()}_API_KEY")
    if not api_key: st.error("❌ Không tìm thấy API Key!"); st.stop()

    status = st.empty(); prog = st.progress(0)
    
    # 1. Lập dàn ý khống chế số lượng mục (Tối đa 6-7 mục để bài không bị quá dài)
    status.info("📝 Bước 1: Lập dàn ý (Đang khống chế độ dài 12-15 trang)...")
    outline_prompt = f"Lập dàn ý bài tập lớn đại học cho chủ đề: {chi_tiet_ai}. Yêu cầu: Chỉ lập 6 mục chính (bao gồm cả Mở đầu, Nội dung chia làm 4 phần, và Kết luận). Chỉ trả về danh sách các mục."
    outline = call_ai(api_key, provider, outline_prompt, model_choice)
    sections = [s for s in outline.split('\n') if len(s.strip()) > 5]
    prog.progress(15)

    # 2. Viết nội dung từng mục (Khống chế 800 từ/mục để tổng đạt ~5000-6000 từ)
    full_content_list = []
    for i, sec in enumerate(sections):
        status.write(f"⏳ Đang viết mục {i+1}/{len(sections)}: {sec}")
        # Yêu cầu độ dài vừa phải để bài đạt 12-15 trang
        part_prompt = f"Viết bài luận học thuật sâu sắc cho mục '{sec}' của đề tài '{chi_tiet_ai}'. Yêu cầu độ dài khoảng 800 từ. Tuyệt đối không markdown."
        part = call_ai(api_key, provider, part_prompt, model_choice)
        full_content_list.append((sec, part))
        prog.progress(15 + int((i+1)/len(sections)*75))

    # 3. Đổ vào Template chuẩn
    status.info("📄 Bước 3: Đang căn lề và đổ dữ liệu bìa...")
    template_path = "DeThi/Template.docx"
    doc = Document(template_path) if os.path.exists(template_path) else Document()
    
    replace_placeholders(doc, {
        "{{CHUYEN_DE}}": ten_mon.upper(),
        "{{TEN_CHU_DE}}": ten_chu_de_bia.upper(),
        "{{HO_TEN}}": ten_hoc_vien,
        "{{SBD}}": so_bao_danh
    })

    # Cấu hình SECTION (Lề & Ẩn số trang bìa)
    section = doc.sections[0]
    section.different_first_page_header_footer = True 
    section.top_margin, section.bottom_margin = Cm(m_top), Cm(m_bottom)
    section.left_margin, section.right_margin = Cm(m_left), Cm(m_right)
    
    # Đánh số trang (Hiển thị từ trang 2)
    target_p = section.header.paragraphs[0] if page_pos == "TOP_CENTER" else (section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph())
    add_page_number(target_p, page_pos)

    # Chèn Mục lục Field Code
    doc.add_page_break()
    p_toc = doc.add_paragraph("MỤC LỤC", style='Heading 1')
    p_toc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_toc(doc.add_paragraph())
    doc.add_page_break()

    # Đổ nội dung (Dùng Heading 1 để mục lục nhận diện)
    for title, content in full_content_list:
        h = doc.add_heading(title, level=1)
        h_run = h.runs[0]
        h_run.font.name, h_run.font.size, h_run.font.color.rgb = "Times New Roman", Pt(14), RGBColor(0, 0, 0)
        h_run._element.rPr.rFonts.set(qn('w:eastAsia'), "Times New Roman")

        for line in content.split('\n'):
            if line.strip():
                p = doc.add_paragraph(line)
                p.paragraph_format.line_spacing, p.alignment = line_sp, WD_ALIGN_PARAGRAPH.JUSTIFY
                run = p.add_run() if not p.runs else p.runs[0]
                run.font.name, run.font.size = "Times New Roman", Pt(font_sz)
                run._element.rPr.rFonts.set(qn('w:eastAsia'), "Times New Roman")

    # Hoàn tất
    buffer = io.BytesIO()
    doc.save(buffer); buffer.seek(0)
    prog.progress(100); status.success("🎉 Đã tạo xong báo cáo chuẩn 12-15 trang!")
    st.download_button("📥 Tải file Word hoàn thiện", buffer, file_name=f"BTL_{ten_hoc_vien}.docx")