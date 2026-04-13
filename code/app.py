import streamlit as st
import os
import io
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
    """Xử lý thay thế chính xác các tag kể cả khi Word chia nhỏ runs (Paragraph-level replacement)"""
    for p in doc.paragraphs:
        # Kiểm tra xem paragraph có chứa bất kỳ key nào không
        for key, value in data_dict.items():
            if key in p.text:
                # Ghi đè lại toàn bộ text của paragraph để xử lý việc run bị chia nhỏ
                new_text = p.text.replace(key, str(value))
                p.text = "" # Xóa nội dung cũ
                run = p.add_run(new_text) # Thêm text mới đã thay thế
                # Giữ định dạng cơ bản của bìa
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
    # Mặc định chọn Groq
    provider = st.selectbox("AI Provider", ["Gemini", "Groq"], index=1)
    
    if provider == "Groq":
        model_map = {"Llama 3.3 70B Versatile": "llama-3.3-70b-versatile", "Qwen 2.5 32B": "qwen-2.5-32b"}
        model_choice = model_map[st.selectbox("Chọn Model", list(model_map.keys()))]
    else:
        model_choice = st.selectbox("Chọn Model", ["gemini-2.5-pro", "gemini-2.5-flash"])

    input_key = st.text_input(f"{provider} Key (Tùy chọn)", type="password", help="Để trống sẽ dùng Key trong Secrets")

    st.divider()
    st.subheader("📏 Cấu hình Lề & Trang")
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
    default_prompt = "Bằng lý luận và thực tiễn, anh (chị) hãy phân tích các chức năng xã hội của giáo dục. Từ đó, anh (chị) hãy liên hệ với việc thực hiện các chức năng này ở Việt Nam."
    chi_tiet_ai = st.text_area("Mô tả đề bài chi tiết để AI viết bài", default_prompt, height=270)

# ===== 4. LOGIC AI (VĂN PHONG CAO CẤP) =====
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

if st.button("🚀 BẮT ĐẦU TẠO TIỂU LUẬN (12-14 TRANG)"):
    api_key = input_key or st.secrets.get(f"{provider.upper()}_API_KEY")
    if not api_key: st.error("❌ Không tìm thấy API Key!"); st.stop()

    status = st.empty(); prog = st.progress(0)
    
    # Bước 1: Lập dàn ý
    status.info("📝 Bước 1: Đang thiết lập cấu trúc tiểu luận học thuật...")
    outline_prompt = f"""Hãy đóng vai một Tiến sĩ Giáo dục, lập dàn ý cho bài tiểu luận chuyên sâu về chủ đề: {chi_tiet_ai}. 
    Yêu cầu: Dàn ý gồm đúng 6 mục chính (Mở đầu, 4 chương nội dung lý luận và thực tiễn, Kết luận). 
    Mỗi mục phải thể hiện tính logic và chuyên môn cao. Chỉ trả về danh sách các đầu mục."""
    
    outline = call_ai(api_key, provider, outline_prompt, model_choice)
    sections = [s for s in outline.split('\n') if len(s.strip()) > 5][:6]
    
    # Bước 2: Viết nội dung
    full_content_list = []
    for i, sec in enumerate(sections):
        status.write(f"⏳ Đang biên soạn nội dung: {sec}")
        # Văn phong chỉn chu, logic, hấp dẫn
        part_prompt = f"""Hãy viết nội dung tiểu luận chuyên sâu cho mục '{sec}' của đề tài '{chi_tiet_ai}'. 
        Yêu cầu: 
        - Văn phong: Học thuật, trang trọng, sử dụng thuật ngữ chuyên ngành Giáo dục học.
        - Lập luận: Logic, có sự liên kết chặt chẽ giữa lý luận và thực tiễn Việt Nam.
        - Độ dài: Khoảng 600 từ. 
        - Hình thức: Không sử dụng ký tự đặc biệt như *, #. Trình bày dưới dạng văn xuôi mạch lạc."""
        
        part = call_ai(api_key, provider, part_prompt, model_choice)
        full_content_list.append((sec, part))
        prog.progress(15 + int((i+1)/len(sections)*75))

    # Bước 3: Đổ vào Template & Định dạng Word
    template_path = "DeThi/Template.docx"
    doc = Document(template_path) if os.path.exists(template_path) else Document()

    # Fill Bìa triệt để
    replace_placeholders(doc, {
        "{{CHUYEN_DE}}": ten_mon_bia.upper(), 
        "{{TEN_CHU_DE}}": ten_chu_de_bia.upper(), 
        "{{HO_TEN}}": ten_hoc_vien, 
        "{{SBD}}": so_bao_danh
    })

    # Căn lề & Ẩn số trang bìa
    section = doc.sections[0]
    section.different_first_page_header_footer = True 
    section.top_margin, section.bottom_margin = Cm(m_top), Cm(m_bottom)
    section.left_margin, section.right_margin = Cm(m_left), Cm(m_right)
    
    # Đánh số trang
    target_para = section.header.paragraphs[0] if page_pos == "TOP_CENTER" else (section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph())
    add_page_number(target_para, page_pos)

    # Trang Mục lục
    doc.add_page_break()
    p_toc_title = doc.add_paragraph("MỤC LỤC", style='Heading 1')
    p_toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_toc(doc.add_paragraph())
    doc.add_page_break()

    # Đổ nội dung: Đầu mục in đậm + Xuống dòng
    for title, content in full_content_list:
        # 1. Thêm đầu mục in đậm
        h = doc.add_paragraph()
        h_run = h.add_run(title.upper())
        h_run.bold = True
        h_run.font.name = "Times New Roman"
        h_run.font.size = Pt(14)
        h_run._element.rPr.rFonts.set(qn('w:eastAsia'), "Times New Roman")
        
        # 2. Thêm nội dung xuống dòng
        for line in content.split('\n'):
            if line.strip():
                p = doc.add_paragraph(line)
                p.paragraph_format.line_spacing = line_sp
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                run = p.add_run() if not p.runs else p.runs[0]
                run.font.name = "Times New Roman"
                run.font.size = Pt(font_sz)
                run._element.rPr.rFonts.set(qn('w:eastAsia'), "Times New Roman")

    # Lưu & Xuất file
    buffer = io.BytesIO()
    doc.save(buffer); buffer.seek(0)
    prog.progress(100); status.success("🎉 Bài tiểu luận đã hoàn thành xuất sắc!")
    st.download_button("📥 Tải Bài Tập Lớn Hoàn Thiện", buffer, file_name=f"BTL_{ten_hoc_vien}.docx")