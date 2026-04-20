import streamlit as st
import os
import io
import re
import unicodedata
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ===== 1. DATABASE 9 HỌC PHẦN (NHẬT MINH DÁN NỘI DUNG EXCEL VÀO ĐÂY) =====
DATA_BTL = {
    "HP_01: Giáo dục học đại cương": {
        "de_tai": "TÊN ĐỀ TÀI TỪ EXCEL CỦA BẠN", 
        "cau_hoi": "NỘI DUNG CÂU HỎI CHI TIẾT TỪ EXCEL CỦA BẠN",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14
    },
    "HP_02: Lý luận dạy học đại học": {
        "de_tai": "TÊN ĐỀ TÀI TỪ EXCEL CỦA BẠN",
        "cau_hoi": "NỘI DUNG CÂU HỎI CHI TIẾT TỪ EXCEL CỦA BẠN",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.3, "font_sz": 14
    },
    "HP_03: Sử dụng phương tiện kỹ thuật KH": {
        "de_tai": "TÊN ĐỀ TÀI TỪ EXCEL CỦA BẠN",
        "cau_hoi": "NỘI DUNG CÂU HỎI CHI TIẾT TỪ EXCEL CỦA BẠN",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14
    },
    "HP_04: Tâm lý học đại cương": {
        "de_tai": "TÊN ĐỀ TÀI TỪ EXCEL CỦA BẠN",
        "cau_hoi": "NỘI DUNG CÂU HỎI CHI TIẾT TỪ EXCEL CỦA BẠN",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14
    },
    "HP_05: Quản lý nhà nước về Giáo dục": {
        "de_tai": "TÊN ĐỀ TÀI TỪ EXCEL CỦA BẠN",
        "cau_hoi": "NỘI DUNG CÂU HỎI CHI TIẾT TỪ EXCEL CỦA BẠN",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14
    },
    "HP_06: Phát triển chương trình đào tạo": {
        "de_tai": "TÊN ĐỀ TÀI TỪ EXCEL CỦA BẠN",
        "cau_hoi": "NỘI DUNG CÂU HỎI CHI TIẾT TỪ EXCEL CỦA BẠN",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14
    },
    "HP_07: Đánh giá trong giáo dục": {
        "de_tai": "TÊN ĐỀ TÀI TỪ EXCEL CỦA BẠN",
        "cau_hoi": "NỘI DUNG CÂU HỎI CHI TIẾT TỪ EXCEL CỦA BẠN",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14
    },
    "HP_08: Phương pháp nghiên cứu khoa học": {
        "de_tai": "TÊN ĐỀ TÀI TỪ EXCEL CỦA BẠN",
        "cau_hoi": "NỘI DUNG CÂU HỎI CHI TIẾT TỪ EXCEL CỦA BẠN",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14
    },
    "HP_09: Kỹ năng mềm trong dạy học": {
        "de_tai": "TÊN ĐỀ TÀI TỪ EXCEL CỦA BẠN",
        "cau_hoi": "NỘI DUNG CÂU HỎI CHI TIẾT TỪ EXCEL CỦA BẠN",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14
    }
}

# ===== 2. TIỆN ÍCH ĐỊNH DẠNG & LÀM SẠCH =====
def remove_vietnamese_accent(s):
    s = s.replace("Đ", "D").replace("đ", "d")
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
    return re.sub(r'[^a-zA-Z0-9_]', '_', s)

def add_toc(paragraph):
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar'); fldChar.set(qn('w:fldCharType'), 'begin'); run._r.append(fldChar)
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    run._r.append(instrText)
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'separate'); run._r.append(fldChar2)
    fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'end'); run._r.append(fldChar3)

def set_font_style(run, size=14, bold=False):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.bold = bold
    r = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), "Times New Roman")
    rFonts.set(qn('w:hAnsi'), "Times New Roman")
    rFonts.set(qn('w:eastAsia'), "Times New Roman")
    r.insert(0, rFonts)

def strictly_clean_content(text):
    """Xóa bỏ các ký tự Markdown và lời dẫn AI"""
    ai_patterns = [r"Dưới đây là.*:", r"Đoạn văn đã được chỉnh sửa.*:", r"Tôi đã thực hiện.*:", r"Nội dung biên tập.*:"]
    for pattern in ai_patterns:
        text = re.sub(pattern, "", text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'[*#_~-]', '', text)
    return text.strip()

# ===== 3. LOGIC AI (VĂN PHONG TIỂU LUẬN CAO CẤP) =====
def call_ai(key, provider, prompt):
    sys_instruction = (
        "Bạn là một Tiến sĩ Giáo dục học dày dặn kinh nghiệm. "
        "NHIỆM VỤ: Viết nội dung tiểu luận học thuật chuyên sâu. "
        "VĂN PHONG: Trang trọng, logic, sử dụng các từ ngữ học thuật chuyên ngành. "
        "CẤM: Không dùng Markdown (*, #), không chào hỏi, không dẫn dắt. "
        "YÊU CẦU: Nội dung phải mang tính lý luận kết hợp thực tiễn sâu sắc."
    )
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=key)
            model = genai.GenerativeModel('gemini-1.5-pro')
            return model.generate_content(f"{sys_instruction}\n\n{prompt}").text
        else:
            from groq import Groq
            client = Groq(api_key=key)
            res = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "system", "content": sys_instruction}, {"role": "user", "content": prompt}]
            )
            return res.choices[0].message.content
    except Exception as e: return f"Lỗi: {e}"

# ===== 4. GIAO DIỆN STREAMLIT =====
st.set_page_config(page_title="AI Essay Pro - VNU", layout="wide")

with st.sidebar:
    st.header("⚙️ Cấu hình")
    provider = st.selectbox("Chọn Model AI", ["Groq", "Gemini"])
    user_key = st.text_input("Dán API Key vào đây", type="password")
    st.divider()
    selected_key = st.selectbox("Chọn Mã bài tập (1-9)", list(DATA_BTL.keys()))
    current_data = DATA_BTL[selected_key]
    btl_code = selected_key.split(":")[0]

st.title("🎓 Hệ thống Tạo Tiểu luận Học thuật chuẩn VNU")

col1, col2 = st.columns([1, 2])
with col1:
    ten_hv = st.text_input("Họ tên học viên", "Đặng Nhật Minh")
    sbd = st.text_input("SBD", "39")
    mon_hoc = st.text_input("Tên môn học", selected_key.split(": ")[1])
    de_tai = st.text_area("Tên đề tài (In bìa)", current_data["de_tai"], height=100)

with col2:
    yeu_cau = st.text_area("Đề bài / Câu hỏi chi tiết", current_data["cau_hoi"], height=220)

if st.button("🚀 BẮT ĐẦU VIẾT TIỂU LUẬN (12-15 TRANG)"):
    api_key = user_key or st.secrets.get(f"{provider.upper()}_API_KEY")
    if not api_key: st.error("Vui lòng cung cấp API Key!"); st.stop()
    if not os.path.exists("Bia.docx"): st.error("Thiếu file Bia.docx!"); st.stop()

    status = st.empty(); prog = st.progress(0)
    
    # Bước 1: Lập dàn ý chương hồi
    status.info("📍 Đang xây dựng dàn ý chương hồi học thuật...")
    outline_p = (
        f"Hãy đóng vai Tiến sĩ Giáo dục, lập dàn ý tiểu luận gồm 6 phần chính cho đề tài: {de_tai}. "
        "Yêu cầu dàn ý gồm: Mở đầu, Các chương nội dung lý luận, Chương thực tiễn tại Việt Nam, Kết luận và Danh mục tài liệu tham khảo. "
        "Chỉ trả về tiêu đề các mục."
    )
    outline = call_ai(api_key, provider, outline_p)
    sections = [s.strip() for s in outline.split('\n') if len(s.strip()) > 10][:6]

    # Bước 2: Viết nội dung chi tiết từng chương
    final_content = []
    for i, sec in enumerate(sections):
        status.warning(f"✍️ Đang biên soạn Chương {i+1}: {sec}")
        write_p = (
            f"Hãy viết nội dung tiểu luận học thuật khoảng 1000 từ cho mục '{sec}' của đề tài '{de_tai}'. "
            f"Nội dung phải giải quyết triệt để yêu cầu: {yeu_cau}. "
            "Yêu cầu văn phong chuyên sâu, không dùng ký tự Markdown, không lời dẫn AI."
        )
        content = call_ai(api_key, provider, write_p)
        final_content.append((sec, strictly_clean_content(content)))
        prog.progress(int((i+1)/len(sections)*100))

    # ===== 5. XUẤT FILE WORD =====
    doc = Document("Bia.docx")
    
    # Fill Bìa
    for p in doc.paragraphs:
        maps = {"{{HO_TEN}}": ten_hv, "{{SBD}}": sbd, "{{MON_HOC}}": mon_hoc.upper(), "{{TEN_DE_TAI}}": de_tai.upper()}
        for k, v in maps.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text:
                        run.text = run.text.replace(k, str(v))
                        set_font_style(run)

    # Trang Mục lục
    doc.add_page_break()
    p_toc = doc.add_paragraph(); p_toc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font_style(p_toc.add_run("MỤC LỤC"), size=16, bold=True)
    add_toc(doc.add_paragraph())

    # Cấu hình lề & Đánh số trang
    doc.add_page_break()
    section = doc.sections[-1]
    section.top_margin, section.bottom_margin, section.left_margin, section.right_margin = Cm(2.0), Cm(2.0), Cm(3.0), Cm(2.0)
    
    doc.sections[0].different_first_page_header_footer = True
    f_p = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
    f_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_pg = f_p.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin'); run_pg._r.append(fld1)
    fld2 = OxmlElement('w:instrText'); fld2.text = "PAGE"; run_pg._r.append(fld2)
    fld3 = OxmlElement('w:fldChar'); fld3.set(qn('w:fldCharType'), 'end'); run_pg._r.append(fld3)

    # Đổ nội dung tiểu luận
    for idx, (title, text) in enumerate(final_content):
        # Tiêu đề chương (Heading 1)
        h = doc.add_paragraph(style='Heading 1')
        clean_title = re.sub(r'^\d+[\.\s\-]+', '', title)
        run_h = h.add_run(f"PHẦN {idx + 1}: {clean_title.upper()}")
        set_font_style(run_h, size=14, bold=True)
        
        # Nội dung văn xuôi
        for line in text.split('\n'):
            if line.strip():
                p = doc.add_paragraph(line.strip())
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.line_spacing = current_data["spacing"]
                run_p = p.add_run() if not p.runs else p.runs[0]
                set_font_style(run_p, size=14)

    # Lưu file
    file_name = f"TieuLuan_{btl_code}_{remove_vietnamese_accent(ten_hv)}.docx"
    buffer = io.BytesIO(); doc.save(buffer); buffer.seek(0)
    status.success(f"✅ Đã tạo xong tiểu luận bài {btl_code}!")
    st.download_button(label=f"📥 Tải xuống {file_name}", data=buffer, file_name=file_name)