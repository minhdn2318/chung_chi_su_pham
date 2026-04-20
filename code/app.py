import streamlit as st
import os
import io
import re
import unicodedata
import time
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ===== 1. DATABASE 9 HỌC PHẦN (PRESETS) =====
DATA_BTL = {
    "HP_01: Giáo dục học đại cương": {"de_tai": "PHÂN TÍCH CÁC CHỨC NĂNG XÃ HỘI CỦA GIÁO DỤC", "cau_hoi": "Phân tích các chức năng xã hội của giáo dục. Liên hệ Việt Nam hiện nay.", "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14},
    "HP_02: Lý luận dạy học đại học": {"de_tai": "VẬN DỤNG NGUYÊN TẮC DẠY HỌC HIỆN ĐẠI", "cau_hoi": "Phân tích hệ thống nguyên tắc dạy học và đề xuất phương án vận dụng.", "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.3, "font_sz": 14},
    "HP_03: Sử dụng phương tiện kỹ thuật KH": {"de_tai": "CÔNG NGHỆ TRONG ĐỔI MỚI DẠY HỌC", "cau_hoi": "Trình bày các phương tiện kỹ thuật hiện đại và thực trạng ứng dụng.", "le": (2.5, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14},
    "HP_04: Tâm lý học đại cương": {"de_tai": "CÁC YẾU TỐ HÌNH THÀNH NHÂN CÁCH", "cau_hoi": "Phân tích cấu trúc nhân cách và các yếu tố ảnh hưởng.", "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14},
    "HP_05: Quản lý nhà nước về Giáo dục": {"de_tai": "QUẢN LÝ GIÁO DỤC KỶ NGUYÊN SỐ", "cau_hoi": "Các chức năng quản lý nhà nước và giải pháp nâng cao hiệu quả.", "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14},
    "HP_06: Phát triển chương trình đào tạo": {"de_tai": "THIẾT KẾ ĐỀ CƯƠNG THEO CHUẨN ĐẦU RA", "cau_hoi": "Quy trình phát triển chương trình và xây dựng ma trận mục tiêu.", "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14},
    "HP_07: Đánh giá trong giáo dục": {"de_tai": "ĐỔI MỚI KIỂM TRA ĐÁNH GIÁ NĂNG LỰC", "cau_hoi": "Các hình thức đánh giá và thiết kế bộ công cụ đánh giá.", "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14},
    "HP_08: Phương pháp nghiên cứu khoa học": {"de_tai": "XÂY DỰNG ĐỀ CƯƠNG NGHIÊN CỨU GIÁO DỤC", "cau_hoi": "Quy trình nghiên cứu khoa học và lập kế hoạch nghiên cứu.", "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14},
    "HP_09: Kỹ năng mềm trong dạy học": {"de_tai": "RÈN LUYỆN KỸ NĂNG GIAO TIẾP SƯ PHẠM", "cau_hoi": "Tầm quan trọng kỹ năng mềm và xử lý tình huống sư phạm.", "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14},
}

# ===== 2. TIỆN ÍCH WORD & LÀM SẠCH =====
def remove_vietnamese_accent(s):
    s = s.replace("Đ", "D").replace("đ", "d")
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
    return re.sub(r'[^a-zA-Z0-9_]', '_', s)

def add_page_number(paragraph, alignment_code):
    """Hỗ trợ Left, Center, Right"""
    if "LEFT" in alignment_code: paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    elif "RIGHT" in alignment_code: paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else: paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin'); run._r.append(fld1)
    instr = OxmlElement('w:instrText'); instr.text = "PAGE"; run._r.append(instr)
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end'); run._r.append(fld2)

def set_font_style(run, font_name="Times New Roman", size=14, bold=False):
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rPr.insert(0, rFonts)

def strictly_clean_content(text):
    ai_patterns = [r"Dưới đây là.*:", r"Đoạn văn đã được chỉnh sửa.*:", r"Chắc chắn rồi.*:", r"Tôi là trợ lý.*:"]
    for pattern in ai_patterns:
        text = re.sub(pattern, "", text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'[*#_~-]', '', text)
    return text.strip()

# ===== 3. GIAO DIỆN (FULL FEATURES) =====
st.set_page_config(page_title="AI Report Ultimate", layout="wide")

with st.sidebar:
    st.header("⚙️ Cấu hình Model")
    provider = st.selectbox("Provider", ["Groq", "Gemini"])
    if provider == "Groq":
        model_choice = st.selectbox("Model", ["llama-3.3-70b-versatile", "qwen-2.5-32b"])
    else:
        model_choice = st.selectbox("Model", ["gemini-1.5-pro", "gemini-2.0-flash"])
    api_key_input = st.text_input("API Key", type="password")

    st.divider()
    st.header("📏 Chỉnh sửa Lề & Font")
    selected_hp = st.selectbox("Chọn Preset Học Phần", list(DATA_BTL.keys()))
    preset = DATA_BTL[selected_hp]
    
    col_l, col_r = st.columns(2)
    with col_l:
        m_top = st.number_input("Trên (cm)", 0.0, 5.0, preset["le"][0])
        m_left = st.number_input("Trái (cm)", 0.0, 5.0, preset["le"][2])
    with col_r:
        m_bottom = st.number_input("Dưới (cm)", 0.0, 5.0, preset["le"][1])
        m_right = st.number_input("Phải (cm)", 0.0, 5.0, preset["le"][3])

    f_sz = st.number_input("Cỡ chữ", 10, 16, preset["font_sz"])
    f_sp = st.number_input("Dãn dòng", 1.0, 2.5, preset["spacing"])
    
    st.divider()
    st.header("🔢 Đánh số trang")
    pg_side = st.selectbox("Vị trí dọc", ["BOTTOM", "TOP"])
    pg_align = st.selectbox("Vị trí ngang", ["LEFT", "CENTER", "RIGHT"])

st.title("🎓 Hệ thống Tạo Tiểu luận Ultimate (VNU-HPU2 Edition)")

col_in1, col_in2 = st.columns([1, 2])
with col_in1:
    st.subheader("👤 Thông tin Bìa")
    ten_hv = st.text_input("Học viên", "Đặng Nhật Minh")
    sbd = st.text_input("SBD", "39")
    mon_hoc = st.text_input("Môn học", selected_hp.split(": ")[1])
    de_tai = st.text_area("Đề tài", preset["de_tai"], height=100)

with col_in2:
    st.subheader("🤖 Nội dung yêu cầu")
    yeu_cau = st.text_area("Câu hỏi chi tiết", preset["cau_hoi"], height=220)

# ===== 4. LOGIC AI (2-STEP: WRITE & REVIEW) =====
def call_ai(key, provider, prompt, model):
    sys = "Bạn là Tiến sĩ Giáo dục. CHỈ xuất nội dung học thuật. KHÔNG dẫn dắt, KHÔNG Markdown."
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=key)
            return genai.GenerativeModel(model).generate_content(f"{sys}\n\n{prompt}").text
        else:
            from groq import Groq
            client = Groq(api_key=key)
            res = client.chat.completions.create(model=model, messages=[{"role": "system", "content": sys}, {"role": "user", "content": prompt}])
            return res.choices[0].message.content
    except Exception as e: return f"Lỗi: {e}"

if st.button("🚀 XUẤT TIỂU LUẬN FULL OPTION"):
    key = api_key_input or st.secrets.get(f"{provider.upper()}_API_KEY")
    if not key: st.error("Thiếu Key!"); st.stop()
    if not os.path.exists("Bia.docx"): st.error("Thiếu Bia.docx!"); st.stop()

    status = st.empty(); prog = st.progress(0)
    
    # 1. Dàn ý
    outline = call_ai(key, provider, f"Lập dàn ý 6 mục cho tiểu luận: {de_tai}", model_choice)
    sections = [s.strip() for s in outline.split('\n') if len(s.strip()) > 10][:6]

    # 2. Viết & Review từng phần
    final_content = []
    for i, sec in enumerate(sections):
        status.info(f"✍️ Đang xử lý Chương {i+1}: {sec}")
        # Bước 1: Viết nháp
        draft = call_ai(key, provider, f"Viết 1000 từ học thuật cho mục '{sec}' đề tài '{de_tai}'. Bám sát câu hỏi: {yeu_cau}", model_choice)
        # Bước 2: Review/Humanize (Làm cho giống người viết)
        review_p = f"Hãy biên tập lại đoạn văn sau: Loại bỏ dấu vết AI, sửa câu văn cho tự nhiên, đảm bảo tính học thuật cao. CẤM dẫn dắt: {draft}"
        polished = call_ai(key, provider, review_p, model_choice)
        
        final_content.append((sec, strictly_clean_content(polished)))
        prog.progress(int((i+1)/len(sections)*100))

    # 3. Tạo Word
    doc = Document("Bia.docx")
    # Thay thế thẻ bìa
    for p in doc.paragraphs:
        maps = {"{{HO_TEN}}": ten_hv, "{{SBD}}": sbd, "{{MON_HOC}}": mon_hoc.upper(), "{{TEN_DE_TAI}}": de_tai.upper()}
        for k, v in maps.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text: run.text = run.text.replace(k, str(v)); set_font_style(run)

    # Mục lục
    doc.add_page_break()
    p_toc = doc.add_paragraph(); p_toc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font_style(p_toc.add_run("MỤC LỤC"), size=16, bold=True)
    run_toc = doc.add_paragraph().add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin'); run_toc._r.append(fld1)
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = 'TOC \\o "1-3" \\h \\z \\u'
    run_toc._r.append(instr)
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'separate'); run_toc._r.append(fld2)
    fld3 = OxmlElement('w:fldChar'); fld3.set(qn('w:fldCharType'), 'end'); run_toc._r.append(fld3)

    # Nội dung
    doc.add_page_break()
    section = doc.sections[-1]
    section.top_margin, section.bottom_margin = Cm(m_top), Cm(m_bottom)
    section.left_margin, section.right_margin = Cm(m_left), Cm(m_right)
    
    # Đánh số trang động (Top/Bottom + Align)
    doc.sections[0].different_first_page_header_footer = True
    target_para = section.header.paragraphs[0] if pg_side == "TOP" else section.footer.paragraphs[0]
    add_page_number(target_para, pg_align)

    for idx, (title, text) in enumerate(final_content):
        h = doc.add_paragraph(style='Heading 1')
        run_h = h.add_run(f"{idx + 1}. {title.upper()}")
        set_font_style(run_h, size=f_sz, bold=True)
        for line in text.split('\n'):
            if line.strip():
                p = doc.add_paragraph(line.strip())
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.line_spacing = f_sp
                run_p = p.add_run() if not p.runs else p.runs[0]
                set_font_style(run_p, size=f_sz)

    file_final = f"BTL_{selected_hp.split(':')[0]}_{remove_vietnamese_accent(ten_hv)}.docx"
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    status.success("🎉 Đã hoàn thành bản Ultimate!")
    st.download_button(label=f"📥 Tải xuống {file_final}", data=buf, file_name=file_final)