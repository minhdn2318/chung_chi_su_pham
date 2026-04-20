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

# ===== 1. DATABASE 9 HỌC PHẦN (CẤU HÌNH ĐỘNG TOÀN DIỆN) =====
# Cấu hình "le": (Trên, Dưới, Trái, Phải)
DATA_BTL = {
    "HP_01: Giáo dục học đại cương": {
        "de_tai": "PHÂN TÍCH CÁC CHỨC NĂNG XÃ HỘI CỦA GIÁO DỤC", 
        "cau_hoi": "Phân tích các chức năng xã hội của giáo dục. Liên hệ thực hiện tại Việt Nam hiện nay.",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14, "page_pos": "LEFT"
    },
    "HP_02: Lý luận dạy học đại học": {
        "de_tai": "VẬN DỤNG NGUYÊN TẮC DẠY HỌC HIỆN ĐẠI",
        "cau_hoi": "Phân tích hệ thống nguyên tắc dạy học đại học và đề xuất phương án vận dụng.",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.3, "font_sz": 14, "page_pos": "LEFT"
    },
    "HP_03: Sử dụng phương tiện kỹ thuật KH": {
        "de_tai": "CÔNG NGHỆ TRONG ĐỔI MỚI DẠY HỌC",
        "cau_hoi": "Trình bày phương tiện hiện đại và thực trạng ứng dụng tại cơ sở.",
        "le": (2.5, 2.0, 3.5, 2.0), "spacing": 1.5, "font_sz": 14, "page_pos": "LEFT"
    },
    "HP_04: Tâm lý học đại cương": {
        "de_tai": "YẾU TỐ HÌNH THÀNH NHÂN CÁCH",
        "cau_hoi": "Phân tích cấu trúc nhân cách và các yếu tố ảnh hưởng đến sự hình thành nhân cách.",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14, "page_pos": "LEFT"
    },
    "HP_05: Quản lý nhà nước về Giáo dục": {
        "de_tai": "QUẢN LÝ GIÁO DỤC KỶ NGUYÊN SỐ",
        "cau_hoi": "Các chức năng quản lý nhà nước và giải pháp nâng cao hiệu quả quản lý giáo dục.",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14, "page_pos": "LEFT"
    },
    "HP_06: Phát triển chương trình đào tạo": {
        "de_tai": "THIẾT KẾ ĐỀ CƯƠNG CHUẨN ĐẦU RA",
        "cau_hoi": "Quy trình phát triển chương trình và xây dựng ma trận mục tiêu cho học phần.",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14, "page_pos": "LEFT"
    },
    "HP_07: Đánh giá trong giáo dục": {
        "de_tai": "ĐỔI MỚI KIỂM TRA ĐÁNH GIÁ NĂNG LỰC",
        "cau_hoi": "Các hình thức đánh giá và thiết kế bộ công cụ đánh giá theo định hướng năng lực.",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14, "page_pos": "LEFT"
    },
    "HP_08: Phương pháp nghiên cứu khoa học": {
        "de_tai": "XÂY DỰNG ĐỀ CƯƠNG NGHIÊN CỨU GIÁO DỤC",
        "cau_hoi": "Quy trình nghiên cứu khoa học và lập kế hoạch nghiên cứu cho đề tài cụ thể.",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14, "page_pos": "LEFT"
    },
    "HP_09: Kỹ năng mềm trong dạy học": {
        "de_tai": "RÈN LUYỆN KỸ NĂNG GIAO TIẾP SƯ PHẠM",
        "cau_hoi": "Tầm quan trọng kỹ năng mềm và các tình huống xử lý sư phạm điển hình.",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5, "font_sz": 14, "page_pos": "LEFT"
    }
}

# ===== 2. TIỆN ÍCH ĐỊNH DẠNG =====
def remove_vietnamese_accent(s):
    s = s.replace("Đ", "D").replace("đ", "d")
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
    return re.sub(r'[^a-zA-Z0-9_]', '_', s)

def add_page_number(paragraph, alignment_code):
    """Đánh số trang động: LEFT, CENTER, RIGHT"""
    if alignment_code == "LEFT":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    elif alignment_code == "RIGHT":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin'); run._r.append(fldChar1)
    instrText = OxmlElement('w:instrText'); instrText.text = "PAGE"; run._r.append(instrText)
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end'); run._r.append(fldChar2)

def set_font_style(run, size=14, bold=False):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), "Times New Roman")
    rFonts.set(qn('w:hAnsi'), "Times New Roman")
    rFonts.set(qn('w:eastAsia'), "Times New Roman")
    rPr.insert(0, rFonts)

def strictly_clean_content(text):
    ai_patterns = [r"Dưới đây là.*:", r"Đoạn văn đã được chỉnh sửa.*:", r"Tôi đã thực hiện.*:", r"Nội dung biên tập.*:"]
    for pattern in ai_patterns:
        text = re.sub(pattern, "", text, flags=re.IGNORECASE | re.DOTALL)
    return re.sub(r'[*#_~-]', '', text).strip()

# ===== 3. GIAO DIỆN VÀ LOGIC AI =====
st.set_page_config(page_title="Hệ thống BTL Pro", layout="wide")

with st.sidebar:
    st.header("⚙️ Cấu hình Hệ thống")
    provider = st.selectbox("AI Model", ["Groq", "Gemini"])
    user_key = st.text_input("API Key (Tùy chọn)", type="password")
    
    st.divider()
    selected_key = st.selectbox("Chọn Mã Học Phần", list(DATA_BTL.keys()))
    conf = DATA_BTL[selected_key]
    
    st.subheader("📏 Thông số mặc định (tự động)")
    st.info(f"Lề: {conf['le']} | Dãn dòng: {conf['spacing']} | Font: {conf['font_sz']}")
    btl_code = selected_key.split(":")[0]

st.title("🎓 Công cụ Tạo BTL chuẩn VNU (Dynamic Config)")

col1, col2 = st.columns([1, 2])
with col1:
    ten_hv = st.text_input("Học viên", "Đặng Nhật Minh")
    sbd = st.text_input("SBD", "39")
    mon_hoc = st.text_input("Môn", selected_key.split(": ")[1])
    de_tai = st.text_area("Tên đề tài", conf["de_tai"], height=100)

with col2:
    yeu_cau = st.text_area("Nội dung câu hỏi Excel", conf["cau_hoi"], height=215)

def call_ai(key, provider, prompt):
    sys = "Bạn là máy soạn thảo tiểu luận VNU. CHỈ xuất nội dung học thuật. KHÔNG dẫn dắt, KHÔNG Markdown."
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=key)
            return genai.GenerativeModel('gemini-1.5-pro').generate_content(f"{sys}\n\n{prompt}").text
        else:
            from groq import Groq
            client = Groq(api_key=key)
            res = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=[{"role": "system", "content": sys}, {"role": "user", "content": prompt}])
            return res.choices[0].message.content
    except Exception as e: return f"Lỗi: {e}"

# ===== 4. QUY TRÌNH XUẤT FILE =====
if st.button("🚀 XUẤT TIỂU LUẬN HOÀN THIỆN"):
    api_key = user_key or st.secrets.get(f"{provider.upper()}_API_KEY")
    if not api_key: st.error("Thiếu Key!"); st.stop()
    if not os.path.exists("Bia.docx"): st.error("Thiếu Bia.docx!"); st.stop()

    status = st.empty(); prog = st.progress(0)
    
    # 1. AI viết nội dung
    outline = call_ai(api_key, provider, f"Lập dàn ý 6 mục cho đề tài: {de_tai}")
    sections = [s.strip() for s in outline.split('\n') if len(s.strip()) > 10][:6]
    final_content = []
    for i, sec in enumerate(sections):
        status.info(f"✍️ Đang biên soạn Chương {i+1}: {sec}")
        raw = call_ai(api_key, provider, f"Viết 1000 từ chuyên sâu cho mục '{sec}' của đề tài '{de_tai}'. Bám sát: {yeu_cau}")
        final_content.append((sec, strictly_clean_content(raw)))
        prog.progress(int((i+1)/len(sections)*100))

    # 2. Xử lý Word
    doc = Document("Bia.docx")
    
    # Fill Bìa
    for p in doc.paragraphs:
        maps = {"{{HO_TEN}}": ten_hv, "{{SBD}}": sbd, "{{MON_HOC}}": mon_hoc.upper(), "{{TEN_DE_TAI}}": de_tai.upper()}
        for k, v in maps.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text:
                        run.text = run.text.replace(k, str(v))
                        set_font_style(run, size=14)

    # 3. Trang Mục lục
    doc.add_page_break()
    p_toc_title = doc.add_paragraph(); p_toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font_style(p_toc_title.add_run("MỤC LỤC"), size=16, bold=True)
    
    # Chèn mã mục lục (Field TOC)
    run_toc = doc.add_paragraph().add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin'); run_toc._r.append(fld1)
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = 'TOC \\o "1-3" \\h \\z \\u'
    run_toc._r.append(instr)
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'separate'); run_toc._r.append(fld2)
    fld3 = OxmlElement('w:fldChar'); fld3.set(qn('w:fldCharType'), 'end'); run_toc._r.append(fld3)

    # 4. Trang Nội dung & Cấu hình lề động
    doc.add_page_break()
    section = doc.sections[-1]
    
    # Lấy thông số từ DATA_BTL
    m = conf["le"]
    section.top_margin, section.bottom_margin = Cm(m[0]), Cm(m[1])
    section.left_margin, section.right_margin = Cm(m[2]), Cm(m[3])
    
    # Đánh số trang động theo vị trí trong DATA_BTL
    doc.sections[0].different_first_page_header_footer = True 
    footer = section.footer
    f_p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    add_page_number(f_p, conf["page_pos"])

    # 5. Đổ nội dung bài làm
    for idx, (title, text) in enumerate(final_content):
        # Tiêu đề mục (Heading 1)
        h = doc.add_paragraph(style='Heading 1')
        clean_title = re.sub(r'^\d+[\.\s\-]+', '', title)
        run_h = h.add_run(f"{idx + 1}. {clean_title.upper()}")
        set_font_style(run_h, size=conf["font_sz"], bold=True)
        
        # Văn bản tiểu luận
        for line in text.split('\n'):
            if line.strip():
                p = doc.add_paragraph(line.strip())
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.line_spacing = conf["spacing"]
                run_p = p.add_run() if not p.runs else p.runs[0]
                set_font_style(run_p, size=conf["font_sz"])

    # Xuất file
    file_name = f"BTL_{btl_code}_{remove_vietnamese_accent(ten_hv)}.docx"
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    status.success(f"✅ Hoàn thành bài {btl_code}!")
    st.download_button(label=f"📥 Tải xuống {file_name}", data=buf, file_name=file_name)