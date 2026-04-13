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

# ===== 1. KHỞI TẠO & TIỆN ÍCH =====
load_dotenv()

def clean_api_key(key):
    """Lọc sạch dấu nháy và tên biến nếu người dùng dán nhầm"""
    if not key: return ""
    key = key.strip().strip("'").strip('"')
    if "API_KEY=" in key:
        key = key.split("=")[-1].strip().strip('"')
    return key

def add_page_number(paragraph, position):
    """Chèn số trang tự động vào Header/Footer theo quy định HPU2"""
    if "CENTER" in position:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif "RIGHT" in position:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)
    instrText = OxmlElement('w:instrText'); instrText.text = "PAGE"
    run._r.append(instrText)
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar2)

def replace_placeholders(doc, data_dict):
    """Thay thế thông tin học viên trên trang bìa của Template"""
    for p in doc.paragraphs:
        for key, value in data_dict.items():
            if key in p.text:
                p.text = p.text.replace(key, str(value))

# ===== 2. CẤU HÌNH GIAO DIỆN & SIDEBAR =====
st.set_page_config(page_title="HPU2 Report Generator PRO", layout="wide", page_icon="🎓")
st.title("🎓 Hệ thống Tạo BTL chuẩn Quy cách HPU2")

with st.sidebar:
    st.header("⚙️ Cấu hình Model")
    provider = st.selectbox("Chọn AI Provider", ["Gemini", "Groq"])
    
    if provider == "Gemini":
        model_choice = st.selectbox("Model ID", ["gemini-2.5-pro", "gemini-2.5-flash"])
    else:
        model_map = {
            "Llama 3.3 70B Versatile": "llama-3.3-70b-versatile",
            "Qwen 2.5 32B": "qwen-2.5-32b",
            "GPT OSS 20B 128k": "GPT OSS 20B 128k"
        }
        m_label = st.selectbox("Model ID", list(model_map.keys()))
        model_choice = model_map[m_label]

    # Lấy API Key (Ưu tiên Secrets của Streamlit)
    default_key = st.secrets.get(f"{provider.upper()}_API_KEY", "")
    api_key_raw = st.text_input(f"Dán {provider} API Key vào đây", value=default_key, type="password")
    api_key = clean_api_key(api_key_raw)

    st.divider()
    st.subheader("📚 Quy cách trình bày")
    hoc_phan_ui = st.selectbox("Học phần mục tiêu", [
        "Giáo dục học đại cương", 
        "Sử dụng phương tiện KH", 
        "Lý Luận dạy học đại học",
        "Khác"
    ])

    # Presets quy chuẩn HPU2
    presets = {
        "Giáo dục học đại cương": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5, "pos": "TOP_CENTER"},
        "Sử dụng phương tiện KH": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5, "pos": "BOTTOM_CENTER"},
        "Lý Luận dạy học đại học": {"m": (2.0, 2.0, 2.0, 2.0), "sp": 1.3, "pos": "BOTTOM_RIGHT"},
        "Khác": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5, "pos": "BOTTOM_RIGHT"}
    }
    cp = presets.get(hoc_phan_ui, presets["Khác"])

    # Điều chỉnh thông số (Biến global cho toàn app)
    col_l, col_r = st.columns(2)
    with col_l:
        m_top = st.number_input("Lề trên (cm)", 0.0, 5.0, cp["m"][0])
        m_left = st.number_input("Lề trái (cm)", 0.0, 5.0, cp["m"][2])
    with col_r:
        m_bottom = st.number_input("Lề dưới (cm)", 0.0, 5.0, cp["m"][1])
        m_right = st.number_input("Lề phải (cm)", 0.0, 5.0, cp["m"][3])

    line_sp = st.number_input("Cách dòng", 1.0, 2.5, cp["sp"])
    font_sz = st.number_input("Cỡ chữ (font_sz)", 12, 16, 14) # Đảm bảo tên biến là font_sz
    page_pos = st.selectbox("Vị trí số trang", ["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"], 
                            index=["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"].index(cp["pos"]))

# ===== 3. VÙNG NHẬP LIỆU CHÍNH =====
col_in1, col_in2 = st.columns([1, 2])
with col_in1:
    ten_hoc_vien = st.text_input("Học viên", "Đặng Nhật Minh")
    ten_mon = st.text_input("Môn bài làm", hoc_phan_ui)
with col_in2:
    default_de = "Bằng lý luận và thực tiễn, anh (chị) hãy phân tích các chức năng xã hội của giáo dục. Từ đó, anh (chị) hãy liên hệ với việc thực hiện các chức năng này ở Việt Nam."
    de_bai = st.text_area("Đề tài chi tiết", default_de, height=120)

# ===== 4. LOGIC GỌI AI =====
def call_ai(key, provider, prompt, model_name):
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=key)
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            # Nghỉ 6s để lách luật 20 request/phút của Gemini 2.5 Flash
            time.sleep(6) 
            return response.text
        else:
            from groq import Groq
            client = Groq(api_key=key)
            response = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": prompt}]
            )
            time.sleep(1)
            return response.choices[0].message.content
    except Exception as e:
        if "429" in str(e):
            st.warning("⚠️ API bận (429). Đang chờ 15s...")
            time.sleep(15)
            return call_ai(key, provider, prompt, model_name)
        return f"ERROR: {str(e)}"

# ===== 5. THỰC THI CHÍNH =====
if st.button("🚀 Bắt đầu tạo Bài tập lớn (12 Trang)"):
    if not api_key:
        st.error("❌ Thiếu API Key!")
        st.stop()

    status = st.empty()
    prog = st.progress(0)
    
    # 1. Gen Dàn ý
    status.info("📝 Bước 1: Đang lập dàn ý chuyên sâu...")
    outline = call_ai(api_key, provider, f"Lập dàn ý bài tập lớn 12 trang: {de_bai}. Chỉ trả về các đầu mục.", model_choice)
    sections = [s for s in outline.split('\n') if len(s.strip()) > 5]
    prog.progress(10)

    # 2. Gen Từng Chương
    full_report = ""
    for idx, sec in enumerate(sections):
        status.write(f"⏳ Đang viết: {sec}")
        content = call_ai(api_key, provider, f"Viết nội dung học thuật (1000 từ) cho mục '{sec}' của đề tài '{de_bai}'. Không dùng markdown.", model_choice)
        full_report += f"\n\n{sec}\n\n{content}"
        prog.progress(10 + int((idx+1)/len(sections)*80))

    # 3. Xuất Word
    status.info("📄 Bước 3: Đang đổ dữ liệu vào Template...")
    template_path = "DeThi/Template.docx" # Khớp với cấu hình folder của bạn
    
    doc = Document(template_path) if os.path.exists(template_path) else Document()
    replace_placeholders(doc, {"Đặng Nhật Minh": ten_hoc_vien, "TÊN CHỦ ĐỀ": de_bai.upper()})

    # Margins
    s = doc.sections[0]
    s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = Cm(m_top), Cm(m_bottom), Cm(m_left), Cm(m_right)
    
    # Page Number
    target_para = s.header.paragraphs[0] if page_pos == "TOP_CENTER" else (s.footer.paragraphs[0] if s.footer.paragraphs else s.footer.add_paragraph())
    add_page_number(target_para, page_pos)

    # Content Formatting
    for line in full_report.split('\n'):
        if line.strip():
            p = doc.add_paragraph(line)
            p.paragraph_format.line_spacing = line_sp
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = p.add_run() if not p.runs else p.runs[0]
            run.font.name = "Times New Roman"
            run.font.size = Pt(font_sz) # <--- ĐÃ ĐẢM BẢO BIẾN font_sz TỒN TẠI
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "Times New Roman")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    prog.progress(100)
    status.success("🎉 Hoàn thành bài tập lớn 12 trang!")
    st.download_button("📥 Tải BTL Hoàn Thiện", buffer, file_name=f"BTL_{ten_hoc_vien}.docx")