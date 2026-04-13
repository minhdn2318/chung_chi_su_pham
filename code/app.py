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
    """Làm sạch API Key nếu dán nhầm dấu nháy hoặc tên biến"""
    if not key: return ""
    key = key.strip().strip("'").strip('"')
    if "API_KEY=" in key:
        key = key.split("=")[-1].strip().strip('"')
    return key

def add_page_number(paragraph, position):
    """Chèn số trang tự động vào Header/Footer"""
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
    """Thay thế các nhãn {{...}} trên trang bìa"""
    for p in doc.paragraphs:
        for key, value in data_dict.items():
            if key in p.text:
                # Thay thế nhưng giữ nguyên định dạng của paragraph
                p.text = p.text.replace(key, str(value))

# ===== 2. CẤU HÌNH SIDEBAR =====
st.set_page_config(page_title="HPU2 Report Generator PRO", layout="wide", page_icon="🎓")
st.title("🎓 Hệ thống Tạo BTL chuẩn Quy cách HPU2")

with st.sidebar:
    st.header("⚙️ Cấu hình AI")
    provider = st.selectbox("Chọn Provider", ["Gemini", "Groq"])
    
    if provider == "Gemini":
        model_choice = st.selectbox("Model ID", ["gemini-2.5-pro", "gemini-2.5-flash"])
        st.info("💡 Hệ thống tự động nghỉ 8s sau mỗi chương để tránh lỗi Quota 429.")
    else:
        model_map = {
            "Llama 3.3 70B Versatile": "llama-3.3-70b-versatile",
            "Qwen 2.5 32B": "qwen-2.5-32b",
            "Llama 3.1 70B": "llama-3.1-70b-versatile"
        }
        model_label = st.selectbox("Model ID", list(model_map.keys()))
        model_choice = model_map[model_label]

    # Lấy API Key từ Secrets hoặc nhập tay
    default_key = st.secrets.get(f"{provider.upper()}_API_KEY", "")
    api_key_raw = st.text_input(f"{provider} API Key", value=default_key, type="password")
    final_api_key = clean_api_key(api_key_raw)

    st.divider()
    st.subheader("📏 Thông số trình bày")
    hoc_phan_ui = st.selectbox("Học phần", [
        "Giáo dục học đại cương", 
        "Sử dụng phương tiện KH", 
        "Lý Luận dạy học đại học",
        "Khác"
    ])

    # Preset mặc định (Mặc định đánh số trang là BOTTOM_RIGHT theo yêu cầu)
    presets = {
        "Giáo dục học đại cương": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5, "pos": "BOTTOM_RIGHT"},
        "Sử dụng phương tiện KH": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5, "pos": "BOTTOM_RIGHT"},
        "Lý Luận dạy học đại học": {"m": (2.0, 2.0, 2.0, 2.0), "sp": 1.3, "pos": "BOTTOM_RIGHT"},
        "Khác": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5, "pos": "BOTTOM_RIGHT"}
    }
    cp = presets.get(hoc_phan_ui, presets["Khác"])

    col_l, col_r = st.columns(2)
    with col_l:
        m_top = st.number_input("Trên (cm)", 1.0, 5.0, cp["m"][0])
        m_left = st.number_input("Trái (cm)", 1.0, 5.0, cp["m"][2])
    with col_r:
        m_bottom = st.number_input("Dưới (cm)", 1.0, 5.0, cp["m"][1])
        m_right = st.number_input("Phải (cm)", 1.0, 5.0, cp["m"][3])

    line_sp = st.number_input("Cách dòng", 1.0, 2.5, cp["sp"])
    font_sz = st.number_input("Cỡ chữ", 12, 16, 14)
    page_pos = st.selectbox("Vị trí số trang", ["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"], 
                            index=["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"].index(cp["pos"]))

# ===== 3. NHẬP LIỆU CHÍNH =====
col_in1, col_in2 = st.columns([1, 2])
with col_in1:
    ten_hoc_vien = st.text_input("Học viên", "Đặng Nhật Minh")
    so_bao_danh = st.text_input("Số báo danh", "SBD-2026-001")
    ten_mon = st.text_input("Tên học phần (CHUYEN_DE)", hoc_phan_ui)
with col_in2:
    default_de = "Bằng lý luận và thực tiễn, anh (chị) hãy phân tích các chức năng xã hội của giáo dục. Từ đó, anh (chị) hãy liên hệ với việc thực hiện các chức năng này ở Việt Nam."
    de_bai = st.text_area("Đề tài chi tiết (TEN_CHU_DE)", default_de, height=145)

# ===== 4. LOGIC AI =====
def call_ai(key, provider, prompt, model_name):
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=key)
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            time.sleep(8) # Chống lỗi 429 cho Gemini 2.5
            return response.text
        else:
            from groq import Groq
            client = Groq(api_key=key)
            response = client.chat.completions.create(model=model_name, messages=[{"role": "user", "content": prompt}])
            time.sleep(2)
            return response.choices[0].message.content
    except Exception as e:
        if "429" in str(e):
            st.warning("⚠️ Đang chạm giới hạn. Nghỉ 15s...")
            time.sleep(15); return call_ai(key, provider, prompt, model_name)
        return f"ERROR: {str(e)}"

# ===== 5. THỰC THI =====
if st.button("🚀 Bắt đầu tạo Bài tập lớn (12 Trang)"):
    if not final_api_key:
        st.error("❌ Vui lòng nhập API Key!"); st.stop()

    status = st.empty(); prog = st.progress(0)
    
    # Bước 1: Dàn ý
    status.info("📝 Bước 1: Lập dàn ý chuyên sâu...")
    outline = call_ai(final_api_key, provider, f"Lập dàn ý bài tập lớn 12 trang: {de_bai}. Chỉ trả về các mục 1, 2, 3...", model_choice)
    sections = [s for s in outline.split('\n') if len(s.strip()) > 5]
    prog.progress(10)

    # Bước 2: Viết chương
    full_text = ""
    for idx, sec in enumerate(sections):
        status.write(f"⏳ Đang viết mục {idx+1}/{len(sections)}: {sec}")
        content = call_ai(final_api_key, provider, f"Viết bài luận chuyên sâu (1000 từ) cho phần '{sec}' của đề tài '{de_bai}'. Phong cách học thuật, không markdown.", model_choice)
        full_text += f"\n\n{sec}\n\n{content}"
        prog.progress(10 + int((idx+1)/len(sections)*80))

    # Bước 3: Đổ vào Template.docx
    status.info("📄 Bước 3: Đang áp dụng Template & Formatting...")
    template_path = "DeThi/Template.docx"
    
    doc = Document(template_path) if os.path.exists(template_path) else Document()
    
    # Đổ dữ liệu vào trang bìa
    data_fill = {
        "{{CHUYEN_DE}}": ten_mon.upper(),
        "{{TEN_CHU_DE}}": de_bai.upper(),
        "{{HO_TEN}}": ten_hoc_vien,
        "{{SBD}}": so_bao_danh
    }
    replace_placeholders(doc, data_fill)

    # Margins & Page Numbers
    s = doc.sections[0]
    s.top_margin, s.bottom_margin = Cm(m_top), Cm(m_bottom)
    s.left_margin, s.right_margin = Cm(m_left), Cm(m_right)
    
    target_p = s.header.paragraphs[0] if page_pos == "TOP_CENTER" else (s.footer.paragraphs[0] if s.footer.paragraphs else s.footer.add_paragraph())
    add_page_number(target_p, page_pos)

    # Nội dung
    for line in full_text.split('\n'):
        if line.strip():
            p = doc.add_paragraph(line)
            p.paragraph_format.line_spacing = line_sp
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = p.runs[0] if p.runs else p.add_run(line)
            run.font.name = "Times New Roman"
            run.font.size = Pt(font_sz)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "Times New Roman")

    buffer = io.BytesIO()
    doc.save(buffer); buffer.seek(0)
    
    prog.progress(100)
    status.success(f"🎉 Đã tạo xong báo cáo môn {ten_mon}!")
    st.download_button("📥 Tải BTL Hoàn Thiện", buffer, file_name=f"BTL_{ten_hoc_vien}.docx")