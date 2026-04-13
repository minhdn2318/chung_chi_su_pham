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
    """Xử lý các trường hợp dán nhầm dấu nháy hoặc tên biến"""
    if not key: return ""
    key = key.strip().strip("'").strip('"')
    if "API_KEY=" in key:
        key = key.split("=")[-1].strip().strip('"')
    return key

def add_page_number(paragraph, position):
    """Chèn số trang tự động (XML Field) vào Header/Footer"""
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
    """Thay thế thông tin trên trang bìa của Template"""
    for p in doc.paragraphs:
        for key, value in data_dict.items():
            if key in p.text:
                p.text = p.text.replace(key, str(value))

# ===== 2. CẤU HÌNH SIDEBAR (MENU TRÁI) =====
st.set_page_config(page_title="HPU2 Report Generator PRO", layout="wide", page_icon="🎓")
st.title("🎓 Hệ thống Tạo BTL chuẩn Quy cách HPU2")

with st.sidebar:
    st.header("⚙️ Cấu hình Model")
    provider = st.selectbox("Chọn Provider", ["Gemini", "Groq"])
    
    if provider == "Gemini":
        model_choice = st.selectbox("Model ID", ["gemini-2.5-pro", "gemini-2.5-flash"])
        st.info("💡 Lưu ý: Free Tier có RPM thấp. Hệ thống sẽ tự nghỉ 10s sau mỗi chương.")
    else:
        model_map = {
            "Llama 3.3 70B Versatile": "llama-3.3-70b-versatile",
            "Qwen 2.5 32B Coder": "qwen-2.5-32b",
            "Llama 3.1 70B": "llama-3.1-70b-versatile"
        }
        model_label = st.selectbox("Model ID", list(model_map.keys()))
        model_choice = model_map[model_label]

    # Lấy key từ st.secrets hoặc nhập tay
    default_key = st.secrets.get(f"{provider.upper()}_API_KEY", "")
    api_key_raw = st.text_input(f"{provider} API Key", value=default_key, type="password")
    final_api_key = clean_api_key(api_key_raw)

    st.divider()
    st.subheader("📏 Quy cách Trình bày")
    hoc_phan_ui = st.selectbox("Chọn Học phần", [
        "Giáo dục học đại cương", 
        "Sử dụng phương tiện KH", 
        "Lý Luận dạy học đại học",
        "Tùy chỉnh khác"
    ])

    # Preset mặc định cho từng môn
    presets = {
        "Giáo dục học đại cương": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5, "pos": "TOP_CENTER"},
        "Sử dụng phương tiện KH": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5, "pos": "BOTTOM_CENTER"},
        "Lý Luận dạy học đại học": {"m": (2.0, 2.0, 2.0, 2.0), "sp": 1.3, "pos": "BOTTOM_RIGHT"},
        "Tùy chỉnh khác": {"m": (2.5, 3.0, 3.0, 2.0), "sp": 1.5, "pos": "BOTTOM_RIGHT"}
    }
    
    current_p = presets.get(hoc_phan_ui)
    
    # Cho phép sửa lại bên Sidebar
    col_l, col_r = st.columns(2)
    with col_l:
        m_top = st.number_input("Trên (cm)", 1.0, 5.0, current_p["m"][0])
        m_left = st.number_input("Trái (cm)", 1.0, 5.0, current_p["m"][2])
    with col_r:
        m_bottom = st.number_input("Dưới (cm)", 1.0, 5.0, current_p["m"][1])
        m_right = st.number_input("Phải (cm)", 1.0, 5.0, current_p["m"][3])

    line_sp = st.number_input("Cách dòng", 1.0, 2.5, current_p["sp"])
    page_pos = st.selectbox("Vị trí số trang", ["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"], 
                            index=["TOP_CENTER", "BOTTOM_CENTER", "BOTTOM_RIGHT"].index(current_p["pos"]))

# ===== 3. VÙNG NHẬP LIỆU CHÍNH =====
col_in1, col_in2 = st.columns([1, 2])
with col_in1:
    ten_hoc_vien = st.text_input("Học viên", "Đặng Nhật Minh")
    ten_mon_hien_thi = st.text_input("Tên môn", hoc_phan_ui)
with col_in2:
    de_bai_default = "Bằng lý luận và thực tiễn, anh (chị) hãy phân tích các chức năng xã hội của giáo dục. Từ đó, anh (chị) hãy liên hệ với việc thực hiện các chức năng này ở Việt Nam."
    de_bai = st.text_area("Đề tài chi tiết", de_bai_default, height=120)

# ===== 4. LOGIC AI (VỚI CƠ CHẾ CHỐNG LỖI 429) =====
def call_ai(key, provider, prompt, model_name):
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=key)
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            # Nghỉ để tránh 429 Free Tier
            time.sleep(10) 
            return response.text
        else:
            from groq import Groq
            client = Groq(api_key=key)
            response = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": prompt}]
            )
            time.sleep(2)
            return response.choices[0].message.content
    except Exception as e:
        if "429" in str(e):
            st.warning("⚠️ Đang chạm giới hạn Request. Đang chờ 15s để thử lại...")
            time.sleep(15)
            return call_ai(key, provider, prompt, model_name)
        return f"ERROR: {str(e)}"

# ===== 5. THỰC THI =====
if st.button("🚀 Bắt đầu tạo Bài tập lớn (12 Trang)"):
    if not final_api_key:
        st.error("❌ Vui lòng nhập API Key!")
        st.stop()

    status = st.empty()
    prog = st.progress(0)
    
    # BƯỚC 1: LẬP DÀN Ý
    status.info("📝 Bước 1: Đang lập dàn ý chi tiết...")
    outline_prompt = f"Hãy lập dàn ý chi tiết bài tập lớn đại học (khoảng 12 trang) cho đề tài: {de_bai}. Chỉ trả về các tiêu đề mục 1, 1.1, 2, 2.1..."
    outline = call_ai(final_api_key, provider, outline_prompt, model_choice)
    sections = [s.strip() for s in outline.split('\n') if len(s.strip()) > 5]
    prog.progress(15)

    # BƯỚC 2: VIẾT NỘI DUNG (Vòng lặp)
    full_report = ""
    for idx, section in enumerate(sections):
        status.write(f"⏳ Đang viết chương {idx+1}/{len(sections)}: {section}")
        chapter_prompt = f"Viết nội dung học thuật sâu sắc (khoảng 800-1000 từ) cho phần '{section}' của đề tài '{de_bai}'. Phong cách học thuật, không markdown."
        chapter = call_ai(final_api_key, provider, chapter_prompt, model_choice)
        full_report += f"\n\n{section}\n\n{chapter}"
        
        # Cập nhật tiến độ
        current_val = 15 + int((idx+1)/len(sections)*75)
        prog.progress(current_val)

    # BƯỚC 3: XUẤT WORD
    status.info("📄 Bước 3: Đang áp dụng quy cách và Template...")
    template_path = "De Thi/template.docx"
    
    if os.path.exists(template_path):
        doc = Document(template_path)
        replace_info_dict = {"Đặng Nhật Minh": ten_hoc_vien, "TÊN CHỦ ĐỀ": de_bai.upper()}
        replace_placeholders(doc, replace_info_dict)
    else:
        doc = Document()
        st.warning("⚠️ Không tìm thấy template.docx tại /De Thi/. Đang tạo file mới.")

    # Cấu hình lề
    sec_docx = doc.sections[0]
    sec_docx.top_margin = Cm(m_top)
    sec_docx.bottom_margin = Cm(m_bottom)
    sec_docx.left_margin = Cm(m_left)
    sec_docx.right_margin = Cm(m_right)

    # Đánh số trang
    if page_pos == "TOP_CENTER":
        add_page_number(sec_docx.header.paragraphs[0], "BOTTOM_RIGHT")
    else:
        # Nếu footer chưa có paragraph thì tạo mới
        target_p = sec_docx.footer.paragraphs[0] if sec_docx.footer.paragraphs else sec_docx.footer.add_paragraph()
        add_page_number(target_p, page_pos)

    # Thêm văn bản
    for line in full_report.split('\n'):
        if line.strip():
            p = doc.add_paragraph(line)
            p.paragraph_format.line_spacing = line_sp
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = p.runs[0] if p.runs else p.add_run(line)
            run.font.name = "Times New Roman"
            run.font.size = Pt(font_sz)
            # Fix lỗi font tiếng Việt khi mở Word
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "Times New Roman")

    # Hoàn tất
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    prog.progress(100)
    status.success(f"🎉 Hoàn thành! Đã tạo xong báo cáo môn {ten_mon_hien_thi}")

    st.download_button(
        label="📥 Tải Bài Tập Lớn Hoàn Thiện",
        data=buffer,
        file_name=f"BTL_{ten_hoc_vien}_{ten_mon_hien_thi}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )