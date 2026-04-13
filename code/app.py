import streamlit as st
import os
import io
import re
from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ===== KHỞI TẠO CẤU HÌNH =====
load_dotenv()

def clean_text(text):
    """Làm sạch văn bản AI nhưng giữ lại cấu trúc phân cấp"""
    text = re.sub(r'[*#]', '', text) 
    return text.strip()

def replace_placeholder(doc, old_text, new_text):
    """Tìm và thay thế văn bản trong toàn bộ file Word (bao gồm bìa)"""
    for p in doc.paragraphs:
        if old_text in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if old_text in inline[i].text:
                    inline[i].text = inline[i].text.replace(old_text, new_text)

# ===== XỬ LÝ AI =====
def call_ai(provider, api_key, prompt, model_name):
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-1.5-flash")
            return model.generate_content(prompt).text
        elif provider == "Groq":
            from groq import Groq
            client = Groq(api_key=api_key)
            response = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.6,
                max_tokens=4096
            )
            return response.choices[0].message.content
    except Exception as e:
        return f"Lỗi gọi AI: {str(e)}"

# ===== GIAO DIỆN STREAMLIT =====
st.set_page_config(page_title="HPU2 Report Generator", layout="wide")
st.title("🎓 Hệ thống Tạo Bài Tập Lớn (Template Support)")

with st.sidebar:
    st.header("⚙️ Cấu hình")
    provider = st.selectbox("Chọn AI Provider", ["Gemini", "Groq"])
    groq_model = "GPT OSS 20B 128k"
    
    api_key_input = st.text_input("API Key (Ghi đè .env)", type="password")
    
    st.subheader("📄 Định dạng")
    target_pages = st.number_input("Số trang mong muốn", 5, 30, 12)
    font_name = "Times New Roman"
    font_size = 14

# ===== NHẬP LIỆU (DEFAULTS THEO YÊU CẦU) =====
col1, col2 = st.columns(2)
with col1:
    ten_hoc_phan = st.text_input("Học phần", "Giáo dục học đại cương")
    ten_hoc_vien = st.text_input("Học viên", "Đặng Nhật Minh")
    template_path = "De Thi/template.docx" # Đường dẫn file của bạn
with col2:
    default_de_bai = "Bằng lý luận và thực tiễn, anh (chị) hãy phân tích các chức năng xã hội của giáo dục. Từ đó, anh (chị) hãy liên hệ với việc thực hiện các chức năng này ở Việt Nam."
    de_bai = st.text_area("Đề tài chi tiết", default_de_bai, height=150)

# ===== LOGIC CHÍNH =====
if st.button("🚀 Bắt đầu tạo bài tập lớn"):
    # 1. Kiểm tra Key
    api_key = api_key_input or (os.getenv("GEMINI_API_KEY") if provider == "Gemini" else os.getenv("GROQ_API_KEY"))
    if not api_key:
        st.error("❌ Thiếu API Key!")
        st.stop()

    try:
        status = st.empty()
        progress = st.progress(0)

        # BƯỚC 1: LẬP DÀN Ý CHI TIẾT
        status.info("📝 Bước 1: Đang lập dàn ý chuyên sâu...")
        outline_prompt = f"""Bạn là giảng viên đại học. Hãy lập dàn ý chi tiết cho bài tập lớn:
        Đề tài: {de_bai}
        Yêu cầu: Có các phần Mở đầu, Nội dung (3-4 chương), Kết luận. 
        Mỗi chương chia nhỏ thành 1.1, 1.2... Chỉ trả về tiêu đề các mục."""
        
        outline_raw = call_ai(provider, api_key, outline_prompt, groq_model)
        sections = [s.strip() for s in outline_raw.split('\n') if len(s.strip()) > 5]
        progress.progress(10)

        # BƯỚC 2: VIẾT TỪNG MỤC (CHỐNG TRÀN TOKEN)
        full_content = []
        status.info(f"✍️ Bước 2: Đang viết chi tiết {len(sections)} chương...")
        
        for idx, section in enumerate(sections):
            status.write(f"⏳ Đang viết: {section}...")
            section_prompt = f"""Hãy viết nội dung học thuật cho mục '{section}' trong đề tài '{de_bai}'.
            Yêu cầu:
            - Độ dài: ít nhất 800 từ.
            - Phong cách: Nghiêm túc, giàu tính lý luận và thực tiễn.
            - Ngôn ngữ: Tiếng Việt.
            - Không dùng markdown."""
            
            part_content = call_ai(provider, api_key, section_prompt, groq_model)
            full_content.append((section, part_content))
            
            # Cập nhật progress
            current_p = 10 + int(((idx + 1) / len(sections)) * 80)
            progress.progress(current_p)

        # BƯỚC 3: ĐỔ DỮ LIỆU VÀO TEMPLATE
        status.info("📄 Bước 3: Đang đổ dữ liệu vào Template và Fix font...")
        
        if not os.path.exists(template_path):
            st.warning(f"Không tìm thấy template tại {template_path}. Sẽ tạo file mới.")
            doc = Document()
        else:
            doc = Document(template_path)

        # Thay thế thông tin trên bìa (Dựa trên template bạn gửi)
        replace_placeholder(doc, "TÊN CHỦ ĐỀ", de_bai.upper())
        replace_placeholder(doc, "Họ và tên học viên:", f"Họ và tên học viên: {ten_hoc_vien}")
        
        # Thêm nội dung vào sau trang bìa/mục lục
        for title, body in full_content:
            # Thêm tiêu đề mục
            h = doc.add_paragraph()
            run_h = h.add_run(clean_text(title))
            run_h.bold = True
            run_h.font.size = Pt(16)
            run_h.font.name = font_name
            run_h._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

            # Thêm nội dung chi tiết
            for para_text in body.split('\n'):
                if para_text.strip():
                    p = doc.add_paragraph(para_text)
                    p.paragraph_format.line_spacing = 1.5
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    
                    run = p.add_run() if not p.runs else p.runs[0]
                    run.text = para_text
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
                    # FIX FONT TIẾNG VIỆT
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

        # Lưu file
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        progress.progress(100)
        status.success("🎉 Đã tạo xong báo cáo 12 trang!")

        st.download_button(
            label="📥 Tải Bài Tập Lớn Hoàn Thiện",
            data=buffer,
            file_name=f"BTL_GiaoDucHoc_{ten_hoc_vien}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"Lỗi hệ thống: {e}")