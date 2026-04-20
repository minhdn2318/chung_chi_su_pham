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

# ===== 1. TIỆN ÍCH HỆ THỐNG =====
def remove_vietnamese_accent(s):
    """Chuyển tiếng Việt thành không dấu để đặt tên file an toàn"""
    s = s.replace("Đ", "D").replace("đ", "d")
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
    return re.sub(r'[^a-zA-Z0-9_]', '_', s)

def strictly_clean_content(text):
    """Xóa bỏ triệt để câu dẫn AI và ký tự Markdown dư thừa"""
    ai_patterns = [
        r"Dưới đây là.*:", r"Đoạn văn đã được chỉnh sửa.*:", 
        r"Đây là nội dung.*:", r"Tôi đã thực hiện.*:", 
        r"Nội dung biên tập.*:", r"Hy vọng bản thảo.*",
        r"Kết luận lại là.*:", r"Sau đây là phần.*:"
    ]
    for pattern in ai_patterns:
        text = re.sub(pattern, "", text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'[*#_~-]', '', text)
    return text.strip()

def replace_placeholders(doc, data_dict):
    """Thay thế thông tin trên trang bìa Bia.docx"""
    for p in doc.paragraphs:
        for key, value in data_dict.items():
            if key in p.text:
                # Thay thế text nhưng giữ nguyên định dạng của Paragraph trong file Bia.docx
                inline = p.runs
                for i in range(len(inline)):
                    if key in inline[i].text:
                        text = inline[i].text.replace(key, str(value))
                        inline[i].text = text

# ===== 2. DATABASE 9 HỌC PHẦN (CẬP NHẬT TỪ EXCEL) =====
DATA_BTL = {
    "HP_01: Giáo dục học đại cương": {
        "de_tai": "PHÂN TÍCH CÁC CHỨC NĂNG XÃ HỘI CỦA GIÁO DỤC",
        "cau_hoi": "Bằng lý luận và thực tiễn, phân tích các chức năng xã hội của giáo dục và liên hệ thực hiện tại Việt Nam.",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5
    },
    "HP_02: Lý luận dạy học đại học": {
        "de_tai": "VẬN DỤNG CÁC NGUYÊN TẮC DẠY HỌC HIỆN ĐẠI TRONG ĐÀO TẠO ĐẠI HỌC",
        "cau_hoi": "Phân tích hệ thống nguyên tắc dạy học và đề xuất phương án vận dụng thực tiễn.",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.3
    },
    "HP_03: Sử dụng phương tiện kỹ thuật KH": {
        "de_tai": "ỨNG DỤNG CÔNG NGHỆ TRONG ĐỔI MỚI PHƯƠNG PHÁP DẠY HỌC",
        "cau_hoi": "Trình bày các phương tiện kỹ thuật hiện đại. Phân tích thực trạng và giải pháp ứng dụng.",
        "le": (2.0, 2.0, 3.0, 2.0), "spacing": 1.5
    },
    # Minh cập nhật tiếp HP_04 đến HP_09 vào đây tương tự...
}

# ===== 3. GIAO DIỆN STREAMLIT =====
st.set_page_config(page_title="VNU Report Tool Pro", layout="wide")

with st.sidebar:
    st.header("⚙️ Cấu hình Hệ thống")
    provider = st.selectbox("AI Model", ["Groq", "Gemini"], index=0)
    user_key = st.text_input(f"Nhập {provider} Key (Optional)", type="password")
    
    st.divider()
    selected_key = st.selectbox("Chọn Bài Tập Lớn (1-9)", list(DATA_BTL.keys()))
    current_data = DATA_BTL[selected_key]
    btl_code = selected_key.split(":")[0]

st.title("🎓 Hệ thống Tạo BTL kèm Bìa tự động")

col1, col2 = st.columns([1, 2])
with col1:
    st.subheader("👤 Thông tin sinh viên")
    ten_hv = st.text_input("Họ và tên học viên", "Đặng Nhật Minh")
    sbd = st.text_input("Số báo danh", "39")
    mon_hoc = st.text_input("Tên học phần", selected_key.split(": ")[1])
    de_tai_input = st.text_area("Tên đề tài (In bìa)", current_data["de_tai"], height=100)

with col2:
    st.subheader("📝 Nội dung yêu cầu (AI)")
    yeu_cau = st.text_area("Đề bài chi tiết", current_data["cau_hoi"], height=220)

# ===== 4. LOGIC XỬ LÝ AI =====
def call_ai_clean(key, provider, prompt):
    system_msg = (
        "Bạn là một máy soạn thảo tiểu luận chuyên nghiệp. "
        "CHỈ xuất nội dung học thuật. KHÔNG lời dẫn, KHÔNG chào hỏi, KHÔNG dùng Markdown."
    )
    try:
        if provider == "Gemini":
            import google.generativeai as genai
            genai.configure(api_key=key)
            model = genai.GenerativeModel('gemini-1.5-pro')
            return model.generate_content(f"{system_msg}\n\n{prompt}").text
        else:
            from groq import Groq
            client = Groq(api_key=key)
            res = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "system", "content": system_msg}, {"role": "user", "content": prompt}]
            )
            return res.choices[0].message.content
    except Exception as e: return f"Lỗi: {e}"

if st.button("🚀 BẮT ĐẦU TẠO BÀI (KÈM TRANG BÌA)"):
    final_api_key = user_key or st.secrets.get(f"{provider.upper()}_API_KEY")
    if not final_api_key: st.error("Vui lòng nhập API Key!"); st.stop()
    
    if not os.path.exists("Bia.docx"):
        st.error("❌ Không tìm thấy file 'Bia.docx' trong thư mục app. Vui lòng kiểm tra lại!")
        st.stop()

    status = st.empty(); prog = st.progress(0)
    
    # Bước 1: Tạo dàn ý học thuật
    status.info("📍 Đang xây dựng dàn ý...")
    outline_p = f"Lập dàn ý 6 mục chuyên sâu cho tiểu luận: {de_tai_input}. Chỉ trả về danh mục."
    outline = call_ai_clean(final_api_key, provider, outline_p)
    sections = [s.strip() for s in outline.split('\n') if len(s.strip()) > 10][:6]

    # Bước 2: Viết nội dung & Làm sạch
    final_content = []
    for i, sec in enumerate(sections):
        status.warning(f"✍️ Đang soạn thảo mục {i+1}/{len(sections)}: {sec}")
        prompt = f"Viết 800 từ chuyên sâu cho mục '{sec}' của đề tài '{de_tai_input}'. Nội dung bám sát: {yeu_cau}. Không dùng Markdown."
        raw_text = call_ai_clean(final_api_key, provider, prompt)
        final_content.append((sec, strictly_clean_content(raw_text)))
        prog.progress(int((i+1)/len(sections)*100))

    # ===== 5. TỔNG HỢP FILE WORD =====
    # Mở file Bia.docx có sẵn
    doc = Document("Bia.docx")
    
    # Thay thế các placeholder trên trang bìa
    # Lưu ý: Trong file Bia.docx, bạn cần để sẵn các cụm: {{HO_TEN}}, {{SBD}}, {{MON_HOC}}, {{TEN_DE_TAI}}
    replace_placeholders(doc, {
        "{{HO_TEN}}": ten_hv,
        "{{SBD}}": sbd,
        "{{MON_HOC}}": mon_hoc.upper(),
        "{{TEN_DE_TAI}}": de_tai_input.upper()
    })
    
    # Ngắt trang để bắt đầu nội dung bài làm
    doc.add_page_break()
    
    # Cấu hình lề cho các trang nội dung
    s = doc.sections[-1] # Áp dụng cho section mới sau ngắt trang
    m = current_data["le"]
    s.top_margin, s.bottom_margin = Cm(m[0]), Cm(m[1])
    s.left_margin, s.right_margin = Cm(m[2]), Cm(m[3])

    # Đổ nội dung bài làm
    for title, text in final_content:
        h = doc.add_paragraph()
        r = h.add_run(strictly_clean_content(title).upper())
        r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(14)
        
        for line in text.split('\n'):
            if line.strip():
                p = doc.add_paragraph(line.strip())
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.line_spacing = current_data["spacing"]
                run = p.runs[0] if p.runs else p.add_run()
                run.font.name = "Times New Roman"; run.font.size = Pt(14)

    # Xử lý tên file tải về
    clean_name = remove_vietnamese_accent(ten_hv)
    file_name_final = f"BTL_{btl_code}_{clean_name}.docx"

    buffer = io.BytesIO()
    doc.save(buffer); buffer.seek(0)
    
    status.success(f"✅ Đã hoàn thành bài {btl_code}!")
    st.download_button(
        label=f"📥 Tải xuống: {file_name_final}",
        data=buffer,
        file_name=file_name_final,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )