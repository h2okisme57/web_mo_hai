import io
import time
import streamlit as st
from google import genai
from docx import Document
from pptx import Presentation

MODEL_CANDIDATES = [
    "gemini-3.8-flash",
    "gemini-2.0-flash",
    "gemini-1.5-flash",
    "gemini-1.5-pro",
]

st.set_page_config(
    page_title="Trợ lý Soạn giảng 7991 - Cô Nguyễn Thị Ngọc Yến",
    page_icon="📚",
    layout="wide"
)

# Header định danh tác giả
st.markdown("## 📚 HỆ THỐNG TRỢ LÝ GIẢNG DẠY & KHẢO THÍ SỐ")
st.markdown("**Giáo viên thực hiện:** Nguyễn Thị Ngọc Yến")
st.caption("Ứng dụng AI biên soạn Kế hoạch bài dạy, Slide bài giảng và Đề kiểm tra định kì theo Công văn 7991/BGDĐT-GDTrH")
st.divider()

# Sidebar cấu hình
with st.sidebar:
    st.header("⚙️ Thiết lập hệ thống")
    api_key_input = st.text_input("Nhập Google Gemini API Key:", type="password")
    api_key = api_key_input.strip() or st.secrets.get("GEMINI_API_KEY", "").strip()
    
    if st.button("🔍 Kiểm tra trạng thái AI", use_container_width=True):
        if not api_key:
            st.error("Chưa nhập API Key!")
        else:
            with st.spinner("Đang kiểm tra kết nối với hệ thống Google..."):
                try:
                    client = genai.Client(api_key=api_key)
                    res = client.models.generate_content(
                        model=MODEL_CANDIDATES[0],
                        contents="ping"
                    )
                    st.success(f"✅ Kết nối thành công!\n\nModel sẵn sàng: `{MODEL_CANDIDATES[0]}`")
                except Exception as e:
                    st.error(f"❌ Lỗi xác thực Key: {e}")

    st.divider()
    st.header("📋 Thông tin bài học")
    gv_name = st.text_input("Họ tên giáo viên:", value="Nguyễn Thị Ngọc Yến")
    mon_hoc = st.selectbox("Môn học phụ trách:", [
        "Ngữ văn", "Toán học", "Tiếng Anh", "Khoa học tự nhiên", 
        "Lịch sử và Địa lí", "Vật lí", "Hóa học", "Sinh học", 
        "Tin học", "Giáo dục công dân", "Công nghệ"
    ])
    khoi_lop = st.selectbox("Khối lớp:", [f"Lớp {i}" for i in range(6, 13)])
    ten_bai_hoc = st.text_input("Tên bài học / Chuyên đề:", value="Phương pháp đọc hiểu văn bản nghị luận")
    thoi_luong = st.number_input("Thời lượng (tiết):", min_value=1, max_value=8, value=2)
    yeu_cau_can_dat = st.text_area(
        "Yêu cầu cần đạt (theo CT GDPT 2018):",
        height=100,
        value="- Nhận biết và phân tích được luận điểm, lí lẽ, bằng chứng tiêu biểu trong văn bản.\n- Liên hệ, rút ra bài học thực tiễn đời sống."
    )

def call_gemini(prompt_text, key):
    if not key:
        st.error("Vui lòng nhập Google Gemini API Key vào thanh bên trái!")
        return None
    try:
        client = genai.Client(api_key=key)
        last_err = None
        for m_name in MODEL_CANDIDATES:
            for _ in range(2):
                try:
                    res = client.models.generate_content(
                        model=m_name,
                        contents=prompt_text
                    )
                    if res and res.text:
                        return res.text
                except Exception as err:
                    last_err = err
                    err_msg = str(err).lower()
                    if "503" in err_msg or "high demand" in err_msg or "unavailable" in err_msg:
                        time.sleep(1)
                        continue
                    else:
                        break
        st.error(f"Lỗi khi xử lý qua AI: {last_err}")
        return None
    except Exception as e:
        st.error(f"Lỗi khởi tạo client: {e}")
        return None

def export_docx(title, content, author):
    doc = Document()
    doc.add_heading(title, level=0)
    p = doc.add_paragraph()
    p.add_run(f"Giáo viên thực hiện: {author}\n").bold = True
    p.add_run("Quy chuẩn chuyên môn: Công văn 7991/BGDĐT-GDTrH\n").italic = True
    for para in content.split("\n"):
        doc.add_paragraph(para)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def export_pptx(title, raw_text, author):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = f"Môn học: {mon_hoc} - {khoi_lop}\nGiáo viên: {author}"
    
    sections = [s for s in raw_text.split("SLIDE:") if s.strip()]
    for sec in sections:
        lines = [line.strip() for line in sec.strip().split("\n") if line.strip()]
        if not lines:
            continue
        slide_title = lines[0].replace("#", "").strip()
        body_lines = lines[1:]
        
        new_slide = prs.slides.add_slide(prs.slide_layouts[1])
        new_slide.shapes.title.text = slide_title
        tf = new_slide.placeholders[1].text_frame
        tf.clear()
        for bl in body_lines:
            p = tf.add_paragraph()
            p.text = bl.lstrip("-*• ")
            p.level = 0
            
    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()

# Bố trí 3 phân hệ chức năng
col_nav1, col_nav2, col_nav3 = st.columns(3)
with col_nav1:
    btn_tab1 = st.button("📖 Kế hoạch bài dạy (KHBD)", use_container_width=True)
with col_nav2:
    btn_tab2 = st.button("📊 Bài giảng PowerPoint (.PPTX)", use_container_width=True)
with col_nav3:
    btn_tab3 = st.button("📝 Đề & Ma trận (CV 7991)", use_container_width=True)

if btn_tab1:
    st.session_state["active_view"] = "khbd"
elif btn_tab2:
    st.session_state["active_view"] = "slide"
elif btn_tab3:
    st.session_state["active_view"] = "exam"
elif "active_view" not in st.session_state:
    st.session_state["active_view"] = "khbd"

cur_view = st.session_state["active_view"]

# 1. KẾ HOẠCH BÀI DẠY
if cur_view == "khbd":
    st.subheader("📋 Thiết kế Kế hoạch bài dạy chuẩn quy định")
    if st.button("🚀 Khởi tạo Kế hoạch bài dạy", type="primary", key="btn_run_khbd"):
        with st.spinner("AI đang biên soạn Kế hoạch bài dạy chi tiết..."):
            prompt = f"""
            Đóng vai trò là cố vấn phương pháp dạy học cấp THCS/THPT. 
            Hãy soạn Kế hoạch bài dạy hoàn chỉnh cho:
            - Môn: {mon_hoc} - {khoi_lop}
            - Bài dạy: {ten_bai_hoc} (Thời lượng: {thoi_luong} tiết)
            - Yêu cầu cần đạt: {yeu_cau_can_dat}
            - Giáo viên thực hiện: {gv_name}

            Cấu trúc bài soạn bắt buộc gồm:
            I. MỤC TIÊU (Năng lực đặc thù, Năng lực chung, Phẩm chất).
            II. THIẾT BỊ DẠY HỌC & HỌC LIỆU (Giáo viên và học sinh).
            III. TIẾN TRÌNH DẠY HỌC:
            1. Hoạt động 1: Mở đầu / Khởi động (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện: Chuyển giao -> Thực hiện -> Báo cáo/Thảo luận -> Kết luận/Nhận định).
            2. Hoạt động 2: Hình thành kiến thức mới (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện 4 bước).
            3. Hoạt động 3: Luyện tập (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện 4 bước).
            4. Hoạt động 4: Vận dụng (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện 4 bước).
            """
            st.session_state["out_khbd"] = call_gemini(prompt, api_key)
            
    if st.session_state.get("out_khbd"):
        st.markdown(st.session_state["out_khbd"])
        f_docx = export_docx(f"KHBD_{ten_bai_hoc}", st.session_state["out_khbd"], gv_name)
        st.download_button(
            "📥 Tải về file Word KHBD (.docx)", 
            data=f_docx, 
            file_name=f"KHBD_{ten_bai_hoc}_{gv_name}.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# 2. SLIDE BÀI GIẢNG
elif cur_view == "slide":
    st.subheader("📊 Thiết kế Slide bài giảng trực quan")
    sl_count = st.slider("Số lượng slide bài giảng:", 5, 15, 8)
    if st.button("🚀 Khởi tạo Slide bài giảng", type="primary", key="btn_run_slide"):
        with st.spinner("Đang xây dựng nội dung bài giảng trình chiếu..."):
            prompt_sl = f"""
            Tạo cấu trúc trình chiếu PowerPoint gồm {sl_count} slide cho:
            - Bài học: {ten_bai_hoc}
            - Môn học: {mon_hoc} ({khoi_lop})
            - Giáo viên: {gv_name}

            QUY ƯỚC ĐỊNH DẠNG:
            Mỗi slide bắt đầu bằng chữ: "SLIDE: [Tên tiêu đề slide]"
            Dưới tiêu đề slide là 3-4 gạch đầu dòng ngắn gọn, súc tích, trực quan.
            Không dùng đoạn văn dài dòng.
            """
            st.session_state["out_slide"] = call_gemini(prompt_sl, api_key)

    if st.session_state.get("out_slide"):
        st.text_area("Dàn ý slide:", value=st.session_state["out_slide"], height=250)
        f_pptx = export_pptx(ten_bai_hoc, st.session_state["out_slide"], gv_name)
        st.download_button(
            "📥 Tải về Slide PowerPoint (.pptx)", 
            data=f_pptx, 
            file_name=f"Slide_{ten_bai_hoc}_{gv_name}.pptx", 
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

# 3. ĐỀ KIỂM TRA ĐỊNH KỲ THEO CÔNG VĂN 7991
elif cur_view == "exam":
    st.subheader("📝 Ma trận, Bản đặc tả & Đề kiểm tra định kỳ (Chuẩn CV 7991)")
    c1, c2 = st.columns(2)
    with c1:
        exam_time = st.selectbox("Thời lượng kiểm tra:", ["45 phút", "60 phút", "90 phút"])
    with c2:
        has_short = st.checkbox("Môn có câu hỏi Trả lời ngắn (Toán, KHTN, Tin...)", value=False)
        
    if st.button("🚀 Khởi tạo Bộ hồ sơ kiểm tra 7991", type="primary", key="btn_run_exam"):
        with st.spinner("Đang lập ma trận và đặc tả theo Công văn 7991/BGDĐT-GDTrH..."):
            prompt_ex = f"""
            Đóng vai trò là chuyên gia khảo thí của Bộ GDĐT. Hãy thiết kế bộ hồ sơ kiểm tra định kì chuẩn xác theo CÔNG VĂN SỐ 7991/BGDĐT-GDTrH:
            - Môn: {mon_hoc} - {khoi_lop}
            - Bài / Chủ đề kiểm tra: {ten_bai_hoc}
            - Thời gian làm bài: {exam_time}
            - Giáo viên thẩm định / ra đề: {gv_name}

            QUY CHUẨN CÔNG VĂN 7991:
            1. Tỉ lệ điểm định dạng câu hỏi:
               - Phần I. Trắc nghiệm Nhiều lựa chọn: 3,0 điểm (30%).
               - Phần II. Trắc nghiệm Đúng - Sai: 2,0 điểm (20%). Gồm các câu có 4 lệnh hỏi a, b, c, d (chọn Đúng hoặc Sai).
               {"- Phần III. Trắc nghiệm Trả lời ngắn: 2,0 điểm (20%)." if has_short else "- Môn không dùng Trả lời ngắn: Chuyển toàn bộ 2,0 điểm sang dạng Đúng - Sai (Tổng Đúng-Sai 4,0 điểm theo chú thích 3 Phụ lục 1)."}
               - Phần IV. Tự luận: 3,0 điểm (30%).
            2. Tỉ lệ mức độ nhận thức:
               - Biết: khoảng 40% (4,0 điểm)
               - Hiểu: khoảng 30% (3,0 điểm)
               - Vận dụng: khoảng 30% (3,0 điểm)

            XUẤT RA CHI TIẾT 4 PHẦN:
            1. KHUNG MA TRẬN ĐỀ KIỂM TRA ĐỊNH KÌ (Vẽ bảng Markdown theo đúng mẫu Phụ lục 1 CV 7991).
            2. BẢN ĐẶC TẢ ĐỀ KIỂM TRA ĐỊNH KÌ (Vẽ bảng theo đúng mẫu Phụ lục 2 CV 7991: TT, Chủ đề, Nội dung, Yêu cầu cần đạt, Số câu theo mức độ Biết/Hiểu/Vận dụng cho từng định dạng).
            3. ĐỀ KIỂM TRA HOÀN CHỈNH (Trình bày rõ từng phần câu hỏi I, II, III, IV).
            4. ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM:
               - Đáp án trắc nghiệm nhiều lựa chọn.
               - Hướng dẫn tính điểm Đúng-Sai: 1 ý đúng = 0.1đ; 2 ý đúng = 0.25đ; 3 ý đúng = 0.5đ; 4 ý đúng = 1.0đ.
               - Đáp án phần trả lời ngắn (nếu có).
               - Barem điểm chi tiết cho bài Tự luận.
            """
            st.session_state["out_exam"] = call_gemini(prompt_ex, api_key)

    if st.session_state.get("out_exam"):
        st.markdown(st.session_state["out_exam"])
        f_exam_docx = export_docx(f"De_Kiem_Tra_7991_{ten_bai_hoc}", st.session_state["out_exam"], gv_name)
        st.download_button(
            "📥 Tải Bộ hồ sơ đề kiểm tra CV 7991 (.docx)", 
            data=f_exam_docx, 
            file_name=f"De_Kiem_Tra_7991_{ten_bai_hoc}_{gv_name}.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
