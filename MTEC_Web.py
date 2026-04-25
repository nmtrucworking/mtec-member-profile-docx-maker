import os
import io
import zipfile
import re
import pandas as pd
from docxtpl import DocxTemplate, RichText
from docx import Document
import streamlit as st
import streamlit.components.v1 as components
from datetime import datetime

# --- Cấu hình trang Web ---
st.set_page_config(
    page_title="MTEC Document Generator | Tự động Excel sang Word",
    page_icon="branding/mtec_logo.svg",
    layout="centered",
    initial_sidebar_state="expanded"
)

# --- Tối ưu SEO ---
# Dùng JS thủ thuật để đẩy tag Meta lên Parent Document của Streamlit
components.html("""
<script>
    const setMeta = (attr, key, value) => {
        let meta = window.parent.document.querySelector(`meta[${attr}="${key}"]`);
        if (!meta) {
            meta = window.parent.document.createElement('meta');
            meta.setAttribute(attr, key);
            window.parent.document.head.appendChild(meta);
        }
        meta.setAttribute('content', value);
    };
    
    // Khai báo các thẻ meta cho công cụ tìm kiếm
    setMeta('name', 'description', 'MTEC Document Generator - Trình tạo văn bản hàng loạt. Giải pháp tự động hóa tạo hàng trăm file Word (.docx) từ dữ liệu Excel/CSV cực kỳ nhanh chóng và chính xác.');
    setMeta('name', 'keywords', 'MTEC, document generator, Mail Merge, trộn thư Word, Excel to Word, tự động tạo file Word, làm hồ sơ tự động, xuất file Word hàng loạt, RPA văn phòng');
    setMeta('name', 'author', 'MTEC');
    setMeta('name', 'robots', 'index, follow');
    
    // Khai báo cho Open Graph / Mạng xã hội FB, Zalo...
    setMeta('property', 'og:title', 'MTEC Document Generator | Tạo Word tự động từ Excel');
    setMeta('property', 'og:description', 'Tự động tạo hàng trăm file Word từ 1 file Excel và 1 file mẫu dễ dàng, miễn phí.');
    setMeta('property', 'og:url', 'https://mtec-profile.streamlit.app/');
    setMeta('property', 'og:type', 'website');
    
    // Gắn thuộc tính lang để bot biết đây là trang đa ngôn ngữ (mặc định ưu tiên hiển thị tiếng Việt)
    window.parent.document.documentElement.lang = 'vi';
</script>
""", height=0, width=0)

# --- Trạng thái ngôn ngữ và giao diện ---
if 'lang' not in st.session_state:
    st.session_state.lang = 'vi'

texts = {
    'vi': {
        'title': "📄 MTEC - Trình Tạo Document Hàng Loạt",
        'subtitle': "Giải pháp tự động hóa trích xuất và khởi tạo tài liệu hàng loạt chuyên nghiệp.",
        'guide': "### 📋 Hướng dẫn sử dụng\n1. Tải lên file **Excel/CSV** chứa danh sách đăng ký.\n2. Tải lên file **Template Word (.docx)** chứa các trường như `{{ ho_ten }}`.\n3. Chọn dòng dữ liệu ở phần xem trước để kiểm tra file mấu.\n4. Sang tab xuất hàng loạt, nhấn **Bắt đầu** để lấy file `ZIP`.",
        'custom_output': "### ⚙️ Tùy chỉnh đầu ra",
        'step1': "1️⃣ Tải lên dữ liệu & Template",
        'upload_data': "📁 Chọn file Excel/CSV",
        'upload_template': "📝 Chọn Word Template",
        'preview_tab': "👁️ Xem trước 1 File",
        'generate_tab': "🚀 Xuất hàng loạt ZIP"
    },
    'en': {
        'title': "📄 MTEC - Mass Document Generator",
        'subtitle': "Professional automated solution for data extraction and mass document generation.",
        'guide': "### 📋 Quick Guide\n1. Upload your **Excel/CSV** data file.\n2. Upload **Word Template (.docx)** containing variables like `{{ ho_ten }}`.\n3. Preview mapping on single row.\n4. Go to batch generation and click **Start** to get `.zip`.",
        'custom_output': "### ⚙️ Output Configuration",
        'step1': "1️⃣ Upload Data & Template",
        'upload_data': "📁 Select Excel/CSV file",
        'upload_template': "📝 Select Word Template",
        'preview_tab': "👁️ Preview 1 File",
        'generate_tab': "🚀 Batch Generate ZIP"
    }
}
t = texts[st.session_state.lang]

# --- CSS Branding ---
st.markdown("""
<style>
    /* Hoạt cảnh hướng người dùng mở sidebar trên mobile */
    @media (max-width: 768px) {
        [data-testid="collapsedControl"] {
            animation: pulse-sidebar 2s infinite;
            border-radius: 50%;
            background-color: rgba(255, 194, 14, 0.3);
        }
    }
    @keyframes pulse-sidebar {
        0% { box-shadow: 0 0 0 0 rgba(255, 194, 14, 0.7); }
        70% { box-shadow: 0 0 0 10px rgba(255, 194, 14, 0); }
        100% { box-shadow: 0 0 0 0 rgba(255, 194, 14, 0); }
    }

    /* Nhấn mạnh các nút bấm chính sang màu Vàng / Xanh sáng */
    button[kind="primary"] {
        background-color: #ffc20e !important;
        color: #061932 !important;
        border: none;
        font-weight: bold;
        border-radius: 8px;
        transition: all 0.3s ease;
    }
    button[kind="primary"]:hover {
        background-color: #ffd859 !important;
        color: #061932 !important;
        transform: scale(1.02);
    }
    
    /* Khung kéo thả file (Uploader) */
    [data-testid="stFileUploadDropzone"] {
        border-radius: 12px;
        border: 2px dashed #1a3c6d;
        transition: border 0.3s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        border: 2px dashed #ffc20e;
    }
</style>
""", unsafe_allow_html=True)

# --- Data Formatting Functions ---
def format_ho_ten(name):
    """Converts name to uppercase."""
    return str(name).upper() if not pd.isna(name) else ""

def format_ngay_sinh(date_str):
    """Formats a date string from m/d/yyyy to dd/mm/yyyy."""
    if pd.isna(date_str): return ""
    try:
        date_obj = pd.to_datetime(date_str)
        return date_obj.strftime('%d/%m/%Y')
    except Exception:
        try:
            day, month, year = str(date_str).split('/')
            return f"{day.zfill(2)}/{month.zfill(2)}/{year}"
        except Exception:
            return str(date_str)

def format_sdt(phone_number):
    """Ensures the phone number is a string, contains only digits, and starts with '0'."""
    if pd.isna(phone_number): return ""
    try:
        phone_str = re.sub(r'\D', '', str(phone_number))
        if not phone_str.startswith('0') and phone_str != "":
            phone_str = '0' + phone_str
        return phone_str
    except Exception:
        return str(phone_number)

formating_columns = ['ho_ten', 'sdt', 'ngay_sinh']
format_ = {
    'ho_ten': format_ho_ten,
    'ngay_sinh': format_ngay_sinh,
    'sdt': format_sdt
}

def clean_data(df):
    """Cleans and standardizes the DataFrame."""
    if 'ho_ten' in df.columns:
        df['ho_ten'] = df['ho_ten'].astype(str).str.strip()
    if 'sdt' in df.columns:
        df['sdt'] = df['sdt'].astype(str).apply(lambda x: re.sub(r'\D', '', x))
        
    prefix_pattern = r'^(khoa khoa|ngành|chuyên ngành)\s*'
    for col in ['khoa', 'chuyen_nganh']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(prefix_pattern, '', regex=True, case=False).str.strip()
    return df

def process_mtec_data_v2(df):
    CHECKED = RichText('☒', font='Segoe UI Symbol')
    UNCHECKED = RichText('☐', font='Segoe UI Symbol')

    mapping = {
        'Họ và tên đầy đủ': 'ho_ten',
        'Giới tính': 'gioi_tinh',
        'Ngày sinh': 'ngay_sinh',
        'Mã số sinh viên (MSSV)': 'mssv',
        'Khoa bạn đang theo học': 'khoa',
        'Chuyên ngành bạn đang theo học': 'chuyen_nganh',
        'Số điện thoại liên hệ': 'sdt',
        'Email sinh viên hoặc Email cá nhân': 'email',
        'Link Facebook cá nhân': 'link_fb',
        'Mục tiêu khi tham gia CLB là gì': 'muc_tieu',
        'Định hướng phát triển cá nhân của bạn khi tham gia CLB?': 'dinh_huong',
        'Bạn có thể cam kết dành thời gian cho CLB ở mức nào?': 'cam_ket_time'
    }
    df = df.rename(columns={k: v for k, v in mapping.items() if k in df.columns})
    df = clean_data(df)
    
    col_ban_raw = 'Bạn muốn tham gia Ban chuyên môn/vị trí nào? (Có thể chọn nhiều)'
    if col_ban_raw in df.columns:
        df['c_ban_cn'] = df[col_ban_raw].apply(lambda x: CHECKED if pd.notna(x) and 'Ban Công nghệ' in str(x) else UNCHECKED)
        df['c_ban_tt'] = df[col_ban_raw].apply(lambda x: CHECKED if pd.notna(x) and 'Ban Truyền thông' in str(x) else UNCHECKED)
        df['c_ban_vh'] = df[col_ban_raw].apply(lambda x: CHECKED if pd.notna(x) and 'Ban Vận hành' in str(x) else UNCHECKED)
        df['c_ban_cnh'] = df[col_ban_raw].apply(lambda x: CHECKED if pd.notna(x) and 'Ban Chủ nhiệm' in str(x) else UNCHECKED)
        df['c_ban_khac'] = UNCHECKED
        
        def extract_vi_tri(text):
            if pd.isna(text): return ''
            positions = re.findall(r'\](.*?)\;', str(text))
            return ', '.join([p.strip() for p in positions])
        df['vi_tri'] = df[col_ban_raw].apply(extract_vi_tri)

    skill_keywords = {
        'tk': 'Thiết kế', 'qd': 'Quay dựng', 'ct': 'Content', 
        'fp': 'Fanpage', 'ca': 'Chụp ảnh', 'lt': 'Lập trình', 'mc': 'MC',
        'gt': 'Giao tiếp', 'lvn': 'Làm việc nhóm', 'qltg': 'thời gian',
        'st': 'Sáng tạo', 'gqvd': 'vấn đề', 'thvp': 'văn phòng'
    }

    for key, keyword in skill_keywords.items():
        target_cols = [c for c in df.columns if keyword.lower() in c.lower()]
        if target_cols:
            col = target_cols[0]
            df[f'{key}_cb']  = df[col].apply(lambda x: CHECKED if str(x).strip() == 'Cơ bản' else UNCHECKED)
            df[f'{key}_tb']  = df[col].apply(lambda x: CHECKED if str(x).strip() == 'Trung bình' else UNCHECKED)
            df[f'{key}_tot'] = df[col].apply(lambda x: CHECKED if str(x).strip() == 'Tốt' else UNCHECKED)
            df[f'{key}'] = df[col].apply(lambda x: CHECKED if str(x).strip() in ['Cơ bản', 'Trung bình', 'Tốt'] else UNCHECKED)

    return df.fillna("")

# --- Giao diện Web ---
# Sidebar
with st.sidebar:
    import builtins
    
    # Toggle Dark Mode thông qua CSS
    dark_mode = st.toggle("🌙 Giao diện tối (Dark Mode)", value=st.session_state.get('dark_mode', True), key="dark_mode_toggle")
    if dark_mode:
        st.session_state.dark_mode = True
        st.markdown('''
        <style>
            [data-testid="stAppViewContainer"] { background-color: #0e1117; color: #fafafa; }
            [data-testid="stSidebar"] { background-color: #262730; color: #fafafa; }
            .stMarkdown, .stText, p, span, label { color: #fafafa !important; }
            h1, h2, h3, h4, h5, h6 { color: #ffc20e !important; }
        </style>
        ''', unsafe_allow_html=True)
    else:
        st.session_state.dark_mode = False
        st.markdown('''
        <style>
            [data-testid="stAppViewContainer"] { background-color: #ffffff; color: #111111; }
            [data-testid="stSidebar"] { background-color: #f0f2f6; color: #111111; }
            .stMarkdown, .stText, p, span, label { color: #111111 !important; }
            h1, h2, h3, h4, h5, h6 { color: #1a3c6d !important; }
        </style>
        ''', unsafe_allow_html=True)
        
    st.selectbox("🌍 Ngôn ngữ / Language", ['vi', 'en'], index=0 if st.session_state.lang == 'vi' else 1, key='lang_selector', on_change=lambda: st.session_state.update(lang=st.session_state.lang_selector))
    t = texts[st.session_state.lang]
    
    try:
        st.image("branding/mtec_logo.png", width=150)
    except:
        st.markdown("## 🛡️ MTEC TOOLS")
        
    st.markdown(t['guide'])
    st.divider()

try:
    st.image("branding/mtec-banner-fb-2026.png", use_container_width=True)
    # Ẩn thẻ h1 để bot Google bot đọc được tiêu đề chính thay vì bỏ qua tag quan trọng này
    st.markdown(f"<h1 style='display: none;'>{t['title']}</h1>", unsafe_allow_html=True)
except Exception:
    st.title(t['title'])

# Dùng h2 thay vì h3 để đảm bảo cấu trúc header hierarchy (SEO)
st.markdown(f"<h2 style='text-align: center; color: #ffc20e; font-size: 1.5rem;'>{t['subtitle']}</h2>", unsafe_allow_html=True)
st.divider()

# Bước 1: Upload File
st.subheader(t['step1'])
col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader(t['upload_data'], type=['csv', 'xlsx'])

with col2:
    template_file = st.file_uploader(t['upload_template'], type=['docx'])

if excel_file and template_file:
    with st.spinner("⏳ Đang đọc dữ liệu..."):
        try:
            if excel_file.name.endswith('.csv'):
                df = pd.read_csv(excel_file, dtype={'Số điện thoại liên hệ': str, 'Mã số sinh viên (MSSV)': str})
            else:
                df = pd.read_excel(excel_file, dtype={'Số điện thoại liên hệ': str, 'Mã số sinh viên (MSSV)': str})
                
            st.success(f"✅ Đã tải file thành công! Tổng cộng **{len(df)}** dòng dữ liệu.")
            with st.expander("🔎 Xem trước dữ liệu Excel/CSV"):
                st.dataframe(df.head(5), use_container_width=True)
                
            # Tuỳ chỉnh tên file trong sidebar sau khi có df
            with st.sidebar:
                st.markdown(t['custom_output'])
                st.info("💡 Điền tên file theo mẫu. Dùng `{Tên Cột}` để chèn dữ liệu tự động. Ví dụ: `MTEC_{Mã số sinh viên (MSSV)}_{Họ và tên đầy đủ}`")
                
                # Hiển thị một số cột làm gợi ý
                sample_cols = [f"`{{{c}}}`" for c in df.columns[:5]]
                st.caption(f"Các cột có sẵn (VD): {', '.join(sample_cols)}...")
                
                default_name = "{Mã số sinh viên (MSSV)}_{Họ và tên đầy đủ}" if "Mã số sinh viên (MSSV)" in df.columns else "HOSO_{index}"
                custom_filename_pattern = st.text_input(
                    "Định dạng tên file tải về:",
                    value=default_name
                )

        except Exception as e:
            st.error(f"❌ Lỗi đọc dữ liệu: {e}")
            st.stop()
            
        # Thêm phần Dashboard Mini thống kê nhanh dữ liệu vừa up lên
        st.markdown("### 📊 Tổng quan dữ liệu")
        metric_col1, metric_col2, metric_col3 = st.columns(3)
        with metric_col1:
            st.metric(label="Tổng số hồ sơ", value=f"{len(df)} người")
        with metric_col2:
            st.metric(label="Số cột dữ liệu", value=f"{len(df.columns)} cột")
        with metric_col3:
            mssv_count = df['Mã số sinh viên (MSSV)'].notna().sum() if 'Mã số sinh viên (MSSV)' in df.columns else 0
            st.metric(label="Mã số sinh viên hợp lệ", value=f"{mssv_count}")
            
    st.divider()

    st.subheader("2️⃣ Thiết lập & Xử lý")
    tab_preview, tab_generate = st.tabs(["👁️ Xem trước 1 File", "🚀 Xuất hàng loạt ZIP"])
    
    with tab_preview:
        st.info("💡 Hãy thử tạo bản xem trước cho 1 cá nhân bất kỳ. Bạn cũng có thể xem dữ liệu đã được 'nhào nặn' trước khi ép vào file Word để biết chắc các biến `{ }` không bị thiếu.")
        sample_index = st.number_input("Chọn thứ tự người để xem trước (Bắt đầu từ 0):", min_value=0, max_value=len(df)-1, value=0)
        
        col_btn_1, col_btn_2 = st.columns(2)
        
        with col_btn_1:
            show_data_dict = st.button("🔍 Xem dữ liệu Mapping (Dictionary)")
            
        with col_btn_2:
            generate_preview = st.button("🔄 Khởi tạo File Xem trước (*.docx)", key="preview_btn")
            
        try:
            row = df.iloc[[sample_index]].copy()
            processed_row = process_mtec_data_v2(row).iloc[0]
            context = processed_row.to_dict()
            
            for field in formating_columns:
                if field in context:
                    context[field] = format_[field](context[field])
                    
            # Xóa các object type đặc biệt để in JSON ra cho UI đẹp hơn
            clean_context_for_json = {k: v for k, v in context.items() if not isinstance(v, RichText)}

            if show_data_dict:
                st.markdown("#### 🛠️ Các biến trúng tuyển (Sẽ áp dụng vào `{{ biến }}` trong Word)")
                st.json(clean_context_for_json, expanded=False)

            if generate_preview:
                template_file.seek(0)
                doc = DocxTemplate(template_file)
                doc.render(context)
                
                doc_io = io.BytesIO()
                doc.save(doc_io)
                
                st.success(f"🎉 Đã tạo thành công bản xem trước cho **'{clean_context_for_json.get('ho_ten', 'Unknown')}'**.")
                
                # Preview via PDF if possible, or fallback to mammoth/text
                doc_io.seek(0)
                try:
                    import tempfile
                    import base64
                    import platform
                    
                    if os.name == 'nt':
                        from docx2pdf import convert
                        
                        temp_dir = tempfile.gettempdir()
                        doc_path = os.path.join(temp_dir, f"preview_{int(datetime.now().timestamp())}.docx")
                        pdf_path = doc_path.replace(".docx", ".pdf")
                        
                        with open(doc_path, "wb") as f:
                            f.write(doc_io.getvalue())
                            
                        with st.spinner("Đang kết xuất PDF để preview..."):
                            convert(doc_path, pdf_path)
                            
                        if os.path.exists(pdf_path):
                            with open(pdf_path, "rb") as f:
                                base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                            pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf"></iframe>'
                            
                            with st.container(border=True):
                                st.markdown("<h4 style='text-align: center;'>📄 Bản Xem Trước (PDF format)</h4>", unsafe_allow_html=True)
                                st.markdown(pdf_display, unsafe_allow_html=True)
                                
                        else:
                            raise Exception("doc2pdf failed")
                    else:
                        raise ImportError("PDF Convert only supported natively on Windows with Word installed.")
                except Exception as e:
                    # Fallback to mammoth.convert_to_html
                    doc_io.seek(0)
                    try:
                        import mammoth
                        result = mammoth.convert_to_html(doc_io)
                        html = result.value
                        with st.container(border=True):
                            st.markdown("<h4 style='text-align: center;'>📄 Bản Xem Trước (HTML Fallback)</h4>", unsafe_allow_html=True)
                            st.components.v1.html(html, height=600, scrolling=True)
                    except ImportError:
                        doc_io.seek(0)
                        doc_preview = Document(doc_io)
                        preview_text = '\\n'.join([p.text for p in doc_preview.paragraphs if p.text.strip()])
                        with st.container(border=True):
                            st.markdown("<h4 style='text-align: center;'>📄 Bản Xem Trước (Preview Text)</h4>", unsafe_allow_html=True)
                            st.text_area("", value=preview_text, height=400, disabled=True)
                        
                st.download_button(
                    label="📥 Tải xuống File Xem Trước",
                    data=doc_io.getvalue(),
                    file_name=f"Preview_{clean_context_for_json.get('mssv', 'Unknown')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="preview_download"
                )
        except Exception as e:
            if generate_preview:
                st.error(f"❌ Lỗi khởi tạo: Có thể Template Word thiếu/sai Object. Chi tiết lỗi: {e}")
                
    with tab_generate:
        st.warning("⚡ Việc tạo hàng chục/trăm file cùng lúc có thể mất từ vài giây đến một phút.")
        if st.button("🚀 Bắt đầu Khởi Tạo & Xuất ZIP", type="primary", use_container_width=True):
            progress_text = "Đang xử lý dữ liệu..."
            my_bar = st.progress(0, text=progress_text)
            zip_buffer = io.BytesIO()
            success_count = 0
            fail_count = 0
            
            try:
                processed_df = process_mtec_data_v2(df)
            except Exception as e:
                st.error(f"Lỗi chuẩn hóa dữ liệu: {e}")
                st.stop()
                
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for index, row in processed_df.iterrows():
                    try:
                        context = row.to_dict()
                        for field in formating_columns:
                            if field in context:
                                context[field] = format_[field](context[field])

                        template_file.seek(0)
                        doc = DocxTemplate(template_file)
                        doc.render(context)
                        
                        # Xử lý tên file tuỳ chỉnh
                        try:
                            # Chuyển đổi định dạng cho an toàn 
                            safe_context = {str(k): (str(v.text) if hasattr(v, 'text') else str(v)) for k, v in context.items()}
                            filename_base = custom_filename_pattern
                            for k, v in safe_context.items():
                                filename_base = filename_base.replace(f"{{{k}}}", str(v))
                            filename_base = filename_base.replace("{index}", str(index))
                        except Exception:
                            filename_base = f"HOSO_{index}"
                            
                        # Clean unusual characters
                        filename_base = re.sub(r'[\\/*?:"<>|]', "", filename_base)
                        filename = f"{filename_base}.docx".replace(" ", "_")
                        
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        zip_file.writestr(filename, doc_io.getvalue())
                        
                        success_count += 1
                    except Exception as e:
                        fail_count += 1
                        st.toast(f"Lỗi dòng {index+1}: {e}", icon="⛔")
                    
                    progress = (index + 1) / len(processed_df)
                    my_bar.progress(progress, text=f"Đang tiến hành: {index+1}/{len(processed_df)} file...")
                    
            if success_count > 0:
                st.success(f"🎉 Khởi tạo hoàn tất! Thành công: **{success_count}**, Thất bại: **{fail_count}**")
                st.balloons()
                st.download_button(
                    label="📦 Tải Toàn Bộ File (Định dạng .zip)",
                    data=zip_buffer.getvalue(),
                    file_name=f"MTEC_OUTPUT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            else:
                st.error("Tất cả đều thất bại, vui lòng kiểm tra lại log!")
