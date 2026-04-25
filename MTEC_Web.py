import os
import io
import zipfile
import re
import pandas as pd
from docxtpl import DocxTemplate, RichText
import streamlit as st
from datetime import datetime

# --- Cấu hình trang Web ---
st.set_page_config(
    page_title="MTEC Document Generator",
    page_icon="branding/mtec_logo.svg",
    layout="centered"
)

# --- CSS Branding ---
st.markdown("""
<style>
    /* Nền chính màu xanh navy MTEC */
    [data-testid="stAppViewContainer"] {
        background-color: #061932;
        color: #ffffff;
    }
    
    [data-testid="stHeader"] {
        background-color: #061932;
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
        background-color: #0a1f3f;
        transition: border 0.3s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        border: 2px dashed #ffc20e;
    }

    /* Expander (Xem trước dữ liệu) */
    [data-testid="stExpander"] {
        border-radius: 10px;
        border: 1px solid #1a3c6d;
        background-color: #0a1f3f;
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
    try:
        st.image("branding/mtec_logo.png", width=150)
    except:
        st.markdown("## 🛡️ MTEC TOOLS")
    st.markdown("### 📋 Hướng dẫn sử dụng")
    st.markdown("""
    1. Tải lên file **Excel/CSV** chứa danh sách đăng ký.
    2. Tải lên file **Template Word (.docx)** chứa các trường như `{{ ho_ten }}`.
    3. Chọn dòng dữ liệu ở phần xem trước để kiểm tra file mấu.
    4. Sang tab xuất hàng loạt, nhấn **Bắt đầu** để lấy file `ZIP`.
    """)
    st.divider()
    st.markdown("### ⚙️ Tùy chỉnh đầu ra")
    file_naming_convention = st.text_input(
        "Format Tên File (Bỏ đuôi .docx):",
        value="HOSO_{mssv}_{ho_ten}",
        help="Dùng các biến {mssv}, {ho_ten}. Ví dụ: MTEC_{mssv}"
    )

try:
    st.image("branding/mtec-banner-fb-2026.png", use_container_width=True)
except Exception:
    st.title("📄 MTEC - Trình Tạo Document Hàng Loạt")

st.markdown("<h3 style='text-align: center; color: #ffc20e;'>Tạo hàng trăm file Word từ 1 file Excel/CSV và 1 file Mẫu (.docx) dễ dàng.</h3>", unsafe_allow_html=True)
st.divider()

# Bước 1: Upload File
st.subheader("1️⃣ Tải lên dữ liệu & Template")
col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("📁 Chọn file Excel/CSV danh sách", type=['csv', 'xlsx'])

with col2:
    template_file = st.file_uploader("📝 Chọn Word Template", type=['docx'])

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
                doc = DocxTemplate(template_file)
                doc.render(context)
                
                doc_io = io.BytesIO()
                doc.save(doc_io)
                
                st.success(f"🎉 Đã tạo thành công bản xem trước cho **'{clean_context_for_json.get('ho_ten', 'Unknown')}'**.")
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

                        doc = DocxTemplate(template_file)
                        doc.render(context)
                        
                        mssv = context.get('mssv', 'Unknown_MSSV')
                        ho_ten = context.get('ho_ten', f'Unknown_Name_{index}')
                        
                        # Generate name based on file_naming_convention provided in sidebar
                        try:
                            # Safely format name ignoring key errors if user typed wrong placeholder
                            filename_base = file_naming_convention.format(mssv=mssv, ho_ten=ho_ten)
                        except:
                            filename_base = f"HOSO_{mssv}_{ho_ten}"
                            
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
