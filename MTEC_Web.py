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
    page_title="MTEC Word Gen Web",
    page_icon="📄",
    layout="centered"
)

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
st.title("📄 MTEC - Trình Tạo Document Hàng Loạt")
st.markdown("Tạo hàng trăm file Word từ 1 file Excel/CSV và 1 file Mẫu (.docx) dễ dàng.")

st.divider()

# Bước 1: Upload File
st.subheader("1. Tải lên dữ liệu & Template")
col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("📁 Chọn file Excel/CSV", type=['csv', 'xlsx'])

with col2:
    template_file = st.file_uploader("📝 Chọn Word Template", type=['docx'])

if excel_file and template_file:
    try:
        if excel_file.name.endswith('.csv'):
            df = pd.read_csv(excel_file, dtype={'Số điện thoại liên hệ': str, 'Mã số sinh viên (MSSV)': str})
        else:
            df = pd.read_excel(excel_file, dtype={'Số điện thoại liên hệ': str, 'Mã số sinh viên (MSSV)': str})
            
        st.success(f"Đã đọc file. Tổng cộng {len(df)} dòng dữ liệu.")
        with st.expander("🔎 Xem trước dữ liệu Excel/CSV"):
            st.dataframe(df.head(5))

    except Exception as e:
        st.error(f"Lỗi đọc dữ liệu: {e}")
        st.stop()
        
    st.divider()

    st.subheader("2. Xem trước Output & Xử lý")
    tab_preview, tab_generate = st.tabs(["👁️ Xem trước 1 File output", "🚀 Tạo hàng loạt"])
    
    with tab_preview:
        sample_index = st.number_input("Chọn dòng dữ liệu để xem trước (Index từ 0):", min_value=0, max_value=len(df)-1, value=0)
        if st.button("Tạo bản xem trước thay thế"):
            try:
                row = df.iloc[[sample_index]].copy()
                processed_row = process_mtec_data_v2(row).iloc[0]
                context = processed_row.to_dict()
                
                for field in formating_columns:
                    if field in context:
                        context[field] = format_[field](context[field])

                doc = DocxTemplate(template_file)
                doc.render(context)
                
                doc_io = io.BytesIO()
                doc.save(doc_io)
                
                st.success(f"Đã tạo thành công bản xem trước cho dòng {sample_index}.")
                st.download_button(
                    label="📥 Tải xuống File Xem Trước (.docx)",
                    data=doc_io.getvalue(),
                    file_name=f"Preview_Row{sample_index}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="preview_download"
                )
            except Exception as e:
                st.error(f"Lỗi tạo file xem trước: {e}")

    with tab_generate:
        if st.button("🚀 Bắt đầu Tạo Toàn Bộ Documents", type="primary", use_container_width=True):
            progress_text = "Đang xử lý dữ liệu..."
            my_bar = st.progress(0, text=progress_text)
            zip_buffer = io.BytesIO()
            success_count = 0
            
            try:
                processed_df = process_mtec_data_v2(df)
            except Exception as e:
                st.error(f"Lỗi xử lý dữ liệu: {e}")
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
                        filename = f"HOSO_{mssv}_{ho_ten}.docx".replace(" ", "_")
                        
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        zip_file.writestr(filename, doc_io.getvalue())
                        
                        success_count += 1
                    except Exception as e:
                        st.toast(f"Lỗi ở dòng {index+1}: {e}", icon="⚠️")
                    
                    progress = (index + 1) / len(processed_df)
                    my_bar.progress(progress, text=f"Đã tạo {index+1}/{len(processed_df)} file...")
                    
            st.success(f"🎉 Đã tạo thành công {success_count} file!")
            st.download_button(
                label="📦 Tải Toàn Bộ File (Định dạng .zip)",
                data=zip_buffer.getvalue(),
                file_name=f"MTEC_HOSO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                use_container_width=True
            )
