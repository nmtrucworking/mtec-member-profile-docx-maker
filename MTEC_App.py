import os
import threading
import pandas as pd
import re
from docxtpl import DocxTemplate, RichText
import customtkinter as ctk
from customtkinter import filedialog

# Set theme and colors for a modern look
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

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

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MTEC Hồ Sơ Generator")
        self.geometry("700x520")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(4, weight=1)

        self.lbl_title = ctk.CTkLabel(self, text="MTEC Document Generator", font=ctk.CTkFont(size=20, weight="bold"))
        self.lbl_title.grid(row=0, column=0, columnspan=3, padx=20, pady=(20, 10))

        self.lbl_excel = ctk.CTkLabel(self, text="Data File (.csv, .xlsx):")
        self.lbl_excel.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        self.ent_excel = ctk.CTkEntry(self, placeholder_text="Select your data source...")
        self.ent_excel.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.btn_excel = ctk.CTkButton(self, text="Browse", width=80, command=self.browse_excel)
        self.btn_excel.grid(row=1, column=2, padx=20, pady=10)

        self.lbl_template = ctk.CTkLabel(self, text="Word Template (.docx):")
        self.lbl_template.grid(row=2, column=0, padx=20, pady=10, sticky="w")
        self.ent_template = ctk.CTkEntry(self, placeholder_text="Select your .docx template...")
        self.ent_template.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.btn_template = ctk.CTkButton(self, text="Browse", width=80, command=self.browse_template)
        self.btn_template.grid(row=2, column=2, padx=20, pady=10)

        self.lbl_output = ctk.CTkLabel(self, text="Output Folder:")
        self.lbl_output.grid(row=3, column=0, padx=20, pady=10, sticky="w")
        self.ent_output = ctk.CTkEntry(self, placeholder_text="Select destination folder...")
        self.ent_output.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        self.btn_output = ctk.CTkButton(self, text="Browse", width=80, command=self.browse_output)
        self.btn_output.grid(row=3, column=2, padx=20, pady=10)

        self.log_box = ctk.CTkTextbox(self, corner_radius=8)
        self.log_box.grid(row=4, column=0, columnspan=3, padx=20, pady=10, sticky="nsew")
        self.log_box.insert("0.0", "Logs and progress will appear here...\n")
        self.log_box.configure(state="disabled")

        self.btn_generate = ctk.CTkButton(self, text="Generate Documents", height=45, font=ctk.CTkFont(size=14, weight="bold"), command=self.start_generation)
        self.btn_generate.grid(row=5, column=0, columnspan=3, padx=20, pady=(10, 20), sticky="ew")

    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Data Files", "*.csv;*.xlsx"), ("All Files", "*.*")])
        if filename:
            self.ent_excel.delete(0, "end")
            self.ent_excel.insert(0, os.path.normpath(filename))

    def browse_template(self):
        filename = filedialog.askopenfilename(filetypes=[("Word Template", "*.docx"), ("All Files", "*.*")])
        if filename:
            self.ent_template.delete(0, "end")
            self.ent_template.insert(0, os.path.normpath(filename))

    def browse_output(self):
        foldername = filedialog.askdirectory()
        if foldername:
            self.ent_output.delete(0, "end")
            self.ent_output.insert(0, os.path.normpath(foldername))

    def start_generation(self):
        excel_file = self.ent_excel.get()
        template_file = self.ent_template.get()
        output_dir = self.ent_output.get()

        if not os.path.exists(excel_file) or not os.path.exists(template_file) or not output_dir:
            self.log("ERROR: Please select valid files and an output directory.")
            return

        self.btn_generate.configure(state="disabled", text="Generating...")
        self.log("--- Starting Generation ---")
        threading.Thread(target=self.generate_docs, args=(excel_file, template_file, output_dir), daemon=True).start()

    def generate_docs(self, excel_file, template_file, output_dir):
        try:
            if excel_file.endswith(".csv"):
                df = pd.read_csv(excel_file, dtype={'Số điện thoại liên hệ': str, 'Mã số sinh viên (MSSV)': str})
            else:
                df = pd.read_excel(excel_file, dtype={'Số điện thoại liên hệ': str, 'Mã số sinh viên (MSSV)': str})

            df = process_mtec_data_v2(df)
            os.makedirs(output_dir, exist_ok=True)
            success_count = 0

            for index, row in df.iterrows():
                context = row.to_dict()
                for field in formating_columns:
                    if field in context:
                        context[field] = format_[field](context[field])

                doc = DocxTemplate(template_file)
                doc.render(context)

                mssv = context.get('mssv', 'Unknown_MSSV')
                ho_ten = context.get('ho_ten', f'Unknown_Name_{index}')
                filename = f"HOSO_{mssv}_{ho_ten}.docx".replace(" ", "_")
                
                output_file = os.path.join(output_dir, filename)
                doc.save(output_file)
                self.log(f"SUCCESS: Saved {filename}")
                success_count += 1

            self.log(f"--- Generation Complete! Successfully created {success_count} files. ---")
        except Exception as e:
            self.log(f"ERROR: {str(e)}")
        finally:
            self.btn_generate.configure(state="normal", text="Generate Documents")

if __name__ == "__main__":
    app = App()
    app.mainloop()