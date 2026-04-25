# MTEC Document Generator

Bộ công cụ hỗ trợ tự động hóa việc tạo hàng loạt file Word (hồ sơ, đơn đăng ký,...) từ dữ liệu Excel/CSV và một file Word mẫu (Template `.docx`). Công cụ này được thiết kế đặc biệt với các logic xử lý dữ liệu, chuẩn hóa thông tin và tự động đánh dấu (checkbox) dựa trên dữ liệu của MTEC Club.

## 📂 Thành phần

- **`MTEC_App.py`**: Ứng dụng Desktop (.exe) với giao diện người dùng được xây dựng bằng `customtkinter`. Phù hợp để chạy cục bộ trên máy tính cá nhân.
- **`MTEC_Web.py`**: Ứng dụng Web được xây dựng bằng `streamlit`. Phù hợp để triển khai nội bộ hoặc chia sẻ qua mạng.
- **`convert.ipynb`**: File Jupyter Notebook gốc chứa toàn bộ quá trình research, thử nghiệm logic format dữ liệu và thao tác tạo file Word.

## 🚀 Tính năng nổi bật

- Chuyển đổi hàng loạt từ dữ liệu CSV/Excel sang các file Word riêng biệt.
- Tự động nhận diện và thay thế các biến (vd: `{{ ho_ten }}`) trong file mẫu.
- Tự động chuẩn hóa dữ liệu:
  - In hoa họ tên.
  - Chuẩn hóa định dạng ngày sinh (dd/mm/yyyy).
  - Chuẩn hóa số điện thoại.
- Tự động sinh Checkbox (`☒`, `☐`) theo lựa chọn Ban, Vị trí và Đánh giá kỹ năng từ form đăng ký.
- Nén toàn bộ output thành 1 file ZIP (đối với bản Web) hoặc lưu trực tiếp vào thư mục chỉ định (đối với bản App).

## 🛠 Cài đặt môi trường

Cài đặt các thư viện Python cần thiết trước khi sử dụng:

```bash
pip install pandas docxtpl customtkinter streamlit openpyxl
```

## 🖥 Hướng dẫn sử dụng

### 1. Chạy ứng dụng Desktop (MTEC_App.py)

Mở Terminal và chạy lệnh:
```bash
python MTEC_App.py
```
**Cách đóng gói thành file `.exe`:**
```bash
pip install pyinstaller
pyinstaller --noconfirm --onedir --windowed MTEC_App.py
```
> Sau khi chạy xong, file `MTEC_App.exe` sẽ nằm trong thư mục `dist/MTEC_App/`.

### 2. Chạy ứng dụng Web (MTEC_Web.py)

Mở Terminal và chạy lệnh:
```bash
streamlit run MTEC_Web.py
```
Trình duyệt sẽ tự động mở trang web ở địa chỉ `http://localhost:8501`.

## 📝 Chuẩn bị file Template (.docx)

File Word mẫu cần chứa các biến môi trường được bọc trong bộ ngoặc nhọn kép. Ví dụ: `{{ ho_ten }}`, `{{ mssv }}`, `{{ sdt }}`, `{{ c_ban_cn }}`...
Tên biến phải khớp chính xác với key/cột đã được code quy định.