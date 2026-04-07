# AppPrintPO

File Excel (đúng thứ tự cột không bắt buộc, chỉ cần đủ tên cột): **Mã đơn**, **Sku**, **Số lượng**, **Mẫu**, **Loại**, **Ngang**, **Cao**. Trên PDF, cột *Tên sản phẩm* hiển thị dạng **Loại - Ngang x Cao - Mẫu**. Khổ in **12×18 cm** (tăng 20 % so với 10×15 cm).

## Cài đặt chạy từ mã nguồn

1. Cài **Python 3.10+** (Windows thường đã có Tkinter).
2. Mở **Command Prompt** hoặc **PowerShell**, vào thư mục project, chạy:

   pip install -r requirements.txt

3. Chạy ứng dụng:

   python main.py

## Tạo file `.exe` có icon

1. Đảm bảo đã cài Python như trên (và đã `pip install -r requirements.txt` nếu bạn chạy app từ mã nguồn).
2. Double-click **`build_exe.bat`** hoặc trong terminal:

   build_exe.bat
