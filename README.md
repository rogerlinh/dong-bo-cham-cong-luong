# Đồng bộ chấm công và lương

Tool gộp bảng chấm công và bảng lương của nhiều hợp đồng về **một file Excel cấu trúc chuẩn**. Mỗi hợp đồng có thể dùng một biểu mẫu khác nhau. Sau khi gộp, tool đối chiếu lại với file gốc và cảnh báo khi số liệu sai hoặc đáng nghi.

## Dùng ngay trên web, không cần cài

Mở **<https://rogerlinh.github.io/dong-bo-cham-cong-luong/>** rồi kéo thả file vào. Tool chạy hẳn trong trình duyệt bằng [Pyodide](https://pyodide.org) (Python dạng WebAssembly), nên file của bạn không tải lên máy chủ nào.

- Lần đầu mở, trang tải bộ xử lý khoảng 20 MB. Những lần sau trình duyệt đã lưu sẵn nên nhanh hơn.
- Đồng bộ 5 file mẫu mất khoảng 30–60 giây, chậm hơn bản cài trên máy.
- Dùng Chrome, Edge hoặc Firefox bản mới.

## Cài và chạy trên máy (không cần internet)

1. Cài **Python 3.10 trở lên** từ <https://www.python.org/downloads/>. Trên Windows, nhớ tick **Add python.exe to PATH** khi cài.
2. Tải source về: bấm **Code → Download ZIP** rồi giải nén, hoặc dùng `git clone`.
3. Mở tool:
   - **Windows:** double-click `tool_dong_bo\Chay_app.bat`.
   - **macOS / Linux:** chạy `sh tool_dong_bo/chay_app.sh`.

   Lần đầu chạy, tool tự cài thư viện `openpyxl`. Trình duyệt sẽ tự mở `http://127.0.0.1:8765`.
4. Kéo thả các file `.xlsx` bảng công/lương vào trang, rồi bấm **Đồng bộ**.
5. Xem bảng đối chiếu và cảnh báo, rồi bấm **Tải file tổng hợp**. File kết quả cũng được lưu trong thư mục `ket_qua/`.

Đóng cửa sổ console là tắt tool.

Mẹo: chép các file `.xlsx` vào thư mục gốc của source, tức là cạnh file README này. Mỗi lần mở, tool sẽ tự nạp sẵn các file đó.

## Biểu mẫu hỗ trợ

Tool hiện đọc được 5 biểu mẫu, gọi là **A, B, C, D, E**, và nhận diện mẫu theo tên sheet trong file. Mã hợp đồng và kỳ lương được lấy từ tên file, ví dụ `Bảng_chi_lương_CTY_A_T09.2026.xlsx`.

Nếu file không thuộc mẫu nào, giao diện báo *Không nhận diện được biểu mẫu* và bỏ qua file đó. Muốn thêm mẫu mới, hãy viết thêm hàm `nhan_dien_X` và `doc_mau_X` trong `tool_dong_bo/dong_bo_luong.py`, rồi khai báo hàm vào `ADAPTERS`.

Chi tiết cấu trúc file kết quả, cách đọc từng mẫu và danh sách quy tắc cảnh báo nằm ở [tool_dong_bo/HUONG_DAN.md](tool_dong_bo/HUONG_DAN.md).

## Cấu trúc source

| File | Vai trò |
|---|---|
| `tool_dong_bo/dong_bo_luong.py` | Lõi xử lý: đọc 5 biểu mẫu, kiểm tra cảnh báo, ghi file chuẩn. Có thể chạy bằng dòng lệnh. |
| `tool_dong_bo/app_web.py` | Máy chủ web nội bộ (chỉ dùng thư viện chuẩn của Python) |
| `tool_dong_bo/giao_dien.html` | Giao diện |
| `tool_dong_bo/app_dong_bo.py` | Cửa sổ tkinter đơn giản, dự phòng |
| `tool_dong_bo/Chay_app.bat`, `chay_app.sh` | Lệnh mở tool |

Chạy bằng dòng lệnh:

```
python tool_dong_bo/dong_bo_luong.py --thu-muc <thư mục chứa xlsx> -o Tong_hop.xlsx [--tach-rieng]
```
