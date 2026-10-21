# Tool đồng bộ bảng chấm công / lương về file chuẩn

## Chạy
- **Giao diện (khuyên dùng):** double-click `Chay_app.bat`. Trình duyệt tự mở trang `http://127.0.0.1:8765`; nếu cổng này bận, tool tự chuyển sang cổng kế tiếp.
  1. Tool tự nạp các file `.xlsx` trong thư mục cha. Có thể kéo thả thêm file hoặc bấm **Chọn file**. Mỗi file hiện mã hợp đồng, biểu mẫu nhận diện được (A–E) và kỳ lương.
  2. Bấm **Đồng bộ N file**. Muốn có thêm mỗi hợp đồng một file riêng thì tick **Xuất thêm mỗi hợp đồng một file riêng** trước khi bấm.
  3. Bảng **Kết quả đối chiếu** hiện cho từng hợp đồng: số NV, tổng lương, thực nhận (file tổng / file gốc), doanh thu, cảnh báo. Dấu **KHỚP** nghĩa là thực nhận của từng nhân viên bằng đúng file gốc; **LỆCH** kèm số chênh. Bấm "N mục" để xem cảnh báo.
  4. Bấm **Tải file tổng hợp**. File cũng được lưu trong thư mục `ket_qua` cạnh các file gốc (nút **Mở thư mục kết quả**).
  - Đóng cửa sổ console đen là tắt tool. Không cần internet.
- **Cửa sổ đơn giản (dự phòng):** `python tool_dong_bo/app_dong_bo.py`.
- **Dòng lệnh:**
  ```
  python tool_dong_bo/dong_bo_luong.py --thu-muc "<thư mục chứa file xlsx>" -o Tong_hop.xlsx [--tach-rieng]
  ```
  `--tach-rieng` xuất thêm mỗi dự án một file chuẩn riêng.

Cần Python 3 và openpyxl.

## File kết quả (cấu trúc theo template Google Sheet)

| Sheet | Nội dung |
|---|---|
| `DM` | A–F là danh mục loại giờ, 8 dòng đầu đúng thứ tự template; 2 loại bổ sung *Tăng ca đêm chủ nhật*, *Tăng ca đêm ngày lễ* để ở cuối. Đơn giá E:F (và hàng 1 D:G của `bang_luong`) chỉ điền khi file có 1 hợp đồng; file gộp nhiều hợp đồng xem đơn giá từng hợp đồng ở X–AC. I–M là danh sách NV như template (mã, tên, tỉnh, CCCD, STK); N–V bổ sung ngân hàng, chủ TK, bộ phận, chức danh, NCC, mã thẻ, ngày vào, ngày nghỉ, hợp đồng. X–AC là đơn giá lương và đơn giá dịch vụ theo từng dự án. |
| `bang_cham_cong` | Chia khối theo **hợp đồng (dự án)**. Mỗi khối giống hệt template: hàng tên hợp đồng + thứ, hàng tiêu đề có **ngày thật đúng kỳ công của bảng chấm công gốc** (A 01–30/09; B, D 01–31/08; C, E 26/07–25/08). Cột G = ngày đầu kỳ. Mỗi NV 4 dòng loại giờ, thêm dòng nếu có giờ CN/lễ. AM–AP: Hợp đồng, Kỳ, Từ ngày, Đến ngày. |
| `bang_luong` | A–U đúng các cột và **đúng công thức của template**: **Tổng lương Q = SUM(H:P)**, **Thực nhận U = Q − R − S − T**. Vì template không có cột riêng nên: **cột O "Phụ cấp khác"** = phụ cấp khác + lương TC ngày nghỉ/CN/lễ + trả lại tiền giữ; **cột R "Trừ đồ"** = trừ đồ + bảo hiểm + khấu trừ khác. V–AP là cột bổ sung **chỉ để xem chi tiết, không cộng thêm vào Q/U**: giờ/lương TC ngày nghỉ & lễ, trả lại tiền giữ, bảo hiểm, khấu trừ khác (các khoản đã nằm trong O/R), Dự án, Kỳ, Thực nhận file gốc, Chênh lệch, Ngày công, Ngày vào, Ngày nghỉ, **Chi tiết khoản gộp** (ứng từng đợt, trừ KTX/thẻ/đồng phục, phụ cấp ăn/công đoạn…), Ghi chú, Cảnh báo. Cuối sheet có bảng tổng theo hợp đồng. Lưu ý: công thức tiền trong template ghép nhầm cột (I "Lương ca đêm" = E "Tăng ca ngày 150%" × đơn giá); tool ghép theo **tên cột**: H = HC ngày, I = HC đêm, J = TC ngày, K = TC đêm. |
| `doanh_thu` | Mỗi dự án một bảng kê công nợ, đặt đúng vị trí dòng như template (tiêu đề bảng ở dòng 14 của khối, TỔNG, THUẾ VAT, TỔNG TIỀN THANH TOÁN, Bằng chữ, thông tin chuyển khoản). Đối chiếu với tổng ở file gốc nằm ở cột H–I. Thông tin chuyển khoản sửa trong `THONG_TIN_CHUYEN_KHOAN` ở đầu `dong_bo_luong.py`. |
| `cham_cong_chi_tiet` | Dạng dọc, mỗi dòng là một NV × ngày thật × loại giờ. Tiện đưa vào tool phân tích, pivot hoặc Power BI. |
| `kiem_tra` | Đối chiếu từng dự án (giờ công so với giờ lương, thực nhận tính lại so với file gốc, doanh thu, số lỗi/nghi ngờ/lưu ý) và danh sách cảnh báo có mức, xem mục **Cảnh báo sai sót**. |

Giá trị được ghi thẳng dạng số, không dùng công thức, để tool phân tích đọc được ngay.

## Cảnh báo sai sót

Mỗi lần đồng bộ, tool tự soát số liệu và chia cảnh báo làm 3 mức:

| Mức | Ý nghĩa | Việc cần làm |
|---|---|---|
| **Lỗi** (đỏ) | Số liệu sai chắc chắn, hoặc không tin được | Sửa file gốc rồi chạy lại trước khi đưa vào tool phân tích |
| **Nghi ngờ** (vàng) | Bất thường, có thể đúng nhưng hiếm | Người phụ trách hợp đồng xem lại |
| **Lưu ý** (xám/xanh) | Thông tin để biết | Không bắt buộc xử lý |

**Xem cảnh báo ở đâu:**
- **Giao diện:** cột *Cảnh báo* hiện số lỗi, nghi ngờ, lưu ý của từng hợp đồng. Bấm vào để xem danh sách, có nút lọc theo mức. Khi có lỗi, khung đỏ phía trên bảng sẽ báo.
- **File Excel, sheet `kiem_tra`:**
  - Bảng tổng có cột *Số lỗi / Số nghi ngờ / Số lưu ý*.
  - Danh sách cảnh báo gồm Mức, Hợp đồng, Mã NV, Tên NV, Loại, Nội dung. Danh sách có sẵn bộ lọc và xếp lỗi lên đầu.
- **File Excel, sheet `bang_luong`:** cột cuối *Cảnh báo* ghi tóm tắt cảnh báo của từng nhân viên. Mã và tên NV có cảnh báo được tô màu theo mức cao nhất.

**Các quy tắc kiểm tra:**

| Nhóm | Lỗi | Nghi ngờ | Lưu ý |
|---|---|---|---|
| File và cấu trúc | Ô lỗi công thức (`#REF!`, `#N/A`, `#VALUE!`…), được tính là 0. Bảng lương trống. Không tìm thấy chấm công ngày. Dòng TỔNG của file gốc khác tổng các dòng. **Có thể sót NV** (dừng đọc bảng lương mà bên dưới vẫn còn mã giống mã NV). **Không có giá trị tiền** (file lưu bằng công cụ không tính công thức). **Kỳ công dài hơn 31 ngày.** **Không xác định được kỳ** (mẫu D, tên file thiếu `Txx.yyyy`). | Kỳ ghi trong file khác kỳ ở tên file. Nhãn loại giờ / hệ số lạ. Trùng mã hợp đồng (2 file cùng mã, tool tự đổi tên). | Hợp đồng khác kỳ với các hợp đồng còn lại (có thể chọn nhầm file tháng khác). Nguồn chấm công đặc biệt (D). |
| Nhân viên | Trùng mã NV trong một bảng lương. Thiếu họ tên. Ngày nghỉ trước ngày vào. | Mã khác dạng các mã còn lại (vd `HMT 140` thừa khoảng trắng). Cùng tên và cùng ngày vào nhưng khác mã. Cùng mã ở 2 hợp đồng. Cùng số tài khoản cho 2 NV khác nhau. Ngày vào sau kỳ công. **Tiền giữ suy ra** (mẫu C, NV không PENDING mà có tiền giữ: ô Thực nhận có thể bị gõ đè). | NV có công nhưng không có trong bảng lương. |
| Giờ và tiền | Có tiền mà không có giờ. Tiền ≠ giờ × đơn giá (so với mọi bộ đơn giá của hợp đồng). Giờ hoặc tiền âm. Lương/giờ dưới tối thiểu vùng IV. Ngày công vượt số ngày của kỳ. Thực nhận tính lại khác file gốc. | Có giờ mà tiền bằng 0. Giờ bảng lương khác tổng chấm công ngày. Thực nhận âm. Khoản trừ âm. | Lương/giờ dưới tối thiểu vùng I. Tổng tăng ca trên 40 giờ/kỳ. |
| Chấm công ngày | Công ghi vào ngày ngoài kỳ hoặc ngoài 31 cột ngày. Một ngày quá 24 giờ. | Một ngày quá 16 giờ. Có công sau ngày nghỉ việc hoặc trước ngày vào. Giờ thường ghi vào Chủ nhật, khi hợp đồng có đơn giá riêng cho Chủ nhật. NV có giờ lương nhưng không có chấm công ngày. | |
| Doanh thu | Doanh thu tính lại khác bảng kê gốc. | Doanh thu thấp hơn quỹ lương. | Thiếu đơn giá dịch vụ cho loại giờ có phát sinh. |

**File không đọc được** (không đúng mẫu nào, hoặc đúng tên sheet nhưng cấu trúc đã đổi) không làm hỏng cả lượt: tool bỏ qua file đó, báo rõ lý do trên giao diện và liệt kê ở đầu danh sách cảnh báo trong sheet `kiem_tra`.

**Dòng trống trong bảng lương:** tool bỏ qua tối đa 2 dòng trống liền nhau; gặp 3 dòng trống liền thì coi là hết bảng (và báo lỗi nếu bên dưới vẫn còn mã NV).

**Ngưỡng:** sửa trong `NGUONG` ở đầu `dong_bo_luong.py`.
- Nghi ngờ khi một ngày quá 16 giờ; lỗi khi quá 24 giờ.
- Tăng ca tối đa 40 giờ/tháng, theo BLLĐ 2019 Điều 107.
- Lương tối thiểu giờ: vùng I 25.500đ, vùng IV 17.800đ, theo NĐ 293/2025/NĐ-CP.
- Lệch tiền cho phép: 1đ cộng 0,1%.

Nếu chính sách của công ty khác, hãy chỉnh các số này.

**Cảnh báo hiện có với 5 file mẫu** (đều do dữ liệu gốc, không phải lỗi tool):
- **B, NV HMT18:** ngày nghỉ (22/03/2026) trước ngày vào (29/05/2026) nhưng vẫn có công tháng 8.
- **B, NV `HMT 140`:** mã có khoảng trắng.
- **C, 5 NV và E, 1 NV:** thực nhận âm.
- **D:**
  - 2 NV lệch giờ OT giữa bảng công và bảng lương.
  - 2 cặp NV dùng chung số tài khoản.
- **E, 2 NV:** có công sau ngày nghỉ việc.

## Cách đọc từng biểu mẫu

| Mẫu | Nhận diện bằng sheet | Chấm công ngày | Bảng lương | Doanh thu |
|---|---|---|---|---|
| A | `BCC…` + `Bảng chi lương` | BCC, 4 dòng/NV | Bảng chi lương | giờ × đơn giá `Bảng phí` |
| B | `Bảng công SK` + `Bảng lương` | Bảng công SK, 6 dòng/NV | Bảng lương (ứng = 4 đợt) | sheet `Tính phí`, tính theo giờ nhà máy |
| C | `Tuần` + `DM ca` + `Bảng lương` | Tuần (1 dòng/NV/ngày, giờ quy từ mã ca) | Bảng lương (giờ đào tạo cộng vào HC ngày) | `Bảng kê T08` |
| D | `Chi lương` + `…明细` | lấy từ cache liên kết ngoài `DS TỔNG` có sẵn trong file | Chi lương + UL (STK) | ước tính theo đơn giá 明细 |
| E | `Công (lương)` + `Bảng lương` | Công (lương), 7 dòng/NV theo hệ số | Bảng lương | ước tính theo đơn giá dòng 3 Công (lương), **chưa xác nhận** |

Mẫu được nhận diện theo tên sheet, không theo tên file. Mã dự án và kỳ lấy từ tên file, ví dụ `CTY_A_T09.2026`.

Khi một dự án đổi biểu mẫu, hoặc có dự án mới: viết thêm hàm `nhan_dien_X` và `doc_mau_X` trong `dong_bo_luong.py`, rồi khai báo trong `ADAPTERS`.

## Lưu ý dữ liệu nguồn
- **Chưa có ở cả 5 file:** CCCD, tỉnh thành, phụ cấp chức vụ, phụ cấp thâm niên. Các cột này để trống hoặc bằng 0. Chỉ CTY D có số tài khoản (sheet UL).
- **Đơn giá trong DM của template bị đảo:** template ghi 37.500 / 55.328 cho "TC ngày 150%" và 45.000 / 63.840 cho "HC đêm 130%". File CTY A lại dùng ngược lại. Tool lấy đơn giá theo file gốc.
- **CTY D:** 2 NV (YB20481, YB20934) được gõ tay OT 46h, trong khi chấm công ngày chỉ có 42h và 44h. Tool giữ số của bảng lương và ghi cảnh báo.
- **CTY E:** 2 NV có công nhưng không nằm trong bảng lương (T0036626, T0036633). Tool loại 2 NV này ra và ghi cảnh báo.
- **CTY C:** NV PENDING bị giữ toàn bộ lương. Khoản này được đưa vào cột "Tiền giữ", nên thực nhận vẫn khớp file gốc.
