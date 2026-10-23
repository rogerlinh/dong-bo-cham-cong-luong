# Tool đồng bộ chấm công về file chuẩn

Tool đọc **bảng chấm công** của từng doanh nghiệp, dựng 2 tab `bang_cham_cong_thang` và `bang_cham_cong_tuan` theo đúng template, rồi đối chiếu với **bảng kê** của chính file đó.

**Tool không đọc bảng lương.** Bảng lương trong file gốc là bản đã tổng hợp. Template tự tạo lại `bang_luong` và `doanh_thu` từ 2 tab chấm công, nên tool không xuất 2 sheet này. Xóa hết các tab lương khỏi file gốc thì kết quả vẫn y hệt; bộ kiểm thử có kiểm tra điều này.

## Chạy
- **Giao diện (khuyên dùng):** double-click `Chay_app.bat`. Trình duyệt tự mở trang `http://127.0.0.1:8765`; nếu cổng này bận, tool tự chuyển sang cổng kế tiếp.
  1. Tool tự nạp các file `.xlsx` trong thư mục cha. Có thể kéo thả thêm file hoặc bấm **Chọn file**. Mỗi file hiện mã hợp đồng, biểu mẫu nhận diện được (A–E) và kỳ công.
  2. Bấm **Đồng bộ N file**. Muốn có thêm mỗi hợp đồng một file riêng thì tick **Xuất thêm mỗi hợp đồng một file riêng** trước khi bấm.
  3. Bảng kết quả hiện cho từng hợp đồng: kỳ công, số NV, **giờ công (chấm công / bảng kê)**, sheet nguồn, số tuần, cảnh báo.
     - **KHỚP:** giờ từng NV, từng loại giờ trong chấm công bằng đúng bảng kê.
     - **LỆCH n NV:** có n NV lệch giờ, có công mà không có trong bảng kê, hoặc ngược lại.
     - **Không có bảng kê:** file không có sheet bảng kê để đối chiếu.
  4. Bấm **Tải file tổng hợp**. File cũng được lưu trong thư mục `ket_qua` cạnh các file gốc (nút **Mở thư mục kết quả**).
  - Đóng cửa sổ console đen là tắt tool. Không cần internet.
- **Cửa sổ đơn giản (dự phòng):** `python tool_dong_bo/app_dong_bo.py`.
- **Dòng lệnh:**
  ```
  python tool_dong_bo/dong_bo_luong.py --thu-muc "<thư mục chứa file xlsx>" -o Tong_hop.xlsx [--tach-rieng]
  ```
  `--tach-rieng` xuất thêm mỗi hợp đồng một file chuẩn riêng.

Cần Python 3 và openpyxl.

## File kết quả (cấu trúc theo template Google Sheet)

| Sheet | Nội dung |
|---|---|
| `DM` | A–F là danh mục loại giờ, 8 dòng đầu đúng thứ tự template; 2 loại bổ sung *Tăng ca đêm chủ nhật*, *Tăng ca đêm ngày lễ* để ở cuối. **E (đơn giá lương) để trống** vì tool không đọc bảng lương: điền tay hoặc để template tính. **F (đơn giá dịch vụ)** lấy từ bảng kê, chỉ điền khi file có 1 hợp đồng; file gộp nhiều hợp đồng xem đơn giá từng hợp đồng từ cột V. I–M là danh sách NV như template (mã, tên, tỉnh, CCCD, STK); N–T bổ sung bộ phận, chức danh, NCC, mã thẻ, ngày vào, ngày nghỉ, hợp đồng. |
| `bang_cham_cong_thang` | Đúng khung template: dòng 1–2 tiêu đề, dòng 3 trống, dữ liệu từ dòng 4. Tiêu đề có **ngày thật đúng kỳ công của bảng chấm công gốc** (A 01–30/09; B, D 01–31/08; C, E 26/07–25/08). Mỗi NV 4 dòng loại giờ (gộp ô mã/tên như template), thêm dòng nếu NV có giờ CN/lễ. F = tổng G:AK. File gộp nhiều hợp đồng thì mỗi hợp đồng một khối, tên hợp đồng ở A1 của khối. AM–AP: Hợp đồng, Kỳ, Từ ngày, Đến ngày. |
| `bang_cham_cong_tuan` | Mỗi **tuần (Thứ 2 → Chủ nhật)** của kỳ công là một khối giống template tuần: G = Thứ 2 … M = Chủ nhật, cột E "Loại công", mỗi NV 4 dòng (không gộp ô), F = tổng 7 ngày. Mọi NV của hợp đồng có mặt ở mọi tuần. Tuần đầu và cuối kỳ có thể chứa ngày ngoài kỳ: vẫn hiện ngày (ô tiêu đề tô nhạt) để cột luôn đúng thứ. O–S: Hợp đồng, Kỳ, Tuần thứ mấy, Từ ngày, Đến ngày. Tổng các tuần của một NV đúng bằng cột F của sheet tháng. |
| `cham_cong_chi_tiet` | Dạng dọc, mỗi dòng là một NV × ngày thật × loại giờ. Tiện đưa vào pivot hoặc Power BI. |
| `kiem_tra` | Bảng tổng từng hợp đồng: file, biểu mẫu, sheet chấm công, kỳ, số NV, số NV có công, giờ từng loại, tổng giờ công, **đối chiếu với sheet nào, tổng giờ bảng kê, chênh lệch, số NV lệch**, số lỗi/nghi ngờ/lưu ý. Bên dưới là danh sách cảnh báo, xem mục **Cảnh báo sai sót**. |

Giá trị được ghi thẳng dạng số, không dùng công thức, để tool phân tích đọc được ngay.

## Cách đọc từng biểu mẫu

Mẫu được nhận diện theo tên sheet, không theo tên file. Mã hợp đồng và kỳ lấy từ tên file, ví dụ `CTY_A_T09.2026`. Danh sách NV lấy từ bảng chấm công, kèm thông tin bổ sung từ bảng kê hoặc danh sách NV nếu có.

| Mẫu | Nhận diện | Bảng chấm công (nguồn 2 tab) | Bảng kê (đối chiếu, đơn giá dịch vụ) |
|---|---|---|---|
| A | sheet bắt đầu bằng `BCC` | `BCC…`, 4 dòng/NV theo nhãn loại giờ | `Bảng kê`: tổng giờ từng loại của cả hợp đồng |
| B | `Bảng công SK` | `Bảng công SK`, 6 dòng/NV | `Tính phí`: giờ từng NV, đơn giá dòng 6 |
| C | `Tuần` + `DM ca` | `Tuần`, 1 dòng/NV/ngày, giờ ở P–W | `Bảng kê T08`: giờ từng NV |
| D | sheet có `明细` | cache liên kết ngoài `DS TỔNG` có sẵn trong file; chỉ lấy NV có trong 明细 | `…明细`: giờ từng NV, đơn giá; cũng là danh sách NV |
| E | `Công (lương)` | `Công (lương)`, mỗi dòng một hệ số (1.0 = HC ngày, 1.3 = HC đêm, 1.5 = TC ngày, 2.0 / 2.3 = CN, 3.0 / 3.3 = lễ) | `Công` (bản gửi khách): giờ từng NV |

Khi một doanh nghiệp đổi biểu mẫu, hoặc có doanh nghiệp mới: viết thêm hàm `nhan_dien_X` và `doc_mau_X` trong `dong_bo_luong.py`, rồi khai báo trong `ADAPTERS`. Bộ đọc chỉ cần đọc bảng chấm công, và bảng kê nếu có.

**Kết quả đối chiếu với 5 file mẫu:**
- **A:** khớp đúng bảng kê (850,5 giờ).
- **B:** tổng giờ bằng `Tính phí` (14.171,5 giờ), nhưng 30 NV chia giờ ngày/đêm khác. Sheet `Bảng công nhà máy` thì khớp đúng `Tính phí`.
- **C:** 12 NV có công trong `Tuần` nhưng chưa có trong bảng kê.
- **D:** 64/66 NV khớp; YB20481 và YB20934 lệch giờ.
- **E:** tổng giờ bằng `Công` (20.213 giờ), 14 NV chia loại giờ khác.

## Cảnh báo sai sót

Mỗi lần đồng bộ, tool tự soát số liệu và chia cảnh báo làm 3 mức:

| Mức | Ý nghĩa | Việc cần làm |
|---|---|---|
| **Lỗi** (đỏ) | Số liệu sai chắc chắn, hoặc không tin được | Sửa file gốc rồi chạy lại |
| **Nghi ngờ** (vàng) | Bất thường, có thể đúng nhưng hiếm | Người phụ trách hợp đồng xem lại |
| **Lưu ý** (xám/xanh) | Thông tin để biết | Không bắt buộc xử lý |

**Xem cảnh báo ở đâu:**
- **Giao diện:** cột *Cảnh báo* hiện số lỗi, nghi ngờ, lưu ý của từng hợp đồng. Bấm vào để xem danh sách, có nút lọc theo mức.
- **File Excel, sheet `kiem_tra`:** bảng tổng có cột số lỗi / nghi ngờ / lưu ý; danh sách cảnh báo gồm Mức, Hợp đồng, Mã NV, Tên NV, Loại, Nội dung, có bộ lọc và xếp lỗi lên đầu.

**Các quy tắc kiểm tra:**

| Nhóm | Lỗi | Nghi ngờ | Lưu ý |
|---|---|---|---|
| File và kỳ | Bảng chấm công trống. Không có giờ công (ô giờ là công thức chưa có giá trị). Ô lỗi công thức (`#REF!`, `#N/A`…), được tính là 0. Kỳ công dài hơn 31 ngày. Không xác định được kỳ. | Kỳ trong file khác kỳ ở tên file. Nhãn loại giờ lạ. Trùng mã hợp đồng (2 file cùng mã, tool tự đổi tên). | Hợp đồng khác kỳ với các hợp đồng còn lại. Nguồn chấm công đặc biệt (D). |
| Nhân viên | Trùng mã NV ở 2 khối của bảng công (giờ không bị cộng đôi). Thiếu họ tên. Ngày nghỉ trước ngày vào. | Mã khác dạng các mã còn lại (vd `HMT 140`). Cùng tên và cùng ngày vào nhưng khác mã. Cùng mã ở 2 hợp đồng. Ngày vào sau kỳ công. | NV không có giờ công. Tên chữ Hán. |
| Chấm công ngày | Công ghi vào ngày ngoài kỳ. Một ngày quá 24 giờ. Giờ âm. | Một ngày quá 16 giờ. Có công sau ngày nghỉ việc hoặc trước ngày vào. Giờ thường ghi vào Chủ nhật. Trùng dòng chấm công. | Làm thêm vượt 40 giờ/tháng. |
| Đối chiếu bảng kê | | Lệch giờ chấm công / bảng kê (từng NV, từng loại giờ). Có công nhưng không có trong bảng kê. Có trong bảng kê nhưng không có chấm công. Lệch tổng giờ với bảng kê. | Không có bảng kê. |

**File không đọc được** (không đúng mẫu nào, hoặc đúng tên sheet nhưng cấu trúc đã đổi) không làm hỏng cả lượt: tool bỏ qua file đó, báo rõ lý do trên giao diện và liệt kê trong sheet `kiem_tra`.

**Ngưỡng:** sửa trong `NGUONG` ở đầu `dong_bo_luong.py`.
- Nghi ngờ khi một ngày quá 16 giờ; lỗi khi quá 24 giờ.
- Làm thêm tối đa 40 giờ/tháng, theo BLLĐ 2019 Điều 107.
- Lệch giờ chấm công / bảng kê trên 0,01 giờ thì tính là lệch.

## Lưu ý dữ liệu nguồn
- **Chưa có ở cả 5 file:** CCCD, tỉnh thành, số tài khoản trong bảng chấm công. Các cột này ở `DM` để trống.
- **Đơn giá trong DM của template bị đảo:** template ghi 37.500 / 55.328 cho "TC ngày 150%" và 45.000 / 63.840 cho "HC đêm 130%". Bảng kê CTY A lại dùng ngược lại (đúng hệ số 1,5 và 1,3). Tool lấy đơn giá theo bảng kê.
- **CTY B, NV HMT18:** ngày nghỉ (22/03/2026) trước ngày vào (29/05/2026) nhưng vẫn có công tháng 8.
