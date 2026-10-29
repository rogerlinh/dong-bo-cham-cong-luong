# Tool đồng bộ chấm công về file chuẩn

Tool đọc **bảng chấm công** của từng doanh nghiệp và tách thành 2 loại công, mỗi loại một tab tháng và một tab tuần theo template:

| Tab | Dùng để |
|---|---|
| `bang_cham_cong_thang_nha_may`, `bang_cham_cong_tuan_nha_may` | Tính **doanh thu** SK thu của nhà máy |
| `bang_cham_cong_thang_SK`, `bang_cham_cong_tuan_SK` | Tính **chi phí** SK trả người lao động |

- **Dự án có 2 bảng công** (một bản của nhà máy, một bản của SK): mỗi bảng được nạp vào tab tương ứng.
- **Dự án chỉ có 1 bảng công:** bảng đó được nạp vào cả tab nhà máy và tab SK.
- Tool tự nhận biết dự án chấm công **theo tuần hay theo tháng** từ các ngày có trong bảng chấm công: kỳ công không quá 7 ngày là **theo tuần**, dài hơn là **theo tháng**. Dự án theo tháng chỉ nạp vào 2 tab **tháng**, dự án theo tuần chỉ nạp vào 2 tab **tuần**. Tab không có dự án nào ghi "Không có dự án chấm công theo tháng/tuần trong lượt đồng bộ này".
- Trong cùng loại tháng/tuần: dự án có 2 bảng công thì công nhà máy vào tab nhà máy, công SK vào tab SK; dự án 1 bảng công thì cùng dữ liệu ở cả tab nhà máy và tab SK.

Sau đó tool đối chiếu công nhà máy với công SK, và công nhà máy với **bảng kê** (bảng tính tiền với nhà máy) của chính file đó.

**Tool không đọc bảng lương.** Bảng lương trong file gốc là bản đã tổng hợp. `bang_luong` và `doanh_thu` được tính ở bước sau từ các tab chấm công, nên tool không xuất 2 sheet này. Xóa hết các tab lương khỏi file gốc thì kết quả vẫn y hệt; bộ kiểm thử có kiểm tra điều này.

## Chạy
- **Giao diện (khuyên dùng):** double-click `Chay_app.bat`. Trình duyệt tự mở trang `http://127.0.0.1:8765`; nếu cổng này bận, tool tự chuyển sang cổng kế tiếp.
  1. Tool tự nạp các file `.xlsx` trong thư mục cha. Có thể kéo thả thêm file hoặc bấm **Chọn file**. Mỗi file hiện mã hợp đồng, biểu mẫu nhận diện được (A–E) và kỳ công.
  2. Bấm **Đồng bộ N file**. Muốn có thêm mỗi hợp đồng một file riêng thì tick **Xuất thêm mỗi hợp đồng một file riêng** trước khi bấm.
  3. Bảng kết quả hiện cho từng hợp đồng: kỳ công, chấm công theo tháng hay tuần, số NV, **tab nguồn và giờ công của nhà máy và SK** (hoặc "Dùng chung" khi chỉ có 1 bảng công), giờ bảng kê, cảnh báo, và 2 con dấu:
     - Con dấu thứ nhất so **công nhà máy với công SK**: **KHỚP**, **LỆCH n NV**, hoặc **1 bảng công**.
     - Con dấu thứ hai so **công nhà máy với bảng kê**: **KHỚP**, **LỆCH n NV**, hoặc **không có bảng kê**.
     - Cả hai xét từng NV, từng loại giờ. LỆCH n NV gồm NV lệch giờ, NV chỉ có ở một bên.
  4. Bấm **Tải file tổng hợp**. File cũng được lưu trong thư mục `ket_qua` cạnh các file gốc (nút **Mở thư mục kết quả**).
  - Đóng cửa sổ console đen là tắt tool. Không cần internet.
- **Cửa sổ đơn giản (dự phòng):** `python tool_dong_bo/app_dong_bo.py`.
- **Dòng lệnh:**
  ```
  python tool_dong_bo/dong_bo_luong.py --thu-muc "<thư mục chứa file xlsx>" -o Tong_hop.xlsx [--tach-rieng]
  ```
  `--tach-rieng` xuất thêm mỗi hợp đồng một file chuẩn riêng.

Cần Python 3 và openpyxl.

## File kết quả (cấu trúc theo template)

| Sheet | Nội dung |
|---|---|
| `DM` | A–F là danh mục loại giờ, 8 dòng đầu đúng thứ tự template; 2 loại bổ sung *Tăng ca đêm chủ nhật*, *Tăng ca đêm ngày lễ* để ở cuối. **E (đơn giá lương) để trống** vì tool không đọc bảng lương: điền tay hoặc để template tính. **F (đơn giá dịch vụ)** lấy từ bảng kê, chỉ điền khi file có 1 hợp đồng; file gộp nhiều hợp đồng xem đơn giá từng hợp đồng từ cột V. I–M là danh sách NV như template (mã, tên, tỉnh, CCCD, STK); N–T bổ sung bộ phận, chức danh, NCC, mã thẻ, ngày vào, ngày nghỉ, hợp đồng. Danh sách NV gồm NV có ở bảng công nhà máy hoặc bảng công SK. |
| `bang_cham_cong_thang_nha_may`, `bang_cham_cong_thang_SK` | Chỉ chứa dự án chấm công theo tháng. Đúng khung template: dòng 1–2 tiêu đề, dòng 3 trống, dữ liệu từ dòng 4. Tiêu đề có **ngày thật đúng kỳ công của bảng chấm công gốc** (A 01–30/09; B, D 01–31/08; C, E 26/07–25/08). Mỗi NV 4 dòng loại giờ (gộp ô mã/tên như template), thêm dòng nếu NV có giờ CN/lễ. F = tổng G:AK. File gộp nhiều hợp đồng thì mỗi hợp đồng một khối, tên hợp đồng ở A1 của khối. AM–AP: Hợp đồng, Kỳ, Từ ngày, Đến ngày. |
| `bang_cham_cong_tuan_nha_may`, `bang_cham_cong_tuan_SK` | Chỉ chứa dự án chấm công theo tuần. Mỗi **tuần (Thứ 2 → Chủ nhật)** của kỳ công là một khối giống template tuần: G = Thứ 2 … M = Chủ nhật, cột E "Loại công", mỗi NV 4 dòng (không gộp ô), F = tổng 7 ngày. Mọi NV của hợp đồng có mặt ở mọi tuần. Tuần đầu và cuối kỳ có thể chứa ngày ngoài kỳ: vẫn hiện ngày (ô tiêu đề tô nhạt) để cột luôn đúng thứ. O–S: Hợp đồng, Kỳ, Tuần thứ mấy, Từ ngày, Đến ngày. |
| Một tab cho mỗi hợp đồng (`CTY A`, `CTY B`…) | Chỉ có ở file gộp nhiều hợp đồng, để đối chiếu với từng file gốc. Hợp đồng theo tháng dùng bố cục tab tháng, theo tuần dùng bố cục tab tuần. Hợp đồng có 2 bảng công: khối **công nhà máy** rồi khối **công SK** ngay dưới, **ô công SK khác công nhà máy tô đỏ** (cả ô ngày và cột F). Hợp đồng 1 bảng công: một khối ghi "CÔNG NHÀ MÁY = CÔNG SK". |
| `cham_cong_chi_tiet` | Dạng dọc, mỗi dòng là một NV × ngày thật × loại giờ, cột K cho biết bảng công (Nhà máy / SK / Nhà máy = SK), cột L là sheet nguồn. Tiện đưa vào pivot hoặc Power BI. |
| `kiem_tra` | Bảng tổng từng hợp đồng: file, biểu mẫu, kỳ, chấm công theo tháng hay tuần, số NV, **sheet và giờ công nhà máy, sheet và giờ công SK, chênh lệch và số NV lệch nhà máy / SK**, bảng kê đối chiếu, giờ bảng kê, chênh lệch và số NV lệch với công nhà máy, số lỗi/nghi ngờ/lưu ý. Bên dưới là danh sách cảnh báo, xem mục **Cảnh báo sai sót**. |

Giá trị được ghi thẳng dạng số, không dùng công thức, để tool phân tích đọc được ngay.

## Cách đọc từng biểu mẫu

Mẫu được nhận diện theo tên sheet, không theo tên file. Mã hợp đồng và kỳ lấy từ tên file, ví dụ `CTY_A_T09.2026`.

| Mẫu | Nhận diện | Công nhà máy | Công SK | Bảng kê (đối chiếu công nhà máy) |
|---|---|---|---|---|
| A | sheet bắt đầu bằng `BCC` | `BCC…` (dùng chung), 4 dòng/NV theo nhãn loại giờ | như nhà máy | `Bảng kê`: tổng giờ từng loại của cả hợp đồng |
| B | `Bảng công SK` | ` Bảng công nhà máy`, 6 dòng/NV (Giờ hành chính, Giờ ban đêm, Ngoài giờ hành chính…) | `Bảng công SK`, 6 dòng/NV (Ca ngày, Ca đêm, Tăng ca ngày…) | `Tính phí`: giờ từng NV, đơn giá dòng 6 |
| C | `Tuần` + `DM ca` | `Tuần`: dữ liệu xuất từ hệ thống nhà máy, 1 dòng/NV/ngày, giờ ở P–W quy đổi chuẩn theo mã ca (`DM ca`) | `DL công (tính lương)` nếu cùng kỳ với `Tuần`, giờ ở W–AD; không thì dùng chung `Tuần` | `Bảng kê T08`: giờ từng NV |
| D | sheet có `明细` | cache liên kết ngoài `DS TỔNG` có sẵn trong file (dùng chung); chỉ lấy NV có trong 明细 | như nhà máy | `…明细`: giờ từng NV, đơn giá; cũng là danh sách NV |
| E | `Công (lương)` | `Công (lương)`: số liệu chết trích từ phần mềm chấm công nhà máy | `Công`: SK dùng công thức SUMIFS trên dữ liệu gốc `DL công` để kiểm tra lại công nhà máy | không có |

Mẫu E ghi giờ theo hệ số: 1.0 = HC ngày, 1.3 = HC đêm, 1.5 = TC ngày, 2.0 = TC ngày CN, 2.3 = TC đêm CN, 3.0 = TC đêm, 3.3 = TC đêm lễ. Riêng hệ số 2.0 thì dữ liệu gốc dùng cho cả TC đêm ngày thường lẫn làm ngày nghỉ, nên cách gán này cần người phụ trách hợp đồng xác nhận. Cách gán này khớp đúng các cột của 'Bảng lương' mẫu E. Trong file tổng hợp, số dòng loại giờ của mỗi NV giống file gốc: mẫu E 7 dòng, mẫu B 6 dòng (dòng không có giờ ghi 0); mẫu A, C, D giữ 4 dòng chính như template, thêm dòng CN/lễ khi NV có giờ. Bảng kê '明细' của mẫu D cộng cả cột thử việc (V, X) với cột chính thức (W, Y) khi đối chiếu TC ngày / TC đêm.

**Bảng công thứ 2 chỉ được dùng khi cùng kỳ.** Tool so ngày và mã NV của bảng công thứ 2 với bảng công chính. Nếu phần lớn ngày nằm ngoài kỳ, hoặc không có NV chung, tool bỏ qua bảng đó và ghi lưu ý *Bỏ qua bảng công khác kỳ*. Với 5 file mẫu:
- **A:** `DL chấm công` và `DL chấm công (2)` là dữ liệu 01–30/06/2026, trong khi file là kỳ T09.
- **C:** `DL công (tính lương)` là dữ liệu 26/09–12/10/2025, không có NV chung với `Tuần`. Bảng kê và bảng lương tháng này đều tính từ `Tuần`; `Pivot` đang lấy dữ liệu `Tuần`.

Khi một doanh nghiệp đổi biểu mẫu, hoặc có doanh nghiệp mới: viết thêm hàm `nhan_dien_X` và `doc_mau_X` trong `dong_bo_luong.py`, rồi khai báo trong `ADAPTERS`. Bộ đọc đọc bảng công SK vào `da.cham_cong`; nếu có bảng công nhà máy riêng thì đọc vào `da.moi_bang_nm(ten_sheet)`.

**Kết quả đối chiếu với 5 file mẫu:**

| Hợp đồng | Nhà máy / SK | Công nhà máy / bảng kê |
|---|---|---|
| A | 1 bảng công | Khớp (850,5 giờ) |
| B | Tổng bằng nhau (14.171,5 giờ), **30 NV chia giờ ngày/đêm khác** | Khớp đúng `Tính phí` |
| C | 1 bảng công | 12 NV có công nhưng chưa có trong bảng kê |
| D | 1 bảng công | 64/66 NV khớp; YB20481 và YB20934 lệch giờ |
| E | Tổng bằng nhau (20.213 giờ), **14 NV chia loại giờ khác** | Không có bảng kê |

Ví dụ CTY B theo tab tháng: HC ngày 7.877 giờ ở công nhà máy nhưng 7.380 giờ ở công SK. Doanh thu và chi phí tính trên 2 con số này.

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

**Các quy tắc kiểm tra** (kiểm tra theo ngày chạy trên công SK):

| Nhóm | Lỗi | Nghi ngờ | Lưu ý |
|---|---|---|---|
| File và kỳ | Bảng chấm công trống. Không có giờ công (ô giờ là công thức chưa có giá trị). Ô lỗi công thức (`#REF!`, `#N/A`…), được tính là 0. Kỳ công dài hơn 31 ngày. Không xác định được kỳ. | Kỳ trong file khác kỳ ở tên file. Nhãn loại giờ lạ. Trùng mã hợp đồng (2 file cùng mã, tool tự đổi tên). | Hợp đồng khác kỳ với các hợp đồng còn lại. Nguồn chấm công đặc biệt (D). Bỏ qua bảng công khác kỳ. |
| Nhân viên | Trùng mã NV ở 2 khối của bảng công (giờ không bị cộng đôi). Thiếu họ tên. Ngày nghỉ trước ngày vào. | Mã khác dạng các mã còn lại (vd `HMT 140`). Cùng tên và cùng ngày vào nhưng khác mã. Cùng mã ở 2 hợp đồng. Ngày vào sau kỳ công. | NV không có giờ công. Tên chữ Hán. |
| Chấm công ngày | Công ghi vào ngày ngoài kỳ. Một ngày quá 24 giờ. Giờ âm. | Một ngày quá 16 giờ. Có công sau ngày nghỉ việc hoặc trước ngày vào. Giờ thường ghi vào Chủ nhật. Trùng dòng chấm công. | Làm thêm vượt 40 giờ/tháng. |
| Nhà máy / SK | | Lệch công nhà máy / SK (từng NV, từng loại giờ). Có công nhà máy nhưng không có công SK. Có công SK nhưng không có công nhà máy. | |
| Bảng kê | | Lệch giờ chấm công / bảng kê (công nhà máy, từng NV, từng loại giờ). Có công nhưng không có trong bảng kê. Có trong bảng kê nhưng không có chấm công. Lệch tổng giờ với bảng kê. | Không có bảng kê. |

**File không đọc được** (không đúng mẫu nào, hoặc đúng tên sheet nhưng cấu trúc đã đổi) không làm hỏng cả lượt: tool bỏ qua file đó, báo rõ lý do trên giao diện và liệt kê trong sheet `kiem_tra`.

**Ngưỡng:** sửa trong `NGUONG` ở đầu `dong_bo_luong.py`.
- Nghi ngờ khi một ngày quá 16 giờ; lỗi khi quá 24 giờ.
- Làm thêm tối đa 40 giờ/tháng, theo BLLĐ 2019 Điều 107.
- Lệch giờ trên 0,01 giờ thì tính là lệch.

## Lưu ý dữ liệu nguồn
- **Chưa có ở cả 5 file:** CCCD, tỉnh thành, số tài khoản trong bảng chấm công. Các cột này ở `DM` để trống.
- **Đơn giá trong DM của template bị đảo:** template ghi 37.500 / 55.328 cho "TC ngày 150%" và 45.000 / 63.840 cho "HC đêm 130%". Bảng kê CTY A lại dùng ngược lại (đúng hệ số 1,5 và 1,3). Tool lấy đơn giá theo bảng kê.
- **CTY B, NV HMT18:** ngày nghỉ (22/03/2026) trước ngày vào (29/05/2026) nhưng vẫn có công tháng 8.
- **CTY E:** bảng lương hiện tại trả theo `Công (lương)` (công nhà máy). Tab SK lấy theo `Công`, nên 14 NV có chi phí khác số đã trả.
