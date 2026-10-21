# -*- coding: utf-8 -*-
"""
Đồng bộ bảng chấm công / bảng lương của nhiều dự án (mỗi bên một biểu mẫu khác nhau)
về 1 file Excel chuẩn theo cấu trúc template (DM, bang_cham_cong_thang, bang_cham_cong_tuan, bang_luong, doanh_thu)
+ 2 sheet phụ: cham_cong_chi_tiet (dạng dọc, có ngày thật) và kiem_tra (đối chiếu).

Cách dùng (dòng lệnh):
    python dong_bo_luong.py <file1.xlsx> <file2.xlsx> ... -o KET_QUA.xlsx
    python dong_bo_luong.py --thu-muc "C:\\...\\template luong" -o KET_QUA.xlsx

Mỗi biểu mẫu được nhận diện tự động theo tên sheet (xem ADAPTERS ở cuối file).
Muốn thêm dự án mới có biểu mẫu khác: viết thêm 1 hàm doc_mau_X() và khai báo trong ADAPTERS.
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import os
import re
import sys
import warnings
import zipfile
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass, field

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# --------------------------------------------------------------------------------------
# Danh mục loại giờ chuẩn (giữ nguyên tên như template; 2 loại cuối là bổ sung)
# --------------------------------------------------------------------------------------
LOAI_GIO = [
    # code, tên trong template (giữ nguyên ký tự xuống dòng / 2 dấu cách như template)
    ("X", "HC Ngày 100%"),
    ("D", "Tăng ca ngày\n150%"),
    ("OT", "HC  đêm 130%"),
    ("OD", "Tăng ca đêm 200%"),
    ("OCN", "Tăng ca ngày chủ nhật"),
    ("OL", "Tăng ca ngày lễ"),
    ("L", "Nghỉ lễ có lương"),
    ("P", "Nghỉ phép có lương"),
    # bổ sung (template chưa có) – để cuối để 8 dòng đầu DM đúng vị trí template
    ("OCND", "Tăng ca đêm chủ nhật"),
    ("OLD", "Tăng ca đêm ngày lễ"),
]
# thứ tự hiển thị các dòng giờ phụ trong bảng chấm công tháng/tuần
THU_TU_GIO_PHU = ["OCN", "OCND", "OL", "OLD", "L", "P"]

# Thông tin chuyển khoản cuối bảng kê công nợ (theo template) – sửa tại đây nếu đổi tài khoản
THONG_TIN_CHUYEN_KHOAN = [
    "Thông tin chuyển khoản:",
    "Số tài khoản: 3338888798888",
    "Tên tài khoản: Công ty TNHH nguồn nhân lực SK Việt Nam",
    "Mở tại: Ngân hàng TMCP Quân Đội - CN Hải Dương",
]
TEN_LOAI = dict(LOAI_GIO)
# template bang_cham_cong_tuan ghi tên loại giờ trên 1 dòng, 1 dấu cách (vd 'Tăng ca ngày 150%', 'HC đêm 130%')
TEN_LOAI_TUAN = {k: " ".join(v.split()) for k, v in TEN_LOAI.items()}
LOAI_CHINH = ["X", "D", "OT", "OD"]           # 4 dòng luôn có trong bảng chấm công tháng/tuần (như template)
LOAI_TIEN = ["X", "D", "OT", "OD", "OCN", "OCND", "OL", "OLD"]
LOAI_TANG_CA = ["D", "OD", "OCN", "OCND", "OL", "OLD"]

# Mức cảnh báo
LOI, NGHI_NGO, LUU_Y = "Lỗi", "Nghi ngờ", "Lưu ý"
THU_TU_MUC = {LOI: 0, NGHI_NGO: 1, LUU_Y: 2}

# Ngưỡng kiểm tra (chỉnh tại đây nếu chính sách khác)
NGUONG = {
    "gio_ngay_nghi_ngo": 16,      # tổng giờ 1 ngày > 16h -> nghi ngờ
    "gio_ngay_loi": 24,           # > 24h/ngày -> lỗi chắc chắn
    "tang_ca_thang": 40,          # BLLĐ 2019 Điều 107: làm thêm không quá 40 giờ/tháng
    # Lương tối thiểu giờ theo NĐ 293/2025/NĐ-CP (từ 01/01/2026): vùng I 25.500đ, vùng IV 17.800đ
    "luong_gio_toi_thieu_vung_I": 25500,
    "luong_gio_toi_thieu_vung_IV": 17800,
    "sai_so_tien": 1,             # lệch tiền > 1đ (cộng 0,1%) coi là sai
}
LOI_CONG_THUC = ("#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!")


@dataclass
class DongLuong:
    ma_nv: str
    ten: str
    gio: dict = field(default_factory=dict)        # code -> giờ
    tien: dict = field(default_factory=dict)       # code -> tiền lương theo loại giờ
    pc_chuc_vu: float = 0
    pc_chuyen_can: float = 0
    pc_tham_nien: float = 0
    pc_khac: float = 0
    thuong: float = 0
    tra_lai_giu: float = 0       # khoản cộng: trả lại tiền giữ kỳ trước / hoàn cọc đồng phục...
    tru_do: float = 0
    tien_giu: float = 0          # giữ lại kỳ này (trừ)
    ung_luong: float = 0
    bhxh: float = 0
    khau_tru_khac: float = 0     # phạt, vi phạm báo trước...
    thuc_nhan_goc: float | None = None
    ghi_chu: str = ""
    ngay_vao: dt.date | None = None
    ngay_nghi: dt.date | None = None
    ngay_cong: float | None = None
    chi_tiet: str = ""           # diễn giải các khoản đã gộp (ứng từng đợt, trừ KTX/thẻ/ĐP...)
    gio_ngoai_cham_cong: dict = field(default_factory=dict)  # giờ được trả lương nhưng không có trong chấm công ngày (vd đào tạo)

    @property
    def tong_luong(self):
        return (sum(self.tien.get(c, 0) for c in LOAI_TIEN) + self.pc_chuc_vu + self.pc_chuyen_can
                + self.pc_tham_nien + self.pc_khac + self.thuong + self.tra_lai_giu)

    @property
    def thuc_nhan(self):
        return (self.tong_luong - self.tru_do - self.tien_giu - self.ung_luong
                - self.bhxh - self.khau_tru_khac)


@dataclass
class DuAn:
    ma: str                      # "CTY A"
    file: str
    mau: str                     # tên biểu mẫu nhận diện
    ky: str                      # "T08/2026"
    tu_ngay: dt.date | None = None
    den_ngay: dt.date | None = None
    nhan_vien: dict = field(default_factory=dict)      # ma_nv -> {ten, tinh, cccd, stk, ngan_hang, bo_phan}
    cham_cong: dict = field(default_factory=lambda: defaultdict(lambda: defaultdict(float)))  # (ma, date) -> {code: h}
    ky_hieu_ngay: dict = field(default_factory=dict)   # (ma, date) -> "N"/"TD"/...
    luong: list = field(default_factory=list)          # [DongLuong]
    don_gia_luong: dict = field(default_factory=dict)  # code -> đơn giá
    don_gia_dv: dict = field(default_factory=dict)     # code -> đơn giá dịch vụ
    ghi_chu_don_gia: str = ""
    doanh_thu: list | None = None    # [(code|nhãn, số lượng, đơn giá, thành tiền)] nếu lấy từ bảng kê gốc
    doanh_thu_goc: float | None = None
    ghi_chu_doanh_thu: str = ""
    van_de: list = field(default_factory=list)         # [(muc, ma_nv, loai, noi_dung)]
    don_gia_luong_khac: list = field(default_factory=list)  # các bộ đơn giá lương khác (vd C có bộ cũ/mới)
    tong_goc: dict = field(default_factory=dict)       # số ở dòng TỔNG của bảng lương gốc
    o_loi: list = field(default_factory=list)          # [(ma_nv, sheet, ô, giá trị lỗi)]
    ngoai_bang_luong: set = field(default_factory=set) # mã có chấm công nhưng không có trong bảng lương

    def bao(self, muc, ma_nv, loai, noi_dung):
        """Ghi 1 cảnh báo. muc: LOI (sai chắc chắn) / NGHI_NGO (cần kiểm tra) / LUU_Y (thông tin)."""
        self.van_de.append((muc, ma_nv, loai, noi_dung))


# --------------------------------------------------------------------------------------
# Hàm tiện ích
# --------------------------------------------------------------------------------------
def so(v) -> float:
    """Chuyển giá trị ô về số; text '0,5' / '12,35' / '1.000' được xử lý; còn lại -> 0."""
    if v is None or isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("\xa0", "").replace(" ", "")
    s = re.sub(r"(?i)(vnđ|vnd|đồng|đ|giờ|h)$", "", s)  # đơn vị dính sau số: 2.500.000đ, 8h
    if not s:
        return 0.0
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", s):     # 1.000.000 / 1.000,5 (kiểu Việt)
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", s):   # 1,000,000 / 1,000.5 (kiểu Anh)
        s = s.replace(",", "")
    else:
        s = s.replace(",", ".")                              # 7,5 giờ
    try:
        return float(s)
    except ValueError:
        return 0.0


def ma(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def ten_sach(v) -> str:
    return re.sub(r"\s+", " ", str(v or "")).strip()


def chuan(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def ci(col: str) -> int:
    return column_index_from_string(col) - 1


def o(row, col: str):
    i = ci(col)
    return row[i] if i < len(row) else None


def la_ngay(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        try:
            m = re.match(r"\s*(\d{4})[/-](\d{1,2})[/-](\d{1,2})", v)
            if m:
                return dt.date(int(m[1]), int(m[2]), int(m[3]))
            m = re.match(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})", v)
            if m:
                return dt.date(int(m[3]), int(m[2]), int(m[1]))
        except ValueError:          # ngày không tồn tại, vd 31/02
            return None
    return None


def co_chu_han(s: str) -> bool:
    return bool(re.search(r"[\u4e00-\u9fff]", s or ""))


def doc_sheet(wb, ten, min_row=1, max_row=None, dung_khi_trong=60):
    """Đọc các dòng (values) của sheet; tự dừng khi gặp chuỗi dài dòng trống (sheet không có dimension)."""
    ws = wb[ten]
    rows, trong = [], 0
    for r in ws.iter_rows(min_row=min_row, max_row=max_row, values_only=True):
        rows.append(r)
        if all(v in (None, "") for v in r):
            trong += 1
            if trong >= dung_khi_trong:
                break
        else:
            trong = 0
    while rows and all(v in (None, "") for v in rows[-1]):
        rows.pop()
    return rows


def dong_trong(row, den_cot) -> bool:
    """Dòng không có giá trị nào từ cột A đến den_cot (dòng trống / dòng ngăn cách)."""
    return all(v is None or (isinstance(v, str) and not v.strip()) for v in row[:ci(den_cot) + 1])


SO_DONG_TRONG_TOI_DA = 3   # bảng lương: bỏ qua tối đa 2 dòng trống liền nhau, gặp 3 dòng trống thì coi là hết bảng


def kiem_con_sot(da, rows, tu_vi_tri, cot_ma, sheet, so_dong_dau):
    """Sau chỗ dừng đọc bảng lương, còn mã có dạng giống mã NV đã đọc -> có thể bị bỏ sót NV."""
    if not da.luong:
        return
    mau = Counter(_mau_ma(d.ma_nv) for d in da.luong).most_common(1)[0][0]
    da_doc = {d.ma_nv for d in da.luong}
    sot = []
    for i, r in enumerate(rows[tu_vi_tri:tu_vi_tri + 300], start=so_dong_dau):
        m = ma(o(r, cot_ma))
        if m and m not in da_doc and _mau_ma(m) == mau:
            sot.append(f"{m} (dòng {i})")
    if sot:
        da.bao(LOI, "", "Có thể sót NV trong bảng lương",
               f"Tool dừng đọc '{sheet}' trước dòng {so_dong_dau}, nhưng bên dưới còn {len(sot)} mã giống mã NV: "
               + _rut_gon(sot, 10) + ". Kiểm tra dòng trống/dòng tổng phụ chen giữa bảng.")


def kiem_o_loi(da, ma_nv, sheet, so_dong, row, den_cot):
    """Ghi lại các ô có lỗi công thức (#REF!, #N/A...) trong 1 dòng, từ cột A đến den_cot."""
    for j, v in enumerate(row[:ci(den_cot) + 1]):
        if isinstance(v, str) and v.strip().upper() in LOI_CONG_THUC:
            da.o_loi.append((ma_nv, sheet, f"{get_column_letter(j + 1)}{so_dong}", v.strip()))


def chi_tiet(*cap):
    """('Nhãn', giá trị[, 'đơn vị']) -> 'Nhãn: 1,000; ...' (bỏ giá trị 0)."""
    ds = []
    for c in cap:
        nhan, v = c[0], c[1]
        if isinstance(v, str):
            if v.strip():
                ds.append(f"{nhan}: {v.strip()}")
            continue
        if v:
            ds.append(f"{nhan}: {v:,.0f}" if len(c) == 2 else f"{nhan}: {v:g} {c[2]}")
    return "; ".join(ds)


def doc_UL(wb, da):
    """Sheet UL (C, D): STK / ngân hàng / chủ TK + các đợt ứng. Trả về {mã: [(nhãn, tiền)]}."""
    if "UL" not in wb.sheetnames:
        return {}
    rows = doc_sheet(wb, "UL", max_row=600)
    hang = next((i for i, r in enumerate(rows[:10]) if chuan(o(r, "B")) in ("mã id", "mã nv")), None)
    if hang is None:
        return {}
    td = rows[hang]
    cot_ung = [(j, ten_sach(v)) for j, v in enumerate(td) if chuan(v).startswith("ứng đợt")]
    kq = {}
    for r in rows[hang + 1:]:
        m = ma(o(r, "B"))
        if not m:
            continue
        nv = da.nhan_vien.setdefault(m, {})
        if o(r, "D") not in (None, ""):
            nv["stk"] = ma(o(r, "D"))
        if ten_sach(o(r, "E")):
            nv["ngan_hang"] = ten_sach(o(r, "E"))
        if ten_sach(o(r, "F")):
            nv["chu_tk"] = ten_sach(o(r, "F"))
        kq[m] = [(t, so(r[j])) for j, t in cot_ung if j < len(r)]
    return kq


def ky_tu_ten_file(path):
    m = re.search(r"T(\d{1,2})[._-](\d{4})", os.path.basename(path))
    return (int(m[1]), int(m[2])) if m else None


def ma_du_an_tu_file(path, mac_dinh=None):
    m = re.search(r"CTY[_ ]?([A-Za-z0-9]+)", os.path.basename(path))
    if m:
        return f"CTY {m[1].upper()}"
    # không có 'CTY...' trong tên file: dùng chính tên file để 2 file khác nhau không bị gộp nhầm 1 mã
    return os.path.splitext(os.path.basename(path))[0][:40]


def nap_luong_kieu_khoi(da: DuAn, rows, cot_ma, cot_nhan, cot_ngay: dict, map_nhan, bat_dau=0,
                        cot_ten=None):
    """Bảng công dạng khối: mỗi NV nhiều dòng (mỗi dòng 1 loại giờ), mã NV chỉ ở dòng đầu khối."""
    ma_hien_tai = None
    for r in rows[bat_dau:]:
        m = ma(o(r, cot_ma))
        nhan = chuan(o(r, cot_nhan))
        if m:
            ma_hien_tai = m
            if cot_ten and ten_sach(o(r, cot_ten)):
                da.nhan_vien.setdefault(m, {}).setdefault("ten", ten_sach(o(r, cot_ten)))
        if not nhan or ma_hien_tai is None:
            continue
        code = map_nhan.get(nhan)
        if code is None:
            da.bao(NGHI_NGO, ma_hien_tai, "Nhãn giờ lạ", f"Không nhận ra loại giờ '{o(r, cot_nhan)}' - bỏ qua dòng này")
            continue
        for col, ngay in cot_ngay.items():
            v = o(r, col)
            if isinstance(v, str) and v.strip().upper() in LOI_CONG_THUC:
                da.o_loi.append((ma_hien_tai, "bảng chấm công", f"{col} ({ngay:%d/%m})", v.strip()))
                continue
            if isinstance(v, str) and v.strip() and so(v) == 0 and not re.fullmatch(r"[\d.,\s]+", v):
                da.ky_hieu_ngay[(ma_hien_tai, ngay)] = v.strip()
                continue
            h = so(v)
            if h:
                da.cham_cong[(ma_hien_tai, ngay)][code] += h


# --------------------------------------------------------------------------------------
# MẪU A: sheet 'BCC LĐ THUÊ LẠI mm.yyyy' (4 dòng/NV) + 'Bảng chi lương' + 'Bảng phí'
# --------------------------------------------------------------------------------------
def nhan_dien_A(sheets):
    return "Bảng chi lương" in sheets and any(s.upper().startswith("BCC") for s in sheets)


def doc_mau_A(path, wb) -> DuAn:
    da = DuAn(ma=ma_du_an_tu_file(path, "Dự án A"), file=path, mau="Mẫu A (BCC thuê lại + Bảng chi lương)", ky="")
    bcc = next(s for s in wb.sheetnames if s.upper().startswith("BCC"))
    rows = doc_sheet(wb, bcc)
    # dòng tiêu đề ngày: tìm dòng có nhiều datetime nhất trong 15 dòng đầu
    hang_ngay = max(range(min(15, len(rows))), key=lambda i: sum(isinstance(v, dt.datetime) for v in rows[i]))
    cot_ngay = {get_column_letter(j + 1): v.date() for j, v in enumerate(rows[hang_ngay]) if isinstance(v, dt.datetime)}
    map_nhan = {"hc ngày 100%": "X", "tăng ca ngày 150%": "D", "hc đêm 130%": "OT", "tăng ca đêm 200%": "OD",
                "tăng ca ngày nghỉ": "OCN", "tăng ca ngày lễ": "OL"}
    du_lieu = [r for r in rows[hang_ngay + 1:] if chuan(o(r, "A")) not in ("tổng", "tổng cộng")]
    nap_luong_kieu_khoi(da, du_lieu, "B", "I", cot_ngay, map_nhan, cot_ten="C")
    for r in du_lieu:
        m = ma(o(r, "B"))
        if m and (ten_sach(o(r, "D")) or ten_sach(o(r, "E"))):
            nv = da.nhan_vien.setdefault(m, {})
            nv.setdefault("bo_phan", ten_sach(o(r, "D")))
            nv.setdefault("chuc_danh", ten_sach(o(r, "E")))
            nv.setdefault("ncc", ten_sach(o(r, "F")))
    da.tu_ngay, da.den_ngay = min(cot_ngay.values()), max(cot_ngay.values())

    # Bảng chi lương: đơn giá dòng 3, tiêu đề dòng 4-5, dữ liệu từ dòng 6
    rows = doc_sheet(wb, "Bảng chi lương")
    dg = rows[2]
    for code, col in zip(LOAI_TIEN, ["Q", "S", "R", "T", "U", "V", "W", "X"]):
        da.don_gia_luong[code] = so(o(dg, col))
    trong = 0
    for i, r in enumerate(rows[5:], start=6):
        m = ma(o(r, "B"))
        if not m:
            if dong_trong(r, "AF") and trong < SO_DONG_TRONG_TOI_DA - 1:
                trong += 1
                continue
            if not dong_trong(r, "AF"):
                da.tong_goc = {"thuc_nhan": o(r, "AF"), "tong_luong": o(r, "AC"), "dong": i}
            kiem_con_sot(da, rows, i, "B", "Bảng chi lương", i + 1)
            break
        trong = 0
        kiem_o_loi(da, m, "Bảng chi lương", i, r, "AF")
        d = DongLuong(ma_nv=m, ten=ten_sach(o(r, "C")))
        d.gio = {"X": so(o(r, "F")), "OT": so(o(r, "H")), "D": so(o(r, "J")), "OD": so(o(r, "K")),
                 "OCN": so(o(r, "L")), "OCND": so(o(r, "M")), "OL": so(o(r, "N")), "OLD": so(o(r, "O"))}
        d.tien = {"X": so(o(r, "Q")), "OT": so(o(r, "R")), "D": so(o(r, "S")), "OD": so(o(r, "T")),
                  "OCN": so(o(r, "U")), "OCND": so(o(r, "V")), "OL": so(o(r, "W")), "OLD": so(o(r, "X"))}
        d.pc_chuyen_can = so(o(r, "Z"))
        d.tra_lai_giu = so(o(r, "AA"))      # "Tiền giữ tháng trước" được CỘNG vào tổng lương
        d.tru_do = so(o(r, "AB"))
        d.ung_luong = so(o(r, "AD"))
        d.tien_giu = so(o(r, "AE"))         # "Giữ lương" kỳ này
        d.thuc_nhan_goc = so(o(r, "AF"))
        d.ghi_chu = ten_sach(o(r, "AG"))
        d.ngay_vao, d.ngay_nghi, d.ngay_cong = la_ngay(o(r, "D")), la_ngay(o(r, "E")), so(o(r, "P"))
        d.chi_tiet = chi_tiet(("Trả lại tiền giữ tháng trước", d.tra_lai_giu), ("Giữ lương", d.tien_giu))
        da.luong.append(d)

    # Bảng phí: đơn giá dịch vụ dòng 3 cùng bố cục cột
    if "Bảng phí" in wb.sheetnames:
        r3 = doc_sheet(wb, "Bảng phí", max_row=3)[2]
        for code, col in zip(LOAI_TIEN, ["Q", "S", "R", "T", "U", "V", "W", "X"]):
            da.don_gia_dv[code] = so(o(r3, col))
    if "Bảng kê" in wb.sheetnames:
        for r in doc_sheet(wb, "Bảng kê"):
            if any("tổng" in chuan(v) and "thanh toán" in chuan(v) for v in r if isinstance(v, str)):
                vals = [so(v) for v in r if isinstance(v, (int, float))]
                if vals:
                    da.doanh_thu_goc = vals[-1]
    da.ghi_chu_doanh_thu = "Doanh thu = giờ bảng lương x đơn giá sheet 'Bảng phí'"
    return da


# --------------------------------------------------------------------------------------
# MẪU B: 'Bảng công SK' (6 dòng/NV) + 'Bảng lương' + 'Tính phí' + 'DSCN'
# --------------------------------------------------------------------------------------
def nhan_dien_B(sheets):
    return "Bảng công SK" in sheets and "Bảng lương" in sheets


def doc_mau_B(path, wb) -> DuAn:
    da = DuAn(ma=ma_du_an_tu_file(path, "Dự án B"), file=path, mau="Mẫu B (Bảng công SK + Bảng lương + Tính phí)", ky="")
    rows = doc_sheet(wb, "Bảng công SK")
    hang_ngay = max(range(min(10, len(rows))), key=lambda i: sum(isinstance(v, dt.datetime) for v in rows[i]))
    cot_ngay = {get_column_letter(j + 1): v.date() for j, v in enumerate(rows[hang_ngay])
                if isinstance(v, dt.datetime) and j >= ci("I")}
    map_nhan = {"ca ngày": "X", "ca đêm": "OT", "tăng ca ngày": "D", "tăng ca đêm": "OD",
                "chủ nhật": "OCN", "ca đêm chủ nhật": "OCND"}
    nap_luong_kieu_khoi(da, rows[hang_ngay + 2:], "C", "H", cot_ngay, map_nhan, cot_ten="D")
    da.tu_ngay, da.den_ngay = min(cot_ngay.values()), max(cot_ngay.values())

    rows = doc_sheet(wb, "Bảng lương")
    dg = rows[4]
    for code, col in zip(["X", "OT", "D", "OD", "OCN", "OCND"], ["O", "P", "Q", "R", "S", "T"]):
        da.don_gia_luong[code] = so(o(dg, col))
    trong = 0
    for i, r in enumerate(rows[6:], start=7):
        m = ma(o(r, "C"))
        if not m and chuan(o(r, "A")) != "tổng" and dong_trong(r, "AA") and trong < SO_DONG_TRONG_TOI_DA - 1:
            trong += 1
            continue
        if not m or chuan(o(r, "A")) == "tổng":
            if chuan(o(r, "A")) == "tổng":
                da.tong_goc = {"thuc_nhan": o(r, "AA"), "tong_luong": o(r, "V"), "dong": i}
            kiem_con_sot(da, rows, i, "C", "Bảng lương", i + 1)
            break
        trong = 0
        kiem_o_loi(da, m, "Bảng lương", i, r, "AA")
        d = DongLuong(ma_nv=m, ten=ten_sach(o(r, "D")))
        d.gio = {"X": so(o(r, "H")), "OT": so(o(r, "I")), "D": so(o(r, "J")), "OD": so(o(r, "K")),
                 "OCN": so(o(r, "L")), "OCND": so(o(r, "M"))}
        d.tien = {"X": so(o(r, "O")), "OT": so(o(r, "P")), "D": so(o(r, "Q")), "OD": so(o(r, "R")),
                  "OCN": so(o(r, "S")), "OCND": so(o(r, "T"))}
        d.pc_chuyen_can = so(o(r, "U"))
        d.ung_luong = sum(so(o(r, c)) for c in ["W", "X", "Y", "Z"])
        d.thuc_nhan_goc = so(o(r, "AA"))
        d.ghi_chu = ten_sach(o(r, "AB"))
        d.ngay_vao, d.ngay_nghi, d.ngay_cong = la_ngay(o(r, "E")), la_ngay(o(r, "F")), so(o(r, "N"))
        d.chi_tiet = chi_tiet(*[(ten_sach(o(rows[5], c)), so(o(r, c))) for c in ["W", "X", "Y", "Z"]],
                              ("Chuyền sản xuất", ten_sach(o(r, "G"))))
        if ma(o(r, "B")):
            da.nhan_vien.setdefault(m, {})["ma_the"] = ma(o(r, "B"))
        da.luong.append(d)

    if "Tính phí" in wb.sheetnames:
        rows = doc_sheet(wb, "Tính phí")
        r6 = rows[5]
        cap = [("X", "H", "I"), ("OT", "J", "K"), ("D", "L", "M"), ("OD", "N", "O"), ("OCN", "P", "Q"), ("OCND", "R", "S")]
        for code, _, cg in cap:
            da.don_gia_dv[code] = so(o(r6, cg))
        tong = next((r for r in rows if chuan(o(r, "A")) == "tổng"), None)
        if tong:
            da.doanh_thu = [(code, so(o(tong, ch)), da.don_gia_dv[code], so(o(tong, ct))) for code, ch, ct in cap]
            da.doanh_thu_goc = so(o(tong, "U"))
            da.ghi_chu_doanh_thu = ("Lấy nguyên từ sheet 'Tính phí' (giờ theo ' Bảng công nhà máy', "
                                    "khác giờ trả lương theo 'Bảng công SK')")
    return da


# --------------------------------------------------------------------------------------
# MẪU C: 'Tuần' (1 dòng/NV/ngày, giờ quy đổi từ mã ca 'DM ca') + 'Bảng lương' + 'Bảng kê Txx'
# --------------------------------------------------------------------------------------
def nhan_dien_C(sheets):
    return "Tuần" in sheets and "DM ca" in sheets and "Bảng lương" in sheets


def doc_mau_C(path, wb) -> DuAn:
    da = DuAn(ma=ma_du_an_tu_file(path, "Dự án C"), file=path, mau="Mẫu C (Tuần + DM ca + Bảng lương + Bảng kê)", ky="")
    rows = doc_sheet(wb, "Tuần")
    hang_td = next(i for i, r in enumerate(rows[:10]) if chuan(o(r, "D")).startswith("mã nhân viên"))
    cot_gio = [("P", "X"), ("Q", "OT"), ("R", "D"), ("S", "OD"), ("T", "OCN"), ("U", "OCND"), ("V", "OL"), ("W", "OLD")]
    for r in rows[hang_td + 1:]:
        m, ngay = ma(o(r, "D")), la_ngay(o(r, "A"))
        if not m or not ngay:
            continue
        da.nhan_vien.setdefault(m, {}).setdefault("ten", ten_sach(o(r, "E")))
        bp = ten_sach(o(r, "AK")) or ten_sach(o(r, "AJ"))     # tên bộ phận, nếu trống thì lấy mã bộ phận
        if bp:
            da.nhan_vien[m].setdefault("bo_phan", bp)
        for col, _ in cot_gio:
            v = o(r, col)
            if isinstance(v, str) and v.strip().upper() in LOI_CONG_THUC:
                da.o_loi.append((m, "Tuần", f"{col} ({ngay:%d/%m})", v.strip()))
        for col, code in cot_gio:
            h = so(o(r, col))
            if h:
                da.cham_cong[(m, ngay)][code] += h
        kh = ten_sach(o(r, "AG"))
        if kh:
            da.ky_hieu_ngay[(m, ngay)] = kh
        da.tu_ngay = min(da.tu_ngay or ngay, ngay)
        da.den_ngay = max(da.den_ngay or ngay, ngay)

    rows = doc_sheet(wb, "Bảng lương")
    hang_td = next(i for i, r in enumerate(rows[:10]) if chuan(o(r, "B")) == "mã nv")
    for code, col in zip(["X", "OT", "D", "OD", "OCN", "OCND", "OL", "OLD"], "QRSTUVWX"):
        da.don_gia_luong[code] = so(o(rows[3], col)) or so(o(rows[1], col))
    da.don_gia_luong_khac = [{code: so(o(rows[1], col)) for code, col in
                              zip(["X", "OT", "D", "OD", "OCN", "OCND", "OL", "OLD"], "QRSTUVWX")}]
    da.ghi_chu_don_gia = "Bảng lương dùng 2 bộ đơn giá (cũ dòng 2, mới dòng 4) tùy NV; đây là bộ mới"
    trong = 0
    for i, r in enumerate(rows[hang_td + 1:], start=hang_td + 2):
        m = ma(o(r, "B"))
        if not m:
            if dong_trong(r, "AL") and trong < SO_DONG_TRONG_TOI_DA - 1:
                trong += 1
                continue
            if not dong_trong(r, "AL"):
                da.tong_goc = {"thuc_nhan": o(r, "AL"), "dong": i}
            kiem_con_sot(da, rows, i - 1, "B", "Bảng lương", i + 1)
            break
        trong = 0
        kiem_o_loi(da, m, "Bảng lương", i, r, "AL")
        d = DongLuong(ma_nv=m, ten=ten_sach(o(r, "C")))
        d.gio = {"X": so(o(r, "F")) + so(o(r, "G")), "OT": so(o(r, "H")), "D": so(o(r, "I")), "OD": so(o(r, "J")),
                 "OCN": so(o(r, "K")), "OCND": so(o(r, "L")), "OL": so(o(r, "M")), "OLD": so(o(r, "N"))}
        d.tien = {"X": so(o(r, "Q")), "OT": so(o(r, "R")), "D": so(o(r, "S")), "OD": so(o(r, "T")),
                  "OCN": so(o(r, "U")), "OCND": so(o(r, "V")), "OL": so(o(r, "W")), "OLD": so(o(r, "X"))}
        if so(o(r, "F")):
            d.gio_ngoai_cham_cong["X"] = so(o(r, "F"))
            d.ghi_chu = f"Gồm {so(o(r, 'F')):g} giờ đào tạo (tính vào HC ngày)"
        d.pc_khac = so(o(r, "Y")) + so(o(r, "Z"))          # lương pending + phụ cấp ca đêm
        d.pc_chuyen_can = so(o(r, "AA"))
        d.tra_lai_giu = so(o(r, "AB")) + so(o(r, "AI"))     # hoàn tiền đồng phục + giữ lương T07 (+/-)
        d.thuong = so(o(r, "AC"))
        d.tru_do = sum(so(o(r, c)) for c in ["AD", "AE", "AF", "AG"])   # KTX, thẻ, đồng phục
        d.ung_luong = so(o(r, "AJ"))
        d.thuc_nhan_goc = so(o(r, "AL"))
        # khoản giữ lại kỳ này = (Tổng lương + giữ T07 - ứng) - thực nhận (bao gồm cả NV PENDING giữ toàn bộ)
        d.tien_giu = round(so(o(r, "AH")) + so(o(r, "AI")) - so(o(r, "AJ")) - d.thuc_nhan_goc, 2)
        if ten_sach(o(r, "AN")).upper() == "PENDING":
            d.ghi_chu = (d.ghi_chu + "; PENDING").strip("; ")
        elif abs(d.tien_giu) > 1:
            da.bao(NGHI_NGO, m, "Tiền giữ suy ra",
                   f"Mẫu C không có cột tiền giữ; tool suy ra {d.tien_giu:,.0f} = Tổng lương + giữ T07 − ứng − Thực nhận (AL). "
                   "NV không PENDING mà có khoản này: kiểm tra ô Thực nhận có bị gõ đè không.")
        if ten_sach(o(r, "AM")):
            d.ghi_chu = (d.ghi_chu + "; " + ten_sach(o(r, "AM"))).strip("; ")
        d.ngay_vao, d.ngay_nghi, d.ngay_cong = la_ngay(o(r, "D")), la_ngay(o(r, "E")), so(o(r, "O"))
        d.chi_tiet = chi_tiet(("Lương pending", so(o(r, "Y"))), ("PC ca đêm", so(o(r, "Z"))),
                              ("Hoàn tiền đồng phục SK giữ", so(o(r, "AB"))), ("Giữ lương T07 (+/-)", so(o(r, "AI"))),
                              ("Trừ KTX", so(o(r, "AD"))), ("Trừ thẻ", so(o(r, "AE"))),
                              ("Trừ đồng phục nhà máy", so(o(r, "AF"))), ("Giữ đồng phục", so(o(r, "AG"))),
                              ("Giữ lương kỳ này", d.tien_giu), ("Tổng công T7", so(o(r, "P")), "công"))
        da.luong.append(d)
    ung = doc_UL(wb, da)
    for d in da.luong:
        if ung.get(d.ma_nv):
            d.chi_tiet = "; ".join(x for x in [d.chi_tiet, chi_tiet(*ung[d.ma_nv])] if x)

    bk = next((s for s in wb.sheetnames if s.lower().startswith("bảng kê")), None)
    if bk:
        rows = doc_sheet(wb, bk)
        cap = [("X", "F", "G"), ("OT", "H", "I"), ("D", "J", "K"), ("OD", "L", "M"),
               ("OCN", "N", "O"), ("OCND", "P", "Q"), ("OL", "R", "S"), ("OLD", "T", "U")]
        sl, tt, kt, dg_dem = defaultdict(float), defaultdict(float), 0.0, defaultdict(lambda: defaultdict(int))
        for r in rows[1:]:
            if not ma(o(r, "B")):
                so_trong_dong = [so(v) for v in r if isinstance(v, (int, float))]
                if (da.doanh_thu_goc is None and so_trong_dong
                        and any("tổng" in chuan(v) for v in r if isinstance(v, str))):
                    da.doanh_thu_goc = max(so_trong_dong)
                continue
            for code, ch, cg in cap:
                h, g = so(o(r, ch)), so(o(r, cg))
                sl[code] += h
                tt[code] += h * g
                dg_dem[code][g] += 1
            kt += so(o(r, "V"))
        for code, _, _ in cap:
            da.don_gia_dv[code] = max(dg_dem[code], key=dg_dem[code].get) if dg_dem[code] else 0
        da.doanh_thu = [(code, sl[code], round(tt[code] / sl[code]) if sl[code] else da.don_gia_dv[code], tt[code])
                        for code, _, _ in cap] + [("Khấu trừ", None, None, -kt)]
        da.ghi_chu_doanh_thu = (f"Lấy từ sheet '{bk}' (đơn giá theo từng NV, ở đây là bình quân; "
                                "bảng kê không gồm NV pending/0 giờ)")
    return da


# --------------------------------------------------------------------------------------
# MẪU D: 'Chi lương' + 'UL' + '明细'; chấm công ngày chỉ còn trong cache liên kết ngoài 'DS TỔNG'
# --------------------------------------------------------------------------------------
def nhan_dien_D(sheets):
    return "Chi lương" in sheets and any("明细" in s for s in sheets)


def _doc_cache_lien_ket(path, ten_sheet="DS TỔNG"):
    """Trả về list[(tên part, {row: {col: value}})] cho các external link có sheet ten_sheet."""
    kq = []
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if not re.fullmatch(r"xl/externalLinks/externalLink\d+\.xml", n):
                continue
            x = z.read(n).decode("utf8")
            names = re.findall(r'<sheetName val="([^"]*)"', x)
            if ten_sheet not in names:
                continue
            sid = names.index(ten_sheet)
            mblock = re.search(rf'<sheetData sheetId="{sid}"[^>]*>(.*?)</sheetData>', x, re.S)
            if not mblock:
                continue
            data = defaultdict(dict)
            for cm in re.finditer(r'<cell r="([A-Z]+)(\d+)"[^>]*>(?:<v>(.*?)</v>)?</cell>', mblock.group(1), re.S):
                col, row, v = cm.group(1), int(cm.group(2)), cm.group(3)
                if v is not None:
                    v = (v.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
                         .replace("&quot;", '"').replace("&apos;", "'"))
                    data[row][col] = v
            kq.append((n, data))
    return kq


def doc_mau_D(path, wb) -> DuAn:
    da = DuAn(ma=ma_du_an_tu_file(path, "Dự án D"), file=path, mau="Mẫu D (Chi lương + UL + 明细; công ngày từ cache 'DS TỔNG')", ky="")
    # ---- Bảng lương
    rows = doc_sheet(wb, "Chi lương", max_row=400)
    hang_td = next(i for i, r in enumerate(rows[:15]) if chuan(o(r, "B")) == "mã nv")
    dg = rows[hang_td - 1]
    for code, col in zip(["X", "OT", "D", "OD", "OCN", "OCND"], "NOPQRS"):
        da.don_gia_luong[code] = so(o(dg, col))
    trong = 0
    for i, r in enumerate(rows[hang_td + 1:], start=hang_td + 2):
        m = ma(o(r, "B"))
        if not m and chuan(o(r, "C")) != "tổng" and dong_trong(r, "AC") and trong < SO_DONG_TRONG_TOI_DA - 1:
            trong += 1
            continue
        if not m or chuan(o(r, "C")) == "tổng":
            if chuan(o(r, "C")) == "tổng":
                da.tong_goc = {"thuc_nhan": o(r, "AC"), "tong_luong": o(r, "AA"), "dong": i}
            kiem_con_sot(da, rows, i - 1, "B", "Chi lương", i + 1)
            break
        trong = 0
        kiem_o_loi(da, m, "Chi lương", i, r, "AC")
        d = DongLuong(ma_nv=m, ten=ten_sach(o(r, "C")))
        d.gio = {"X": so(o(r, "F")), "OT": so(o(r, "G")), "D": so(o(r, "H")), "OD": so(o(r, "I")),
                 "OCN": so(o(r, "J")), "OCND": so(o(r, "K"))}
        d.tien = {"X": so(o(r, "N")), "OT": so(o(r, "O")), "D": so(o(r, "P")), "OD": so(o(r, "Q")),
                  "OCN": so(o(r, "R")), "OCND": so(o(r, "S"))}
        d.pc_khac = so(o(r, "T")) + so(o(r, "U"))        # phụ cấp bữa ăn + trợ cấp công đoạn
        d.thuong = so(o(r, "V"))
        d.pc_chuyen_can = so(o(r, "W"))
        d.bhxh = so(o(r, "X"))
        d.khau_tru_khac = so(o(r, "Y"))                  # phạt nghỉ tự do
        d.tru_do = so(o(r, "Z"))
        d.ung_luong = so(o(r, "AB"))
        d.thuc_nhan_goc = so(o(r, "AC"))
        d.ghi_chu = ten_sach(o(r, "AD"))
        d.ngay_vao, d.ngay_nghi, d.ngay_cong = la_ngay(o(r, "D")), la_ngay(o(r, "E")), so(o(r, "M"))
        d.chi_tiet = chi_tiet(("Số suất ăn", so(o(r, "L")), "suất"), ("PC bữa ăn", so(o(r, "T"))),
                              ("Trợ cấp công đoạn", so(o(r, "U"))), ("Phạt nghỉ tự do", d.khau_tru_khac))
        da.luong.append(d)
    # ---- UL: tài khoản ngân hàng + tên tiếng Việt + ứng từng đợt
    ung = doc_UL(wb, da)
    for d in da.luong:
        if ung.get(d.ma_nv):
            d.chi_tiet = "; ".join(x for x in [d.chi_tiet, chi_tiet(*ung[d.ma_nv])] if x)
        if co_chu_han(d.ten) and da.nhan_vien.get(d.ma_nv, {}).get("chu_tk"):
            d.ghi_chu = (d.ghi_chu + f"; tên gốc: {d.ten}").strip("; ")
            d.ten = da.nhan_vien[d.ma_nv]["chu_tk"].title()
    # ---- đơn giá phía khách hàng (明细) – tham khảo
    mx = next((s for s in wb.sheetnames if "明细" in s and "(" not in s), None) or \
        next(s for s in wb.sheetnames if "明细" in s)
    rows = doc_sheet(wb, mx, max_row=6)
    if len(rows) >= 5:
        for code, col in zip(LOAI_TIEN[:4] + ["OCN", "OCND", "OL", "OLD"],
                             ["AH", "AL", "AJ", "AN", "AP", "AR", "AT", "AV"]):
            da.don_gia_dv[code] = so(o(rows[4], col))
    da.ghi_chu_doanh_thu = ("Ước tính = giờ bảng lương x đơn giá ở sheet '" + mx +
                            "' (file không có bảng kê doanh thu)")
    # ---- chấm công ngày từ cache liên kết ngoài
    th = ky_tu_ten_file(path)
    luong_gio = {d.ma_nv: d.gio["X"] for d in da.luong}
    ung_vien = _doc_cache_lien_ket(path)
    tot, diem_tot = None, -1
    for ten_part, data in ung_vien:
        khoi = {}
        for row in sorted(data):
            mm = data[row].get("D", "").strip()
            if mm and re.fullmatch(r"[A-Z]{1,4}\d{3,}", mm):
                khoi[mm] = row
        diem = sum(1 for k, r0 in khoi.items() if k in luong_gio and
                   abs(sum(so(data[r0].get(get_column_letter(ci("N") + 1 + i))) for i in range(31)) - luong_gio[k]) < 0.01)
        if diem > diem_tot:
            tot, diem_tot = (ten_part, data, khoi), diem
    if not tot:
        da.bao(LOI, "", "Thiếu chấm công ngày", "Không tìm thấy dữ liệu 'DS TỔNG' trong file")
        return da
    _, data, khoi = tot
    if th:
        nam, thang = th[1], th[0]
    else:
        nam, thang = dt.date.today().year, dt.date.today().month
        da.bao(LOI, "", "Không xác định được kỳ", "Tên file không có kỳ dạng 'Txx.yyyy' (vd T08.2026); mẫu D cần kỳ "
               f"trong tên file để gán ngày cho chấm công – đang tạm dùng tháng hiện tại {thang:02d}/{nam}. Đổi tên file rồi chạy lại.")
    da.tu_ngay = dt.date(nam, thang, 1)
    da.den_ngay = dt.date(nam + thang // 12, thang % 12 + 1, 1) - dt.timedelta(days=1)
    offset_loai = {0: "X", 1: "OT", 2: "D", 3: "OD"}
    for mm, r0 in khoi.items():
        if data[r0].get("E"):
            da.nhan_vien.setdefault(mm, {}).setdefault("ten_cc", data[r0]["E"])
        for i in range(31):
            try:
                ngay = dt.date(nam, thang, i + 1)
            except ValueError:
                continue
            col = get_column_letter(ci("N") + 1 + i)
            for off, code in offset_loai.items():
                v = data.get(r0 + off, {}).get(col)
                if v is None:
                    continue
                if re.fullmatch(r"-?[\d.]+", v.strip()):
                    if float(v):
                        da.cham_cong[(mm, ngay)][code] += float(v)
                elif v.strip():
                    da.ky_hieu_ngay[(mm, ngay)] = v.strip()
    da.bao(LUU_Y, "", "Nguồn chấm công", "Chấm công ngày lấy từ bộ nhớ đệm liên kết ngoài 'DS TỔNG' "
                      "(bảng-công-thời-vụ) lưu trong file; N = nghỉ, TD = nghỉ tự do")
    return da


# --------------------------------------------------------------------------------------
# MẪU E: 'Công (lương)' (7 dòng/NV theo hệ số) + 'Bảng lương' + 'DS CN'
# --------------------------------------------------------------------------------------
def nhan_dien_E(sheets):
    return "Công (lương)" in sheets and "Bảng lương" in sheets


def doc_mau_E(path, wb) -> DuAn:
    da = DuAn(ma=ma_du_an_tu_file(path, "Dự án E"), file=path, mau="Mẫu E (Công (lương) + Bảng lương + DS CN)", ky="")
    rows = doc_sheet(wb, "Công (lương)")
    cot_ngay = {get_column_letter(j + 1): v.date() for j, v in enumerate(rows[1]) if isinstance(v, dt.datetime)}
    da.tu_ngay, da.den_ngay = min(cot_ngay.values()), max(cot_ngay.values())
    he_so = {1.0: "X", 1.3: "OT", 1.5: "D", 2.0: "OCN", 2.3: "OCND", 3.0: "OD", 3.3: "OLD"}
    hien_tai = None
    for r in rows[4:]:
        if chuan(o(r, "A")) == "total":
            break
        if ma(o(r, "B")):
            hien_tai = ma(o(r, "B"))
            nv = da.nhan_vien.setdefault(hien_tai, {})
            nv.setdefault("ten", ten_sach(o(r, "C")))
            if ten_sach(o(r, "D")):
                nv["bo_phan"] = ten_sach(o(r, "D"))
        if hien_tai is None or o(r, "F") in (None, ""):
            continue
        code = he_so.get(round(so(o(r, "F")), 2))
        if not code:
            if so(o(r, "F")) and any(so(o(r, c)) for c in cot_ngay):
                da.bao(NGHI_NGO, hien_tai, "Nhãn giờ lạ", f"Hệ số '{o(r, 'F')}' không thuộc bảng hệ số đã biết – bỏ qua dòng này")
            continue
        for col, ngay in cot_ngay.items():
            v = o(r, col)
            if isinstance(v, str) and v.strip().upper() in LOI_CONG_THUC:
                da.o_loi.append((hien_tai, "Công (lương)", f"{col} ({ngay:%d/%m})", v.strip()))
            h = so(v)
            if h:
                da.cham_cong[(hien_tai, ngay)][code] += h

    rows = doc_sheet(wb, "Bảng lương")
    hang_td = next(i for i, r in enumerate(rows[:15]) if chuan(o(r, "B")) == "mã nv")
    dg = rows[hang_td - 1]
    for code, col in zip(["X", "OT", "D", "OD", "OCN", "OCND"], "PQRSTU"):
        da.don_gia_luong[code] = so(o(dg, col))
    for i, r in enumerate(rows[hang_td + 1:], start=hang_td + 2):
        m = ma(o(r, "B"))
        if chuan(o(r, "C")) == "tổng":
            da.tong_goc = {"thuc_nhan": o(r, "AD"), "dong": i}
            break
        if not m:
            continue
        kiem_o_loi(da, m, "Bảng lương", i, r, "AD")
        d = DongLuong(ma_nv=m, ten=ten_sach(o(r, "C")))
        d.gio = {"X": so(o(r, "G")), "OT": so(o(r, "H")), "D": so(o(r, "I")), "OD": so(o(r, "J")),
                 "OCN": so(o(r, "K")), "OCND": so(o(r, "L")), "OLD": so(o(r, "M"))}
        d.tien = {"X": so(o(r, "P")), "OT": so(o(r, "Q")), "D": so(o(r, "R")), "OD": so(o(r, "S")),
                  "OCN": so(o(r, "T")), "OCND": so(o(r, "U"))}
        d.pc_chuyen_can = so(o(r, "V"))
        d.pc_khac = so(o(r, "W"))               # trợ cấp tiền ăn
        d.tra_lai_giu = so(o(r, "X"))           # trả lại tiền giữ vi phạm báo trước
        d.khau_tru_khac = so(o(r, "Y")) + so(o(r, "AC"))   # trừ vi phạm báo trước
        d.tru_do = so(o(r, "Z"))                # trừ đồng phục
        d.tien_giu = so(o(r, "AA"))             # giữ lại tiền vi phạm báo trước
        d.ung_luong = so(o(r, "AB"))
        d.thuc_nhan_goc = so(o(r, "AD"))
        d.ghi_chu = ten_sach(o(r, "AE"))
        d.ngay_vao, d.ngay_cong = la_ngay(o(r, "D")), so(o(r, "N"))
        d.ngay_nghi = la_ngay(o(r, "E")) or la_ngay(o(r, "F"))
        d.chi_tiet = chi_tiet(("Trả lại tiền giữ vi phạm báo trước", d.tra_lai_giu),
                              ("Trừ vi phạm báo trước", so(o(r, "Y"))), ("Giữ lại tiền vi phạm báo trước", d.tien_giu),
                              ("Công HC", so(o(r, "O")), "công"))
        da.luong.append(d)

    if "DS CN" in wb.sheetnames:
        for r in doc_sheet(wb, "DS CN"):
            m = ma(o(r, "B"))
            if m and chuan(m) != "mã nv" and ten_sach(o(r, "D")):
                da.nhan_vien.setdefault(m, {})["ten_chuan"] = ten_sach(o(r, "D"))
    # Đơn giá ở dòng 3 sheet 'Công (lương)' cao hơn đơn giá lương -> nhiều khả năng là đơn giá dịch vụ
    r3 = doc_sheet(wb, "Công (lương)", max_row=3)[2]
    for code, col in zip(["X", "OT", "D", "OCN", "OCND", "OD", "OLD"], ["AL", "AM", "AN", "AO", "AP", "AQ", "AR"]):
        da.don_gia_dv[code] = so(o(r3, col))
    da.ghi_chu_doanh_thu = ("Ước tính = giờ bảng lương x đơn giá dòng 3 sheet 'Công (lương)' "
                            "(CHƯA XÁC NHẬN là đơn giá dịch vụ; file không có bảng kê)")
    return da


ADAPTERS = [
    ("A", nhan_dien_A, doc_mau_A),
    ("B", nhan_dien_B, doc_mau_B),
    ("C", nhan_dien_C, doc_mau_C),
    ("D", nhan_dien_D, doc_mau_D),
    ("E", nhan_dien_E, doc_mau_E),
]


def doc_file(path, log=print) -> DuAn:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ten, nd, doc in ADAPTERS:
            if nd(wb.sheetnames):
                log(f"  -> nhận diện biểu mẫu {ten}")
                da = doc(path, wb)
                break
        else:
            raise ValueError(f"Không nhận diện được biểu mẫu của file {os.path.basename(path)} "
                             f"(sheets: {wb.sheetnames})")
    except ValueError:
        raise
    except Exception as e:      # StopIteration (không thấy dòng tiêu đề), sheet rỗng, ...
        raise ValueError(f"File {os.path.basename(path)} nhận diện là mẫu {ten} nhưng đọc không được "
                         f"({type(e).__name__}: {e or 'không tìm thấy dòng tiêu đề/dữ liệu'}). "
                         "Có thể biểu mẫu đã bị đổi cấu trúc.") from e
    finally:
        wb.close()
    hoan_thien(da)
    kiem_tra_du_an(da)
    return da


def hoan_thien(da: DuAn):
    """Xác định kỳ, lọc chấm công theo NV trong bảng lương, gom thông tin NV, đối chiếu."""
    th = ky_tu_ten_file(da.file)
    if not da.tu_ngay:
        ngay = sorted({k[1] for k in da.cham_cong} | {k[1] for k in da.ky_hieu_ngay})
        if ngay:
            da.tu_ngay, da.den_ngay = ngay[0], ngay[-1]
        elif th:
            da.tu_ngay = dt.date(th[1], th[0], 1)
            da.den_ngay = dt.date(th[1] + th[0] // 12, th[0] % 12 + 1, 1) - dt.timedelta(days=1)
    if not th and da.den_ngay:
        th = (da.den_ngay.month, da.den_ngay.year)     # kỳ 26/07 - 25/08 => T08
    da.ky = f"T{th[0]:02d}/{th[1]}" if th else ""

    trong_luong = {d.ma_nv for d in da.luong}
    co_cong = {k[0] for k in da.cham_cong}
    da.ngoai_bang_luong = co_cong - trong_luong
    # trùng mã
    dem = defaultdict(int)
    for d in da.luong:
        dem[d.ma_nv] += 1
    for m, c in dem.items():
        if c > 1:
            da.bao(LOI, m, "Trùng mã NV", f"Mã xuất hiện {c} lần trong bảng lương")
    # đối chiếu giờ chấm công vs bảng lương
    tong_cc = defaultdict(lambda: defaultdict(float))
    for (m, _), g in da.cham_cong.items():
        for c, h in g.items():
            tong_cc[m][c] += h
    for d in da.luong:
        gio_cc_can = {c: d.gio.get(c, 0) - d.gio_ngoai_cham_cong.get(c, 0) for c in LOAI_TIEN}
        if d.ma_nv not in co_cong:
            if sum(gio_cc_can.values()) > 0.01:
                da.bao(NGHI_NGO, d.ma_nv, "Không có chấm công ngày", "Có giờ trong bảng lương nhưng không có chấm công ngày")
        elif dem[d.ma_nv] == 1:        # trùng mã đã báo lỗi riêng, không so giờ 2 lần
            lech = [f"{TEN_LOAI[c].replace(chr(10), ' ')}: công {tong_cc[d.ma_nv].get(c, 0):g} / lương {d.gio.get(c, 0):g}"
                    for c in LOAI_TIEN if abs(tong_cc[d.ma_nv].get(c, 0) - gio_cc_can[c]) > 0.01]
            if lech:
                da.bao(NGHI_NGO, d.ma_nv, "Lệch giờ công/lương", "; ".join(lech))
        cl = d.thuc_nhan - (d.thuc_nhan_goc or 0)
        if d.thuc_nhan_goc is not None and abs(cl) > 1:
            da.bao(LOI, d.ma_nv, "Lệch thực nhận", f"Tính lại {d.thuc_nhan:,.0f} / file gốc {d.thuc_nhan_goc:,.0f}")
    if da.ngoai_bang_luong:
        da.bao(LUU_Y, "", "NV có công nhưng không có trong bảng lương",
                f"{len(da.ngoai_bang_luong)} mã (đã loại khỏi file chuẩn): " + ", ".join(sorted(da.ngoai_bang_luong)[:40])
                + (" ..." if len(da.ngoai_bang_luong) > 40 else ""))
    for d in da.luong:
        nv = da.nhan_vien.setdefault(d.ma_nv, {})
        if nv.get("ten_chuan") and chuan(nv["ten_chuan"]) == chuan(d.ten):
            d.ten = nv["ten_chuan"]
        nv["ten"] = d.ten


# --------------------------------------------------------------------------------------
# KIỂM TRA LỖI / NGHI NGỜ
# --------------------------------------------------------------------------------------
def doanh_thu_tinh(da: DuAn) -> float:
    if da.doanh_thu is not None:
        return sum(x[3] or 0 for x in da.doanh_thu)
    return sum(d.gio.get(c, 0) * da.don_gia_dv.get(c, 0) for d in da.luong for c in LOAI_TIEN)


def _rut_gon(ds, n=15):
    ds = list(ds)
    return ", ".join(ds[:n]) + (f" ... (+{len(ds) - n})" if len(ds) > n else "")


def _mau_ma(m):
    # chỉ so cấu trúc chữ/số (HMT02 và HMT474 cùng dạng), không so độ dài dãy số
    return re.sub(r"9+", "9", re.sub(r"[A-Za-z]+", "A", re.sub(r"\d", "9", m)))


def kiem_tra_du_an(da: DuAn):
    """Các quy tắc phát hiện sai sót / bất thường trong 1 hợp đồng."""
    ten_nv = {d.ma_nv: d.ten for d in da.luong}

    # 1. File / cấu trúc
    if not da.luong:
        da.bao(LOI, "", "Bảng lương trống", "Không đọc được nhân viên nào trong bảng lương - kiểm tra lại biểu mẫu")
    th = ky_tu_ten_file(da.file)
    if th and da.tu_ngay and da.den_ngay and th[0] not in (da.tu_ngay.month, da.den_ngay.month):
        da.bao(NGHI_NGO, "", "Kỳ lương không khớp", f"Tên file ghi T{th[0]:02d}/{th[1]} nhưng bảng công là "
               f"{da.tu_ngay:%d/%m/%Y} – {da.den_ngay:%d/%m/%Y}")
    if da.luong:
        co_gio = sum(sum(d.gio.values()) for d in da.luong)
        co_thuc_nhan = sum(abs(d.thuc_nhan_goc or 0) for d in da.luong)
        if co_gio > 0 and co_thuc_nhan == 0:
            da.bao(LOI, "", "Không có giá trị tiền",
                   "Có giờ công nhưng cột Thực nhận của file gốc trống ở mọi NV. Thường do file được tạo/lưu bằng công cụ không tính công "
                   "thức (giá trị chưa được lưu). Mở file bằng Excel, bấm Lưu, rồi chạy lại.")
    if da.tu_ngay and da.den_ngay and (da.den_ngay - da.tu_ngay).days + 1 > 31:
        da.bao(LOI, "", "Kỳ công dài hơn 31 ngày",
               f"Kỳ {da.tu_ngay:%d/%m/%Y} – {da.den_ngay:%d/%m/%Y} dài {(da.den_ngay - da.tu_ngay).days + 1} ngày; sheet "
               "bang_cham_cong_thang chỉ có 31 cột ngày nên công sau ngày thứ 31 không hiện trong lưới. Kiểm tra ngày gõ nhầm trong bảng công.")
    if da.o_loi:
        theo_nv = defaultdict(list)
        for m, sh, cell, v in da.o_loi:
            theo_nv[m].append(f"{sh}!{cell}={v}")
        for m, ds in theo_nv.items():
            if m in ten_nv:
                da.bao(LOI, m, "Ô lỗi công thức", "Ô bị lỗi, giá trị được tính là 0: " + "; ".join(ds[:8])
                       + (" ..." if len(ds) > 8 else ""))

    # 2. Dòng TỔNG của file gốc so với tổng các dòng đã đọc (phát hiện dòng bị sót / ẩn / gõ đè)
    v, gia_tri, ten = da.tong_goc.get("thuc_nhan"), sum(d.thuc_nhan_goc or 0 for d in da.luong), "Thực nhận"
    if isinstance(v, (int, float)) and abs(v - gia_tri) > NGUONG["sai_so_tien"]:
            da.bao(LOI, "", "Lệch dòng TỔNG file gốc",
                   f"{ten} ở dòng TỔNG (dòng {da.tong_goc.get('dong')}) = {v:,.0f} nhưng cộng các dòng NV = {gia_tri:,.0f} "
                   f"(chênh {v - gia_tri:,.0f}) - có thể sót dòng, dòng bị ẩn hoặc gõ đè số tổng")

    # 3. Từng nhân viên
    bo_don_gia = [da.don_gia_luong] + da.don_gia_luong_khac
    mau_pho_bien = defaultdict(int)
    for d in da.luong:
        mau_pho_bien[_mau_ma(d.ma_nv)] += 1
    mau_chinh = max(mau_pho_bien, key=mau_pho_bien.get) if mau_pho_bien else None
    vd_ma = next((d.ma_nv for d in da.luong if _mau_ma(d.ma_nv) == mau_chinh), "")
    so_ngay_ky = (da.den_ngay - da.tu_ngay).days + 1 if da.tu_ngay and da.den_ngay else 31
    vuot_tc = []
    trung_ten = defaultdict(list)
    for d in da.luong:
        m = d.ma_nv
        if not d.ten:
            da.bao(LOI, m, "Thiếu tên NV", "Dòng lương không có họ tên")
        else:
            trung_ten[(chuan(d.ten), d.ngay_vao)].append(m)
        if mau_chinh and mau_pho_bien[mau_chinh] >= 5 and _mau_ma(m) != mau_chinh:
            da.bao(NGHI_NGO, m, "Mã NV khác định dạng", f"Mã '{m}' khác dạng của các mã khác trong hợp đồng (vd '{vd_ma}') – có thể gõ thừa khoảng trắng/ký tự")
        # giờ <-> tiền <-> đơn giá
        for c in LOAI_TIEN:
            h, tien = d.gio.get(c, 0), d.tien.get(c, 0)
            ten_loai = TEN_LOAI[c].replace("\n", " ")
            if tien and not h:
                da.bao(LOI, m, "Có tiền không có giờ", f"{ten_loai}: tiền {tien:,.0f} nhưng số giờ = 0")
            elif h and not tien and any(b.get(c) for b in bo_don_gia):
                da.bao(NGHI_NGO, m, "Có giờ không có tiền", f"{ten_loai}: {h:g} giờ nhưng tiền lương = 0")
            elif h and tien:
                dgs = [b[c] for b in bo_don_gia if b.get(c)]
                if dgs and not any(abs(tien - h * g) <= NGUONG["sai_so_tien"] + 0.001 * tien for g in dgs):
                    da.bao(LOI, m, "Sai đơn giá", f"{ten_loai}: {h:g} giờ x đơn giá {'/'.join(f'{g:,.0f}' for g in dgs)} "
                           f"= {h * dgs[0]:,.0f} nhưng bảng lương ghi {tien:,.0f} (≈ {tien / h:,.0f}đ/giờ)")
        if d.gio.get("X") and d.tien.get("X"):
            dg = d.tien["X"] / d.gio["X"]
            if dg < NGUONG["luong_gio_toi_thieu_vung_IV"]:
                da.bao(LOI, m, "Dưới lương tối thiểu", f"Lương giờ HC {dg:,.0f}đ thấp hơn mức tối thiểu vùng IV "
                       f"({NGUONG['luong_gio_toi_thieu_vung_IV']:,}đ/giờ)")
            elif dg < NGUONG["luong_gio_toi_thieu_vung_I"]:
                da.bao(LUU_Y, m, "Có thể dưới lương tối thiểu", f"Lương giờ HC {dg:,.0f}đ thấp hơn mức vùng I "
                       f"({NGUONG['luong_gio_toi_thieu_vung_I']:,}đ/giờ) - kiểm tra vùng của nơi làm việc")
        # tiền âm / bất thường
        if d.thuc_nhan < -NGUONG["sai_so_tien"]:
            da.bao(NGHI_NGO, m, "Thực nhận âm", f"Thực nhận {d.thuc_nhan:,.0f} (ứng/khấu trừ lớn hơn lương)")
        for ten_k, v in [("Ứng lương", d.ung_luong), ("Trừ đồ", d.tru_do), ("Bảo hiểm", d.bhxh),
                         ("Khấu trừ khác", d.khau_tru_khac), ("Tiền giữ", d.tien_giu)]:
            if v < 0:
                da.bao(NGHI_NGO, m, "Khoản trừ âm", f"{ten_k} = {v:,.0f} (khoản trừ mang dấu âm)")
        for c in LOAI_TIEN:
            if d.gio.get(c, 0) < 0 or d.tien.get(c, 0) < 0:
                da.bao(LOI, m, "Giờ/tiền âm", f"{TEN_LOAI[c].replace(chr(10), ' ')}: giờ {d.gio.get(c, 0):g}, tiền {d.tien.get(c, 0):,.0f}")
        if d.ngay_cong and d.ngay_cong > so_ngay_ky + 0.01:
            da.bao(LOI, m, "Ngày công vượt kỳ", f"Ngày công {d.ngay_cong:g} > {so_ngay_ky} ngày của kỳ công")
        tc = sum(d.gio.get(c, 0) for c in LOAI_TANG_CA)
        if tc > NGUONG["tang_ca_thang"]:
            vuot_tc.append((tc, m))
        # ngày vào / ngày nghỉ
        if d.ngay_vao and d.ngay_nghi and d.ngay_nghi < d.ngay_vao:
            da.bao(LOI, m, "Ngày nghỉ trước ngày vào", f"Ngày vào {d.ngay_vao:%d/%m/%Y}, ngày nghỉ {d.ngay_nghi:%d/%m/%Y}")
        if d.ngay_vao and da.den_ngay and d.ngay_vao > da.den_ngay and sum(d.gio.values()):
            da.bao(NGHI_NGO, m, "Ngày vào sau kỳ công", f"Ngày vào {d.ngay_vao:%d/%m/%Y} nhưng có giờ trong kỳ đến {da.den_ngay:%d/%m/%Y}")

    if vuot_tc:
        vuot_tc.sort(reverse=True)
        da.bao(LUU_Y, "", "Làm thêm vượt 40 giờ/tháng",
               f"{len(vuot_tc)} NV có tổng giờ tăng ca > {NGUONG['tang_ca_thang']} giờ/kỳ (BLLĐ 2019 Điều 107). "
               f"Cao nhất: " + _rut_gon(f"{m} ({h:g}h)" for h, m in vuot_tc[:10]))
    for (ten, nv), ds in trung_ten.items():
        if len(ds) > 1:
            da.bao(NGHI_NGO, ds[0], "Có thể trùng người", f"{len(ds)} mã cùng họ tên '{ten_nv[ds[0]]}'"
                   + (f" và cùng ngày vào {nv:%d/%m/%Y}" if nv else "") + ": " + ", ".join(ds))

    # 4. Chấm công theo ngày
    ngay_vao = {d.ma_nv: d.ngay_vao for d in da.luong}
    ngay_nghi = {d.ma_nv: d.ngay_nghi for d in da.luong}
    co_tang_ca_cn = any(d.gio.get("OCN") or d.gio.get("OCND") for d in da.luong) or bool(da.don_gia_luong.get("OCN"))
    ngoai_ky, qua_gio, sau_nghi, truoc_vao, hc_cn = (defaultdict(list) for _ in range(5))
    for (m, ngay), g in da.cham_cong.items():
        if m not in ten_nv:
            continue
        tong = sum(g.values())
        if da.tu_ngay and tong and (not (da.tu_ngay <= ngay <= da.den_ngay) or (ngay - da.tu_ngay).days >= 31):
            ngoai_ky[m].append(ngay)
        if tong > NGUONG["gio_ngay_nghi_ngo"]:
            qua_gio[m].append((ngay, tong))
        if ngay_nghi.get(m) and ngay > ngay_nghi[m] and tong:
            sau_nghi[m].append(ngay)
        if ngay_vao.get(m) and ngay < ngay_vao[m] and tong:
            truoc_vao[m].append(ngay)
        if co_tang_ca_cn and ngay.weekday() == 6 and any(g.get(c) for c in ("X", "OT", "D", "OD")):
            hc_cn[m].append(ngay)
    for m, ds in ngoai_ky.items():
        da.bao(LOI, m, "Công ngoài kỳ", f"{len(ds)} ngày có công nằm ngoài kỳ / ngoài 31 cột ngày: "
               + _rut_gon(f"{n:%d/%m/%Y}" for n in sorted(ds)))
    for m, ds in qua_gio.items():
        nang = [x for x in ds if x[1] > NGUONG["gio_ngay_loi"]]
        da.bao(LOI if nang else NGHI_NGO, m, "Giờ làm trong ngày quá cao",
               f"{len(ds)} ngày > {NGUONG['gio_ngay_nghi_ngo']} giờ: " + _rut_gon(f"{n:%d/%m} ({h:g}h)" for n, h in sorted(ds)))
    for m, ds in sau_nghi.items():
        da.bao(NGHI_NGO, m, "Có công sau ngày nghỉ việc", f"Nghỉ việc {ngay_nghi[m]:%d/%m/%Y} nhưng còn công {len(ds)} ngày: "
               + _rut_gon(f"{n:%d/%m}" for n in sorted(ds)))
    for m, ds in truoc_vao.items():
        da.bao(NGHI_NGO, m, "Có công trước ngày vào", f"Vào làm {ngay_vao[m]:%d/%m/%Y} nhưng có công {len(ds)} ngày trước đó: "
               + _rut_gon(f"{n:%d/%m}" for n in sorted(ds)))
    for m, ds in hc_cn.items():
        da.bao(NGHI_NGO, m, "Giờ thường ghi vào Chủ nhật", f"{len(ds)} ngày Chủ nhật có giờ HC/tăng ca ngày thường "
               "(hợp đồng có đơn giá tăng ca CN riêng): " + _rut_gon(f"{n:%d/%m}" for n in sorted(ds)))

    # 5. Doanh thu
    dthu = doanh_thu_tinh(da)
    tong_luong = sum(d.tong_luong for d in da.luong)
    if da.doanh_thu_goc is not None and abs(da.doanh_thu_goc - dthu) > NGUONG["sai_so_tien"]:
        da.bao(LOI, "", "Lệch doanh thu", f"Doanh thu tính được {dthu:,.0f} khác bảng kê gốc {da.doanh_thu_goc:,.0f}")
    if dthu and tong_luong and dthu < tong_luong:
        da.bao(NGHI_NGO, "", "Doanh thu thấp hơn quỹ lương", f"Doanh thu {dthu:,.0f} < tổng lương {tong_luong:,.0f}")
    if not da.don_gia_dv or not any(da.don_gia_dv.values()):
        da.bao(LUU_Y, "", "Thiếu đơn giá dịch vụ", "Không tìm thấy đơn giá dịch vụ, chưa tính được doanh thu")


def kiem_tra_cheo(ds: list[DuAn]):
    """Kiểm tra chéo giữa các hợp đồng (chạy sau khi đã đọc hết file)."""
    theo_ma = defaultdict(list)
    theo_stk = defaultdict(set)
    for da in ds:
        for d in da.luong:
            theo_ma[d.ma_nv].append(da)
            stk = da.nhan_vien.get(d.ma_nv, {}).get("stk")
            if stk:
                theo_stk[re.sub(r"\D", "", stk)].add((da.ma, d.ma_nv, d.ten))
    for m, das in theo_ma.items():
        if len({x.ma for x in das}) > 1:
            for da in das:
                da.bao(NGHI_NGO, m, "Trùng mã giữa các hợp đồng",
                       f"Mã {m} có trong {', '.join(sorted({x.ma for x in das}))} - kiểm tra có phải cùng một người")
    for stk, nvs in theo_stk.items():
        if len({x[1] for x in nvs}) > 1:
            for da in ds:
                for hd, m, t in nvs:
                    if hd == da.ma:
                        da.bao(NGHI_NGO, m, "Trùng số tài khoản",
                               f"STK {stk} dùng cho {len(nvs)} NV: " + "; ".join(f"{x[1]} {x[2]} ({x[0]})" for x in sorted(nvs)))
    dem_ky = Counter(da.ky for da in ds if da.ky)
    if len(dem_ky) > 1:
        ky_chinh = dem_ky.most_common(1)[0][0]
        for da in ds:
            if da.ky and da.ky != ky_chinh:
                da.bao(LUU_Y, "", "Khác kỳ lương", f"Hợp đồng này kỳ {da.ky}, các hợp đồng khác phần lớn kỳ {ky_chinh}. "
                       "Kiểm tra có chọn nhầm file tháng khác không.")
    for da in ds:
        da.van_de.sort(key=lambda x: (THU_TU_MUC[x[0]], x[2], x[1]))


def dem_muc(da: DuAn) -> dict:
    dem = {LOI: 0, NGHI_NGO: 0, LUU_Y: 0}
    for x in da.van_de:
        dem[x[0]] += 1
    return dem


def canh_bao_nv(da: DuAn) -> dict:
    """ma_nv -> (mức cao nhất, 'Lỗi: ...; Nghi ngờ: ...')"""
    kq = defaultdict(list)
    for muc, m, loai, _ in da.van_de:
        if m:
            kq[m].append((muc, loai))
    out = {}
    for m, ds in kq.items():
        ds.sort(key=lambda x: THU_TU_MUC[x[0]])
        nhom = defaultdict(list)
        for muc, loai in ds:
            if loai not in nhom[muc]:
                nhom[muc].append(loai)
        out[m] = (ds[0][0], "; ".join(f"{muc}: {', '.join(l)}" for muc, l in nhom.items()))
    return out


# --------------------------------------------------------------------------------------
# GHI FILE CHUẨN
# --------------------------------------------------------------------------------------
FONT = "Arial"
F_TD = Font(name=FONT, bold=True, size=10)
F_TH = Font(name=FONT, size=10)
F_GHI = Font(name=FONT, size=9, italic=True, color="666666")
NEN_TD = PatternFill("solid", fgColor="D9E1F2")
NEN_PHU = PatternFill("solid", fgColor="FFF2CC")
NEN_DO = PatternFill("solid", fgColor="F8CBAD")
NEN_MUC = {"Lỗi": PatternFill("solid", fgColor="F4B6A6"), "Nghi ngờ": PatternFill("solid", fgColor="FCE4B0"),
           "Lưu ý": PatternFill("solid", fgColor="DDEBF7")}
VIEN = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
CANH_GIUA = Alignment(horizontal="center", vertical="center", wrap_text=True)
SO = "#,##0;-#,##0;-"
GIO = "#,##0.##;-#,##0.##;-"
THU = {0: "Thứ 2", 1: "Thứ 3", 2: "Thứ 4", 3: "Thứ 5", 4: "Thứ 6", 5: "Thứ 7", 6: "Chủ nhật"}


def _tieu_de(ws, hang, ds, nen=NEN_TD, bat_dau=1):
    for j, t in enumerate(ds, start=bat_dau):
        c = ws.cell(hang, j, t)
        c.font, c.fill, c.alignment, c.border = F_TD, nen, CANH_GIUA, VIEN


def _o(ws, r, c, v, fmt=None, font=F_TH):
    cell = ws.cell(r, c, v)
    # gán font/viền từng ô rất chậm ở openpyxl (băm lại style mỗi lần); chép mảng style của ô đầu tiên
    # cùng kiểu. Mảng style trỏ vào danh sách style của workbook nên bộ đệm lưu theo từng workbook.
    # Chỉ áp cho ô chưa có style, để không xóa nền/căn lề đã gán trước đó.
    dem = ws.parent.__dict__.setdefault("_kieu_o", {})
    mau = dem.get((id(font), fmt))
    if mau is not None and not cell.has_style:
        cell._style = copy(mau)
        return cell
    moi = not cell.has_style
    cell.font, cell.border = font, VIEN
    if fmt:
        cell.number_format = fmt
    if moi:
        dem[(id(font), fmt)] = copy(cell._style)
    return cell


def _khoi_cham_cong(ws, r, stt, da, ngay, ky, tieu_de, cot_phu, ten_phu, gt_phu, gop=True,
                    ten_cot_e="Tổng công ca ngày", nhan=TEN_LOAI):
    """Ghi 1 khối chấm công theo template (hàng thứ, hàng tiêu đề có ngày, 1 hàng trống, rồi dữ liệu từ hàng 4).
    ngay: các ngày ứng với cột G trở đi. ky=None (bảng tháng): cột F = toàn bộ giờ của NV trong kỳ;
    ky=(từ, đến) (bảng tuần): cột F = giờ trong khối, ngày ngoài kỳ tô nhạt ở hàng tiêu đề.
    gop: gộp ô Mã NV/Tên theo NV (template tháng có gộp, template tuần không).
    ten_cot_e, nhan: tiêu đề cột E và tên loại giờ (template tuần ghi khác template tháng).
    Trả về (hàng tiếp theo sau khối, stt)."""
    cc_nv = defaultdict(dict)                    # mã NV -> {ngày: giờ theo loại}
    for (ma, n), g in da.cham_cong.items():
        cc_nv[ma][n] = g
    cot_cua_ngay = {n: 7 + i for i, n in enumerate(ngay)}
    if tieu_de:
        ws.cell(r, 1, tieu_de).font = Font(name=FONT, bold=True, size=10, color="C00000")
    for n, col in cot_cua_ngay.items():
        ws.cell(r, col, THU[n.weekday()]).font = F_TH
    r += 1
    _tieu_de(ws, r, ["STT", "Mã theo dõi", "Mã NV", "Tên nhân viên", ten_cot_e, "TỔNG CỘNG"])
    for n, col in cot_cua_ngay.items():
        c = ws.cell(r, col, dt.datetime.combine(n, dt.time()))
        c.number_format = "yyyy-mm-dd"
        ngoai = ky and not (ky[0] <= n <= ky[1])
        c.font, c.fill, c.alignment, c.border = F_TD, NEN_PHU if ngoai else NEN_TD, CANH_GIUA, VIEN
    _tieu_de(ws, r, ten_phu, nen=NEN_PHU, bat_dau=cot_phu)
    r += 2      # bỏ trống 1 hàng như template (dữ liệu bắt đầu ở hàng thứ 4 của khối)
    da_ghi = set()
    for d in da.luong:
        # mã trùng: chấm công ngày chỉ ghi 1 lần (ở dòng đầu tiên) để không nhân đôi giờ
        ngay_nv = {} if d.ma_nv in da_ghi else cc_nv.get(d.ma_nv, {})
        da_ghi.add(d.ma_nv)
        loai = LOAI_CHINH + [c for c in LOAI_GIO_PHU(da, d.ma_nv)]
        stt += 1
        r0 = r
        for code in loai:
            _o(ws, r, 1, stt if r == r0 else None)
            _o(ws, r, 2, d.ma_nv, "@")
            _o(ws, r, 3, d.ma_nv if r == r0 else None, "@")
            _o(ws, r, 4, d.ten if r == r0 else None)
            _o(ws, r, 5, nhan[code])
            tong = 0
            for n, g in (ngay_nv.items() if ky is None else ((n, ngay_nv[n]) for n in ngay if n in ngay_nv)):
                h = g.get(code, 0)
                if h:
                    if n in cot_cua_ngay:
                        _o(ws, r, cot_cua_ngay[n], h, GIO)
                    tong += h
            for col in cot_cua_ngay.values():
                if ws.cell(r, col).value is None:
                    _o(ws, r, col, None)
            _o(ws, r, 6, tong, GIO)
            for j, v in enumerate(gt_phu, start=cot_phu):
                _o(ws, r, j, v, "dd/mm/yyyy" if isinstance(v, dt.date) else None)
            r += 1
        if gop and r - r0 > 1:
            ws.merge_cells(start_row=r0, end_row=r - 1, start_column=3, end_column=3)
            ws.merge_cells(start_row=r0, end_row=r - 1, start_column=4, end_column=4)
            for c in (3, 4):
                ws.cell(r0, c).alignment = Alignment(vertical="center", wrap_text=True)
    return r + 2, stt      # cách 2 hàng trước khối tiếp theo


def cot_O(d) -> float:
    """Cột O 'Phụ cấp khác' của template = phụ cấp khác + lương TC CN/lễ + trả lại tiền giữ."""
    return d.pc_khac + sum(d.tien.get(c, 0) for c in ("OCN", "OCND", "OL", "OLD")) + d.tra_lai_giu


def cot_R(d) -> float:
    """Cột R 'Trừ đồ' của template = trừ đồ + bảo hiểm + khấu trừ khác."""
    return d.tru_do + d.bhxh + d.khau_tru_khac


def ghi_file_chuan(ds: list[DuAn], out_path, log=print, file_loi=()):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    mot_ky = len({(d.tu_ngay, d.den_ngay) for d in ds}) == 1

    # ================= DM =================
    ws = wb.create_sheet("DM")
    _tieu_de(ws, 1, ["STT", "Ký hiệu", "Nội dung", "Đơn vị tính", "Đơn giá lương", "Đơn giá dịch vụ"])
    for i, (code, ten) in enumerate(LOAI_GIO, start=1):
        _o(ws, i + 1, 1, i); _o(ws, i + 1, 2, code); _o(ws, i + 1, 3, ten); _o(ws, i + 1, 4, "giờ")
        if len(ds) == 1:
            _o(ws, i + 1, 5, ds[0].don_gia_luong.get(code) or None, SO)
            _o(ws, i + 1, 6, ds[0].don_gia_dv.get(code) or None, SO)
        else:
            _o(ws, i + 1, 5, None); _o(ws, i + 1, 6, None)
    if len(ds) > 1:
        ws.cell(len(LOAI_GIO) + 3, 1, "Đơn giá khác nhau theo dự án → xem bảng 'Đơn giá theo dự án' (cột X trở đi)").font = F_GHI
    _tieu_de(ws, 1, ["Mã Nhân viên", "Họ tên", "Tỉnh thành", "CMT/CCCD", "Số TK"], bat_dau=9)
    _tieu_de(ws, 1, ["Ngân hàng", "Chủ TK", "Bộ phận", "Chức danh", "NCC", "Mã thẻ", "Ngày vào", "Ngày nghỉ",
                     "Hợp đồng / Dự án"], nen=NEN_PHU, bat_dau=14)
    r = 2
    for da in ds:
        for d in da.luong:
            nv = da.nhan_vien.get(d.ma_nv, {})
            vals = [d.ma_nv, d.ten, nv.get("tinh"), nv.get("cccd"), nv.get("stk"), nv.get("ngan_hang"),
                    nv.get("chu_tk"), nv.get("bo_phan"), nv.get("chuc_danh"), nv.get("ncc"), nv.get("ma_the"),
                    d.ngay_vao, d.ngay_nghi, da.ma]
            for j, v in enumerate(vals, start=9):
                _o(ws, r, j, v or None, "dd/mm/yyyy" if j in (20, 21) else ("@" if j in (9, 12, 13) else None))
            r += 1
    _tieu_de(ws, 1, ["Hợp đồng / Dự án", "Ký hiệu", "Nội dung", "Đơn giá lương", "Đơn giá dịch vụ", "Ghi chú"], bat_dau=24)
    r = 2
    for da in ds:
        for code in LOAI_TIEN:
            if da.don_gia_luong.get(code) or da.don_gia_dv.get(code):
                for j, v in enumerate([da.ma, code, TEN_LOAI[code], da.don_gia_luong.get(code) or None,
                                       da.don_gia_dv.get(code) or None, da.ghi_chu_don_gia], start=24):
                    _o(ws, r, j, v, SO if j in (27, 28) else None)
                r += 1
    for col, w in {"A": 5, "B": 8, "C": 24, "D": 9, "E": 13, "F": 14, "I": 13, "J": 26, "K": 12, "L": 14,
                   "M": 16, "N": 12, "O": 22, "P": 22, "Q": 14, "R": 8, "S": 10, "T": 11, "U": 11, "V": 12,
                   "X": 9, "Y": 8, "Z": 24, "AA": 13, "AB": 14, "AC": 40}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"

    # ================= bang_cham_cong_thang =================
    # Mỗi hợp đồng (dự án) là 1 khối giống hệt template: hàng thứ + hàng tiêu đề có NGÀY THẬT theo đúng
    # kỳ công của bảng chấm công gốc (vd 26/07 -> 25/08), cột G = ngày đầu kỳ, H = ngày thứ 2...
    ws = wb.create_sheet("bang_cham_cong_thang")
    r, stt = 1, 0
    for da in ds:
        ngay_ky = []
        if da.tu_ngay:
            n = da.tu_ngay
            while n <= da.den_ngay and len(ngay_ky) < 31:
                ngay_ky.append(n)
                n += dt.timedelta(days=1)
        # file 1 hợp đồng để trống A1 như template
        tieu_de = None
        if len(ds) > 1:
            tieu_de = (f"HỢP ĐỒNG: {da.ma} | Kỳ {da.ky} | {da.tu_ngay:%d/%m/%Y} – {da.den_ngay:%d/%m/%Y}"
                       if da.tu_ngay else f"HỢP ĐỒNG: {da.ma}")
        r, stt = _khoi_cham_cong(ws, r, stt, da, ngay_ky, None, tieu_de, 39,
                                 ["Hợp đồng / Dự án", "Kỳ lương", "Từ ngày", "Đến ngày"],
                                 [da.ma, da.ky, da.tu_ngay, da.den_ngay])
    for col, w in {"A": 5, "B": 12, "C": 12, "D": 24, "E": 20, "F": 10}.items():
        ws.column_dimensions[col].width = w
    for k in range(7, 38):
        ws.column_dimensions[get_column_letter(k)].width = 10.5
    ws.column_dimensions["AM"].width = 9
    ws.column_dimensions["AN"].width = 10
    ws.column_dimensions["AO"].width = 11
    ws.column_dimensions["AP"].width = 11
    ws.freeze_panes = "G1"

    # ================= bang_cham_cong_tuan =================
    # Mỗi tuần (Thứ 2 -> Chủ nhật) của kỳ công là 1 khối giống template: G = Thứ 2 ... M = Chủ nhật.
    # Tuần đầu/cuối kỳ có thể chứa ngày ngoài kỳ: vẫn hiện ngày (tô nhạt) để cột luôn đúng thứ.
    ws = wb.create_sheet("bang_cham_cong_tuan")
    r, stt = 1, 0
    for da in ds:
        if not da.tu_ngay:
            continue
        dau = da.tu_ngay - dt.timedelta(days=da.tu_ngay.weekday())
        so_tuan = 0
        while dau <= da.den_ngay:
            so_tuan += 1
            tuan = [dau + dt.timedelta(days=i) for i in range(7)]
            # file 1 hợp đồng để trống A1 như template (tuần đã thể hiện ở hàng ngày và cột O:S)
            tieu_de = (f"HỢP ĐỒNG: {da.ma} | Kỳ {da.ky} | TUẦN {so_tuan}: {max(dau, da.tu_ngay):%d/%m/%Y} – "
                       f"{min(tuan[-1], da.den_ngay):%d/%m/%Y}") if len(ds) > 1 else None
            r, _ = _khoi_cham_cong(ws, r, 0, da, tuan, (da.tu_ngay, da.den_ngay), tieu_de, 15,
                                   ["Hợp đồng / Dự án", "Kỳ lương", "Tuần", "Từ ngày", "Đến ngày"],
                                   [da.ma, da.ky, so_tuan, tuan[0], tuan[-1]], gop=False,
                                   ten_cot_e="Loại công", nhan=TEN_LOAI_TUAN)
            dau += dt.timedelta(days=7)
    for col, w in {"A": 5, "B": 12, "C": 12, "D": 24, "E": 20, "F": 10, "N": 3,
                   "O": 9, "P": 10, "Q": 6, "R": 11, "S": 11}.items():
        ws.column_dimensions[col].width = w
    for k in range(7, 14):
        ws.column_dimensions[get_column_letter(k)].width = 10.5
    ws.freeze_panes = "G1"

    # ================= bang_luong =================
    ws = wb.create_sheet("bang_luong")
    td = ["STT", "Mã nhân viên", "Tên nhân viên", TEN_LOAI["X"], TEN_LOAI["D"], TEN_LOAI["OT"], TEN_LOAI["OD"],
          "Lương ca ngày", "Lương ca đêm", "Lương tăng ca ngày thường", "Lương tăng ca đêm ngày thường",
          "Phụ cấp chức vụ", "Phụ cấp chuyên cần", "Phụ cấp thâm niên", "Phụ cấp khác", "Thưởng kỳ", "Tổng lương",
          "Trừ đồ", "Tiền giữ tháng trước", "Ứng lương", "Thực nhận"]
    phu = ["Giờ TC ngày nghỉ/CN", "Giờ TC đêm ngày nghỉ/CN", "Giờ TC ngày lễ", "Giờ TC đêm ngày lễ",
           "Lương TC ngày nghỉ/CN (đã gồm trong O)", "Lương TC đêm ngày nghỉ/CN (đã gồm trong O)",
           "Lương TC ngày lễ (đã gồm trong O)", "Lương TC đêm ngày lễ (đã gồm trong O)",
           "Trả lại tiền giữ / hoàn cọc (đã gồm trong O)", "Bảo hiểm (đã gồm trong R)", "Khấu trừ khác (đã gồm trong R)",
           "Hợp đồng / Dự án", "Kỳ lương", "Thực nhận (file gốc)", "Chênh lệch",
           "Ngày công", "Ngày vào", "Ngày nghỉ", "Chi tiết khoản gộp (ứng từng đợt, trừ KTX/thẻ/ĐP, phụ cấp...)", "Ghi chú",
           "Cảnh báo (xem sheet kiem_tra)"]
    _tieu_de(ws, 2, td)
    _tieu_de(ws, 2, phu, nen=NEN_PHU, bat_dau=22)
    nhom = {12: "Cố định", 13: "Ổn định", 14: "Cố định", 15: "Biến đổi theo kỳ", 16: "Biến đổi theo kỳ",
            18: "Điền tay", 19: "Điền tay", 20: "Điền tay"}
    for c, t in nhom.items():
        ws.cell(1, c, t).font = F_GHI
    if len(ds) == 1:
        for c, code in zip([4, 5, 6, 7], LOAI_CHINH):
            _o(ws, 1, c, ds[0].don_gia_luong.get(code) or None, SO)
    else:
        ws.cell(1, 1, "Đơn giá theo dự án: xem DM!X:AC").font = F_GHI
    ws.cell(1, 22, "Cột bổ sung để xem chi tiết, KHÔNG cộng thêm vào Q/U. Đúng công thức template: Q = SUM(H:P), "
                   "U = Q − R − S − T. Cột O đã gồm Z:AD (lương TC CN/lễ, trả lại tiền giữ); cột R đã gồm AE:AF "
                   "(bảo hiểm, khấu trừ khác)").font = F_GHI
    r, stt = 4, 0
    tong = defaultdict(float)
    for da in ds:
        cb_nv = canh_bao_nv(da)
        for d in da.luong:
            stt += 1
            g, t = d.gio, d.tien
            vals = [stt, d.ma_nv, d.ten, g.get("X", 0), g.get("D", 0), g.get("OT", 0), g.get("OD", 0),
                    t.get("X", 0), t.get("OT", 0), t.get("D", 0), t.get("OD", 0),
                    d.pc_chuc_vu, d.pc_chuyen_can, d.pc_tham_nien, cot_O(d), d.thuong, d.tong_luong,
                    cot_R(d), d.tien_giu, d.ung_luong, d.thuc_nhan,
                    g.get("OCN", 0), g.get("OCND", 0), g.get("OL", 0), g.get("OLD", 0),
                    t.get("OCN", 0), t.get("OCND", 0), t.get("OL", 0), t.get("OLD", 0),
                    d.tra_lai_giu, d.bhxh, d.khau_tru_khac, da.ma, da.ky, d.thuc_nhan_goc,
                    (d.thuc_nhan - d.thuc_nhan_goc) if d.thuc_nhan_goc is not None else None,
                    d.ngay_cong or None, d.ngay_vao, d.ngay_nghi, d.chi_tiet or None, d.ghi_chu or None,
                    cb_nv.get(d.ma_nv, (None, None))[1]]
            for j, v in enumerate(vals, start=1):
                if isinstance(v, float):
                    v = round(v, 2)
                    if v.is_integer():
                        v = int(v)
                fmt = ("@" if j == 2 else GIO if j in (4, 5, 6, 7, 22, 23, 24, 25, 37) else
                       ("dd/mm/yyyy" if j in (38, 39) else (SO if 8 <= j <= 36 and j not in (33, 34) else None)))
                c = _o(ws, r, j, v, fmt)
                if j == 36 and v not in (None, 0) and abs(v) > 1:
                    c.fill = NEN_DO
                if j == 42 and v:
                    c.fill = NEN_MUC[cb_nv[d.ma_nv][0]]
                    for jj in (2, 3):                     # tô màu mã + tên NV có cảnh báo
                        ws.cell(r, jj).fill = NEN_MUC[cb_nv[d.ma_nv][0]]
                if isinstance(v, (int, float)) and j > 3 and j not in (33, 34) and j <= 37:
                    tong[j] += v
            r += 1
    r += 1
    _o(ws, r, 2, "TỔNG CỘNG", font=F_TD)
    for j, v in tong.items():
        _o(ws, r, j, v, GIO if j in (4, 5, 6, 7, 22, 23, 24, 25) else SO, font=F_TD)
    # tổng theo từng hợp đồng
    if len(ds) > 1:
        r += 2
        ws.cell(r, 2, "TỔNG THEO HỢP ĐỒNG").font = Font(name=FONT, bold=True, size=11)
        r += 1
        for da in ds:
            _o(ws, r, 2, da.ma, font=F_TD); _o(ws, r, 3, f"{len(da.luong)} NV – {da.ky}")
            for j, f in [(4, lambda d: d.gio.get("X", 0)), (5, lambda d: d.gio.get("D", 0)),
                         (6, lambda d: d.gio.get("OT", 0)), (7, lambda d: d.gio.get("OD", 0)),
                         (17, lambda d: d.tong_luong), (18, cot_R), (19, lambda d: d.tien_giu),
                         (20, lambda d: d.ung_luong), (21, lambda d: d.thuc_nhan)]:
                _o(ws, r, j, round(sum(f(d) for d in da.luong), 2), GIO if j <= 7 else SO)
            r += 1
    widths = [5, 13, 24] + [10] * 4 + [13] * 14 + [10] * 4 + [13] * 7 + [8, 9, 14, 11, 9, 11, 11, 60, 30, 50]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[2].height = 45
    ws.freeze_panes = "D3"

    # ================= doanh_thu =================
    ws = wb.create_sheet("doanh_thu")
    r = 2
    tong_tat_ca = 0
    for da in ds:
        ws.merge_cells(start_row=r, end_row=r, start_column=1, end_column=6)
        c = ws.cell(r, 1, f"BẢNG KÊ CHI TIẾT CÔNG NỢ THÁNG {da.ky.replace('T', '')} – {da.ma}")
        c.font, c.alignment = Font(name=FONT, bold=True, size=12), Alignment(horizontal="center")
        r += 1
        nam = da.ky.split("/")[-1] if da.ky and "/" in da.ky else "...."
        ws.merge_cells(start_row=r, end_row=r, start_column=1, end_column=6)
        ws.cell(r, 1, f"Kèm theo hóa đơn .............................. ngày .../.../{nam}").font = F_TH
        r += 2
        for t in ["BÊN A:", "Địa chỉ:", "MST:", "", "BÊN B:", "Địa chỉ:", "MST:"]:
            ws.cell(r, 1, t or None).font = F_TH
            r += 1
        r += 2
        _tieu_de(ws, r, ["STT", "Diễn giải", "ĐVT", "Số lượng", "Đơn giá", "Thành tiền"])
        r += 1
        dien_giai = {"X": "Lương giờ làm ban ngày ngày thường", "OT": "Lương giờ làm ban đêm ngày thường",
                     "D": "Lương giờ tăng ca ngày thường", "OD": "Lương làm thêm giờ ca đêm của ngày làm việc bình thường",
                     "OCN": "Lương giờ tăng ca ngày nghỉ/chủ nhật", "OCND": "Lương giờ tăng ca đêm ngày nghỉ/chủ nhật",
                     "OL": "Lương giờ tăng ca ngày lễ", "OLD": "Lương giờ tăng ca đêm ngày lễ"}
        if da.doanh_thu is not None:
            dong = da.doanh_thu
        else:
            gio = defaultdict(float)
            for d in da.luong:
                for c_, h in d.gio.items():
                    gio[c_] += h
            dong = [(c_, gio[c_], da.don_gia_dv.get(c_, 0), gio[c_] * da.don_gia_dv.get(c_, 0)) for c_ in
                    ["X", "OT", "D", "OD", "OCN", "OCND", "OL", "OLD"]]
        dong = [x for x in dong if (x[1] or x[3]) or x[0] in ("X", "OT", "D", "OD")]
        r_dau, stt, ts, tt = r, 0, 0, 0
        for code, sl, dg, tt_ in dong:
            stt += 1
            _o(ws, r, 1, stt); _o(ws, r, 2, dien_giai.get(code, code)); _o(ws, r, 3, "Giờ" if code in dien_giai else None)
            _o(ws, r, 4, sl, GIO); _o(ws, r, 5, dg, SO); _o(ws, r, 6, round(tt_), SO)
            ts += sl or 0
            tt += tt_ or 0
            r += 1
        r_tong = r
        _o(ws, r, 2, "TỔNG", font=F_TD); _o(ws, r, 4, ts, GIO, F_TD); _o(ws, r, 6, round(tt), SO, F_TD); r += 1
        _o(ws, r, 2, "THUẾ VAT", font=F_TD); _o(ws, r, 6, None); r += 1
        _o(ws, r, 2, "TỔNG TIỀN THANH TOÁN", font=F_TD); _o(ws, r, 6, round(tt), SO, F_TD); r += 1
        tong_tat_ca += tt
        ws.cell(r, 1, "Bằng chữ: " + "." * 143).font = F_TH
        r += 2
        for t in THONG_TIN_CHUYEN_KHOAN:
            ws.cell(r, 1, t).font = F_TH
            r += 1
        # cột bổ sung bên phải (ngoài vùng A:F của template): đối chiếu với file gốc
        if da.doanh_thu_goc is not None:
            _o(ws, r_tong, 8, "Đối chiếu: tổng theo file gốc", font=F_TD)
            c = _o(ws, r_tong, 9, round(da.doanh_thu_goc), SO)
            if abs(da.doanh_thu_goc - tt) > 1:
                c.fill = NEN_DO
        ws.cell(r_tong + 1, 8, "Ghi chú: " + da.ghi_chu_doanh_thu).font = F_GHI
        r += 3
    if len(ds) > 1:
        _o(ws, r, 2, "TỔNG DOANH THU TẤT CẢ DỰ ÁN", font=F_TD); _o(ws, r, 6, round(tong_tat_ca), SO, F_TD)
    for col, w in {"A": 6, "B": 55, "C": 7, "D": 12, "E": 12, "F": 16, "H": 30, "I": 16}.items():
        ws.column_dimensions[col].width = w

    # ================= cham_cong_chi_tiet (dạng dọc cho phân tích) =================
    ws = wb.create_sheet("cham_cong_chi_tiet")
    _tieu_de(ws, 1, ["Hợp đồng / Dự án", "Kỳ lương", "Mã NV", "Tên nhân viên", "Ngày", "Thứ", "Ký hiệu loại giờ",
                     "Loại giờ", "Số giờ", "Ký hiệu ngày (nghỉ/phép...)"])
    r = 2
    for da in ds:
        ten = {d.ma_nv: d.ten for d in da.luong}
        keys = sorted({k for k in da.cham_cong if k[0] in ten} | {k for k in da.ky_hieu_ngay if k[0] in ten},
                      key=lambda k: (list(ten).index(k[0]), k[1]))
        for k in keys:
            g = da.cham_cong.get(k, {})
            kh = da.ky_hieu_ngay.get(k)
            items = [(c, g[c]) for c in TEN_LOAI if g.get(c)] or [(None, 0)]
            for code, h in items:
                vals = [da.ma, da.ky, k[0], ten[k[0]], dt.datetime.combine(k[1], dt.time()), THU[k[1].weekday()],
                        code, TEN_LOAI.get(code, "").replace("\n", " ") if code else None, h, kh]
                for j, v in enumerate(vals, start=1):
                    c = ws.cell(r, j, v)
                    c.font = F_TH
                    if j == 5:
                        c.number_format = "dd/mm/yyyy"
                r += 1
    for col, w in {"A": 9, "B": 10, "C": 12, "D": 24, "E": 11, "F": 9, "G": 8, "H": 22, "I": 8, "J": 18}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{max(r - 1, 1)}"

    # ================= kiem_tra =================
    ws = wb.create_sheet("kiem_tra")
    td = ["Dự án", "File nguồn", "Biểu mẫu", "Kỳ", "Từ ngày", "Đến ngày", "Số NV bảng lương", "Số NV có chấm công",
          "Giờ HC ngày (công)", "Giờ HC ngày (lương)", "Giờ TC ngày (công)", "Giờ TC ngày (lương)",
          "Giờ HC đêm (công)", "Giờ HC đêm (lương)", "Giờ TC đêm (công)", "Giờ TC đêm (lương)",
          "Giờ khác CN/lễ (công)", "Giờ khác CN/lễ (lương)", "Giờ trả lương ngoài chấm công (đào tạo...)",
          "Tổng lương", "Thực nhận (tính lại)", "Thực nhận (file gốc)", "Chênh lệch", "Doanh thu (file chuẩn)", "Doanh thu (file gốc)",
          "Số lỗi", "Số nghi ngờ", "Số lưu ý"]
    _tieu_de(ws, 1, td)
    r = 2
    for da in ds:
        trong = {d.ma_nv for d in da.luong}
        cc = defaultdict(float)
        for (m, _), g in da.cham_cong.items():
            if m in trong:
                for c, h in g.items():
                    cc[c] += h
        ll = defaultdict(float)
        for d in da.luong:
            for c, h in d.gio.items():
                ll[c] += h
        khac = ["OCN", "OCND", "OL", "OLD", "L", "P"]
        tn = sum(d.thuc_nhan for d in da.luong)
        tng = sum(d.thuc_nhan_goc or 0 for d in da.luong)
        dt_chuan = None
        # lấy lại doanh thu đã ghi
        if da.doanh_thu is not None:
            dt_chuan = sum(x[3] or 0 for x in da.doanh_thu)
        else:
            dt_chuan = sum(d.gio.get(c, 0) * da.don_gia_dv.get(c, 0) for d in da.luong for c in LOAI_TIEN)
        vals = [da.ma, os.path.basename(da.file), da.mau, da.ky, da.tu_ngay, da.den_ngay, len(da.luong),
                len({k[0] for k in da.cham_cong} & trong),
                cc["X"], ll["X"], cc["D"], ll["D"], cc["OT"], ll["OT"], cc["OD"], ll["OD"],
                sum(cc[c] for c in khac), sum(ll[c] for c in khac),
                sum(sum(d.gio_ngoai_cham_cong.values()) for d in da.luong),
                sum(d.tong_luong for d in da.luong), tn, tng, tn - tng, dt_chuan, da.doanh_thu_goc]
        dm = dem_muc(da)
        vals += [dm[LOI], dm[NGHI_NGO], dm[LUU_Y]]
        for j, v in enumerate(vals, start=1):
            if isinstance(v, float):
                v = round(v, 2)
            c = _o(ws, r, j, v, "dd/mm/yyyy" if j in (5, 6) else (GIO if 9 <= j <= 19 else (SO if 20 <= j <= 25 else None)))
            if j >= 26 and v:
                c.fill = NEN_MUC[[LOI, NGHI_NGO, LUU_Y][j - 26]]
        ngoai = defaultdict(float)
        for d in da.luong:
            for c_, h in d.gio_ngoai_cham_cong.items():
                ngoai[c_] += h
        bu = {9: ngoai["X"], 11: ngoai["D"], 13: ngoai["OT"], 15: ngoai["OD"], 17: sum(ngoai[c] for c in khac)}
        for a, b in [(9, 10), (11, 12), (13, 14), (15, 16), (17, 18), (21, 22)]:
            if abs((vals[a - 1] or 0) + bu.get(a, 0) - (vals[b - 1] or 0)) > 0.01:
                ws.cell(r, a).fill = ws.cell(r, b).fill = NEN_DO
        r += 1
    r += 2
    if file_loi:
        ws.cell(r, 1, "FILE BỊ BỎ QUA (không đọc được)").font = Font(name=FONT, bold=True, size=11, color="C00000")
        r += 1
        for ten, l in file_loi:
            c = _o(ws, r, 1, "Lỗi"); c.fill = NEN_MUC[LOI]
            _o(ws, r, 2, ten); _o(ws, r, 6, l)
            r += 1
        r += 1
    ws.cell(r, 1, "DANH SÁCH CẢNH BÁO").font = Font(name=FONT, bold=True, size=11)
    r += 1
    for muc, giai_thich in [(LOI, "sai chắc chắn hoặc số liệu không tin được – phải sửa file gốc trước khi dùng"),
                            (NGHI_NGO, "bất thường, cần người kiểm tra lại"),
                            (LUU_Y, "thông tin để biết, không nhất thiết là sai")]:
        c = ws.cell(r, 1, muc)
        c.font, c.fill = F_TD, NEN_MUC[muc]
        ws.cell(r, 2, giai_thich).font = F_GHI
        r += 1
    r += 1
    r_td = r
    _tieu_de(ws, r, ["Mức", "Hợp đồng", "Mã NV", "Tên NV", "Loại", "Nội dung"])
    r += 1
    tat_ca = sorted(((x, da) for da in ds for x in da.van_de), key=lambda t: (THU_TU_MUC[t[0][0]], t[1].ma, t[0][2]))
    for (muc, m, loai, nd), da in tat_ca:
        ten = next((d.ten for d in da.luong if d.ma_nv == m), None) if m else None
        for j, v in enumerate([muc, da.ma, m or None, ten, loai, nd], start=1):
            c = _o(ws, r, j, v)
            if j == 1:
                c.fill = NEN_MUC[muc]
            if j == 6:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    if not tat_ca:
        ws.cell(r, 1, "Không có cảnh báo nào").font = F_TH
    ws.auto_filter.ref = f"A{r_td}:F{max(r - 1, r_td)}"
    for j, w in enumerate([10, 40, 40, 24, 30, 90] + [11] * 13 + [15] * 6 + [9, 11, 9], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.column_dimensions["D"].width = 12
    ws.row_dimensions[1].height = 45
    ws.freeze_panes = "B2"

    wb.save(out_path)
    log(f"Đã ghi: {out_path}")


def LOAI_GIO_PHU(da: DuAn, ma_nv):
    """Các loại giờ ngoài 4 loại chính mà NV có phát sinh (để thêm dòng trong bảng chấm công)."""
    co = set()
    for (m, _), g in da.cham_cong.items():
        if m == ma_nv:
            co |= {c for c, h in g.items() if h}
    d = next((x for x in da.luong if x.ma_nv == ma_nv), None)
    if d:
        co |= {c for c, h in d.gio.items() if h}
    return [c for c in THU_TU_GIO_PHU if c in co]


def chay(files, out_path, log=print, tach_rieng=False, file_loi=None):
    """file_loi: list (tùy chọn) nhận các (tên file, lỗi) bị bỏ qua."""
    ds = []
    file_loi = [] if file_loi is None else file_loi
    for f in files:
        log(f"Đọc: {os.path.basename(f)}")
        try:
            da = doc_file(f, log)
        except Exception as e:
            file_loi.append((os.path.basename(f), str(e)))
            log(f"  !! BỎ QUA file này: {e}")
            continue
        dm = dem_muc(da)
        log(f"  {da.ma} | {da.ky} | {len(da.luong)} NV | thực nhận {sum(d.thuc_nhan for d in da.luong):,.0f}"
            f" (gốc {sum(d.thuc_nhan_goc or 0 for d in da.luong):,.0f}) | "
            f"{dm[LOI]} lỗi, {dm[NGHI_NGO]} nghi ngờ, {dm[LUU_Y]} lưu ý")
        ds.append(da)
    if not ds:
        raise ValueError("Không đọc được file nào: " + "; ".join(f"{t}: {l}" for t, l in file_loi))
    # 2 file cùng mã hợp đồng -> đổi tên để không ghi đè / gộp nhầm
    dem_ma = Counter(da.ma for da in ds)
    for da in ds:
        if dem_ma[da.ma] > 1:
            cu = da.ma
            da.ma = f"{cu} ({os.path.splitext(os.path.basename(da.file))[0][:30]})"
            da.bao(NGHI_NGO, "", "Trùng mã hợp đồng", f"Nhiều file cùng mã '{cu}' – đã đổi thành '{da.ma}'. "
                   "Kiểm tra có tải trùng file, hoặc đặt tên file chứa 'CTY_<mã>'.")
    ds.sort(key=lambda d: d.ma)
    kiem_tra_cheo(ds)
    tong = {k: sum(dem_muc(da)[k] for da in ds) for k in (LOI, NGHI_NGO, LUU_Y)}
    log(f"Kiểm tra: {tong[LOI]} lỗi, {tong[NGHI_NGO]} nghi ngờ, {tong[LUU_Y]} lưu ý (chi tiết ở sheet kiem_tra)")
    if file_loi:
        log(f"Bỏ qua {len(file_loi)} file lỗi: " + ", ".join(t for t, _ in file_loi))
    ghi_file_chuan(ds, out_path, log, file_loi)
    if tach_rieng:
        goc, _ = os.path.splitext(out_path)
        for da in ds:
            ten = re.sub(r'[\\/:*?"<>|() ]+', "_", da.ma).strip("_")
            ghi_file_chuan([da], f"{goc}_{ten}_{da.ky.replace('/', '-')}.xlsx", log)
    return ds


def main(argv=None):
    ap = argparse.ArgumentParser(description="Đồng bộ bảng chấm công/lương nhiều dự án về 1 file chuẩn")
    ap.add_argument("files", nargs="*", help="Các file .xlsx nguồn")
    ap.add_argument("--thu-muc", help="Thư mục chứa các file .xlsx nguồn (lấy tất cả)")
    ap.add_argument("-o", "--output", default=None, help="File kết quả (.xlsx)")
    ap.add_argument("--tach-rieng", action="store_true",
                    help="Xuất thêm 1 file chuẩn riêng cho từng dự án (cột ngày là ngày thật như template)")
    a = ap.parse_args(argv)
    files = list(a.files)
    if a.thu_muc:
        files += [f for f in sorted(glob.glob(os.path.join(a.thu_muc, "*.xlsx")))
                  if not os.path.basename(f).startswith(("~$", "KET_QUA", "Tong_hop"))]
    if not files:
        ap.error("Chưa chọn file nguồn")
    out = a.output or os.path.join(os.path.dirname(os.path.abspath(files[0])),
                                   f"Tong_hop_cham_cong_luong_{dt.datetime.now():%Y%m%d_%H%M}.xlsx")
    chay(files, out, tach_rieng=a.tach_rieng)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
