# -*- coding: utf-8 -*-
"""
Giao diện web chạy trên máy (không cần cài thêm thư viện) cho dong_bo_luong.py.
Chạy:  python app_web.py   -> tự mở http://127.0.0.1:8765
"""
import datetime as dt
import json
import os
import sys
import threading
import traceback
import urllib.parse
import subprocess
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import html
import re
import zipfile

THU_MUC_TOOL = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, THU_MUC_TOOL)
import dong_bo_luong as core  # noqa: E402

THU_MUC_GOC = os.path.dirname(THU_MUC_TOOL)                    # thư mục chứa các file bảng lương
THU_MUC_TAI_LEN = os.path.join(THU_MUC_TOOL, "_tai_len")
THU_MUC_KET_QUA = os.path.join(THU_MUC_GOC, "ket_qua")
CONG = int(os.environ.get("PORT", 8765))

files = []          # [{path, ten, mau, ky, loi}]
job = {"dang_chay": False, "log": [], "ket_qua": None, "loi": None}
khoa = threading.Lock()


def quet(path):
    """Nhận diện nhanh biểu mẫu của file (chỉ đọc danh sách sheet)."""
    tt = {"path": path, "ten": os.path.basename(path), "mau": None, "ky": None, "loi": None,
          "kich_thuoc": os.path.getsize(path)}
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("xl/workbook.xml").decode("utf8")
        sheets = [html.unescape(x) for x in re.findall(r'<sheet\s[^>]*?name="([^"]*)"', xml)]
        tt["mau"] = next((ten for ten, nd, _ in core.ADAPTERS if nd(sheets)), None)
        if not tt["mau"]:
            tt["loi"] = "Chưa có biểu mẫu phù hợp (cần viết thêm bộ đọc cho mẫu này)"
        th = core.ky_tu_ten_file(path)
        tt["ky"] = f"T{th[0]:02d}/{th[1]}" if th else None
        tt["hop_dong"] = core.ma_du_an_tu_file(path, os.path.splitext(tt["ten"])[0])
    except Exception as e:  # file hỏng, không phải xlsx...
        tt["loi"] = f"Không mở được file: {e}"
    return tt


def them_file(path):
    with khoa:
        for i, f in enumerate(files):
            if os.path.basename(f["path"]) == os.path.basename(path):
                files[i] = quet(path)
                return files[i]
        tt = quet(path)
        files.append(tt)
        return tt


def nap_mac_dinh():
    for f in sorted(os.listdir(THU_MUC_GOC)):
        if f.lower().endswith(".xlsx") and not f.startswith(("~$", "Tong_hop")):
            them_file(os.path.join(THU_MUC_GOC, f))


def tom_tat(ds, outputs, file_loi=()):
    kq = []
    for da in ds:
        tn = sum(d.thuc_nhan for d in da.luong)
        goc = sum(d.thuc_nhan_goc or 0 for d in da.luong)
        ten = {d.ma_nv: d.ten for d in da.luong}
        canh_bao = [{"muc": muc, "ma": m, "ten": ten.get(m, ""), "loai": l, "noi_dung": n}
                    for muc, m, l, n in da.van_de]
        dm = core.dem_muc(da)
        dthu = core.doanh_thu_tinh(da)
        kq.append({
            "hop_dong": da.ma, "file": os.path.basename(da.file), "mau": da.mau, "ky": da.ky,
            "tu": da.tu_ngay.strftime("%d/%m/%Y") if da.tu_ngay else "",
            "den": da.den_ngay.strftime("%d/%m/%Y") if da.den_ngay else "",
            "so_nv": len(da.luong), "tong_gio": sum(sum(d.gio.values()) for d in da.luong),
            "tong_luong": sum(d.tong_luong for d in da.luong), "thuc_nhan": tn, "thuc_nhan_goc": goc,
            "lech": tn - goc, "so_nv_lech": sum(1 for d in da.luong if d.thuc_nhan_goc is not None
                                                and abs(d.thuc_nhan - d.thuc_nhan_goc) > 1),
            "doanh_thu": dthu, "doanh_thu_goc": da.doanh_thu_goc, "canh_bao": canh_bao,
            "dem": {"loi": dm[core.LOI], "nghi_ngo": dm[core.NGHI_NGO], "luu_y": dm[core.LUU_Y]},
        })
    return {"hop_dong": kq, "files": outputs, "luc": dt.datetime.now().strftime("%H:%M %d/%m/%Y"),
            "file_loi": [{"ten": t, "loi": l} for t, l in file_loi]}


def chay_dong_bo(tach_rieng):
    def log(s):
        job["log"].append(s)
    try:
        os.makedirs(THU_MUC_KET_QUA, exist_ok=True)
        nguon = [f["path"] for f in files if f["mau"] and not f["loi"]]
        out = os.path.join(THU_MUC_KET_QUA, f"Tong_hop_cham_cong_luong_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx")
        truoc = set(os.listdir(THU_MUC_KET_QUA))
        file_loi = []
        ds = core.chay(nguon, out, log=log, tach_rieng=tach_rieng, file_loi=file_loi)
        moi = sorted(set(os.listdir(THU_MUC_KET_QUA)) - truoc)
        outputs = [os.path.basename(out)] + [f for f in moi if f != os.path.basename(out)]
        job["ket_qua"] = tom_tat(ds, outputs, file_loi)
    except PermissionError:
        job["loi"] = "Không ghi được file kết quả. Có thể file đang mở trong Excel – đóng lại rồi chạy lại."
    except Exception as e:
        job["loi"] = f"{e}" or type(e).__name__
        log(traceback.format_exc())
    finally:
        job["dang_chay"] = False


class MayChu(ThreadingHTTPServer):
    # Trên Windows SO_REUSEADDR cho phép 2 tiến trình cùng chiếm 1 cổng -> tắt để biết cổng bận
    allow_reuse_address = os.name != "nt"
    daemon_threads = True


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _json(self, obj, code=200):
        b = json.dumps(obj, ensure_ascii=False, default=str).encode("utf8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def _cong_khai(self, f):
        return {k: v for k, v in f.items() if k != "path"}

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(u.query)
        if u.path == "/":
            b = open(os.path.join(THU_MUC_TOOL, "giao_dien.html"), "rb").read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(b)))
            self.end_headers()
            self.wfile.write(b)
        elif u.path == "/api/files":
            self._json([self._cong_khai(f) for f in files])
        elif u.path == "/api/status":
            self._json({"dang_chay": job["dang_chay"], "log": job["log"], "ket_qua": job["ket_qua"], "loi": job["loi"]})
        elif u.path == "/tai":
            ten = os.path.basename(q.get("f", [""])[0])
            p = os.path.join(THU_MUC_KET_QUA, ten)
            if not ten or not os.path.isfile(p):
                return self._json({"loi": "Không tìm thấy file"}, 404)
            b = open(p, "rb").read()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename*=UTF-8''" + urllib.parse.quote(ten))
            self.send_header("Content-Length", str(len(b)))
            self.end_headers()
            self.wfile.write(b)
        else:
            self._json({"loi": "Không có trang này"}, 404)

    def do_POST(self):
        u = urllib.parse.urlparse(self.path)
        n = int(self.headers.get("Content-Length") or 0)
        if u.path == "/api/upload":
            ten = os.path.basename(urllib.parse.unquote(self.headers.get("X-Ten-File", "")))
            if not ten.lower().endswith(".xlsx"):
                self.rfile.read(n)
                return self._json({"loi": f"'{ten}' không phải file .xlsx"}, 400)
            os.makedirs(THU_MUC_TAI_LEN, exist_ok=True)
            p = os.path.join(THU_MUC_TAI_LEN, ten)
            with open(p, "wb") as fo:
                con = n
                while con > 0:
                    chunk = self.rfile.read(min(con, 1 << 20))
                    if not chunk:
                        break
                    fo.write(chunk)
                    con -= len(chunk)
            return self._json(self._cong_khai(them_file(p)))
        body = json.loads(self.rfile.read(n) or b"{}")
        if u.path == "/api/xoa":
            with khoa:
                files[:] = [f for f in files if f["ten"] != body.get("ten")]
            return self._json({"ok": True})
        if u.path == "/api/xoa_het":
            with khoa:
                files.clear()
            return self._json({"ok": True})
        if u.path == "/api/nap_mac_dinh":
            nap_mac_dinh()
            return self._json({"ok": True})
        if u.path == "/api/chay":
            if job["dang_chay"]:
                return self._json({"loi": "Đang đồng bộ, vui lòng chờ"}, 409)
            if not any(f["mau"] and not f["loi"] for f in files):
                return self._json({"loi": "Chưa có file nào đọc được"}, 400)
            job.update(dang_chay=True, log=[], ket_qua=None, loi=None)
            threading.Thread(target=chay_dong_bo, args=(bool(body.get("tach_rieng")),), daemon=True).start()
            return self._json({"ok": True})
        if u.path == "/api/mo_thu_muc":
            os.makedirs(THU_MUC_KET_QUA, exist_ok=True)
            if hasattr(os, "startfile"):
                os.startfile(THU_MUC_KET_QUA)
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", THU_MUC_KET_QUA])
            return self._json({"ok": True, "thu_muc": THU_MUC_KET_QUA})
        self._json({"loi": "Không có API này"}, 404)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    srv = None
    for cong in range(CONG, CONG + 20):          # cổng bận thì thử cổng kế tiếp
        try:
            srv = MayChu(("127.0.0.1", cong), Handler)
            break
        except OSError:
            continue
    if srv is None:
        sys.exit(f"Không mở được cổng nào từ {CONG} đến {CONG + 19}")
    nap_mac_dinh()
    url = f"http://127.0.0.1:{cong}"
    print(f"Giao diện đang chạy tại {url}  (đóng cửa sổ này để tắt)", flush=True)
    if "--khong-mo" not in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
