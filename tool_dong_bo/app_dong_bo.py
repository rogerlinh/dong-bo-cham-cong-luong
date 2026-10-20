# -*- coding: utf-8 -*-
"""Mini app (giao diện) cho dong_bo_luong.py – chọn file nguồn, bấm Đồng bộ."""
import datetime as dt
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dong_bo_luong as core  # noqa: E402


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Đồng bộ bảng chấm công / lương về file chuẩn")
        self.geometry("860x560")
        self.minsize(700, 450)
        self.q = queue.Queue()
        self.files = []

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Button(top, text="Thêm file...", command=self.them_file).pack(side="left")
        ttk.Button(top, text="Thêm cả thư mục...", command=self.them_thu_muc).pack(side="left", padx=6)
        ttk.Button(top, text="Xóa chọn", command=self.xoa_chon).pack(side="left")
        ttk.Button(top, text="Xóa hết", command=self.xoa_het).pack(side="left", padx=6)

        mid = ttk.Frame(self, padding=(10, 0))
        mid.pack(fill="both", expand=True)
        ttk.Label(mid, text="File nguồn (mỗi file 1 dự án, biểu mẫu được nhận diện tự động):").pack(anchor="w")
        self.lb = tk.Listbox(mid, height=7, selectmode="extended")
        self.lb.pack(fill="x")

        out = ttk.Frame(self, padding=10)
        out.pack(fill="x")
        ttk.Label(out, text="File kết quả:").pack(side="left")
        self.out_var = tk.StringVar()
        ttk.Entry(out, textvariable=self.out_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(out, text="Chọn...", command=self.chon_out).pack(side="left")

        opt = ttk.Frame(self, padding=(10, 0))
        opt.pack(fill="x")
        self.tach = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="Xuất thêm file riêng cho từng dự án", variable=self.tach).pack(side="left")
        self.btn = ttk.Button(opt, text="▶  ĐỒNG BỘ", command=self.chay)
        self.btn.pack(side="right")
        self.btn_mo = ttk.Button(opt, text="Mở file kết quả", command=self.mo_ket_qua, state="disabled")
        self.btn_mo.pack(side="right", padx=6)

        logf = ttk.Frame(self, padding=10)
        logf.pack(fill="both", expand=True)
        self.log = tk.Text(logf, height=12, font=("Consolas", 9))
        self.log.pack(fill="both", expand=True)
        self.after(100, self.poll)

        mac_dinh = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self._them([os.path.join(mac_dinh, f) for f in sorted(os.listdir(mac_dinh))
                    if f.lower().endswith(".xlsx") and not f.startswith(("~$", "Tong_hop"))])

    # ---- chọn file
    def _them(self, paths):
        for p in paths:
            p = os.path.abspath(p)
            if p not in self.files:
                self.files.append(p)
                self.lb.insert("end", os.path.basename(p))
        if self.files and not self.out_var.get():
            self.out_var.set(os.path.join(os.path.dirname(self.files[0]),
                                          f"Tong_hop_cham_cong_luong_{dt.datetime.now():%Y%m%d_%H%M}.xlsx"))

    def them_file(self):
        self._them(filedialog.askopenfilenames(title="Chọn file bảng lương", filetypes=[("Excel", "*.xlsx")]))

    def them_thu_muc(self):
        d = filedialog.askdirectory(title="Chọn thư mục chứa file bảng lương")
        if d:
            self._them([os.path.join(d, f) for f in sorted(os.listdir(d))
                        if f.lower().endswith(".xlsx") and not f.startswith(("~$", "Tong_hop"))])

    def xoa_chon(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i)
            del self.files[i]

    def xoa_het(self):
        self.lb.delete(0, "end")
        self.files.clear()

    def chon_out(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                         initialfile=os.path.basename(self.out_var.get() or "Tong_hop.xlsx"))
        if p:
            self.out_var.set(p)

    # ---- chạy
    def ghi(self, s):
        self.q.put(s)

    def poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m == "__XONG__":
                    self.btn.config(state="normal")
                    self.btn_mo.config(state="normal")
                    messagebox.showinfo("Xong", "Đã đồng bộ xong. Xem sheet 'kiem_tra' để đối chiếu.")
                elif m == "__LOI__":
                    self.btn.config(state="normal")
                else:
                    self.log.insert("end", m + "\n")
                    self.log.see("end")
        except queue.Empty:
            pass
        self.after(100, self.poll)

    def chay(self):
        if not self.files:
            messagebox.showwarning("Thiếu file", "Chưa chọn file nguồn")
            return
        out = self.out_var.get().strip()
        if not out:
            messagebox.showwarning("Thiếu file kết quả", "Chưa chọn nơi lưu file kết quả")
            return
        self.btn.config(state="disabled")
        self.log.delete("1.0", "end")

        def work():
            try:
                core.chay(list(self.files), out, log=self.ghi, tach_rieng=self.tach.get())
                self.q.put("__XONG__")
            except PermissionError:
                self.ghi("LỖI: Không ghi được file kết quả (file đang mở trong Excel?). Đóng file rồi chạy lại.")
                self.q.put("__LOI__")
            except Exception:
                self.ghi("LỖI:\n" + traceback.format_exc())
                self.q.put("__LOI__")

        threading.Thread(target=work, daemon=True).start()

    def mo_ket_qua(self):
        p = self.out_var.get()
        if os.path.exists(p):
            os.startfile(p)


if __name__ == "__main__":
    App().mainloop()
