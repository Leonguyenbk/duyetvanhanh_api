# -*- coding: utf-8 -*-
"""
Tool tự động cập nhật trạng thái gói tin đồng bộ thi công (KHÔNG có dữ liệu vận hành)
- Tkinter UI: nhập username/password đăng nhập lấy token+session, nhập "dotBanGiaoNId"
- Luồng: lặp lại tra cứu 20 bản ghi có trangThai = 0 → lọc bản ghi có
  ThongTinDangKyVanHanh RỖNG → gửi update (trangThai = 1) từng bản ghi →
  tiếp tục tra cứu lại (start=0) cho đến khi không còn bản ghi trangThai=0
  nào xử lý được nữa (đã hết hoặc chỉ còn bản ghi có dữ liệu vận hành → bỏ qua)
Cài đặt: pip install selenium webdriver-manager requests
"""

import os
import json
import time
import threading
from datetime import datetime

import requests

import tkinter as tk
from tkinter import ttk, messagebox

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# ============================ CẤU HÌNH ============================

BASE_URL = "https://dla.mplis.gov.vn"

# Trang dùng để đăng nhập + lấy __RequestVerificationToken
REFERER_LOGIN = f"{BASE_URL}/dc/DonDangKy/KeKhaiDangKyV2"
# Referer dùng khi gọi các API TichHopDongBoDuLieuAjax
REFERER_API = f"{BASE_URL}/dc/"

URL_KIEM_TRA = f"{BASE_URL}/dc/TichHopDongBoDuLieuAjax/KiemTraDuLieuChuyenDoi"
URL_CAP_NHAT = f"{BASE_URL}/dc/TichHopDongBoDuLieuAjax/UpdateStatusGoiTinDongBoThiCong"

SO_BAN_GHI_MOI_TRANG = 20
TIMEOUT = 120


# ============================ HELPER ============================

def lay_token_tu_trang(driver):
    js = """
    return (
        document.querySelector('input[name="__RequestVerificationToken"]')?.value ||
        document.querySelector('input[name="__requestverificationtoken"]')?.value ||
        document.querySelector('meta[name="__RequestVerificationToken"]')?.content ||
        document.querySelector('meta[name="__requestverificationtoken"]')?.content ||
        document.querySelector('meta[name="RequestVerificationToken"]')?.content ||
        ''
    );
    """
    return driver.execute_script(js)


# ============================ CORE API ============================

class MplisSyncClient:
    def __init__(self, log_fn):
        self.log = log_fn
        self.session = None
        self.driver = None

    # ---------- login ----------
    def open_browser_and_fill_login(self, username, password):
        options = Options()
        options.add_argument("--start-maximized")
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.get(REFERER_LOGIN)
        time.sleep(2)

        try:
            inputs = self.driver.find_elements(By.CSS_SELECTOR, "input")
            user_box = None
            pass_box = None
            for inp in inputs:
                typ = (inp.get_attribute("type") or "").lower()
                if not user_box and typ in ["text", "email"]:
                    user_box = inp
                if not pass_box and typ == "password":
                    pass_box = inp
            if user_box and pass_box:
                user_box.clear()
                user_box.send_keys(username)
                pass_box.clear()
                pass_box.send_keys(password)
                pass_box.send_keys(Keys.ENTER)
                self.log("Đã điền thông tin đăng nhập, chờ trang load...")
        except Exception as e:
            self.log(f"Không tự điền được form login ({e}), hãy đăng nhập tay trên Chrome.")

    def build_session_from_browser(self):
        if not self.driver:
            raise RuntimeError("Chưa mở trình duyệt.")

        token = lay_token_tu_trang(self.driver)
        if not token:
            raise RuntimeError("Không lấy được token. Kiểm tra đã đăng nhập và đang ở đúng trang chưa.")

        session = requests.Session()
        user_agent = self.driver.execute_script("return navigator.userAgent;")

        session.headers.update({
            "User-Agent": user_agent,
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Content-Type": "application/json; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": BASE_URL,
            "Referer": REFERER_API,
            "__requestverificationtoken": token,
        })

        for c in self.driver.get_cookies():
            session.cookies.set(
                name=c["name"],
                value=c["value"],
                domain=c.get("domain"),
                path=c.get("path", "/"),
            )

        self.session = session
        self.log("✅ Đã lấy session + token thành công.")

    def close_browser(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    # ---------- API ----------
    def kiem_tra_du_lieu_chuyen_doi(self, dot_ban_giao_id, start=0, length=SO_BAN_GHI_MOI_TRANG):
        payload = {
            "dotBanGiaoNId": dot_ban_giao_id,
            "requestQueries": {
                "soPhatHanh": "",
                "soThuTuThua": "",
                "soHieuToBanDo": "",
                "hoTenChu": "",
                "soGiayTo": "",
                "trangThai": "0",
            },
            "start": start,
            "length": length,
        }

        res = self.session.post(
            URL_KIEM_TRA, json=payload, timeout=TIMEOUT, allow_redirects=False
        )

        if res.status_code == 302 or res.headers.get("Location"):
            raise RuntimeError(
                f"Bị redirect (status {res.status_code}) — có thể phiên đăng nhập đã hết hạn."
            )
        if res.status_code == 404:
            raise RuntimeError(f"URL không tồn tại (404): {URL_KIEM_TRA}")

        try:
            js = res.json()
        except Exception:
            raise RuntimeError(f"API kiểm tra không trả JSON (status {res.status_code}): {res.text[:500]}")

        return {
            "items": js.get("value") or [],
            "total": js.get("recordsTotal"),
        }

    def cap_nhat_goi_tin(self, goi_tin_id):
        payload = {
            "goiTinDongBoNId": goi_tin_id,
            "actionProcessSyncDataThiCong": 1,
            "processSyncDatas": [],
            "trangThai": 1,
        }

        res = self.session.post(
            URL_CAP_NHAT, json=payload, timeout=TIMEOUT, allow_redirects=False
        )

        if res.status_code == 302 or res.headers.get("Location"):
            raise RuntimeError(
                f"Bị redirect (status {res.status_code}) — có thể phiên đăng nhập đã hết hạn."
            )
        if res.status_code == 404:
            raise RuntimeError(f"URL không tồn tại (404): {URL_CAP_NHAT}")

        try:
            js = res.json()
        except Exception:
            raise RuntimeError(f"API cập nhật không trả JSON (status {res.status_code}): {res.text[:500]}")

        if not js.get("success"):
            raise RuntimeError("Cập nhật lỗi: " + str(js)[:500])

        return js


# ============================ TKINTER UI ============================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Cập nhật gói tin đồng bộ thi công MPLIS")
        root.geometry("860x600")

        self.client = MplisSyncClient(self.log)
        self.running = False
        self.stop_flag = False

        frm = ttk.Frame(root, padding=10)
        frm.pack(fill="x")

        ttk.Label(frm, text="Username:").grid(row=0, column=0, sticky="w")
        self.ent_user = ttk.Entry(frm, width=30)
        self.ent_user.grid(row=0, column=1, sticky="w", padx=5, pady=3)

        ttk.Label(frm, text="Password:").grid(row=0, column=2, sticky="w")
        self.ent_pass = ttk.Entry(frm, width=30, show="*")
        self.ent_pass.grid(row=0, column=3, sticky="w", padx=5, pady=3)

        ttk.Label(frm, text="dotBanGiaoNId:").grid(row=1, column=0, sticky="w")
        self.ent_dot_ban_giao = ttk.Entry(frm, width=45)
        self.ent_dot_ban_giao.grid(row=1, column=1, columnspan=2, sticky="w", padx=5, pady=3)

        ttk.Label(frm, text="Giới hạn số lượt (test, để trống = chạy hết):").grid(row=2, column=0, sticky="w")
        self.ent_gioi_han_lap = ttk.Entry(frm, width=10)
        self.ent_gioi_han_lap.grid(row=2, column=1, sticky="w", padx=5, pady=3)

        btn_frm = ttk.Frame(root, padding=(10, 0))
        btn_frm.pack(fill="x")

        self.btn_login = ttk.Button(btn_frm, text="1. Mở Chrome đăng nhập", command=self.mo_chrome)
        self.btn_login.pack(side="left", padx=5)

        self.btn_confirm = ttk.Button(
            btn_frm, text="2. Đã đăng nhập xong → Lấy session", command=self.lay_session, state="disabled"
        )
        self.btn_confirm.pack(side="left", padx=5)

        self.btn_run = ttk.Button(btn_frm, text="3. Bắt đầu xử lý", command=self.chay, state="disabled")
        self.btn_run.pack(side="left", padx=5)

        self.btn_stop = ttk.Button(btn_frm, text="Dừng", command=self.dung, state="disabled")
        self.btn_stop.pack(side="left", padx=5)

        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(8, 0))

        status_frm = ttk.Frame(root)
        status_frm.pack(fill="x", padx=10)
        self.lbl_status = ttk.Label(status_frm, text="Chưa chạy")
        self.lbl_status.pack(side="left")
        self.lbl_so_lan_lap = ttk.Label(status_frm, text="", foreground="blue")
        self.lbl_so_lan_lap.pack(side="right")

        self.txt = tk.Text(root, wrap="word", height=26)
        self.txt.pack(fill="both", expand=True, padx=10, pady=8)
        scroll = ttk.Scrollbar(self.txt, command=self.txt.yview)
        self.txt.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")

        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- UI helpers ----------
    def log(self, msg):
        def _append():
            self.txt.insert("end", f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n")
            self.txt.see("end")
        self.root.after(0, _append)

    def set_status(self, msg):
        self.root.after(0, lambda: self.lbl_status.config(text=msg))

    def set_so_lan_lap(self, so_lan, gioi_han=None):
        text = f"Lượt: {so_lan}" + (f" / {gioi_han}" if gioi_han else "")
        self.root.after(0, lambda: self.lbl_so_lan_lap.config(text=text))

    # ---------- actions ----------
    def mo_chrome(self):
        username = self.ent_user.get().strip()
        password = self.ent_pass.get()
        if not username or not password:
            messagebox.showwarning("Thiếu thông tin", "Nhập username và password trước.")
            return

        self.btn_login.config(state="disabled")

        def _work():
            try:
                self.log("Đang mở Chrome...")
                self.client.open_browser_and_fill_login(username, password)
                self.log("Chrome đã mở. Hoàn tất đăng nhập (OTP, captcha... nếu có) rồi bấm nút 2.")
                self.root.after(0, lambda: self.btn_confirm.config(state="normal"))
            except Exception as e:
                self.log(f"❌ Lỗi mở Chrome: {e}")
                self.root.after(0, lambda: self.btn_login.config(state="normal"))

        threading.Thread(target=_work, daemon=True).start()

    def lay_session(self):
        def _work():
            try:
                self.client.build_session_from_browser()
                self.root.after(0, lambda: self.btn_run.config(state="normal"))
            except Exception as e:
                self.log(f"❌ {e}")

        threading.Thread(target=_work, daemon=True).start()

    def dung(self):
        self.stop_flag = True
        self.log("⏸ Đã yêu cầu dừng, sẽ dừng sau khi xử lý xong lượt hiện tại...")

    def chay(self):
        if self.running:
            return

        dot_ban_giao_id = self.ent_dot_ban_giao.get().strip()
        if not dot_ban_giao_id:
            messagebox.showwarning("Thiếu thông tin", "Nhập dotBanGiaoNId trước.")
            return

        gioi_han_nhap = self.ent_gioi_han_lap.get().strip()
        gioi_han_lap = None
        if gioi_han_nhap:
            if not gioi_han_nhap.isdigit() or int(gioi_han_nhap) <= 0:
                messagebox.showwarning("Sai giá trị", "Giới hạn số lượt phải là số nguyên dương.")
                return
            gioi_han_lap = int(gioi_han_nhap)

        dong_gioi_han = f"\nCHỈ CHẠY TỐI ĐA {gioi_han_lap} LƯỢT (chế độ test)." if gioi_han_lap else ""
        if not messagebox.askyesno(
            "Xác nhận",
            f"Sẽ tự động tra cứu (mỗi lần {SO_BAN_GHI_MOI_TRANG} bản ghi, trạng thái = 0) "
            f"và cập nhật (trangThai = 1) cho các gói tin KHÔNG có dữ liệu vận hành "
            f"thuộc dotBanGiaoNId:\n{dot_ban_giao_id}"
            f"{dong_gioi_han}\n\n"
            f"Lặp lại cho đến khi không còn bản ghi nào xử lý được nữa. Tiếp tục?",
        ):
            return

        self.running = True
        self.stop_flag = False
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progress.config(value=0, maximum=100)
        self.set_so_lan_lap(0, gioi_han_lap)

        threading.Thread(target=self._run_loop, args=(dot_ban_giao_id, gioi_han_lap), daemon=True).start()

    def _run_loop(self, dot_ban_giao_id, gioi_han_lap=None):
        da_bo_qua = set()   # goiTinDongBoNId có dữ liệu vận hành → không update, không truy lại
        thanh_cong = []
        that_bai = []
        so_lan_lap = 0
        tong_ban_dau = None   # recordsTotal ở lượt đầu tiên, dùng làm mốc 100% cho thanh tiến độ

        while not self.stop_flag:
            so_lan_lap += 1
            self.set_so_lan_lap(so_lan_lap, gioi_han_lap)
            self.set_status(f"Đang tra cứu lượt {so_lan_lap}...")
            self.log(f"--- Lượt {so_lan_lap}: tra cứu {SO_BAN_GHI_MOI_TRANG} bản ghi (trangThai=0) ---")

            try:
                ket_qua = self.client.kiem_tra_du_lieu_chuyen_doi(dot_ban_giao_id, start=0, length=SO_BAN_GHI_MOI_TRANG)
            except Exception as e:
                self.log(f"❌ Lỗi tra cứu: {e}")
                break

            danh_sach = ket_qua["items"]
            recordsTotal = ket_qua["total"]

            if tong_ban_dau is None and recordsTotal:
                tong_ban_dau = recordsTotal
                self.root.after(0, lambda t=tong_ban_dau: self.progress.config(maximum=t, value=0))

            self.log(f"   → Nhận được {len(danh_sach)} bản ghi, còn lại trangThai=0: {recordsTotal}")

            # recordsTotal = 0 (hoặc không còn bản ghi) → đã xử lý hết
            if not danh_sach or recordsTotal == 0:
                self.log("✅ Không còn bản ghi nào có trangThai=0. Hoàn tất.")
                break

            if tong_ban_dau:
                da_xu_ly_uoc_tinh = max(0, tong_ban_dau - recordsTotal)
                self.root.after(0, lambda v=da_xu_ly_uoc_tinh: self.progress.config(value=v))

            # Lọc các bản ghi CHƯA xử lý (chưa từng bị bỏ qua) trong lượt này
            can_xu_ly = []
            for item in danh_sach:
                goi_tin_id = item.get("goiTinDongBoNId") or item.get("GoiTinDongBoNId")
                if not goi_tin_id or goi_tin_id in da_bo_qua:
                    continue
                can_xu_ly.append((goi_tin_id, item))

            if not can_xu_ly:
                self.log(
                    "✅ Toàn bộ bản ghi còn lại (trangThai=0) đều đã có dữ liệu vận hành "
                    "(đã bỏ qua trước đó). Dừng để tránh lặp vô hạn."
                )
                break

            co_cap_nhat_lan_nay = False

            for goi_tin_id, item in can_xu_ly:
                if self.stop_flag:
                    self.log("⏹ Đã dừng theo yêu cầu.")
                    break

                van_hanh = item.get("ThongTinDangKyVanHanh") or []
                if van_hanh:
                    da_bo_qua.add(goi_tin_id)
                    self.log(f"   ⏭ Bỏ qua {goi_tin_id}: đã có dữ liệu vận hành ({len(van_hanh)} mục)")
                    continue

                try:
                    self.client.cap_nhat_goi_tin(goi_tin_id)
                    thanh_cong.append(goi_tin_id)
                    co_cap_nhat_lan_nay = True
                    self.log(f"   ✅ Cập nhật OK: {goi_tin_id} (còn ~{recordsTotal - len(thanh_cong) - len(da_bo_qua)} bản ghi)")
                except Exception as e:
                    that_bai.append({"goiTinDongBoNId": goi_tin_id, "loi": str(e)})
                    da_bo_qua.add(goi_tin_id)  # tránh lặp vô hạn với bản ghi luôn lỗi
                    self.log(f"   ❌ Lỗi cập nhật {goi_tin_id}: {e}")

            if self.stop_flag:
                break

            if not co_cap_nhat_lan_nay:
                self.log("✅ Lượt này không cập nhật được bản ghi nào mới. Dừng để tránh lặp vô hạn.")
                break

            self.set_status(
                f"Đã xử lý {len(thanh_cong)} thành công, {len(that_bai)} lỗi. "
                f"Còn khoảng {recordsTotal} bản ghi trangThai=0. Tiếp tục lượt sau..."
            )

            if gioi_han_lap and so_lan_lap >= gioi_han_lap:
                self.log(f"⏸ Đã chạy đủ {gioi_han_lap} lượt theo giới hạn test. Dừng lại.")
                break

        self.log("=" * 60)
        self.log(f"HOÀN THÀNH sau {so_lan_lap} lượt tra cứu" + (f" (giới hạn: {gioi_han_lap})" if gioi_han_lap else ""))
        self.log(f"Thành công: {len(thanh_cong)}")
        self.log(f"Lỗi: {len(that_bai)}")
        self.log(f"Bỏ qua (đã có dữ liệu vận hành): {len(da_bo_qua) - len(that_bai)}")
        if that_bai:
            self.log("Danh sách lỗi:")
            self.log(json.dumps(that_bai, ensure_ascii=False, indent=2))

        self.set_status(f"Hoàn tất: {len(thanh_cong)} OK, {len(that_bai)} lỗi")
        self.set_so_lan_lap(so_lan_lap, gioi_han_lap)
        self.running = False
        self.root.after(0, lambda: self.progress.config(value=self.progress["maximum"]))
        self.root.after(0, lambda: (self.btn_run.config(state="normal"), self.btn_stop.config(state="disabled")))

    def on_close(self):
        if self.running and not messagebox.askyesno("Đang chạy", "Đang xử lý, thoát luôn?"):
            return
        self.client.close_browser()
        self.root.destroy()


def kiem_tra_dieu_kien_khoi_dong(root):
    """Hỏi 2 câu trước khi cho vào app. Trả về True nếu được phép dùng tiếp."""
    root.withdraw()  # ẩn cửa sổ chính trong lúc hỏi

    beo = messagebox.askyesno("Câu hỏi 1", "Anh Tuấn có béo không?")
    if beo:
        messagebox.showerror("Kết thúc", "Bạn đã làm tổn thương anh Tuấn")
        return False

    dep_trai = messagebox.askyesno("Câu hỏi 2", "Anh Tuấn có đẹp trai không?")
    if not dep_trai:
        messagebox.showerror("Kết thúc", "Bạn không trung thực")
        return False

    messagebox.showinfo("OK", "Cảm ơn bạn đã trung thực")
    root.deiconify()  # hiện lại cửa sổ chính
    return True


if __name__ == "__main__":
    root = tk.Tk()
    if not kiem_tra_dieu_kien_khoi_dong(root):
        root.destroy()
    else:
        app = App(root)
        root.mainloop()