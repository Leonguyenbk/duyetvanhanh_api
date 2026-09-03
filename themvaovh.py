# -*- coding: utf-8 -*-
"""
Tool tự động cập nhật trạng thái gói tin đồng bộ thi công MPLIS
- Đăng nhập: qua Chrome Selenium hoặc dán trực tiếp Cookie + __RequestVerificationToken.
- Tra cứu: gửi song song 3-4 request 1 lần đến KiemTraDuLieuChuyenDoi với length = 5
  (vì payload từ server rất nặng, ~1MB/5 bản ghi, nên chia nhỏ và gửi đồng thời để tối ưu tốc độ).
- Xử lý: kiểm tra goiTinDongBoNId nào có trạng thái 0 (trangThaiGoiTinDongBo == 0) thì tiến hành update.
- Cập nhật: gọi UpdateStatusGoiTinDongBoThiCong với trangThai = 1 theo lô đa luồng.
- Cài đặt: pip install selenium webdriver-manager requests
"""

import os
import json
import time
import threading
import concurrent.futures
from datetime import datetime
from http.cookies import SimpleCookie

import requests
from requests.adapters import HTTPAdapter

import tkinter as tk
from tkinter import ttk, messagebox

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# ============================ CẤU HÌNH MẶC ĐỊNH ============================

BASE_URL = "https://dla.mplis.gov.vn"

REFERER_LOGIN = f"{BASE_URL}/dc/DonDangKy/KeKhaiDangKyV2"
REFERER_API = f"{BASE_URL}/dc/"

URL_KIEM_TRA = f"{BASE_URL}/dc/TichHopDongBoDuLieuAjax/KiemTraDuLieuChuyenDoi"
URL_CAP_NHAT = f"{BASE_URL}/dc/TichHopDongBoDuLieuAjax/UpdateStatusGoiTinDongBoThiCong"

DOT_BAN_GIAO_MAC_DINH = "7497d062-a144-4b17-b945-16e09d7b6c93"
SO_BAN_GHI_MOI_REQUEST = 5     # length mặc định = 5 theo yêu cầu
SO_REQUEST_DONG_THOI = 4       # Gửi 3-4 request 1 lần để lấy nhiều thông tin
SO_WORKER_UPDATE = 4           # Số worker cập nhật đồng thời
TIMEOUT = 120
SO_LAN_THU_LAI = 3


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


# ============================ CORE API CLIENT ============================

class MplisSyncClient:
    def __init__(self, log_fn):
        self.log = log_fn
        self.session = None
        self.driver = None

    # ---------- Session Setup ----------
    def _tao_session_base(self, token):
        session = requests.Session()
        adapter = HTTPAdapter(pool_connections=30, pool_maxsize=30, max_retries=2)
        session.mount("https://", adapter)
        session.mount("http://", adapter)

        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Content-Type": "application/json; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": BASE_URL,
            "Referer": REFERER_API,
            "__requestverificationtoken": token,
        })
        return session

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

        session = self._tao_session_base(token)
        try:
            user_agent = self.driver.execute_script("return navigator.userAgent;")
            if user_agent:
                session.headers["User-Agent"] = user_agent
        except Exception:
            pass

        for c in self.driver.get_cookies():
            session.cookies.set(
                name=c["name"],
                value=c["value"],
                domain=c.get("domain"),
                path=c.get("path", "/"),
            )

        self.session = session
        self.log("✅ Đã lấy session + token từ trình duyệt thành công.")

    def build_session_from_manual(self, cookie_str, token):
        cookie_str = (cookie_str or "").strip()
        token = (token or "").strip()
        if not cookie_str:
            raise ValueError("Cookie không được để trống.")
        if not token:
            raise ValueError("Token không được để trống.")

        session = self._tao_session_base(token)
        parsed = SimpleCookie()
        parsed.load(cookie_str)
        if not parsed:
            raise ValueError("Không đọc được cookie nào từ chuỗi nhập vào.")

        for name, morsel in parsed.items():
            session.cookies.set(
                name=name,
                value=morsel.value,
                domain="dla.mplis.gov.vn",
                path="/",
            )

        self.session = session
        self.log("✅ Đã thiết lập session từ Cookie & Token nhập tay.")

    def close_browser(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    # ---------- API Tra Cứu (Kiểm Tra Dữ Liệu) ----------
    def kiem_tra_du_lieu_chuyen_doi(self, dot_ban_giao_id, start=0, length=SO_BAN_GHI_MOI_REQUEST, retry=SO_LAN_THU_LAI):
        """
        Gửi payload kiểm tra dữ liệu chuyển đổi:
        {
            "dotBanGiaoNId": "...",
            "length": 5,
            "requestQueries": {
                "soPhatHanh": "",
                "soThuTuThua": "",
                "soHieuToBanDo": "",
                "hoTenChu": "",
                "soGiayTo": "",
                "trangThai": "0"
            },
            "start": 0
        }
        """
        if not self.session:
            raise RuntimeError("Chưa có session. Vui lòng đăng nhập hoặc dán cookie trước.")

        payload = {
            "dotBanGiaoNId": dot_ban_giao_id,
            "length": length,
            "requestQueries": {
                "soPhatHanh": "",
                "soThuTuThua": "",
                "soHieuToBanDo": "",
                "hoTenChu": "",
                "soGiayTo": "",
                "trangThai": "0",
            },
            "start": start,
        }

        loi_cuoi = None
        for attempt in range(1, retry + 1):
            try:
                res = self.session.post(
                    URL_KIEM_TRA,
                    json=payload,
                    timeout=TIMEOUT,
                    allow_redirects=False,
                )

                if res.status_code in (301, 302) or res.headers.get("Location"):
                    raise RuntimeError(
                        f"Bị redirect (HTTP {res.status_code}) — phiên đăng nhập có thể đã hết hạn."
                    )
                if res.status_code == 404:
                    raise RuntimeError(f"URL không tồn tại (404): {URL_KIEM_TRA}")
                if res.status_code != 200:
                    raise RuntimeError(f"HTTP {res.status_code}: {res.text[:300]}")

                try:
                    js = res.json()
                except Exception:
                    raise RuntimeError(f"API kiểm tra không trả JSON (HTTP {res.status_code}): {res.text[:300]}")

                return {
                    "start": start,
                    "items": js.get("value") or [],
                    "total": js.get("recordsTotal"),
                    "success": js.get("success", False),
                }
            except Exception as e:
                loi_cuoi = e
                if attempt < retry:
                    time.sleep(1.0)
                else:
                    raise loi_cuoi

    def tra_cuu_nhieu_request_song_song(self, dot_ban_giao_id, start_offsets, length=SO_BAN_GHI_MOI_REQUEST):
        """
        Gửi đồng thời 3-4 request KiemTraDuLieuChuyenDoi với các offset start khác nhau.
        """
        ket_qua_danh_sach = []
        tong_records = None

        with concurrent.futures.ThreadPoolExecutor(max_workers=len(start_offsets)) as executor:
            future_to_start = {
                executor.submit(self.kiem_tra_du_lieu_chuyen_doi, dot_ban_giao_id, s, length): s
                for s in start_offsets
            }
            for future in concurrent.futures.as_completed(future_to_start):
                s = future_to_start[future]
                try:
                    res = future.result()
                    ket_qua_danh_sach.append(res)
                    if res.get("total") is not None:
                        tong_records = res["total"]
                except Exception as e:
                    self.log(f"   ❌ Lỗi tra cứu start={s}: {e}")

        # Sắp xếp lại theo start để duy trì thứ tự danh sách
        ket_qua_danh_sach.sort(key=lambda x: x.get("start", 0))

        tat_ca_items = []
        for r in ket_qua_danh_sach:
            tat_ca_items.extend(r.get("items") or [])

        return {
            "items": tat_ca_items,
            "total": tong_records,
            "sub_results": ket_qua_danh_sach,
        }

    # ---------- API Cập Nhật Gói Tin ----------
    def cap_nhat_goi_tin(self, goi_tin_id, retry=SO_LAN_THU_LAI):
        """
        Gửi payload cập nhật gói tin:
        {
            "goiTinDongBoNId": goi_tin_id,
            "actionProcessSyncDataThiCong": 1,
            "processSyncDatas": [],
            "trangThai": 1
        }
        """
        if not self.session:
            raise RuntimeError("Chưa có session.")

        payload = {
            "goiTinDongBoNId": goi_tin_id,
            "actionProcessSyncDataThiCong": 1,
            "processSyncDatas": [],
            "trangThai": 1,
        }

        loi_cuoi = None
        for attempt in range(1, retry + 1):
            try:
                res = self.session.post(
                    URL_CAP_NHAT,
                    json=payload,
                    timeout=TIMEOUT,
                    allow_redirects=False,
                )

                if res.status_code in (301, 302) or res.headers.get("Location"):
                    raise RuntimeError(
                        f"Bị redirect (HTTP {res.status_code}) — phiên đăng nhập có thể đã hết hạn."
                    )
                if res.status_code == 404:
                    raise RuntimeError(f"URL không tồn tại (404): {URL_CAP_NHAT}")
                if res.status_code != 200:
                    raise RuntimeError(f"HTTP {res.status_code}: {res.text[:300]}")

                try:
                    js = res.json()
                except Exception:
                    raise RuntimeError(f"API cập nhật không trả JSON (HTTP {res.status_code}): {res.text[:300]}")

                if not js.get("success"):
                    raise RuntimeError("API trả success != True: " + str(js)[:300])

                return js
            except Exception as e:
                loi_cuoi = e
                if attempt < retry:
                    time.sleep(0.5)
                else:
                    raise loi_cuoi

    def cap_nhat_nhieu_goi_tin_song_song(self, list_goi_tin_id, max_workers=SO_WORKER_UPDATE):
        """Cập nhật song song nhiều gói tin để tiết kiệm thời gian."""
        ket_qua = {"thanh_cong": [], "that_bai": []}
        if not list_goi_tin_id:
            return ket_qua

        with concurrent.futures.ThreadPoolExecutor(max_workers=min(max_workers, len(list_goi_tin_id))) as executor:
            future_to_id = {
                executor.submit(self.cap_nhat_goi_tin, gid): gid
                for gid in list_goi_tin_id
            }
            for future in concurrent.futures.as_completed(future_to_id):
                gid = future_to_id[future]
                try:
                    future.result()
                    ket_qua["thanh_cong"].append(gid)
                except Exception as e:
                    ket_qua["that_bai"].append({"id": gid, "error": str(e)})

        return ket_qua


# ============================ TKINTER UI ============================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Cập nhật gói tin đồng bộ thi công MPLIS (Multi-Request)")
        root.geometry("920x680")

        self.client = MplisSyncClient(self.log)
        self.running = False
        self.stop_flag = False

        notebook = ttk.Notebook(root)
        notebook.pack(fill="x", padx=10, pady=5)

        # Tab 1: Đăng nhập bằng Chrome
        tab_browser = ttk.Frame(notebook, padding=8)
        notebook.add(tab_browser, text="1. Đăng nhập Chrome tự động")

        ttk.Label(tab_browser, text="Username:").grid(row=0, column=0, sticky="w")
        self.ent_user = ttk.Entry(tab_browser, width=25)
        self.ent_user.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(tab_browser, text="Password:").grid(row=0, column=2, sticky="w", padx=(10, 0))
        self.ent_pass = ttk.Entry(tab_browser, width=25, show="*")
        self.ent_pass.grid(row=0, column=3, sticky="w", padx=5, pady=2)

        btn_browser_box = ttk.Frame(tab_browser)
        btn_browser_box.grid(row=1, column=0, columnspan=4, sticky="w", pady=5)

        self.btn_login = ttk.Button(btn_browser_box, text="Mở Chrome đăng nhập", command=self.mo_chrome)
        self.btn_login.pack(side="left", padx=5)

        self.btn_confirm = ttk.Button(
            btn_browser_box, text="Đã đăng nhập xong → Lấy session", command=self.lay_session, state="disabled"
        )
        self.btn_confirm.pack(side="left", padx=5)

        # Tab 2: Dán Cookie / Token trực tiếp
        tab_manual = ttk.Frame(notebook, padding=8)
        notebook.add(tab_manual, text="2. Dán Cookie & Token (thủ công)")

        ttk.Label(tab_manual, text="Cookie:").grid(row=0, column=0, sticky="w")
        self.ent_cookie = ttk.Entry(tab_manual, width=65)
        self.ent_cookie.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(tab_manual, text="Token:").grid(row=1, column=0, sticky="w")
        self.ent_token = ttk.Entry(tab_manual, width=65)
        self.ent_token.grid(row=1, column=1, sticky="w", padx=5, pady=2)

        self.btn_apply_manual = ttk.Button(tab_manual, text="Áp dụng Cookie & Token", command=self.ap_dung_manual)
        self.btn_apply_manual.grid(row=1, column=2, padx=10, pady=2)

        # Frame Cấu hình thực thi
        frm_cfg = ttk.LabelFrame(root, text="Cấu hình gửi request & Cập nhật", padding=8)
        frm_cfg.pack(fill="x", padx=10, pady=5)

        ttk.Label(frm_cfg, text="dotBanGiaoNId:").grid(row=0, column=0, sticky="w")
        self.ent_dot_ban_giao = ttk.Entry(frm_cfg, width=45)
        self.ent_dot_ban_giao.insert(0, DOT_BAN_GIAO_MAC_DINH)
        self.ent_dot_ban_giao.grid(row=0, column=1, columnspan=3, sticky="w", padx=5, pady=3)

        ttk.Label(frm_cfg, text="Số request gửi 1 lần (song song):").grid(row=1, column=0, sticky="w")
        self.ent_so_request = ttk.Entry(frm_cfg, width=10)
        self.ent_so_request.insert(0, str(SO_REQUEST_DONG_THOI))
        self.ent_so_request.grid(row=1, column=1, sticky="w", padx=5, pady=3)

        ttk.Label(frm_cfg, text="length mỗi request:").grid(row=1, column=2, sticky="w", padx=(15, 0))
        self.ent_length = ttk.Entry(frm_cfg, width=10)
        self.ent_length.insert(0, str(SO_BAN_GHI_MOI_REQUEST))
        self.ent_length.grid(row=1, column=3, sticky="w", padx=5, pady=3)

        ttk.Label(frm_cfg, text="Giới hạn lượt lặp (test, để trống = chạy hết):").grid(row=2, column=0, sticky="w")
        self.ent_gioi_han_lap = ttk.Entry(frm_cfg, width=10)
        self.ent_gioi_han_lap.grid(row=2, column=1, sticky="w", padx=5, pady=3)

        self.var_bo_qua_van_hanh = tk.BooleanVar(value=False)
        self.chk_van_hanh = ttk.Checkbutton(
            frm_cfg,
            text="Bỏ qua nếu đã có dữ liệu vận hành (Mặc định TẮT - Cập nhật tất cả gói tin trạng thái 0)",
            variable=self.var_bo_qua_van_hanh,
        )
        self.chk_van_hanh.grid(row=2, column=2, columnspan=2, sticky="w", padx=15, pady=3)

        # Thanh nút điều khiển
        btn_frm = ttk.Frame(root, padding=(10, 0))
        btn_frm.pack(fill="x")

        self.btn_run = ttk.Button(btn_frm, text="▶ Bắt đầu xử lý", command=self.chay, state="disabled")
        self.btn_run.pack(side="left", padx=5)

        self.btn_stop = ttk.Button(btn_frm, text="⏹ Dừng", command=self.dung, state="disabled")
        self.btn_stop.pack(side="left", padx=5)

        self.btn_test_local = ttk.Button(btn_frm, text="🔍 Kiểm tra file mẫu @respone_kiemtraduliue.json", command=self.test_file_mau)
        self.btn_test_local.pack(side="left", padx=10)

        # Tiến trình & Trạng thái
        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(8, 0))

        status_frm = ttk.Frame(root)
        status_frm.pack(fill="x", padx=10)
        self.lbl_status = ttk.Label(status_frm, text="Sẵn sàng")
        self.lbl_status.pack(side="left")
        self.lbl_so_lan_lap = ttk.Label(status_frm, text="", foreground="blue")
        self.lbl_so_lan_lap.pack(side="right")

        # Khung Text Log
        self.txt = tk.Text(root, wrap="word", height=20)
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
                self.log("Chrome đã mở. Hoàn tất đăng nhập (OTP, captcha nếu có) rồi bấm nút Lấy session.")
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

    def ap_dung_manual(self):
        cookie_str = self.ent_cookie.get().strip()
        token = self.ent_token.get().strip()
        try:
            self.client.build_session_from_manual(cookie_str, token)
            self.btn_run.config(state="normal")
            messagebox.showinfo("Thành công", "Đã nhận Session từ Cookie & Token!")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def dung(self):
        self.stop_flag = True
        self.log("⏸ Đã yêu cầu dừng, sẽ dừng sau khi hoàn thành lượt hiện tại...")

    def test_file_mau(self):
        """Đọc và kiểm tra file mẫu respone_kiemtraduliue.json"""
        json_path = os.path.join(os.path.dirname(__file__), "respone_kiemtraduliue.json")
        if not os.path.exists(json_path):
            messagebox.showwarning("Không tìm thấy", f"Không tìm thấy file: {json_path}")
            return

        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            items = data.get("value") or []
            total = data.get("recordsTotal")
            self.log(f"📁 Đọc file mẫu '{os.path.basename(json_path)}':")
            self.log(f"   - recordsTotal: {total}")
            self.log(f"   - Số bản ghi trong payload: {len(items)}")

            id_trang_thai_0 = []
            for idx, item in enumerate(items, 1):
                gid = item.get("goiTinDongBoNId")
                status = item.get("trangThaiGoiTinDongBo")
                has_vh = bool(item.get("ThongTinDangKyVanHanh"))
                if str(status) == "0" and gid:
                    id_trang_thai_0.append(gid)
                self.log(f"   [{idx}] ID: {gid} | trangThai: {status} | ThongTinDangKyVanHanh: {has_vh}")

            self.log(f"   👉 Số gói tin có trạng thái 0 sẵn sàng update: {len(id_trang_thai_0)} / {len(items)}")
            messagebox.showinfo(
                "Kiểm tra file mẫu",
                f"Đã đọc file mẫu thành công!\n"
                f"recordsTotal: {total}\n"
                f"Số bản ghi trong payload: {len(items)}\n"
                f"Số gói tin trạng thái 0: {len(id_trang_thai_0)}"
            )
        except Exception as e:
            self.log(f"❌ Lỗi đọc file mẫu: {e}")
            messagebox.showerror("Lỗi", str(e))

    def chay(self):
        if self.running:
            return

        dot_ban_giao_id = self.ent_dot_ban_giao.get().strip()
        if not dot_ban_giao_id:
            messagebox.showwarning("Thiếu thông tin", "Nhập dotBanGiaoNId trước.")
            return

        try:
            so_request_dong_thoi = int(self.ent_so_request.get().strip() or SO_REQUEST_DONG_THOI)
            length = int(self.ent_length.get().strip() or SO_BAN_GHI_MOI_REQUEST)
        except ValueError:
            messagebox.showwarning("Sai định dạng", "Số request và length phải là số nguyên dương.")
            return

        gioi_han_nhap = self.ent_gioi_han_lap.get().strip()
        gioi_han_lap = None
        if gioi_han_nhap:
            if not gioi_han_nhap.isdigit() or int(gioi_han_nhap) <= 0:
                messagebox.showwarning("Sai giá trị", "Giới hạn số lượt phải là số nguyên dương.")
                return
            gioi_han_lap = int(gioi_han_nhap)

        bo_qua_van_hanh = self.var_bo_qua_van_hanh.get()

        dong_gioi_han = f"\nCHỈ CHẠY TỐI ĐA {gioi_han_lap} LƯỢT (chế độ test)." if gioi_han_lap else ""
        if not messagebox.askyesno(
            "Xác nhận",
            f"Sẽ gửi đồng thời {so_request_dong_thoi} request (mỗi request length={length})\n"
            f"đến KiemTraDuLieuChuyenDoi để tra cứu các gói tin có trangThai = 0,\n"
            f"rồi tiến hành update (trangThai = 1) cho dotBanGiaoNId:\n{dot_ban_giao_id}\n"
            f"{'⚠ Chế độ: Bỏ qua nếu có dữ liệu vận hành' if bo_qua_van_hanh else '⚡ Chế độ: Cập nhật tất cả gói tin có trạng thái 0'}"
            f"{dong_gioi_han}\n\nTiếp tục?",
        ):
            return

        self.running = True
        self.stop_flag = False
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progress.config(value=0, maximum=100)
        self.set_so_lan_lap(0, gioi_han_lap)

        threading.Thread(
            target=self._run_loop,
            args=(dot_ban_giao_id, length, so_request_dong_thoi, bo_qua_van_hanh, gioi_han_lap),
            daemon=True,
        ).start()

    def _run_loop(self, dot_ban_giao_id, length, so_request_dong_thoi, bo_qua_van_hanh, gioi_han_lap=None):
        da_thanh_cong = set()
        da_that_bai = {}
        da_bo_qua = set()
        so_lan_lap = 0
        tong_ban_dau = None

        while not self.stop_flag:
            if gioi_han_lap and so_lan_lap >= gioi_han_lap:
                self.log(f"⏸ Đã chạy đủ {gioi_han_lap} lượt theo giới hạn test. Dừng lại.")
                break

            so_lan_lap += 1
            self.set_so_lan_lap(so_lan_lap, gioi_han_lap)

            # Các gói tin lỗi hoặc bỏ qua sẽ kẹt lại ở đầu danh sách trangThai=0,
            # do đó bắt đầu lấy từ vị trí sau các gói tin đã bỏ qua/lỗi.
            base_offset = len(da_that_bai) + len(da_bo_qua)
            start_offsets = [base_offset + i * length for i in range(so_request_dong_thoi)]

            # Nếu đã biết recordsTotal và toàn bộ start_offsets vượt quá recordsTotal thì dừng
            if tong_ban_dau is not None and base_offset >= tong_ban_dau:
                self.log("✅ Toàn bộ bản ghi đã được duyệt qua. Hoàn tất.")
                break

            # Lọc bỏ các offset vượt quá tổng số bản ghi nếu đã biết
            if tong_ban_dau is not None:
                start_offsets = [s for s in start_offsets if s < tong_ban_dau]
                if not start_offsets:
                    self.log("✅ Không còn offset hợp lệ. Hoàn tất.")
                    break

            self.set_status(f"Đang gửi {len(start_offsets)} request tra cứu đồng thời (lượt {so_lan_lap})...")
            self.log(f"--- Lượt {so_lan_lap}: Gửi {len(start_offsets)} request tra cứu đồng thời (start={start_offsets}, length={length}) ---")

            try:
                ket_qua = self.client.tra_cuu_nhieu_request_song_song(dot_ban_giao_id, start_offsets, length=length)
            except Exception as e:
                self.log(f"❌ Lỗi tra cứu đồng thời: {e}")
                break

            danh_sach = ket_qua["items"]
            records_total = ket_qua["total"]

            if tong_ban_dau is None and records_total is not None:
                tong_ban_dau = records_total
                self.root.after(0, lambda t=tong_ban_dau: self.progress.config(maximum=max(1, t), value=0))

            self.log(f"   → Thu thập {len(danh_sach)} bản ghi từ {len(start_offsets)} request; recordsTotal: {records_total}")

            if not danh_sach or (records_total is not None and records_total == 0):
                self.log("✅ Không còn bản ghi nào có trangThai=0. Hoàn tất.")
                break

            # Lọc danh sách gói tin cần cập nhật
            id_can_update = []
            id_da_thay_trong_luot = set()

            for item in danh_sach:
                goi_tin_id = item.get("goiTinDongBoNId") or item.get("Id") or item.get("id")
                if not goi_tin_id:
                    continue
                goi_tin_id = str(goi_tin_id).strip()

                if goi_tin_id in id_da_thay_trong_luot:
                    continue
                id_da_thay_trong_luot.add(goi_tin_id)

                if goi_tin_id in da_thanh_cong or goi_tin_id in da_that_bai or goi_tin_id in da_bo_qua:
                    continue

                # Kiểm tra trạng thái gói tin: chỉ xử lý nếu trạng thái là 0
                trang_thai = item.get("trangThaiGoiTinDongBo")
                if trang_thai is None:
                    trang_thai = item.get("trangThai")
                if str(trang_thai) != "0":
                    self.log(f"   ⏭ Bỏ qua {goi_tin_id}: trangThai={trang_thai} (khác 0)")
                    da_bo_qua.add(goi_tin_id)
                    continue

                # Nếu bật tùy chọn bỏ qua khi có dữ liệu vận hành
                if bo_qua_van_hanh:
                    van_hanh = item.get("ThongTinDangKyVanHanh") or []
                    if van_hanh:
                        self.log(f"   ⏭ Bỏ qua {goi_tin_id}: đã có dữ liệu vận hành ({len(van_hanh)} mục)")
                        da_bo_qua.add(goi_tin_id)
                        continue

                id_can_update.append(goi_tin_id)

            if not id_can_update:
                self.log("⚠ Không tìm thấy thêm gói tin mới nào cần update trong lượt này.")
                # Nếu tất cả các bản ghi trong đợt này đều đã được duyệt qua hoặc lỗi
                if records_total is not None and (len(da_thanh_cong) + len(da_that_bai) + len(da_bo_qua)) >= records_total:
                    self.log("✅ Toàn bộ gói tin đều đã được xử lý hoặc ghi nhận. Hoàn tất.")
                    break
                # Tránh lặp vô hạn
                self.log("Dừng để tránh lặp vô hạn.")
                break

            self.log(f"   ⚡ Tiến hành update {len(id_can_update)} gói tin đồng thời...")
            self.set_status(f"Đang update {len(id_can_update)} gói tin...")

            kq_update = self.client.cap_nhat_nhieu_goi_tin_song_song(id_can_update, max_workers=SO_WORKER_UPDATE)

            for ok_id in kq_update["thanh_cong"]:
                da_thanh_cong.add(ok_id)
                self.log(f"   ✅ Update OK: {ok_id}")

            for err in kq_update["that_bai"]:
                da_that_bai[err["id"]] = err["error"]
                self.log(f"   ❌ Update LỖI {err['id']}: {err['error']}")

            # Cập nhật thanh tiến độ
            if tong_ban_dau:
                da_xong = len(da_thanh_cong) + len(da_that_bai) + len(da_bo_qua)
                self.root.after(0, lambda v=da_xong: self.progress.config(value=v))

            self.set_status(
                f"Đã xử lý: {len(da_thanh_cong)} thành công, {len(da_that_bai)} lỗi, "
                f"{len(da_bo_qua)} bỏ qua. (Tổng ban đầu: {tong_ban_dau or '?'})"
            )

            # Nghỉ ngắn giữa các đợt để giảm tải áp lực cho server
            time.sleep(0.5)

        self.log("=" * 60)
        self.log(f"HOÀN THÀNH sau {so_lan_lap} lượt tra cứu" + (f" (giới hạn: {gioi_han_lap})" if gioi_han_lap else ""))
        self.log(f"Thành công: {len(da_thanh_cong)}")
        self.log(f"Lỗi: {len(da_that_bai)}")
        self.log(f"Bỏ qua: {len(da_bo_qua)}")
        if da_that_bai:
            self.log("Danh sách lỗi:")
            self.log(json.dumps(da_that_bai, ensure_ascii=False, indent=2))

        self.set_status(f"Hoàn tất: {len(da_thanh_cong)} OK, {len(da_that_bai)} lỗi, {len(da_bo_qua)} bỏ qua")
        self.set_so_lan_lap(so_lan_lap, gioi_han_lap)
        self.running = False
        self.root.after(0, lambda: self.progress.config(value=self.progress["maximum"]))
        self.root.after(0, lambda: (self.btn_run.config(state="normal"), self.btn_stop.config(state="disabled")))

    def on_close(self):
        if self.running and not messagebox.askyesno("Đang chạy", "Đang xử lý, bạn có chắc chắn muốn thoát?"):
            return
        self.client.close_browser()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()