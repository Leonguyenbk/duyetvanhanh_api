# -*- coding: utf-8 -*-
"""
Tool tự động cập nhật trạng thái gói tin đồng bộ thi công MPLIS.

Luồng xử lý:
1. Mở Chrome và đăng nhập.
2. Lấy cookie + __RequestVerificationToken từ Chrome sang requests.Session.
3. Gọi KiemTraDuLieuChuyenDoi bằng payload DataTables dạng
   application/x-www-form-urlencoded.
4. Lấy trường "Id" trong từng bản ghi làm goiTinDongBoNId.
5. Update toàn bộ bản ghi nhận được, không kiểm tra ThongTinDangKyVanHanh.
6. Sau mỗi lô lại tra cứu từ start=0 cho đến khi hết trangThai=0.

Cài đặt:
    pip install requests selenium webdriver-manager
"""

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

REFERER_LOGIN = f"{BASE_URL}/dc/DonDangKy/KeKhaiDangKyV2"
REFERER_API = f"{BASE_URL}/dc/"

URL_KIEM_TRA = (
    f"{BASE_URL}/dc/TichHopDongBoDuLieuAjax/GetGoiTinDongBo"
)
URL_CAP_NHAT = (
    f"{BASE_URL}/dc/TichHopDongBoDuLieuAjax/UpdateStatusGoiTinDongBoThiCong"
)

SO_BAN_GHI_MOI_TRANG = 10
TIMEOUT = 60
SO_LAN_THU_LAI_UPDATE = 3
THOI_GIAN_NGHI_GIUA_UPDATE = 0.25

# Giá trị người dùng cung cấp; vẫn có thể sửa trực tiếp trên giao diện.
DOT_BAN_GIAO_MAC_DINH = ""


# ============================ HELPER ============================

def lay_token_tu_trang(driver):
    """Lấy anti-forgery token từ input hoặc meta của trang hiện tại."""
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
    return (driver.execute_script(js) or "").strip()


def rut_gon_response_text(response, gioi_han=3000):
    text = response.text or ""
    if len(text) <= gioi_han:
        return text
    return text[:gioi_han] + "..."


# ============================ CORE API ============================

class MplisSyncClient:
    def __init__(self, log_fn):
        self.log = log_fn
        self.session = None
        self.driver = None
        self.draw_counter = 0

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
                input_type = (inp.get_attribute("type") or "").lower()
                if user_box is None and input_type in ("text", "email"):
                    user_box = inp
                if pass_box is None and input_type == "password":
                    pass_box = inp

            if user_box is not None and pass_box is not None:
                user_box.clear()
                user_box.send_keys(username)
                pass_box.clear()
                pass_box.send_keys(password)
                pass_box.send_keys(Keys.ENTER)
                self.log("Đã điền thông tin đăng nhập, chờ trang tải...")
            else:
                self.log("Không tìm thấy đủ ô đăng nhập; hãy đăng nhập tay trên Chrome.")

        except Exception as exc:
            self.log(
                f"Không tự điền được form đăng nhập ({exc}); "
                "hãy đăng nhập tay trên Chrome."
            )

    def build_session_from_browser(self):
        if self.driver is None:
            raise RuntimeError("Chưa mở trình duyệt.")

        token = lay_token_tu_trang(self.driver)
        if not token:
            raise RuntimeError(
                "Không lấy được __RequestVerificationToken. "
                "Hãy bảo đảm đã đăng nhập và đang mở trang MPLIS."
            )

        session = requests.Session()
        user_agent = self.driver.execute_script("return navigator.userAgent;")

        # Không đặt Content-Type cố định ở session vì:
        # - API tra cứu dùng form-urlencoded.
        # - API update dùng JSON.
        session.headers.update({
            "User-Agent": user_agent,
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": BASE_URL,
            "Referer": REFERER_API,
            "__requestverificationtoken": token,
        })

        for cookie in self.driver.get_cookies():
            session.cookies.set(
                name=cookie["name"],
                value=cookie["value"],
                domain=cookie.get("domain"),
                path=cookie.get("path", "/"),
            )

        self.session = session
        self.log("✅ Đã lấy session + token thành công.")

    def close_browser(self):
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    def _kiem_tra_session(self):
        if self.session is None:
            raise RuntimeError("Chưa lấy session từ Chrome.")

    @staticmethod
    def _kiem_tra_redirect(response):
        if 300 <= response.status_code < 400 or response.headers.get("Location"):
            raise RuntimeError(
                f"Bị redirect HTTP {response.status_code}; "
                f"Location={response.headers.get('Location')}. "
                "Có thể phiên đăng nhập đã hết hạn."
            )

    # ---------- payload DataTables ----------
    def tao_payload_kiem_tra(self, dot_ban_giao_id, start, length):
        self.draw_counter += 1

        return {
            "draw": str(self.draw_counter),

            "columns[0][data]": "#",
            "columns[0][name]": "",
            "columns[0][searchable]": "true",
            "columns[0][orderable]": "false",
            "columns[0][search][value]": "",
            "columns[0][search][regex]": "false",

            "columns[1][data]": "thuaDat",
            "columns[1][name]": "",
            "columns[1][searchable]": "true",
            "columns[1][orderable]": "false",
            "columns[1][search][value]": "",
            "columns[1][search][regex]": "false",

            "columns[2][data]": "chuSuDung",
            "columns[2][name]": "",
            "columns[2][searchable]": "true",
            "columns[2][orderable]": "false",
            "columns[2][search][value]": "",
            "columns[2][search][regex]": "false",

            "columns[3][data]": "ngayDangKy",
            "columns[3][name]": "",
            "columns[3][searchable]": "true",
            "columns[3][orderable]": "false",
            "columns[3][search][value]": "",
            "columns[3][search][regex]": "false",

            "columns[4][data]": "nguoiDuyet",
            "columns[4][name]": "",
            "columns[4][searchable]": "true",
            "columns[4][orderable]": "false",
            "columns[4][search][value]": "",
            "columns[4][search][regex]": "false",

            "columns[5][data]": "nguoiDongBo",
            "columns[5][name]": "",
            "columns[5][searchable]": "true",
            "columns[5][orderable]": "false",
            "columns[5][search][value]": "",
            "columns[5][search][regex]": "false",

            "columns[6][data]": "ghiChu",
            "columns[6][name]": "",
            "columns[6][searchable]": "true",
            "columns[6][orderable]": "false",
            "columns[6][search][value]": "",
            "columns[6][search][regex]": "false",

            "columns[7][data]": "trangThai",
            "columns[7][name]": "",
            "columns[7][searchable]": "true",
            "columns[7][orderable]": "false",
            "columns[7][search][value]": "",
            "columns[7][search][regex]": "false",

            "start": str(start),
            "length": str(length),
            "search[value]": "",
            "search[regex]": "false",

            "dotBanGiaoNId": dot_ban_giao_id,
            "requestQueries[soPhatHanh]": "",
            "requestQueries[soThuTuThua]": "",
            "requestQueries[soHieuToBanDo]": "",
            "requestQueries[hoTenChu]": "",
            "requestQueries[soGiayTo]": "",
            "requestQueries[trangThai]": "0",
        }

    # ---------- API tra cứu ----------
    def kiem_tra_du_lieu_chuyen_doi(
        self,
        dot_ban_giao_id,
        start=0,
        length=SO_BAN_GHI_MOI_TRANG,
    ):
        self._kiem_tra_session()

        payload = self.tao_payload_kiem_tra(
            dot_ban_giao_id=dot_ban_giao_id,
            start=start,
            length=length,
        )

        response = self.session.post(
            URL_KIEM_TRA,
            data=payload,  # QUAN TRỌNG: payload form-urlencoded, không phải JSON
            headers={
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
            },
            timeout=TIMEOUT,
            allow_redirects=False,
        )

        self._kiem_tra_redirect(response)

        if response.status_code == 404:
            raise RuntimeError(f"URL không tồn tại (404): {URL_KIEM_TRA}")

        if not 200 <= response.status_code < 300:
            raise RuntimeError(
                f"API tra cứu lỗi HTTP {response.status_code}: "
                f"{rut_gon_response_text(response)}"
            )

        try:
            body = response.json()
        except ValueError as exc:
            raise RuntimeError(
                f"API tra cứu không trả JSON, HTTP {response.status_code}: "
                f"{rut_gon_response_text(response)}"
            ) from exc

        # DataTables thường trả danh sách trong "data".
        # Giữ thêm "value" để tương thích nếu server trả theo kiểu cũ.
        items = body.get("data")
        if items is None:
            items = body.get("value")
        if items is None:
            items = body.get("items")
        if not isinstance(items, list):
            items = []

        total = body.get("recordsFiltered")
        if total is None:
            total = body.get("recordsTotal")
        if total is None:
            total = len(items)

        try:
            total = int(total)
        except (TypeError, ValueError):
            total = len(items)

        return {
            "items": items,
            "total": total,
            "raw": body,
        }

    # ---------- API update ----------
    def cap_nhat_goi_tin(self, goi_tin_id):
        self._kiem_tra_session()

        payload = {
            "goiTinDongBoNId": goi_tin_id,
            "actionProcessSyncDataThiCong": 1,
            "processSyncDatas": [],
            "trangThai": 1,
        }

        loi_cuoi = None

        for lan_thu in range(1, SO_LAN_THU_LAI_UPDATE + 1):
            try:
                response = self.session.post(
                    URL_CAP_NHAT,
                    json=payload,
                    headers={"Content-Type": "application/json; charset=UTF-8"},
                    timeout=TIMEOUT,
                    allow_redirects=False,
                )

                self._kiem_tra_redirect(response)

                if response.status_code == 404:
                    raise RuntimeError(f"URL không tồn tại (404): {URL_CAP_NHAT}")

                if not 200 <= response.status_code < 300:
                    raise RuntimeError(
                        f"HTTP {response.status_code}: "
                        f"{rut_gon_response_text(response)}"
                    )

                try:
                    body = response.json()
                except ValueError as exc:
                    raise RuntimeError(
                        "Response update không phải JSON: "
                        f"{rut_gon_response_text(response)}"
                    ) from exc

                if body.get("success") is not True:
                    raise RuntimeError(
                        "API trả success khác true: "
                        + json.dumps(body, ensure_ascii=False)[:3000]
                    )

                return body

            except Exception as exc:
                loi_cuoi = exc
                if lan_thu < SO_LAN_THU_LAI_UPDATE:
                    self.log(
                        f"      ⚠ Update {goi_tin_id} lỗi lần "
                        f"{lan_thu}/{SO_LAN_THU_LAI_UPDATE}: {exc}"
                    )
                    time.sleep(lan_thu * 1.5)

        raise RuntimeError(
            f"Update thất bại sau {SO_LAN_THU_LAI_UPDATE} lần: {loi_cuoi}"
        )


# ============================ TKINTER UI ============================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Cập nhật gói tin đồng bộ thi công MPLIS")
        self.root.geometry("900x650")

        self.client = MplisSyncClient(self.log)
        self.running = False
        self.stop_flag = False

        form = ttk.Frame(root, padding=10)
        form.pack(fill="x")

        ttk.Label(form, text="Username:").grid(row=0, column=0, sticky="w")
        self.ent_user = ttk.Entry(form, width=30)
        self.ent_user.grid(row=0, column=1, sticky="ew", padx=5, pady=3)

        ttk.Label(form, text="Password:").grid(row=0, column=2, sticky="w")
        self.ent_pass = ttk.Entry(form, width=30, show="*")
        self.ent_pass.grid(row=0, column=3, sticky="ew", padx=5, pady=3)

        ttk.Label(form, text="dotBanGiaoNId:").grid(row=1, column=0, sticky="w")
        self.ent_dot_ban_giao = ttk.Entry(form, width=48)
        self.ent_dot_ban_giao.grid(
            row=1,
            column=1,
            columnspan=2,
            sticky="ew",
            padx=5,
            pady=3,
        )
        self.ent_dot_ban_giao.insert(0, DOT_BAN_GIAO_MAC_DINH)

        ttk.Label(
            form,
            text="Giới hạn số lượt (để trống = chạy hết):",
        ).grid(row=2, column=0, sticky="w")
        self.ent_gioi_han_lap = ttk.Entry(form, width=10)
        self.ent_gioi_han_lap.grid(row=2, column=1, sticky="w", padx=5, pady=3)

        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        buttons = ttk.Frame(root, padding=(10, 0))
        buttons.pack(fill="x")

        self.btn_login = ttk.Button(
            buttons,
            text="1. Mở Chrome đăng nhập",
            command=self.mo_chrome,
        )
        self.btn_login.pack(side="left", padx=5)

        self.btn_confirm = ttk.Button(
            buttons,
            text="2. Lấy session",
            command=self.lay_session,
            state="disabled",
        )
        self.btn_confirm.pack(side="left", padx=5)

        self.btn_run = ttk.Button(
            buttons,
            text="3. Bắt đầu xử lý",
            command=self.chay,
            state="disabled",
        )
        self.btn_run.pack(side="left", padx=5)

        self.btn_stop = ttk.Button(
            buttons,
            text="Dừng",
            command=self.dung,
            state="disabled",
        )
        self.btn_stop.pack(side="left", padx=5)

        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(8, 0))

        status_frame = ttk.Frame(root)
        status_frame.pack(fill="x", padx=10, pady=(3, 0))

        self.lbl_status = ttk.Label(status_frame, text="Chưa chạy")
        self.lbl_status.pack(side="left")

        self.lbl_so_lan_lap = ttk.Label(status_frame, text="", foreground="blue")
        self.lbl_so_lan_lap.pack(side="right")

        log_frame = ttk.Frame(root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=8)

        self.txt = tk.Text(log_frame, wrap="word", height=28)
        self.txt.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.txt.yview,
        )
        scrollbar.pack(side="right", fill="y")
        self.txt.configure(yscrollcommand=scrollbar.set)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- UI helpers ----------
    def log(self, message):
        def append_log():
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.txt.insert("end", f"{timestamp}  {message}\n")
            self.txt.see("end")

        self.root.after(0, append_log)

    def set_status(self, message):
        self.root.after(0, lambda: self.lbl_status.config(text=message))

    def set_so_lan_lap(self, so_lan, gioi_han=None):
        text = f"Lượt: {so_lan}"
        if gioi_han:
            text += f" / {gioi_han}"
        self.root.after(0, lambda: self.lbl_so_lan_lap.config(text=text))

    def set_progress(self, value=None, maximum=None):
        def update_progress():
            if maximum is not None:
                self.progress.configure(maximum=max(1, maximum))
            if value is not None:
                self.progress.configure(value=max(0, value))

        self.root.after(0, update_progress)

    # ---------- actions ----------
    def mo_chrome(self):
        username = self.ent_user.get().strip()
        password = self.ent_pass.get()

        if not username or not password:
            messagebox.showwarning(
                "Thiếu thông tin",
                "Nhập username và password trước.",
            )
            return

        self.btn_login.config(state="disabled")

        def work():
            try:
                self.log("Đang mở Chrome...")
                self.client.open_browser_and_fill_login(username, password)
                self.log(
                    "Chrome đã mở. Hoàn tất đăng nhập nếu có OTP/captcha, "
                    "sau đó bấm nút '2. Lấy session'."
                )
                self.root.after(
                    0,
                    lambda: self.btn_confirm.config(state="normal"),
                )
            except Exception as exc:
                self.log(f"❌ Lỗi mở Chrome: {exc}")
                self.root.after(
                    0,
                    lambda: self.btn_login.config(state="normal"),
                )

        threading.Thread(target=work, daemon=True).start()

    def lay_session(self):
        def work():
            try:
                self.client.build_session_from_browser()
                self.root.after(0, lambda: self.btn_run.config(state="normal"))
            except Exception as exc:
                self.log(f"❌ Lỗi lấy session: {exc}")

        threading.Thread(target=work, daemon=True).start()

    def dung(self):
        self.stop_flag = True
        self.log("⏸ Đã yêu cầu dừng; chương trình sẽ dừng sau request hiện tại.")

    def chay(self):
        if self.running:
            return

        dot_ban_giao_id = self.ent_dot_ban_giao.get().strip()
        if not dot_ban_giao_id:
            messagebox.showwarning(
                "Thiếu thông tin",
                "Nhập dotBanGiaoNId trước.",
            )
            return

        gioi_han_lap = None
        gioi_han_nhap = self.ent_gioi_han_lap.get().strip()
        if gioi_han_nhap:
            if not gioi_han_nhap.isdigit() or int(gioi_han_nhap) <= 0:
                messagebox.showwarning(
                    "Sai giá trị",
                    "Giới hạn số lượt phải là số nguyên dương.",
                )
                return
            gioi_han_lap = int(gioi_han_nhap)

        gioi_han_text = ""
        if gioi_han_lap:
            gioi_han_text = f"\nChỉ chạy tối đa {gioi_han_lap} lượt."

        confirmed = messagebox.askyesno(
            "Xác nhận",
            f"Tra cứu mỗi lần {SO_BAN_GHI_MOI_TRANG} bản ghi có trangThai=0.\n"
            "Lấy trường Id làm goiTinDongBoNId và update toàn bộ, "
            "không kiểm tra ThongTinDangKyVanHanh.\n\n"
            f"dotBanGiaoNId:\n{dot_ban_giao_id}"
            f"{gioi_han_text}\n\nTiếp tục?",
        )
        if not confirmed:
            return

        self.running = True
        self.stop_flag = False
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.set_progress(value=0, maximum=100)
        self.set_so_lan_lap(0, gioi_han_lap)

        threading.Thread(
            target=self._run_loop,
            args=(dot_ban_giao_id, gioi_han_lap),
            daemon=True,
        ).start()

    def _run_loop(self, dot_ban_giao_id, gioi_han_lap=None):
        thanh_cong = []
        loi_theo_id = {}
        id_da_loi = set()

        so_lan_lap = 0
        tong_ban_dau = None
        ket_thuc_vi_het_du_lieu = False

        while not self.stop_flag:
            if gioi_han_lap and so_lan_lap >= gioi_han_lap:
                self.log(
                    f"⏸ Đã chạy đủ {gioi_han_lap} lượt theo giới hạn."
                )
                break

            so_lan_lap += 1
            self.set_so_lan_lap(so_lan_lap, gioi_han_lap)
            self.set_status(f"Đang tra cứu lượt {so_lan_lap}...")
            self.log(
                f"--- Lượt {so_lan_lap}: tra cứu start=0, "
                f"length={SO_BAN_GHI_MOI_TRANG}, trangThai=0 ---"
            )

            try:
                result = self.client.kiem_tra_du_lieu_chuyen_doi(
                    dot_ban_giao_id=dot_ban_giao_id,
                    start=0,
                    length=SO_BAN_GHI_MOI_TRANG,
                )
            except Exception as exc:
                self.log(f"❌ Lỗi tra cứu: {exc}")
                break

            danh_sach = result["items"]
            records_total = result["total"]

            if tong_ban_dau is None:
                tong_ban_dau = max(records_total, len(danh_sach))
                self.set_progress(value=0, maximum=max(1, tong_ban_dau))

            self.log(
                f"   → Nhận {len(danh_sach)} bản ghi; "
                f"recordsTotal/recordsFiltered={records_total}"
            )

            if not danh_sach or records_total == 0:
                self.log("✅ Không còn bản ghi trangThai=0. Hoàn tất.")
                ket_thuc_vi_het_du_lieu = True
                break

            # QUAN TRỌNG: lấy trường Id làm goiTinDongBoNId.
            # Không đọc và không kiểm tra ThongTinDangKyVanHanh.
            danh_sach_id = []
            id_trong_lo = set()

            for item in danh_sach:
                goi_tin_id = item.get("Id") or item.get("id")

                if not goi_tin_id:
                    self.log(
                        "   ⚠ Bỏ qua một bản ghi vì response không có trường Id: "
                        + json.dumps(item, ensure_ascii=False)[:800]
                    )
                    continue

                goi_tin_id = str(goi_tin_id).strip()
                if not goi_tin_id or goi_tin_id in id_trong_lo:
                    continue

                id_trong_lo.add(goi_tin_id)
                danh_sach_id.append(goi_tin_id)

            if not danh_sach_id:
                self.log(
                    "❌ Response có bản ghi nhưng không lấy được Id nào. "
                    "Kiểm tra lại tên trường trong response."
                )
                self.log(
                    "Response mẫu: "
                    + json.dumps(result["raw"], ensure_ascii=False)[:3000]
                )
                break

            # Những ID đã lỗi sau 3 lần sẽ không gọi lại trong cùng lần chạy,
            # tránh lặp vô hạn khi start luôn bằng 0.
            id_can_update = [
                goi_tin_id
                for goi_tin_id in danh_sach_id
                if goi_tin_id not in id_da_loi
            ]

            if not id_can_update:
                self.log(
                    "⚠ Lô start=0 hiện chỉ còn các ID đã lỗi sau nhiều lần thử. "
                    "Dừng để tránh lặp vô hạn."
                )
                break

            so_ok_trong_luot = 0

            for index, goi_tin_id in enumerate(id_can_update, start=1):
                if self.stop_flag:
                    self.log("⏹ Đã dừng theo yêu cầu.")
                    break

                self.set_status(
                    f"Lượt {so_lan_lap}: update {index}/{len(id_can_update)} - "
                    f"{goi_tin_id}"
                )

                try:
                    response_update = self.client.cap_nhat_goi_tin(goi_tin_id)
                    thanh_cong.append(goi_tin_id)
                    so_ok_trong_luot += 1

                    trang_thai_tra_ve = response_update.get("trangThai")
                    action_tra_ve = response_update.get(
                        "actionProcessSyncDataThiCong"
                    )

                    self.log(
                        f"   ✅ [{index}/{len(id_can_update)}] Update OK: "
                        f"{goi_tin_id} | trangThai={trang_thai_tra_ve} | "
                        f"action={action_tra_ve}"
                    )

                except Exception as exc:
                    id_da_loi.add(goi_tin_id)
                    loi_theo_id[goi_tin_id] = str(exc)
                    self.log(
                        f"   ❌ [{index}/{len(id_can_update)}] "
                        f"Update lỗi {goi_tin_id}: {exc}"
                    )

                if tong_ban_dau:
                    self.set_progress(
                        value=min(
                            len(thanh_cong) + len(id_da_loi),
                            tong_ban_dau,
                        )
                    )

                time.sleep(THOI_GIAN_NGHI_GIUA_UPDATE)

            if self.stop_flag:
                break

            if so_ok_trong_luot == 0:
                self.log(
                    "⚠ Lượt này không có bản ghi nào update thành công. "
                    "Dừng để tránh gọi lại vô hạn cùng một lô."
                )
                break

            self.set_status(
                f"Đã update {len(thanh_cong)} bản ghi; "
                f"{len(id_da_loi)} bản ghi lỗi. Đang lấy lại start=0..."
            )

        self.log("=" * 70)
        self.log(f"KẾT THÚC SAU {so_lan_lap} LƯỢT TRA CỨU")
        self.log(f"Update thành công: {len(thanh_cong)}")
        self.log(f"Update lỗi: {len(loi_theo_id)}")

        if loi_theo_id:
            self.log("Danh sách lỗi:")
            self.log(
                json.dumps(
                    [
                        {
                            "goiTinDongBoNId": goi_tin_id,
                            "loi": error,
                        }
                        for goi_tin_id, error in loi_theo_id.items()
                    ],
                    ensure_ascii=False,
                    indent=2,
                )
            )

        if ket_thuc_vi_het_du_lieu and tong_ban_dau:
            self.set_progress(value=tong_ban_dau, maximum=tong_ban_dau)

        self.set_status(
            f"Hoàn tất: {len(thanh_cong)} OK, {len(loi_theo_id)} lỗi"
        )
        self.set_so_lan_lap(so_lan_lap, gioi_han_lap)

        self.running = False
        self.root.after(0, lambda: self.btn_run.config(state="normal"))
        self.root.after(0, lambda: self.btn_stop.config(state="disabled"))

    def on_close(self):
        if self.running:
            confirmed = messagebox.askyesno(
                "Đang chạy",
                "Chương trình đang xử lý. Thoát luôn?",
            )
            if not confirmed:
                return

        self.stop_flag = True
        self.client.close_browser()
        self.root.destroy()


def kiem_tra_dieu_kien_khoi_dong(root):
    """Giữ lại hai câu hỏi khởi động từ phiên bản cũ."""
    root.withdraw()

    beo = messagebox.askyesno("Câu hỏi 1", "Anh Tuấn có béo không?")
    if beo:
        messagebox.showerror("Kết thúc", "Bạn đã làm tổn thương anh Tuấn")
        return False

    dep_trai = messagebox.askyesno("Câu hỏi 2", "Anh Tuấn có đẹp trai không?")
    if not dep_trai:
        messagebox.showerror("Kết thúc", "Bạn không trung thực")
        return False

    messagebox.showinfo("OK", "Cảm ơn bạn đã trung thực")
    root.deiconify()
    return True


if __name__ == "__main__":
    root = tk.Tk()

    if not kiem_tra_dieu_kien_khoi_dong(root):
        root.destroy()
    else:
        app = App(root)
        root.mainloop()
