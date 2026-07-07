# -*- coding: utf-8 -*-
"""
Tool duyệt vận hành hàng loạt MPLIS
- Tkinter UI: nhập username/password, chọn folder hồ sơ quét, chọn file Excel
- Excel cần các cột: "Mã xã", "Số tờ", "Số thửa", "Số phát hành", "Tên file"
- Các bản ghi trùng "Số phát hành" chỉ xử lý 1 bản ghi đại diện (bản ghi đầu tiên)
Cài đặt: pip install selenium webdriver-manager requests pandas openpyxl
"""

import os
import re
import json
import uuid
import copy
import time
import threading
import traceback
from datetime import datetime, timezone

import requests
import pandas as pd

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# ============================ CẤU HÌNH ============================

REFERER_URL = "https://dla.mplis.gov.vn/dc/DonDangKy/KeKhaiDangKyV2"

URL_SEARCH_THU_THAP_GCN = "https://dla.mplis.gov.vn/dc/ThuThapThongTinAjax/GetThuThapThongTinGiayChungNhan"
URL_UPLOAD_THU_THAP_GCN = "https://dla.mplis.gov.vn/dc/ThuThapThongTinAjax/UpdateTapTinThuThapThongTin"
URL_BUILD_DUYET = "https://dla.mplis.gov.vn/dc/ThuThapThongTinAjax/ConvertThongTinDangKy"
URL_CHECK_DUYET = "https://dla.mplis.gov.vn/dc/DangKyAjax/ValidateThongTinDangKy"
URL_SAVE_DUYET = "https://dla.mplis.gov.vn/dc/ThuThapThongTinAjax/SaveThongTinDangKy"

TIMEOUT = 180
KIEM_TRA_TRUOC_KHI_DUYET = True

# Tên cột trong Excel
COL_MA_XA = "Mã xã"
COL_SO_TO = "Số tờ"           # tùy chọn, chỉ để hiển thị/log
COL_SO_THUA = "Số thửa"       # tùy chọn, chỉ để hiển thị/log
COL_SO_PHAT_HANH = "Số phát hành"
COL_TEN_MO_TA = "Tên mô tả"   # = tenTapTin, nhiều giá trị cách nhau bởi , hoặc ;
COL_TEN_FILE = "Tên file"     # nhiều file cách nhau bởi , hoặc ; theo đúng thứ tự Tên mô tả

REQUIRED_COLS = [COL_MA_XA, COL_SO_PHAT_HANH, COL_TEN_MO_TA, COL_TEN_FILE]
OPTIONAL_COLS = [COL_SO_TO, COL_SO_THUA]


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


def dotnet_date_to_iso(value):
    if not isinstance(value, str):
        return value
    m = re.search(r"/Date\((-?\d+)\)/", value)
    if not m:
        return value
    ms = int(m.group(1))
    if ms <= -62135596800000:
        return "0001-01-01T00:00:00.000Z"
    dt = datetime.fromtimestamp(ms / 1000, tz=timezone.utc)
    return dt.strftime("%Y-%m-%dT%H:%M:%S.") + f"{int(dt.microsecond / 1000):03d}Z"


def convert_dates_recursive(obj):
    if isinstance(obj, dict):
        return {k: convert_dates_recursive(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [convert_dates_recursive(x) for x in obj]
    return dotnet_date_to_iso(obj)


def add_id_recursive(obj):
    if isinstance(obj, list):
        for i, item in enumerate(obj, start=1):
            if isinstance(item, dict):
                item.setdefault("_id", i)
                add_id_recursive(item)
    elif isinstance(obj, dict):
        for v in obj.values():
            add_id_recursive(v)
    return obj


def clean_cell(v):
    """Chuẩn hóa giá trị đọc từ Excel: bỏ NaN, bỏ .0 của số, strip khoảng trắng."""
    if v is None:
        return ""
    if isinstance(v, float):
        if pd.isna(v):
            return ""
        if v == int(v):
            return str(int(v))
        return str(v)
    s = str(v).strip()
    if s.lower() == "nan":
        return ""
    return s


def tach_danh_sach(cell):
    """Tách chuỗi nhiều giá trị cách nhau bởi , hoặc ; thành list, bỏ phần tử rỗng."""
    s = clean_cell(cell)
    if not s:
        return []
    return [p.strip() for p in re.split(r"[,;]", s) if p.strip()]


def xac_dinh_loai_ho_so_quet(ten_mo_ta):
    """
    Xác định loaiHoSoQuet theo HẬU TỐ cuối tên mô tả (sau dấu - hoặc _):
      -GT = 0 (giấy tờ), -GCN = 1 (giấy chứng nhận),
      -DDK = 2 (đơn đăng ký), -TBXN = 3 (thông báo xác nhận).
    Bỏ qua đuôi .pdf nếu có. Không nhận diện được → mặc định 2 (DDK).
    """
    s = (ten_mo_ta or "").strip().upper()
    s = re.sub(r"\.PDF$", "", s)  # bỏ đuôi .pdf nếu có
    m = re.search(r"[-_](GT|GCN|DDK|TBXN)$", s)
    if not m:
        return 2
    return {"GT": 0, "GCN": 1, "DDK": 2, "TBXN": 3}[m.group(1)]


# ============================ CORE API ============================

class BoQuaBanGhi(Exception):
    """Bản ghi đã được duyệt vào vận hành rồi → bỏ qua, không phải lỗi."""
    pass


class MplisClient:
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
        self.driver.get(REFERER_URL)
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
        """Gọi sau khi user xác nhận đã đăng nhập xong trên Chrome."""
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
            "X-Requested-With": "XMLHttpRequest",
            "Origin": "https://dla.mplis.gov.vn",
            "Referer": REFERER_URL,
            "__requestverificationtoken": token,
            "__RequestVerificationToken": token,
            "RequestVerificationToken": token,
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

    # ---------- request helpers ----------
    def _check_response(self, res, name, url):
        if res.status_code == 404:
            raise RuntimeError(f"{name}: URL không tồn tại (404): {url}")
        ct = res.headers.get("Content-Type") or ""
        if "application/json" not in ct and "text/json" not in ct:
            raise RuntimeError(
                f"{name}: server không trả JSON (status {res.status_code}), "
                f"có thể mất session hoặc sai URL. Nội dung: {res.text[:400]}"
            )

    def post_json(self, url, payload, name):
        headers = dict(self.session.headers)
        headers.update({
            "Content-Type": "application/json; charset=UTF-8",
        })
        res = self.session.post(
            url,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers=headers,
            timeout=TIMEOUT,
        )
        self._check_response(res, name, url)
        return res.json()

    # ---------- business ----------
    def tra_cuu_thu_thap_gcn(self, xa_id, so_phat_hanh, so_to=None, so_thua=None):
        # Tìm CHỈ theo số phát hành (+ xã), không dùng tờ/thửa
        headers = dict(self.session.headers)
        headers.update({
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        })

        data = {
            "traCuu[tuNgay]": "",
            "traCuu[denNgay]": "",
            "traCuu[soPhatHanh]": clean_cell(so_phat_hanh),
            "traCuu[soVaoSo]": "",
            "traCuu[hoTenChu]": "",
            "traCuu[soGiayTo]": "",
            "traCuu[soThuTuThua]": "",
            "traCuu[soHieuToBanDo]": "",
            "traCuu[soThuTuThuaCu]": "",
            "traCuu[soHieuToBanDoCu]": "",
            "traCuu[type]": "-1",
            "traCuu[coFileHoSoQuet]": "",
            "traCuu[query]": "",
            "traCuu[xaId]": clean_cell(xa_id),
            "start": "0",
            "length": "10",
        }

        res = self.session.post(URL_SEARCH_THU_THAP_GCN, data=data, headers=headers, timeout=120)
        self._check_response(res, "TRA CỨU", URL_SEARCH_THU_THAP_GCN)
        js = res.json()

        if not js.get("success"):
            raise RuntimeError("Tra cứu lỗi: " + str(js)[:500])

        rows = js.get("data") or []
        if not rows:
            raise RuntimeError("Không tìm thấy bản ghi thu thập.")

        if len(rows) > 1:
            # Lọc lại theo soPhatHanh khớp CHÍNH XÁC (server có thể match gần đúng)
            sph_can_tim = clean_cell(so_phat_hanh).upper()
            khop = [
                r for r in rows
                if clean_cell(r.get("soPhatHanh")).upper() == sph_can_tim
            ]
            if len(khop) == 1:
                self.log(
                    f"   → Tra cứu ra {len(rows)} bản ghi, đã lọc còn 1 khớp chính xác soPhatHanh."
                )
                rows = khop
            else:
                ds_sph = ", ".join(clean_cell(r.get("soPhatHanh")) for r in rows[:5])
                raise RuntimeError(
                    f"Tìm thấy {len(rows)} bản ghi ({len(khop)} khớp chính xác soPhatHanh), "
                    f"dừng để tránh nhầm. Các soPhatHanh: {ds_sph}"
                )

        item = rows[0]
        return {
            "thu_thap_id": item["thuThapThongTinGiayChungNhanId"],
            "so_phat_hanh": item.get("soPhatHanh"),
            "ten_file_hsq": item.get("tenFileHoSoQuet"),
            "tinh_hinh_dang_ky_id": item.get("tinhHinhDangKyId"),
            "type": item.get("type"),
            "raw": item,
        }

    def upload_files(self, thu_thap_id, danh_sach):
        """
        Upload nhiều hồ sơ quét trong 1 request.
        danh_sach: list các tuple (file_path, ten_tap_tin)
        """
        # Kiểm tra tồn tại file trước khi mở
        for file_path, _ in danh_sach:
            if not os.path.isfile(file_path):
                raise FileNotFoundError(f"Không thấy file: {file_path}")

        list_tap_tin = []
        file_parts = []  # list of (field_name, (filename, bytes, mimetype))

        for file_path, ten_tap_tin in danh_sach:
            temp_id = str(uuid.uuid4())
            loai = xac_dinh_loai_ho_so_quet(ten_tap_tin)
            list_tap_tin.append({
                "Id": temp_id,
                "tenTapTin": ten_tap_tin,
                "laGiayChungNhan": loai == 1,
                "loaiHoSoQuet": loai,
            })
            self.log(f"      • {ten_tap_tin} → loaiHoSoQuet={loai}")
            with open(file_path, "rb") as f:
                data = f.read()
            file_parts.append((temp_id, (os.path.basename(file_path), data, "application/pdf")))

        payload_json = {
            "thuThapThongTinGiayChungNhanId": int(thu_thap_id),
            "ListTapTin": list_tap_tin,
        }

        headers = dict(self.session.headers)
        headers.pop("Content-Type", None)

        files = file_parts + [
            (
                "ThuThapThongTinGiayChungNhan",
                (None, json.dumps(payload_json, ensure_ascii=False), "application/json"),
            )
        ]

        res = self.session.post(
            URL_UPLOAD_THU_THAP_GCN, files=files, headers=headers, timeout=TIMEOUT
        )

        self._check_response(res, "UPLOAD", URL_UPLOAD_THU_THAP_GCN)
        js = res.json()
        if not js.get("success"):
            raise RuntimeError("Upload lỗi: " + str(js)[:500])
        return js

    def lay_payload_duyet(self, thu_thap_id):
        payload = {
            "thuThapThongTinGiayChungNhanIds": [int(thu_thap_id)],
            "thayTheThuaDatKhongCoDon": False,
        }
        js = self.post_json(URL_BUILD_DUYET, payload, "BUILD DUYỆT")
        if not js.get("success"):
            raise RuntimeError("Lấy payload duyệt lỗi: " + str(js)[:500])

        value = js.get("value")
        if not isinstance(value, dict):
            raise RuntimeError("Không có value object để duyệt.")

        payload_duyet = convert_dates_recursive(copy.deepcopy(value))
        payload_duyet = add_id_recursive(payload_duyet)
        if payload_duyet.get("doiTuongThayTheKeys") is None:
            payload_duyet["doiTuongThayTheKeys"] = []
        return payload_duyet

    def check_duyet(self, payload_duyet):
        js = self.post_json(URL_CHECK_DUYET, payload_duyet, "CHECK DUYỆT")
        if not js.get("success"):
            raise RuntimeError("Check duyệt lỗi: " + str(js)[:500])
        if js.get("hasExistSTTHSTTDK") is True:
            raise RuntimeError("Đã tồn tại STT hồ sơ/thông tin đăng ký.")
        if js.get("taiSanDaDuocDangKy") is True:
            raise RuntimeError("Tài sản đã được đăng ký.")
        return js

    def save_duyet(self, payload_duyet):
        payload = copy.deepcopy(payload_duyet)
        payload["thayTheThuaDatKhongCoDon"] = False
        js = self.post_json(URL_SAVE_DUYET, payload, "SAVE DUYỆT")
        if not js.get("success"):
            raise RuntimeError("Duyệt vận hành lỗi: " + str(js)[:800])
        return js

    def xu_ly_1_ban_ghi(self, xa_id, so_to, so_thua, so_phat_hanh, folder_pdf, ds_ten_mo_ta, ds_ten_file, upload_hsq=True):
        danh_sach = []
        if upload_hsq:
            # Ghép cặp Tên mô tả ↔ Tên file theo thứ tự
            if not ds_ten_file:
                raise RuntimeError("Cột 'Tên file' rỗng.")
            if not ds_ten_mo_ta:
                # Không có mô tả → dùng chính tên file làm tenTapTin
                ds_ten_mo_ta = list(ds_ten_file)
            if len(ds_ten_mo_ta) != len(ds_ten_file):
                raise RuntimeError(
                    f"Số lượng 'Tên mô tả' ({len(ds_ten_mo_ta)}) khác số lượng 'Tên file' ({len(ds_ten_file)})."
                )

            # Bỏ đuôi .pdf khỏi tên mô tả (nếu người dùng lỡ gõ kèm)
            ds_ten_mo_ta = [re.sub(r"\.pdf$", "", mt, flags=re.IGNORECASE) for mt in ds_ten_mo_ta]

            danh_sach = [
                (os.path.join(folder_pdf, ten_file), ten_mo_ta)
                for ten_mo_ta, ten_file in zip(ds_ten_mo_ta, ds_ten_file)
            ]

        found = self.tra_cuu_thu_thap_gcn(
            xa_id=xa_id, so_phat_hanh=so_phat_hanh, so_to=so_to, so_thua=so_thua
        )
        thu_thap_id = found["thu_thap_id"]
        thdk_id = found["tinh_hinh_dang_ky_id"]
        rec_type = found["type"]
        self.log(
            f"   → Tìm thấy thuThapId={thu_thap_id}, soPhatHanh={found['so_phat_hanh']}, "
            f"tinhHinhDangKyId={thdk_id}, type={rec_type}"
        )

        # Chỉ duyệt khi tinhHinhDangKyId = 0 (chưa duyệt vận hành).
        # tinhHinhDangKyId != 0 → đã được duyệt vào vận hành rồi → bỏ qua.
        thdk_id_num = int(thdk_id) if thdk_id not in (None, "") else 0
        if thdk_id_num != 0:
            raise BoQuaBanGhi(
                f"Đã duyệt vào vận hành (tinhHinhDangKyId={thdk_id})"
            )

        if upload_hsq:
            self.upload_files(thu_thap_id, danh_sach)
            self.log(f"   → Upload {len(danh_sach)} hồ sơ quét OK")
        else:
            self.log("   → Bỏ qua upload hồ sơ quét (đã tắt), chỉ duyệt vận hành")

        payload_duyet = self.lay_payload_duyet(thu_thap_id)

        if KIEM_TRA_TRUOC_KHI_DUYET:
            self.check_duyet(payload_duyet)
            self.log("   → Check duyệt OK")

        self.save_duyet(payload_duyet)
        self.log("   → ✅ Duyệt vận hành thành công")


# ============================ TKINTER UI ============================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Duyệt vận hành MPLIS hàng loạt")
        root.geometry("860x640")

        self.client = MplisClient(self.log)
        self.df = None
        self.running = False
        self.stop_flag = False

        frm = ttk.Frame(root, padding=10)
        frm.pack(fill="x")

        # Username / password
        ttk.Label(frm, text="Username:").grid(row=0, column=0, sticky="w")
        self.ent_user = ttk.Entry(frm, width=30)
        self.ent_user.grid(row=0, column=1, sticky="w", padx=5, pady=3)

        ttk.Label(frm, text="Password:").grid(row=0, column=2, sticky="w")
        self.ent_pass = ttk.Entry(frm, width=30, show="*")
        self.ent_pass.grid(row=0, column=3, sticky="w", padx=5, pady=3)

        # Folder hồ sơ quét
        ttk.Label(frm, text="Folder hồ sơ quét:").grid(row=1, column=0, sticky="w")
        self.var_folder = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_folder, width=60).grid(
            row=1, column=1, columnspan=2, sticky="we", padx=5, pady=3
        )
        ttk.Button(frm, text="Chọn folder...", command=self.chon_folder).grid(row=1, column=3, sticky="w")

        # Excel
        ttk.Label(frm, text="File Excel:").grid(row=2, column=0, sticky="w")
        self.var_excel = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_excel, width=60).grid(
            row=2, column=1, columnspan=2, sticky="we", padx=5, pady=3
        )
        ttk.Button(frm, text="Chọn Excel...", command=self.chon_excel).grid(row=2, column=3, sticky="w")

        # Bật/tắt upload hồ sơ quét
        self.var_upload_hsq = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frm,
            text="Đẩy hồ sơ quét trước khi duyệt vận hành (tắt = chỉ duyệt vận hành)",
            variable=self.var_upload_hsq,
        ).grid(row=3, column=0, columnspan=4, sticky="w", pady=(5, 0))

        # Buttons
        btn_frm = ttk.Frame(root, padding=(10, 0))
        btn_frm.pack(fill="x")

        self.btn_login = ttk.Button(btn_frm, text="1. Mở Chrome đăng nhập", command=self.mo_chrome)
        self.btn_login.pack(side="left", padx=5)

        self.btn_confirm = ttk.Button(
            btn_frm, text="2. Đã đăng nhập xong → Lấy session", command=self.lay_session, state="disabled"
        )
        self.btn_confirm.pack(side="left", padx=5)

        self.btn_run = ttk.Button(btn_frm, text="3. Chạy duyệt hàng loạt", command=self.chay, state="disabled")
        self.btn_run.pack(side="left", padx=5)

        self.btn_stop = ttk.Button(btn_frm, text="Dừng", command=self.dung, state="disabled")
        self.btn_stop.pack(side="left", padx=5)

        # Progress
        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(8, 0))
        self.lbl_status = ttk.Label(root, text="Chưa chạy")
        self.lbl_status.pack(anchor="w", padx=10)

        # Log box
        self.txt = tk.Text(root, wrap="word", height=24)
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

    def chon_folder(self):
        d = filedialog.askdirectory(title="Chọn folder chứa hồ sơ quét")
        if d:
            self.var_folder.set(d)

    def chon_excel(self):
        f = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tất cả", "*.*")],
        )
        if f:
            self.var_excel.set(f)

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
        self.log("⏸ Đã yêu cầu dừng, sẽ dừng sau bản ghi hiện tại...")

    def doc_excel(self, path):
        df = pd.read_excel(path, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]

        missing = [c for c in REQUIRED_COLS if c not in df.columns]
        if missing:
            raise RuntimeError(f"Excel thiếu cột: {', '.join(missing)}. Cần đủ: {', '.join(REQUIRED_COLS)}")

        # Cột tùy chọn: nếu không có thì tự thêm cột rỗng
        for c in OPTIONAL_COLS:
            if c not in df.columns:
                df[c] = ""

        # Chuẩn hóa dữ liệu
        for c in REQUIRED_COLS + OPTIONAL_COLS:
            df[c] = df[c].map(clean_cell)

        # Bỏ dòng không có Số phát hành
        df = df[df[COL_SO_PHAT_HANH] != ""].copy()

        # Dedupe theo Số phát hành: giữ bản ghi đầu tiên
        truoc = len(df)
        df = df.drop_duplicates(subset=[COL_SO_PHAT_HANH], keep="first").reset_index(drop=True)
        sau = len(df)
        if truoc != sau:
            self.log(f"Đã gộp {truoc - sau} bản ghi trùng 'Số phát hành' (giữ bản ghi đầu tiên). Còn {sau} bản ghi.")
        return df

    def chay(self):
        if self.running:
            return

        folder = self.var_folder.get().strip()
        excel = self.var_excel.get().strip()
        upload_hsq = self.var_upload_hsq.get()

        if upload_hsq and (not folder or not os.path.isdir(folder)):
            messagebox.showwarning("Thiếu folder", "Đang bật đẩy hồ sơ quét — chọn folder hồ sơ quét hợp lệ.")
            return
        if not excel or not os.path.isfile(excel):
            messagebox.showwarning("Thiếu Excel", "Chọn file Excel hợp lệ.")
            return

        try:
            df = self.doc_excel(excel)
        except Exception as e:
            messagebox.showerror("Lỗi Excel", str(e))
            return

        if df.empty:
            messagebox.showwarning("Excel rỗng", "Không có bản ghi nào để xử lý.")
            return

        che_do = "ĐẨY HỒ SƠ QUÉT + DUYỆT VẬN HÀNH" if upload_hsq else "CHỈ DUYỆT VẬN HÀNH (không đẩy hồ sơ quét)"
        if not messagebox.askyesno(
            "Xác nhận",
            f"Chế độ: {che_do}\n"
            f"Sẽ xử lý {len(df)} bản ghi (sau khi gộp trùng Số phát hành).\n"
            f"Thao tác DUYỆT VẬN HÀNH không dễ hoàn tác. Tiếp tục?",
        ):
            return

        self.running = True
        self.stop_flag = False
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progress.config(maximum=len(df), value=0)

        threading.Thread(target=self._run_batch, args=(df, folder, excel, upload_hsq), daemon=True).start()

    def _run_batch(self, df, folder, excel_path, upload_hsq):
        results = []
        total = len(df)

        for i, row in df.iterrows():
            if self.stop_flag:
                self.log("⏹ Đã dừng theo yêu cầu.")
                break

            sph = row[COL_SO_PHAT_HANH]
            self.set_status(f"Đang xử lý {i + 1}/{total}: {sph}")
            self.log(f"[{i + 1}/{total}] Số phát hành: {sph} | Tờ {row[COL_SO_TO]} Thửa {row[COL_SO_THUA]} | Xã {row[COL_MA_XA]}")

            ket_qua = "OK"
            loi = ""
            try:
                self.client.xu_ly_1_ban_ghi(
                    xa_id=row[COL_MA_XA],
                    so_to=row[COL_SO_TO],
                    so_thua=row[COL_SO_THUA],
                    so_phat_hanh=sph,
                    folder_pdf=folder,
                    ds_ten_mo_ta=tach_danh_sach(row[COL_TEN_MO_TA]),
                    ds_ten_file=tach_danh_sach(row[COL_TEN_FILE]),
                    upload_hsq=upload_hsq,
                )
            except BoQuaBanGhi as e:
                ket_qua = "BỎ QUA"
                loi = str(e)
                self.log(f"   → ⏭ Bỏ qua: {loi}")
            except Exception as e:
                ket_qua = "LỖI"
                loi = str(e)
                self.log(f"   → ❌ {loi}")

            results.append({
                COL_MA_XA: row[COL_MA_XA],
                COL_SO_TO: row[COL_SO_TO],
                COL_SO_THUA: row[COL_SO_THUA],
                COL_SO_PHAT_HANH: sph,
                COL_TEN_MO_TA: row[COL_TEN_MO_TA],
                COL_TEN_FILE: row[COL_TEN_FILE],
                "Kết quả": ket_qua,
                "Lỗi": loi,
            })

            self.root.after(0, lambda v=i + 1: self.progress.config(value=v))

        # Xuất kết quả
        try:
            out_path = os.path.join(
                os.path.dirname(excel_path),
                f"ket_qua_duyet_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            )
            pd.DataFrame(results).to_excel(out_path, index=False)
            self.log(f"📄 Đã xuất file kết quả: {out_path}")
        except Exception as e:
            self.log(f"⚠ Không xuất được file kết quả: {e}")

        ok = sum(1 for r in results if r["Kết quả"] == "OK")
        bo_qua = sum(1 for r in results if r["Kết quả"] == "BỎ QUA")
        loi = sum(1 for r in results if r["Kết quả"] == "LỖI")
        self.log(f"===== HOÀN TẤT: {ok} thành công | {bo_qua} bỏ qua (đã duyệt) | {loi} lỗi / tổng {len(results)} =====")
        self.set_status(f"Hoàn tất: {ok} OK | {bo_qua} bỏ qua | {loi} lỗi / {len(results)}")

        self.running = False
        self.root.after(0, lambda: (self.btn_run.config(state="normal"), self.btn_stop.config(state="disabled")))

    def on_close(self):
        if self.running and not messagebox.askyesno("Đang chạy", "Đang xử lý, thoát luôn?"):
            return
        self.client.close_browser()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()