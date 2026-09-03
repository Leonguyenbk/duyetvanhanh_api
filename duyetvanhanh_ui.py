# -*- coding: utf-8 -*-
"""
Tool duyệt vận hành hàng loạt MPLIS
- Tkinter UI: nhập username/password, chọn folder hồ sơ quét, chọn file Excel,
  checkbox bật/tắt đẩy hồ sơ quét (tắt = chỉ duyệt vận hành)
- Excel BẮT BUỘC các cột: "Mã xã", "Số phát hành", "Tên mô tả", "Tên file"
  (tùy chọn: "Số tờ", "Số thửa" — chỉ để hiển thị/log)
- Tra cứu theo Số phát hành + Mã xã; chỉ duyệt khi tinhHinhDangKyId = 0,
  khác 0 nghĩa là đã duyệt vận hành → bỏ qua
- "Tên mô tả" = tenTapTin; nhiều tập tin trong 1 ô cách nhau bởi , hoặc ;
  ghép cặp với "Tên file" theo thứ tự; đuôi .pdf trong mô tả tự bị cắt;
  hậu tố cuối mô tả quyết định loaiHoSoQuet: -GT=0, -GCN=1, -DDK=2, -TBXN=3
- Các bản ghi trùng "Số phát hành" chỉ xử lý 1 bản ghi đại diện (dòng đầu tiên)
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
from datetime import datetime, timezone, timedelta

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


# Mốc thời gian → loaiGiayChungNhanId (ngày theo giờ Việt Nam)
BANG_LOAI_GCN = [
    # (từ ngày,            đến ngày,             ma_loai, tên)
    (datetime(1993, 10, 15), datetime(2004, 6, 30), 2,  "GCN QSDĐ theo Luật Đất Đai 1993"),
    (datetime(2004, 7, 1),  datetime(2009, 12, 9),  1,  "GCN QSDĐ theo Luật Đất Đai 2003"),
    (datetime(2009, 12, 10), datetime(2014, 6, 30), 6,  "GCN theo NĐ 88/2009/NĐ-CP"),
    (datetime(2014, 7, 1),  datetime(2024, 7, 31),  11, "GCN theo NĐ 43/2014/NĐ-CP"),
    (datetime(2024, 8, 1),  datetime(9999, 12, 31), 98, "GCN theo Luật Đất đai 2024"),
]


def tinh_loai_gcn_theo_ngay(iso_str):
    """
    Nhận chuỗi ISO UTC (vd '2018-12-04T17:00:00.000Z'), đổi sang giờ VN (+7)
    rồi tra bảng mốc thời gian → trả về loaiGiayChungNhanId, hoặc None nếu không xác định.
    """
    if not iso_str or not isinstance(iso_str, str):
        return None
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})", iso_str)
    if not m:
        return None
    y, mo, d, h, mi = map(int, m.groups())
    if y <= 1:  # ngày 0001-01-01 = không có ngày
        return None
    dt_utc = datetime(y, mo, d, h, mi)
    dt_vn = dt_utc + timedelta(hours=7)  # UTC → giờ Việt Nam
    ngay = datetime(dt_vn.year, dt_vn.month, dt_vn.day)
    for tu_ngay, den_ngay, ma_loai, _ten in BANG_LOAI_GCN:
        if tu_ngay <= ngay <= den_ngay:
            return ma_loai
    return None


def gan_loai_giay_chung_nhan(obj, log_fn):
    """
    Duyệt đệ quy payload: dict nào có key 'loaiGiayChungNhanId' đang trống (None/0)
    thì tính từ ngayCapGiay (ưu tiên) hoặc ngayVaoSo rồi gán vào.
    Trả về số chỗ đã gán.
    """
    so_gan = 0
    if isinstance(obj, dict):
        if "loaiGiayChungNhanId" in obj and not obj.get("loaiGiayChungNhanId"):
            ngay = obj.get("ngayCapGiay") or obj.get("ngayVaoSo")
            loai = tinh_loai_gcn_theo_ngay(ngay)
            if loai is not None:
                obj["loaiGiayChungNhanId"] = loai
                so_gan += 1
                log_fn(f"      • Gán loaiGiayChungNhanId={loai} (theo ngày {str(ngay)[:10]}, soPhatHanh={obj.get('soPhatHanh')})")
            else:
                log_fn(f"      • ⚠ Không xác định được loaiGiayChungNhanId (không có ngày cấp/vào sổ, soPhatHanh={obj.get('soPhatHanh')})")
        for v in obj.values():
            so_gan += gan_loai_giay_chung_nhan(v, log_fn)
    elif isinstance(obj, list):
        for x in obj:
            so_gan += gan_loai_giay_chung_nhan(x, log_fn)
    return so_gan


def xac_dinh_loai_ho_so_quet(ten_mo_ta):
    """
    Xác định loaiHoSoQuet theo HẬU TỐ cuối tên mô tả (sau dấu - hoặc _):
      -GT = 0 (giấy tờ, cho phép kèm số thứ tự: -GT1, -GT2...),
      -GCN = 1 (giấy chứng nhận), -DDK = 2 (đơn đăng ký), -TBXN = 3 (thông báo xác nhận).
    Bỏ qua đuôi .pdf nếu có. Không nhận diện được → mặc định 2 (DDK).
    """
    s = (ten_mo_ta or "").strip().upper()
    s = re.sub(r"\.PDF$", "", s)  # bỏ đuôi .pdf nếu có
    m = re.search(r"[-_](GT\d*|GCN|DDK|TBXN)$", s)
    if not m:
        return 2
    hau_to = m.group(1)
    if hau_to.startswith("GT"):
        return 0
    return {"GCN": 1, "DDK": 2, "TBXN": 3}[hau_to]


def ngay_vn_sang_iso(ddmmyyyy):
    """
    Đổi ngày nhập dd/mm/yyyy (hiểu là 00:00 giờ VN) sang chuỗi ISO UTC.
    Ví dụ: 24/04/2026 → '2026-04-23T17:00:00.000Z' (lùi 7 tiếng).
    """
    s = (ddmmyyyy or "").strip()
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$", s)
    if not m:
        raise ValueError(f"Ngày không đúng định dạng dd/mm/yyyy: '{ddmmyyyy}'")
    d, mo, y = map(int, m.groups())
    dt_vn = datetime(y, mo, d)  # 00:00 giờ VN
    dt_utc = dt_vn - timedelta(hours=7)
    return dt_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")


CHU_SU_DUNG_JSON = "chu_su_dung.json"  # nằm cùng thư mục với script


def nap_chu_su_dung(chu_su_dung_id, duong_dan_json=None):
    """
    Đọc file JSON chứa thông tin chủ, trả về NGUYÊN object ChuSoHuu
    (gồm CaNhans, ToChucs, HoGiaDinhs...) ứng với toChucId.
    duong_dan_json trống → mặc định dùng chu_su_dung.json cùng thư mục script.
    """
    duong_dan = duong_dan_json or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), CHU_SU_DUNG_JSON
    )
    if not os.path.isfile(duong_dan):
        raise RuntimeError(
            f"Không thấy file {duong_dan}. "
            f"Tạo file này: key = toChucId, value = nguyên object ChuSoHuu "
            f"bắt từ request SaveThongTinDangKy thật trên web."
        )
    with open(duong_dan, "r", encoding="utf-8") as f:
        data = json.load(f)
    key = str(chu_su_dung_id)
    if key not in data:
        raise RuntimeError(
            f"toChucId {chu_su_dung_id} chưa có trong {os.path.basename(duong_dan)}. "
            f"Các id hiện có: {', '.join(data.keys()) or '(trống)'}."
        )
    tpl = data[key]
    to_chucs = tpl.get("ToChucs") if isinstance(tpl, dict) else None
    if not isinstance(to_chucs, list) or not to_chucs or int(to_chucs[0].get("toChucId", -1)) != int(chu_su_dung_id):
        raise RuntimeError(
            f"Object của id {chu_su_dung_id} trong {os.path.basename(duong_dan)} không hợp lệ: "
            f"phải là nguyên ChuSoHuu có ToChucs[0].toChucId = {chu_su_dung_id}."
        )
    return tpl


def ap_dung_tuy_chinh_payload(payload, thoi_diem_iso, co_quyen_quan_ly, chu_template, log_fn):
    """
    Duyệt đệ quy payload duyệt:
    - dict key 'TinhHinhDangKy' → gán thoiDiemDangKy + coQuyenQuanLy (nếu được cung cấp)
    - dict key 'ChuSoHuu' → THAY NGUYÊN OBJECT bằng chu_template (nếu được cung cấp),
      đảm bảo thông tin chủ đúng và đầy đủ theo dữ liệu chuẩn đã lưu.
    """
    def _ten_chu(cso):
        if not isinstance(cso, dict):
            return "?"
        cac_ten = []
        for loai in ("ToChucs", "CaNhans", "HoGiaDinhs", "CongDongs", "NhomNguois", "VoChongs"):
            for item in (cso.get(loai) or []):
                if isinstance(item, dict):
                    cac_ten.append(str(item.get("tenToChuc") or item.get("hoTen") or "?"))
        return ", ".join(cac_ten) or "(trống)"

    def _walk(obj):
        if isinstance(obj, dict):
            for k, v in obj.items():
                if k == "TinhHinhDangKy" and isinstance(v, dict):
                    if thoi_diem_iso:
                        v["thoiDiemDangKy"] = thoi_diem_iso
                        log_fn(f"      • TinhHinhDangKy.thoiDiemDangKy = {thoi_diem_iso}")
                    if co_quyen_quan_ly is not None:
                        v["coQuyenQuanLy"] = bool(co_quyen_quan_ly)
                        log_fn(f"      • TinhHinhDangKy.coQuyenQuanLy = {bool(co_quyen_quan_ly)}")
                if k == "ChuSoHuu" and isinstance(v, dict) and chu_template:
                    cu = _ten_chu(v)
                    obj[k] = copy.deepcopy(chu_template)
                    log_fn(f"      • Thay nguyên ChuSoHuu: [{cu}] → [{_ten_chu(chu_template)}]")
                    continue  # không đi sâu vào object vừa thay
                _walk(v)
        elif isinstance(obj, list):
            for x in obj:
                _walk(x)

    _walk(payload)


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

    def xu_ly_1_ban_ghi(self, xa_id, so_to, so_thua, so_phat_hanh, folder_pdf, ds_ten_mo_ta, ds_ten_file,
                        upload_hsq=True, thoi_diem_iso=None, co_quyen_quan_ly=None, chu_template=None):
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

        # Gán loaiGiayChungNhanId theo mốc thời gian ngày cấp/vào sổ
        # (cần khi hồ sơ có giấy chứng nhận -GCN; chỉ điền vào chỗ đang trống)
        so_gan = gan_loai_giay_chung_nhan(payload_duyet, self.log)
        if so_gan:
            self.log(f"   → Đã gán loaiGiayChungNhanId cho {so_gan} giấy chứng nhận")

        # Áp tùy chỉnh từ UI: thời điểm đăng ký, coQuyenQuanLy, thông tin chủ sử dụng (tổ chức)
        if thoi_diem_iso or co_quyen_quan_ly is not None or chu_template:
            ap_dung_tuy_chinh_payload(
                payload_duyet, thoi_diem_iso, co_quyen_quan_ly, chu_template, self.log
            )

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

        # Thời điểm đăng ký + coQuyenQuanLy + id chủ sử dụng
        ttk.Label(frm, text="Thời điểm đăng ký (dd/mm/yyyy):").grid(row=4, column=0, sticky="w")
        self.ent_thoi_diem = ttk.Entry(frm, width=15)
        self.ent_thoi_diem.grid(row=4, column=1, sticky="w", padx=5, pady=3)

        self.var_co_quyen_quan_ly = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frm, text="coQuyenQuanLy", variable=self.var_co_quyen_quan_ly
        ).grid(row=4, column=2, sticky="w", padx=5)

        ttk.Label(frm, text="ID chủ sử dụng (toChucId):").grid(row=5, column=0, sticky="w")
        self.ent_chu_su_dung = ttk.Entry(frm, width=15)
        self.ent_chu_su_dung.grid(row=5, column=1, sticky="w", padx=5, pady=3)
        ttk.Label(frm, text="(để trống = giữ nguyên payload)").grid(row=5, column=2, columnspan=2, sticky="w")

        # File JSON chứa thông tin chủ
        ttk.Label(frm, text="File JSON chủ sử dụng:").grid(row=6, column=0, sticky="w")
        self.var_json_chu = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_json_chu, width=60).grid(
            row=6, column=1, columnspan=2, sticky="we", padx=5, pady=3
        )
        ttk.Button(frm, text="Chọn JSON...", command=self.chon_json_chu).grid(row=6, column=3, sticky="w")

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

    def chon_json_chu(self):
        f = filedialog.askopenfilename(
            title="Chọn file JSON chứa thông tin chủ sử dụng",
            filetypes=[("JSON", "*.json"), ("Tất cả", "*.*")],
        )
        if f:
            self.var_json_chu.set(f)

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
        co_quyen_quan_ly = self.var_co_quyen_quan_ly.get()

        # Thời điểm đăng ký (bắt buộc nhập, dd/mm/yyyy)
        try:
            thoi_diem_iso = ngay_vn_sang_iso(self.ent_thoi_diem.get())
        except ValueError as e:
            messagebox.showwarning("Sai ngày", str(e))
            return

        # ID chủ sử dụng (tùy chọn): nạp object chuẩn từ file JSON
        chu_su_dung_id = self.ent_chu_su_dung.get().strip()
        chu_template = None
        if chu_su_dung_id:
            if not chu_su_dung_id.isdigit():
                messagebox.showwarning("Sai ID", "ID chủ sử dụng (toChucId) phải là số.")
                return
            try:
                json_path = self.var_json_chu.get().strip() or None
                chu_template = nap_chu_su_dung(int(chu_su_dung_id), json_path)
            except Exception as e:
                messagebox.showerror("Lỗi chủ sử dụng", str(e))
                return

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
            f"Thời điểm đăng ký: {self.ent_thoi_diem.get().strip()} (→ {thoi_diem_iso})\n"
            f"coQuyenQuanLy: {co_quyen_quan_ly}\n"
            f"Chủ sử dụng: {chu_template['ToChucs'][0].get('tenToChuc') + ' (id ' + chu_su_dung_id + ')' if chu_template else '(giữ nguyên)'}\n"
            f"Sẽ xử lý {len(df)} bản ghi (sau khi gộp trùng Số phát hành).\n"
            f"Thao tác DUYỆT VẬN HÀNH không dễ hoàn tác. Tiếp tục?",
        ):
            return

        self.running = True
        self.stop_flag = False
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progress.config(maximum=len(df), value=0)

        threading.Thread(
            target=self._run_batch,
            args=(df, folder, excel, upload_hsq, thoi_diem_iso, co_quyen_quan_ly, chu_template),
            daemon=True,
        ).start()

    def _run_batch(self, df, folder, excel_path, upload_hsq, thoi_diem_iso, co_quyen_quan_ly, chu_template):
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
                    thoi_diem_iso=thoi_diem_iso,
                    co_quyen_quan_ly=co_quyen_quan_ly,
                    chu_template=chu_template,
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