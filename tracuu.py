# -*- coding: utf-8 -*-
"""
Script dòng lệnh: tra cứu Tình hình đăng ký MPLIS theo danh sách Số tờ/Số thửa
  HOẶC theo danh sách tinhHinhDangKyId có sẵn (tự phát hiện qua cột Excel)
- Nếu Excel có cột "tinhHinhDangKyId"/"tinhHinhDangKyIds" (ID có sẵn, mỗi ô có thể
  chứa nhiều ID cách nhau bởi dấu , ; hoặc khoảng trắng): BỎ QUA bước tra Số tờ/Số
  thửa, lấy thẳng các ID đó để gọi GetThongTinDangKyByTinhHinhDangKyIds.
- Nếu không có cột đó: Excel đầu vào BẮT BUỘC các cột "Mã xã", "Số tờ", "Số thửa"
  (Mã xã có thể khác nhau theo từng dòng, không cố định)
- Với mỗi dòng (chế độ Số tờ/Số thửa), gọi AdvancedSearchTinhHinhDangKy để lấy:
  maDon, soThuTu, dieuKienCapGiay, ngayTiepNhan, thoiDiemDangKyLanDau, thoiDiemDangKy,
  và tinhHinhDangKyId (bổ sung để phân biệt khi 1 thửa có nhiều bản ghi)
- Nếu 1 dòng tra ra nhiều bản ghi thì xuất mỗi bản ghi thành 1 dòng kết quả
- Kết quả được gộp theo Mã xã, mỗi xã xuất 1 file: "<MaXa>_<tên gốc>.xlsx"
  (tên gốc do người dùng nhập khi chạy), lưu cùng thư mục với Excel đầu vào
- KHÔNG dùng Selenium: tự đọc Cookie từ Chrome/Edge đang đăng nhập MPLIS sẵn
  trên máy rồi tự tải trang để bóc __RequestVerificationToken ra; nếu không
  đọc được thì hỏi dán Cookie + token tay (giống cách cũ) để dự phòng.
- Sau khi tra cứu xong, tự gọi GetThongTinDangKyByTinhHinhDangKyIds với toàn bộ
  tinhHinhDangKyId tìm được để lấy thông tin chi tiết đơn đăng ký + hồ sơ quét,
  lưu nguyên JSON trả về ra file "<tên gốc>_ThongTinChiTiet.json" (chưa xử lý
  cập nhật gì cả — bước cập nhật sẽ bổ sung sau khi biết rõ cấu trúc dữ liệu).

Chạy: python tra_cuu_tinh_hinh_dang_ky.py <duong_dan_excel.xlsx> [ten_goc]
Cài đặt: pip install requests pandas openpyxl browser_cookie3 pycryptodomex pywin32
"""

import os
import re
import sys
import json
import copy
import time
import threading
import traceback
import unicodedata
from datetime import datetime, timezone, timedelta
from http.cookies import SimpleCookie

import requests
import pandas as pd
import browser_cookie3

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ============================ CẤU HÌNH ============================

REFERER_URL = "https://dla.mplis.gov.vn/dc/DonDangKy/KeKhaiDangKyV2"
URL_ADVANCED_SEARCH_THDK = "https://dla.mplis.gov.vn/dc/DangKyAjax/AdvancedSearchTinhHinhDangKy"
URL_GET_THONG_TIN_DANG_KY = "https://dla.mplis.gov.vn/dc/DangKyAjax/GetThongTinDangKyByTinhHinhDangKyIds"
URL_UPDATE_HO_SO_QUET = "https://dla.mplis.gov.vn/dc/HoSoQuetAjax/UpdateHoSoQuetExistFile"
URL_UPDATE_THONG_TIN_DANG_KY = "https://dla.mplis.gov.vn/dc/DangKyAjax/UpdateThongTinDangKy"
URL_GUI_YEU_CAU_PHAN_LOAI_LAI = "https://dla.mplis.gov.vn/dc/LamSachDuLieuAjax/GuiYeuCauPhanLoaiLai"

TIMEOUT = 120
PAGE_LENGTH = 100  # đủ lớn để lấy hết bản ghi khớp 1 tờ/thửa trong 1 lần gọi
REQUEST_DELAY_SECONDS = 0.15
THONG_TIN_CHUNK_SIZE = 200  # số tinhHinhDangKyId tối đa gửi trong 1 lần gọi GetThongTinDangKyByTinhHinhDangKyIds
SAVE_EVERY_ROWS = 30  # tự lưu Excel kết quả sau mỗi 30 bản ghi, tránh mất tiến trình khi file lớn

COL_MA_XA = "Mã xã"
COL_SO_TO = "Số tờ"
COL_SO_THUA = "Số thửa"
REQUIRED_COLS = [COL_MA_XA, COL_SO_TO, COL_SO_THUA]

RESULT_HEADERS = [
    "STT", "Mã xã", "Số tờ", "Số thửa", "tinhHinhDangKyId", "maDon", "soThuTu",
    "dieuKienCapGiay", "ngayTiepNhan", "thoiDiemDangKyLanDau", "thoiDiemDangKy",
    "Kết quả", "Chi tiết",
]

DIEU_KIEN_CAP_GIAY_MAP = {
    0: "Đủ điều kiện cấp giấy",
    1: "Không đủ điều kiện cấp giấy",
    2: "Không có nhu cầu cấp giấy",
    3: "Chưa đủ điều kiện cấp giấy",
}


# ============================ HELPER ============================

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


EPOCH_UTC = datetime(1970, 1, 1, tzinfo=timezone.utc)


def _dotnet_ms_to_datetime_utc(ms):
    """epoch + timedelta thay vì datetime.fromtimestamp: fromtimestamp lỗi OSError trên Windows
    với mốc âm (ngày trước 1970, vd ngày sinh cũ), còn epoch+timedelta thì không giới hạn đó."""
    return EPOCH_UTC + timedelta(milliseconds=ms)


def dotnet_date_to_ddmmyyyy(value):
    """Đổi '/Date(ms)/' (giờ UTC) sang chuỗi dd/mm/yyyy theo giờ Việt Nam (+7)."""
    if not isinstance(value, str):
        return ""
    m = re.search(r"/Date\((-?\d+)\)/", value)
    if not m:
        return ""
    ms = int(m.group(1))
    if ms <= -62135596800000:  # mốc "chưa có ngày" của .NET
        return ""
    dt_vn = _dotnet_ms_to_datetime_utc(ms) + timedelta(hours=7)
    return dt_vn.strftime("%d/%m/%Y")


def dotnet_date_to_iso(value):
    """Đổi '/Date(ms)/' sang chuỗi ISO UTC 'YYYY-MM-DDTHH:MM:SS.sssZ' (dùng khi build payload cập nhật)."""
    if not isinstance(value, str):
        return value
    m = re.search(r"/Date\((-?\d+)\)/", value)
    if not m:
        return value
    ms = int(m.group(1))
    if ms <= -62135596800000:  # mốc "chưa có ngày" của .NET
        return "0001-01-01T00:00:00.000Z"
    dt = _dotnet_ms_to_datetime_utc(ms)
    return dt.strftime("%Y-%m-%dT%H:%M:%S.") + f"{int(dt.microsecond / 1000):03d}Z"


def convert_dates_recursive(obj):
    if isinstance(obj, dict):
        return {k: convert_dates_recursive(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [convert_dates_recursive(x) for x in obj]
    return dotnet_date_to_iso(obj)


def add_id_recursive(obj):
    """Gán _id (thứ tự 1..n) cho mọi dict trong các list con — API cập nhật cần field này."""
    if isinstance(obj, list):
        for i, item in enumerate(obj, start=1):
            if isinstance(item, dict):
                item.setdefault("_id", i)
                add_id_recursive(item)
    elif isinstance(obj, dict):
        for v in obj.values():
            add_id_recursive(v)
    return obj


def build_advanced_search_payload(xa_id, so_to, so_thua, length=PAGE_LENGTH):
    """Payload form-urlencoded giống đúng request DataTables của AdvancedSearchTinhHinhDangKy."""
    columns = [
        ("", "", True, False),
        ("tinhHinhDangKyId", "tinhHinhDangKyId", True, True),
        ("maDon", "maDon", True, True),
        ("soThuTu", "soThuTu", True, True),
        ("DaiDienKhaiTrinh", "DaiDienKhaiTrinh", True, False),
        ("ngayTiepNhan", "ngayTiepNhan", True, True),
        ("thoiDiemDangKy", "thoiDiemDangKy", True, True),
    ]

    payload = {"draw": "1"}
    for i, (data, name, searchable, orderable) in enumerate(columns):
        payload[f"columns[{i}][data]"] = data
        payload[f"columns[{i}][name]"] = name
        payload[f"columns[{i}][searchable]"] = "true" if searchable else "false"
        payload[f"columns[{i}][orderable]"] = "true" if orderable else "false"
        payload[f"columns[{i}][search][value]"] = ""
        payload[f"columns[{i}][search][regex]"] = "false"

    payload.update({
        "order[0][column]": "5",
        "order[0][dir]": "desc",
        "start": "0",
        "length": str(length),
        "search[value]": "",
        "search[regex]": "false",
    })

    model_fields = [
        "huyenId", "tinhHinhDangKyId", "maDon", "soThuTu", "ngayTiepNhan",
        "thoiDiemDangKy", "loaiGiayChungNhanId", "soPhatHanh", "maVach",
        "soVaoSo", "soVaoSoCu", "ngayVaoSo", "soHoSoGoc", "soHoSoGocCu",
        "hoTen", "soGiayTo", "namSinh", "soThuTuThuaCu", "soHieuToBanDoCu",
        "soNha", "diaChiChiTiet", "dieuKienCapGiay",
    ]
    payload["model[xaId]"] = clean_cell(xa_id)
    for field in model_fields:
        payload[f"model[{field}]"] = ""
    payload["model[soThuTuThua]"] = clean_cell(so_thua)
    payload["model[soHieuToBanDo]"] = clean_cell(so_to)
    payload["model[phucHoiDuLieu]"] = "false"
    return payload


# ============================ LOGIN / SESSION ============================

DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

TOKEN_PATTERNS = [
    re.compile(r'name=["\']__RequestVerificationToken["\'][^>]*value=["\']([^"\']+)["\']', re.IGNORECASE),
    re.compile(r'value=["\']([^"\']+)["\'][^>]*name=["\']__RequestVerificationToken["\']', re.IGNORECASE),
    re.compile(r'name=["\']__RequestVerificationToken["\'][^>]*content=["\']([^"\']+)["\']', re.IGNORECASE),
]


def chia_nhom(danh_sach, kich_thuoc):
    for i in range(0, len(danh_sach), kich_thuoc):
        yield danh_sach[i:i + kich_thuoc]


def trich_token_tu_html(html):
    for pattern in TOKEN_PATTERNS:
        m = pattern.search(html)
        if m:
            return m.group(1)
    return None


def lay_cookiejar_tu_trinh_duyet(domain="mplis.gov.vn"):
    """Đọc cookie đã đăng nhập sẵn từ Chrome, rồi Edge nếu Chrome không có."""
    loi = []
    for ten, ham in (("Chrome", browser_cookie3.chrome), ("Edge", browser_cookie3.edge)):
        try:
            cookiejar = ham(domain_name=domain)
            if len(cookiejar) > 0:
                return cookiejar, ten
            loi.append(f"{ten}: không có cookie nào cho {domain}")
        except Exception as e:
            loi.append(f"{ten}: {e}")
    raise RuntimeError("Không đọc được cookie từ trình duyệt nào. " + "; ".join(loi))


class MplisClient:
    def __init__(self):
        self.session = None

    def build_session_from_browser_cookies(self, user_agent=None):
        """Tự lấy Cookie từ Chrome/Edge đang đăng nhập MPLIS, rồi tải trang để bóc token."""
        cookiejar, trinh_duyet = lay_cookiejar_tu_trinh_duyet()

        session = requests.Session()
        session.cookies.update(cookiejar)
        session.headers.update({
            "User-Agent": user_agent or DEFAULT_USER_AGENT,
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": "https://dla.mplis.gov.vn",
            "Referer": REFERER_URL,
        })

        res = session.get(REFERER_URL, timeout=TIMEOUT)
        res.raise_for_status()
        token = trich_token_tu_html(res.text)
        if not token:
            raise RuntimeError(
                "Không tìm thấy __RequestVerificationToken trong trang tải về "
                f"(cookie lấy từ {trinh_duyet}). Có thể chưa đăng nhập MPLIS trên trình duyệt đó, "
                "hoặc phiên đăng nhập đã hết hạn."
            )

        session.headers.update({
            "__requestverificationtoken": token,
            "__RequestVerificationToken": token,
            "RequestVerificationToken": token,
        })
        self.session = session
        return trinh_duyet

    def build_session_from_manual_input(self, cookie_str, token, user_agent=None):
        """Dựng session trực tiếp từ Cookie header và token dán tay (không Selenium)."""
        cookie_str = (cookie_str or "").strip()
        token = (token or "").strip()
        if not cookie_str:
            raise ValueError("Cookie không được rỗng.")
        if not token:
            raise ValueError("Token không được rỗng.")

        session = requests.Session()
        session.headers.update({
            "User-Agent": user_agent or DEFAULT_USER_AGENT,
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": "https://dla.mplis.gov.vn",
            "Referer": REFERER_URL,
            "__requestverificationtoken": token,
            "__RequestVerificationToken": token,
            "RequestVerificationToken": token,
        })

        parsed = SimpleCookie()
        parsed.load(cookie_str)
        if not parsed:
            raise ValueError("Không đọc được cookie nào từ chuỗi đã dán.")
        for name, morsel in parsed.items():
            session.cookies.set(name=name, value=morsel.value, domain="dla.mplis.gov.vn", path="/")

        self.session = session

    def tra_cuu_tinh_hinh_dang_ky(self, xa_id, so_to, so_thua):
        headers = dict(self.session.headers)
        headers["Content-Type"] = "application/x-www-form-urlencoded; charset=UTF-8"
        payload = build_advanced_search_payload(xa_id, so_to, so_thua)

        res = self.session.post(
            URL_ADVANCED_SEARCH_THDK, data=payload, headers=headers, timeout=TIMEOUT,
        )
        res.raise_for_status()
        ct = res.headers.get("Content-Type") or ""
        if "json" not in ct:
            raise RuntimeError(
                f"Server không trả JSON (status {res.status_code}), có thể mất session. "
                f"Nội dung: {res.text[:400]}"
            )

        js = res.json()
        rows = js.get("data")
        if not isinstance(rows, list):
            raise RuntimeError("Response không có danh sách 'data' hợp lệ: " + str(js)[:500])
        return rows

    def lay_thong_tin_dang_ky(self, tinh_hinh_dang_ky_ids, get_ho_so_quet=True):
        """Gọi GetThongTinDangKyByTinhHinhDangKyIds để lấy chi tiết đơn đăng ký + hồ sơ quét."""
        payload = {
            "tinhHinhDangKyIds": [int(x) for x in tinh_hinh_dang_ky_ids],
            "getHoSoQuet": bool(get_ho_so_quet),
        }
        headers = dict(self.session.headers)
        headers["Content-Type"] = "application/json; charset=UTF-8"

        res = self.session.post(
            URL_GET_THONG_TIN_DANG_KY,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers=headers,
            timeout=TIMEOUT,
        )
        res.raise_for_status()
        ct = res.headers.get("Content-Type") or ""
        if "json" not in ct:
            raise RuntimeError(
                f"Server không trả JSON (status {res.status_code}), có thể mất session. "
                f"Nội dung: {res.text[:400]}"
            )
        return res.json()

    def update_ho_so_quet_exist_file(self, payload):
        """Gọi UpdateHoSoQuetExistFile với payload dạng form (hoSoQuet, infoHoSoQuet_N, count, isLuuKhoHoSoQuet)."""
        headers = dict(self.session.headers)
        headers["Content-Type"] = "application/x-www-form-urlencoded; charset=UTF-8"
        res = self.session.post(URL_UPDATE_HO_SO_QUET, data=payload, headers=headers, timeout=TIMEOUT)
        res.raise_for_status()
        ct = res.headers.get("Content-Type") or ""
        if "json" not in ct:
            raise RuntimeError(
                f"Server không trả JSON (status {res.status_code}), có thể mất session. "
                f"Nội dung: {res.text[:400]}"
            )
        ket_qua = res.json()
        if isinstance(ket_qua, dict) and ket_qua.get("success") is False:
            raise RuntimeError("UpdateHoSoQuetExistFile thất bại: " + str(ket_qua)[:500])
        return ket_qua

    def update_thong_tin_dang_ky(self, payload):
        """Gọi UpdateThongTinDangKy với payload JSON đầy đủ (TinhHinhDangKy, ChuSoHuu, TaiSan, ...)."""
        headers = dict(self.session.headers)
        headers["Content-Type"] = "application/json; charset=UTF-8"
        res = self.session.post(
            URL_UPDATE_THONG_TIN_DANG_KY,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers=headers,
            timeout=TIMEOUT,
        )
        res.raise_for_status()
        ct = res.headers.get("Content-Type") or ""
        if "json" not in ct:
            raise RuntimeError(
                f"Server không trả JSON (status {res.status_code}), có thể mất session. "
                f"Nội dung: {res.text[:400]}"
            )
        ket_qua = res.json()
        if isinstance(ket_qua, dict) and ket_qua.get("success") is False:
            raise RuntimeError("UpdateThongTinDangKy thất bại: " + str(ket_qua)[:500])
        return ket_qua

    def gui_yeu_cau_phan_loai_lai(self, thua_dat_ids):
        """Gọi GuiYeuCauPhanLoaiLai với danh sách thuaDatId (bước cuối sau khi cập nhật xong)."""
        payload = {"thuaDatIds": [int(x) for x in thua_dat_ids]}
        headers = dict(self.session.headers)
        headers["Content-Type"] = "application/json; charset=UTF-8"
        res = self.session.post(
            URL_GUI_YEU_CAU_PHAN_LOAI_LAI,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers=headers,
            timeout=TIMEOUT,
        )
        res.raise_for_status()
        ct = res.headers.get("Content-Type") or ""
        if "json" not in ct:
            raise RuntimeError(
                f"Server không trả JSON (status {res.status_code}), có thể mất session. "
                f"Nội dung: {res.text[:400]}"
            )
        ket_qua = res.json()
        if isinstance(ket_qua, dict) and ket_qua.get("success") is False:
            raise RuntimeError("GuiYeuCauPhanLoaiLai thất bại: " + str(ket_qua)[:500])
        return ket_qua


# ============================ EXCEL ============================

def doc_excel(path):
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise RuntimeError(f"Excel thiếu cột: {', '.join(missing)}. Cần đủ: {', '.join(REQUIRED_COLS)}")

    for c in REQUIRED_COLS:
        df[c] = df[c].map(clean_cell)

    df = df[(df[COL_SO_TO] != "") | (df[COL_SO_THUA] != "")].reset_index(drop=True)
    if df.empty:
        raise RuntimeError("Excel không có dữ liệu Số tờ/Số thửa để xử lý.")
    return df


COL_TINH_HINH_DANG_KY_ID_ALIASES = {"tinhhinhdangkyid", "tinhhinhdangkyids"}


def chuan_hoa_ten_cot(v):
    return re.sub(r"[^a-z0-9]", "", str(v or "").strip().lower())


def tim_cot_tinh_hinh_dang_ky_id(columns):
    for c in columns:
        if chuan_hoa_ten_cot(c) in COL_TINH_HINH_DANG_KY_ID_ALIASES:
            return c
    return None


def xac_dinh_che_do(path):
    """Tự phát hiện chế độ: Excel có cột tinhHinhDangKyId(s) → tra thẳng theo ID,
    không có thì tra theo Mã xã/Số tờ/Số thửa như cũ. Trả về ("id", tên_cột) hoặc ("to_thua", None)."""
    header = pd.read_excel(path, nrows=0)
    cot_id = tim_cot_tinh_hinh_dang_ky_id([str(c).strip() for c in header.columns])
    if cot_id:
        return "id", cot_id
    return "to_thua", None


def doc_excel_ids(path, cot_id):
    """Đọc danh sách tinhHinhDangKyId từ 1 cột Excel (1 ô có thể chứa nhiều ID
    cách nhau bởi , ; hoặc khoảng trắng), trả về list int không trùng, giữ thứ tự."""
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    ids = []
    da_thay = set()
    for gia_tri in df[cot_id]:
        s = clean_cell(gia_tri)
        if not s:
            continue
        for phan in re.split(r"[,;\s]+", s):
            phan = phan.strip()
            if not phan:
                continue
            try:
                id_int = int(float(phan))
            except ValueError:
                continue
            if id_int not in da_thay:
                da_thay.add(id_int)
                ids.append(id_int)

    if not ids:
        raise RuntimeError(f"Không tìm thấy tinhHinhDangKyId hợp lệ trong cột '{cot_id}'.")
    return ids


def nhan_dieu_kien_cap_giay(value):
    if value is None or value == "":
        return ""
    try:
        key = int(value)
    except (TypeError, ValueError):
        return str(value)
    return DIEU_KIEN_CAP_GIAY_MAP.get(key, str(value))


def trich_dong_ket_qua(stt, ma_xa, so_to, so_thua, item):
    return {
        "STT": stt,
        "Mã xã": ma_xa,
        "Số tờ": so_to,
        "Số thửa": so_thua,
        "tinhHinhDangKyId": item.get("tinhHinhDangKyId"),
        "maDon": item.get("maDon"),
        "soThuTu": item.get("soThuTu"),
        "dieuKienCapGiay": nhan_dieu_kien_cap_giay(item.get("dieuKienCapGiay")),
        "ngayTiepNhan": dotnet_date_to_ddmmyyyy(item.get("ngayTiepNhan")),
        "thoiDiemDangKyLanDau": dotnet_date_to_ddmmyyyy(item.get("thoiDiemDangKyLanDau")),
        "thoiDiemDangKy": dotnet_date_to_ddmmyyyy(item.get("thoiDiemDangKy")),
        "Kết quả": "OK",
        "Chi tiết": "",
    }


def dong_trong(stt, ma_xa, so_to, so_thua, ket_qua, chi_tiet):
    return {
        "STT": stt, "Mã xã": ma_xa, "Số tờ": so_to, "Số thửa": so_thua,
        "tinhHinhDangKyId": "", "maDon": "", "soThuTu": "", "dieuKienCapGiay": "",
        "ngayTiepNhan": "", "thoiDiemDangKyLanDau": "", "thoiDiemDangKy": "",
        "Kết quả": ket_qua, "Chi tiết": chi_tiet,
    }


# ============================ XỬ LÝ CHUNG (dùng cho cả CLI và Tkinter) ============================

def trich_dong_tu_chi_tiet(item):
    """Trích 1 hoặc nhiều dòng kết quả (theo từng thửa đất trong TaiSan.ThuaDats) từ 1 item chi tiết
    trả về bởi GetThongTinDangKyByTinhHinhDangKyIds — dùng khi tra thẳng theo tinhHinhDangKyId
    (không đi qua bước AdvancedSearchTinhHinhDangKy nên không có sẵn Số tờ/Số thửa/Mã xã)."""
    thdk = item.get("TinhHinhDangKy") or {}
    ma_xa = thdk.get("xaId")
    thua_dats = [t for t in ((item.get("TaiSan") or {}).get("ThuaDats") or []) if isinstance(t, dict)]

    def _dong(so_to, so_thua, ma_xa_dong):
        return {
            "STT": "",
            "Mã xã": ma_xa_dong,
            "Số tờ": so_to,
            "Số thửa": so_thua,
            "tinhHinhDangKyId": thdk.get("tinhHinhDangKyId"),
            "maDon": thdk.get("maDon"),
            "soThuTu": thdk.get("soThuTu"),
            "dieuKienCapGiay": nhan_dieu_kien_cap_giay(thdk.get("dieuKienCapGiay")),
            "ngayTiepNhan": dotnet_date_to_ddmmyyyy(thdk.get("ngayTiepNhan")),
            "thoiDiemDangKyLanDau": dotnet_date_to_ddmmyyyy(thdk.get("thoiDiemDangKyLanDau")),
            "thoiDiemDangKy": dotnet_date_to_ddmmyyyy(thdk.get("thoiDiemDangKy")),
            "Kết quả": "OK",
            "Chi tiết": "",
        }

    if not thua_dats:
        return [_dong("", "", ma_xa)]
    return [
        _dong(t.get("soHieuToBanDo"), t.get("soThuTuThua"), t.get("xaId", ma_xa))
        for t in thua_dats
    ]


def xu_ly_theo_id(tinh_hinh_dang_ky_ids, ten_goc, output_dir, client, log_fn=print, cap_nhat_ngay=None):
    """Chế độ tinhHinhDangKyId có sẵn: lấy chi tiết theo từng nhóm ≤200 ID, xuất ngay ra 1 file
    Excel tổng hợp "<tên gốc>.xlsx" (tự lưu lại sau mỗi nhóm lấy được), đồng thời cập nhật ngay
    (nếu có cap_nhat_ngay) — trước đây chế độ này chỉ lưu JSON, không có Excel."""
    danh_sach = []
    out_path = os.path.join(output_dir, f"{ten_goc}.xlsx")

    def xu_ly_1_nhom(items_chunk):
        for item in items_chunk:
            for dong in trich_dong_tu_chi_tiet(item):
                dong["STT"] = len(danh_sach) + 1
                danh_sach.append(dong)
        pd.DataFrame(danh_sach, columns=RESULT_HEADERS).to_excel(out_path, index=False)
        log_fn(f"Đã lưu tạm kết quả ({len(danh_sach)} dòng): {out_path}")
        if cap_nhat_ngay:
            cap_nhat_ngay(items_chunk)

    lay_va_luu_thong_tin_chi_tiet(
        client, tinh_hinh_dang_ky_ids, output_dir, ten_goc, log_fn=log_fn, on_chunk_fn=xu_ly_1_nhom
    )

    if danh_sach:
        pd.DataFrame(danh_sach, columns=RESULT_HEADERS).to_excel(out_path, index=False)
        log_fn(f"Đã xuất file kết quả: {out_path}")
    return out_path


def tao_ham_cap_nhat_tang_dan(client, dieu_kien_cap_giay, log_fn, xac_nhan_fn):
    """Trả về (cap_nhat_ngay, lay_tong_ket).
    cap_nhat_ngay(items_chunk) dùng làm callback: cập nhật NGAY một nhóm bản ghi vừa lấy chi tiết
    xong (1 xã, hoặc 1 nhóm ID), thay vì đợi lấy chi tiết xong TOÀN BỘ rồi mới cập nhật 1 lần.
    Chỉ hỏi xác nhận (qua xac_nhan_fn) đúng 1 LẦN DUY NHẤT, trước lần cập nhật đầu tiên."""
    trang_thai = {"da_hoi": False, "tiep_tuc": True, "thanh_cong": 0, "loi": 0}

    def cap_nhat_ngay(items_chunk):
        if not dieu_kien_cap_giay or not items_chunk or not trang_thai["tiep_tuc"]:
            return
        if not trang_thai["da_hoi"]:
            trang_thai["da_hoi"] = True
            if xac_nhan_fn:
                trang_thai["tiep_tuc"] = xac_nhan_fn(dieu_kien_cap_giay)
            if not trang_thai["tiep_tuc"]:
                log_fn("Đã hủy bước cập nhật.")
                return
        log_fn(f"Đang cập nhật {len(items_chunk)} bản ghi...")
        tc, l = cap_nhat_hang_loat(client, items_chunk, dieu_kien_cap_giay, log_fn=log_fn)
        trang_thai["thanh_cong"] += tc
        trang_thai["loi"] += l

    def lay_tong_ket():
        return trang_thai["thanh_cong"], trang_thai["loi"]

    return cap_nhat_ngay, lay_tong_ket


def xu_ly_theo_xa(df, ten_goc, input_file, client, log_fn=print, progress_fn=None, cap_nhat_ngay=None):
    """Tra cứu từng dòng, xử lý THEO TỪNG MÃ XÃ MỘT: tra xong hết các dòng của 1 xã → xuất Excel
    xã đó → lấy thông tin chi tiết xã đó → gọi cap_nhat_ngay(items) cho xã đó ngay lập tức, rồi
    mới sang xã tiếp theo (không đợi tra cứu xong TOÀN BỘ file rồi mới cập nhật — quan trọng với
    file lớn hàng nghìn dòng). Tự lưu lại Excel của xã đang xử lý sau mỗi SAVE_EVERY_ROWS dòng để
    không mất tiến trình nếu chương trình bị dừng giữa chừng.
    Trả về list đường dẫn Excel đã xuất."""
    output_dir = os.path.dirname(os.path.abspath(input_file))
    tong = len(df)
    da_xu_ly = 0
    duong_dan_ket_qua = []

    for ma_xa, nhom_df in df.groupby(COL_MA_XA, sort=False):
        danh_sach = []
        ids_xa = []
        da_thay = set()
        ten_xa_an_toan = re.sub(r'[\\/*?:"<>|]', "_", str(ma_xa)) or "KhongRoXa"
        out_path = os.path.join(output_dir, f"{ten_xa_an_toan}_{ten_goc}.xlsx")

        for i, row in nhom_df.iterrows():
            so_to = row[COL_SO_TO]
            so_thua = row[COL_SO_THUA]
            stt = i + 1
            da_xu_ly += 1
            log_fn(f"[{da_xu_ly}/{tong}] Xã {ma_xa} - Tờ {so_to} - Thửa {so_thua}")
            try:
                if not ma_xa or not so_to or not so_thua:
                    raise ValueError("Thiếu Mã xã, Số tờ hoặc Số thửa.")
                rows = client.tra_cuu_tinh_hinh_dang_ky(ma_xa, so_to, so_thua)
                if not rows:
                    danh_sach.append(dong_trong(stt, ma_xa, so_to, so_thua, "KHÔNG TÌM THẤY", ""))
                    log_fn("   → Không tìm thấy bản ghi.")
                else:
                    for item in rows:
                        dong = trich_dong_ket_qua(stt, ma_xa, so_to, so_thua, item)
                        danh_sach.append(dong)
                        thdk_id = dong["tinhHinhDangKyId"]
                        if thdk_id not in (None, "") and thdk_id not in da_thay:
                            da_thay.add(thdk_id)
                            ids_xa.append(thdk_id)
                    log_fn(f"   → Tìm thấy {len(rows)} bản ghi.")
            except Exception as e:
                danh_sach.append(dong_trong(stt, ma_xa, so_to, so_thua, "LỖI", str(e)))
                log_fn(f"   → LỖI: {e}")

            if progress_fn:
                progress_fn(da_xu_ly, tong)

            if da_xu_ly % SAVE_EVERY_ROWS == 0:
                pd.DataFrame(danh_sach, columns=RESULT_HEADERS).to_excel(out_path, index=False)
                log_fn(f"   Đã tự lưu tạm sau {da_xu_ly} bản ghi: {out_path}")

            time.sleep(REQUEST_DELAY_SECONDS)

        pd.DataFrame(danh_sach, columns=RESULT_HEADERS).to_excel(out_path, index=False)
        log_fn(f"Đã xuất file kết quả cho xã {ma_xa}: {out_path}")
        duong_dan_ket_qua.append(out_path)

        if not ids_xa:
            continue

        try:
            _out_path, items_xa = lay_va_luu_thong_tin_chi_tiet(
                client, ids_xa, output_dir, f"{ten_xa_an_toan}_{ten_goc}", log_fn=log_fn
            )
        except Exception as e:
            log_fn(f"Lỗi khi lấy thông tin chi tiết xã {ma_xa}: {e}")
            continue

        if cap_nhat_ngay:
            cap_nhat_ngay(items_xa)

    return duong_dan_ket_qua


def lay_va_luu_thong_tin_chi_tiet(client, tinh_hinh_dang_ky_ids, output_dir, ten_goc, log_fn=print, on_chunk_fn=None):
    """Gọi GetThongTinDangKyByTinhHinhDangKyIds cho toàn bộ tinhHinhDangKyId đã tra được,
    lưu nguyên JSON trả về ra file (để tham chiếu/khôi phục sau này nếu cần), đồng thời trả về
    (duong_dan_file, list_item) — list_item dùng ngay cho bước cập nhật trong CÙNG 1 lần chạy,
    không cần đọc lại file. Nếu có on_chunk_fn(items_chunk) thì gọi ngay sau mỗi nhóm (≤200 ID)
    vừa lấy xong, để bước cập nhật có thể bắt đầu sớm thay vì đợi lấy hết toàn bộ."""
    if not tinh_hinh_dang_ky_ids:
        log_fn("Không có tinhHinhDangKyId nào để lấy thông tin chi tiết.")
        return None, []

    log_fn(f"Đang lấy thông tin chi tiết (đơn đăng ký + hồ sơ quét) cho {len(tinh_hinh_dang_ky_ids)} tinhHinhDangKyId...")
    tat_ca_ban_ghi = []
    cac_nhom = list(chia_nhom(tinh_hinh_dang_ky_ids, THONG_TIN_CHUNK_SIZE))
    for idx, nhom in enumerate(cac_nhom, start=1):
        ket_qua = client.lay_thong_tin_dang_ky(nhom, get_ho_so_quet=True)
        # Response thực tế có dạng {"value": [...], "success": true} — KHÔNG có key "data".
        du_lieu = ket_qua.get("value") if isinstance(ket_qua, dict) else ket_qua
        chunk_items = du_lieu if isinstance(du_lieu, list) else [ket_qua]
        tat_ca_ban_ghi.extend(chunk_items)
        log_fn(f"   → Nhóm {idx}/{len(cac_nhom)}: đã lấy thông tin {len(nhom)} bản ghi.")
        if on_chunk_fn:
            on_chunk_fn(chunk_items)

    out_path = os.path.join(output_dir, f"{ten_goc}_ThongTinChiTiet.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(tat_ca_ban_ghi, f, ensure_ascii=False, indent=2)
    log_fn(f"Đã lưu thông tin chi tiết: {out_path}")
    return out_path, tat_ca_ban_ghi


# ============================ CẬP NHẬT (hồ sơ quét + thông tin đăng ký) ============================
# Dựa đúng theo payload mẫu bạn cung cấp cho UpdateHoSoQuetExistFile và cấu trúc "value"
# trả về từ GetThongTinDangKyByTinhHinhDangKyIds. CHƯA tự động chạy trong main()/GUI vì đây là
# thao tác GHI vào dữ liệu thật — cần gọi tường minh với tinhHinhDangKyId + dieuKienCapGiay cụ thể.

HO_SO_QUET_FIELDS = [
    "hoSoQuetId", "thongTinHoSoId", "tuiHoSoId", "tinhHinhDangKyId",
    "bienDongId", "xaId", "CreatedDate", "ModifiedDate",
    "Id", "Title", "Name", "Path", "ParentPath",
]


def build_ho_so_quet_field(ho_so_quet_item):
    """Rút gọn 1 bản ghi trong ListHoSoQuet về đúng field 'hoSoQuet' mà UpdateHoSoQuetExistFile cần."""
    ket_qua = {key: ho_so_quet_item.get(key) for key in HO_SO_QUET_FIELDS}
    ket_qua["CreatedDate"] = dotnet_date_to_iso(ket_qua["CreatedDate"])
    ket_qua["ModifiedDate"] = dotnet_date_to_iso(ket_qua["ModifiedDate"])
    ket_qua["_id"] = 1
    ket_qua["TuiHoSo"] = None
    return ket_qua


def build_info_ho_so_quet_field(file_item, idx, loai_ho_so_quet=2, la_giay_to_ve_nguon_goc=True):
    """Build field 'infoHoSoQuet_{idx}' từ 1 phần tử ListFileHoSoQuet, ép loaiHoSoQuet/laGiayToVeNguonGoc
    về giá trị chỉ định (mặc định 2 = đơn đăng ký, true) bất kể giá trị gốc là gì."""
    ket_qua = dict(file_item)
    ket_qua["tenGiayTo"] = ket_qua.get("tenGiayTo") or ""
    ket_qua["giayChungNhanId"] = ket_qua.get("giayChungNhanId") or ""
    ket_qua["loaiHoSoQuet"] = loai_ho_so_quet
    ket_qua["laGiayToVeNguonGoc"] = la_giay_to_ve_nguon_goc
    ket_qua["_id"] = idx
    ket_qua["files"] = None
    return ket_qua


def _khong_dau_upper(s):
    s = str(s or "").upper().replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def chon_file_ho_so_quet_de_cap_nhat(file_items):
    """Chọn 1 file đại diện trong ListFileHoSoQuet để cập nhật:
    - Chỉ 1 file → lấy file đó.
    - Nhiều file → ưu tiên moTa chứa "ĐƠN"/"DON", không có thì ưu tiên chứa "GT",
      không có nữa thì lấy file đầu tiên."""
    if not file_items:
        return None
    if len(file_items) == 1:
        return file_items[0]

    for item in file_items:
        if "DON" in _khong_dau_upper(item.get("moTa")):
            return item
    for item in file_items:
        if "GT" in _khong_dau_upper(item.get("moTa")):
            return item
    return file_items[0]


def build_update_ho_so_quet_payload(ho_so_quet_item, loai_ho_so_quet=2, la_giay_to_ve_nguon_goc=True):
    """Build payload form cho UpdateHoSoQuetExistFile từ 1 bản ghi trong ListHoSoQuet: chọn 1 file đại
    diện trong ListFileHoSoQuet (xem chon_file_ho_so_quet_de_cap_nhat) rồi ép loaiHoSoQuet/laGiayToVeNguonGoc
    cho file đó."""
    file_items = (ho_so_quet_item.get("ListFileHoSoQuet") or {}).get("ListFileHoSoQuet") or []
    file_can_cap_nhat = chon_file_ho_so_quet_de_cap_nhat(file_items)
    if file_can_cap_nhat is None:
        raise ValueError(f"hoSoQuetId {ho_so_quet_item.get('hoSoQuetId')} không có ListFileHoSoQuet để cập nhật.")

    payload = {
        "hoSoQuet": json.dumps(build_ho_so_quet_field(ho_so_quet_item), ensure_ascii=False),
        "count": "1",
        "isLuuKhoHoSoQuet": "false",
        "infoHoSoQuet_1": json.dumps(
            build_info_ho_so_quet_field(
                file_can_cap_nhat, 1, loai_ho_so_quet=loai_ho_so_quet,
                la_giay_to_ve_nguon_goc=la_giay_to_ve_nguon_goc,
            ),
            ensure_ascii=False,
        ),
    }
    return payload


def build_update_thong_tin_dang_ky_payload(item, dieu_kien_cap_giay):
    """Build payload UpdateThongTinDangKy từ 1 phần tử value[] của GetThongTinDangKyByTinhHinhDangKyIds:
    đổi hết '/Date(ms)/' sang ISO, gán _id đệ quy cho các list con, rồi gán TinhHinhDangKy.dieuKienCapGiay
    (gửi dạng chuỗi, đúng như payload mẫu, vd "3")."""
    payload = convert_dates_recursive(copy.deepcopy(item))
    add_id_recursive(payload)
    if not isinstance(payload.get("TinhHinhDangKy"), dict):
        raise ValueError("Bản ghi không có object TinhHinhDangKy để cập nhật.")
    payload["TinhHinhDangKy"]["dieuKienCapGiay"] = str(dieu_kien_cap_giay)
    return payload


def lay_thua_dat_ids(item):
    """Lấy danh sách thuaDatId trong TaiSan.ThuaDats của 1 bản ghi chi tiết, không trùng, giữ thứ tự."""
    thua_dats = (item.get("TaiSan") or {}).get("ThuaDats") or []
    ids = []
    da_thay = set()
    for thua_dat in thua_dats:
        if not isinstance(thua_dat, dict):
            continue
        thua_dat_id = thua_dat.get("thuaDatId")
        if thua_dat_id not in (None, "") and thua_dat_id not in da_thay:
            da_thay.add(thua_dat_id)
            ids.append(thua_dat_id)
    return ids


def cap_nhat_ho_so_quet_va_tinh_hinh_dang_ky(
    client, item, dieu_kien_cap_giay, log_fn=print,
    loai_ho_so_quet=2, la_giay_to_ve_nguon_goc=True,
):
    """Với 1 bản ghi chi tiết đầy đủ (1 phần tử value[] của GetThongTinDangKyByTinhHinhDangKyIds):
    1) Cập nhật từng hồ sơ quét trong ListHoSoQuet (ép loaiHoSoQuet/laGiayToVeNguonGoc)
    2) Cập nhật thông tin đăng ký (đổi dieuKienCapGiay)
    3) Gửi yêu cầu phân loại lại cho các thuaDatId có trong đơn (TaiSan.ThuaDats)
    LƯU Ý: đây là thao tác GHI vào dữ liệu thật trên MPLIS, hãy kiểm tra kỹ trước khi gọi hàng loạt."""
    thdk_id = (item.get("TinhHinhDangKy") or {}).get("tinhHinhDangKyId")
    danh_sach_ho_so_quet = item.get("ListHoSoQuet") or []
    if not danh_sach_ho_so_quet:
        log_fn(f"[{thdk_id}] Không có ListHoSoQuet, bỏ qua bước cập nhật hồ sơ quét.")
    for ho_so_quet in danh_sach_ho_so_quet:
        payload = build_update_ho_so_quet_payload(
            ho_so_quet, loai_ho_so_quet=loai_ho_so_quet, la_giay_to_ve_nguon_goc=la_giay_to_ve_nguon_goc
        )
        client.update_ho_so_quet_exist_file(payload)
        log_fn(f"[{thdk_id}] Đã cập nhật hồ sơ quét {ho_so_quet.get('hoSoQuetId')}.")

    payload_dang_ky = build_update_thong_tin_dang_ky_payload(item, dieu_kien_cap_giay)
    client.update_thong_tin_dang_ky(payload_dang_ky)
    log_fn(f"[{thdk_id}] Đã cập nhật thông tin đăng ký (dieuKienCapGiay={dieu_kien_cap_giay}).")

    thua_dat_ids = lay_thua_dat_ids(item)
    if not thua_dat_ids:
        log_fn(f"[{thdk_id}] Không có thuaDatId nào trong đơn, bỏ qua bước gửi yêu cầu phân loại lại.")
    else:
        client.gui_yeu_cau_phan_loai_lai(thua_dat_ids)
        log_fn(f"[{thdk_id}] Đã gửi yêu cầu phân loại lại cho thuaDatIds={thua_dat_ids}.")


def doc_thong_tin_chi_tiet(path):
    """Đọc file '..._ThongTinChiTiet.json' đã lưu, trả về list các item (mỗi item ứng 1 tinhHinhDangKyId,
    có TinhHinhDangKy/ChuSoHuu/TaiSan/ListHoSoQuet...). Hỗ trợ cả dạng list phẳng item lẫn dạng cũ
    list các {"value": [...]}."""
    with open(path, "r", encoding="utf-8") as f:
        du_lieu = json.load(f)
    if not isinstance(du_lieu, list):
        raise ValueError("File JSON không đúng định dạng (phải là 1 list).")

    items = []
    for phan_tu in du_lieu:
        if not isinstance(phan_tu, dict):
            continue
        if isinstance(phan_tu.get("value"), list):
            items.extend(x for x in phan_tu["value"] if isinstance(x, dict))
        elif "TinhHinhDangKy" in phan_tu:
            items.append(phan_tu)

    if not items:
        raise ValueError("Không tìm thấy bản ghi tinhHinhDangKy nào trong file JSON.")
    return items


def cap_nhat_hang_loat(
    client, items, dieu_kien_cap_giay, log_fn=print,
    loai_ho_so_quet=2, la_giay_to_ve_nguon_goc=True,
):
    """Cập nhật hồ sơ quét + thông tin đăng ký + gửi yêu cầu phân loại lại cho TỪNG bản ghi trong
    'items' (list các item value[] đã lấy từ GetThongTinDangKyByTinhHinhDangKyIds), cùng 1
    dieuKienCapGiay. Lỗi ở 1 bản ghi không dừng các bản ghi còn lại.
    LƯU Ý: đây là thao tác GHI vào dữ liệu thật trên MPLIS."""
    thanh_cong = 0
    loi = 0
    for idx, item in enumerate(items, start=1):
        thdk_id = (item.get("TinhHinhDangKy") or {}).get("tinhHinhDangKyId")
        log_fn(f"[{idx}/{len(items)}] tinhHinhDangKyId={thdk_id}")
        try:
            cap_nhat_ho_so_quet_va_tinh_hinh_dang_ky(
                client, item, dieu_kien_cap_giay, log_fn=log_fn,
                loai_ho_so_quet=loai_ho_so_quet, la_giay_to_ve_nguon_goc=la_giay_to_ve_nguon_goc,
            )
            thanh_cong += 1
        except Exception as e:
            loi += 1
            log_fn(f"   → LỖI: {e}")
        time.sleep(REQUEST_DELAY_SECONDS)

    log_fn(f"HOÀN TẤT CẬP NHẬT: {thanh_cong} thành công, {loi} lỗi / tổng {len(items)}.")
    return thanh_cong, loi


def cap_nhat_hang_loat_tu_file(
    client, json_path, dieu_kien_cap_giay, log_fn=print,
    loai_ho_so_quet=2, la_giay_to_ve_nguon_goc=True,
):
    """Đọc file '..._ThongTinChiTiet.json' đã lưu rồi cập nhật hàng loạt — dùng để KHÔI PHỤC/CHẠY LẠI
    khi cần (vd lần chạy trước bị lỗi giữa chừng), KHÔNG cần thiết trong luồng chạy bình thường vì
    main()/GUI đã tự cập nhật ngay trong cùng 1 lần chạy."""
    items = doc_thong_tin_chi_tiet(json_path)
    log_fn(f"Đã đọc {len(items)} bản ghi từ {json_path}.")
    return cap_nhat_hang_loat(
        client, items, dieu_kien_cap_giay, log_fn=log_fn,
        loai_ho_so_quet=loai_ho_so_quet, la_giay_to_ve_nguon_goc=la_giay_to_ve_nguon_goc,
    )


def lay_session_console(client, log_fn=print):
    """Tự lấy Cookie/token từ trình duyệt; nếu lỗi thì hỏi dán tay qua console (input())."""
    try:
        log_fn("Đang tự lấy Cookie + token từ trình duyệt (Chrome/Edge) đã đăng nhập MPLIS...")
        trinh_duyet = client.build_session_from_browser_cookies()
        log_fn(f"Đã tự lấy session + token thành công từ {trinh_duyet}.")
    except Exception as e:
        log_fn(f"Không tự lấy được ({e}).")
        log_fn("Mở DevTools (F12) > tab Network trên trang MPLIS đã đăng nhập, bấm 1 request bất kỳ,")
        log_fn("copy phần Cookie và giá trị __RequestVerificationToken rồi dán vào đây.")
        cookie_str = input("Dán Cookie: ").strip()
        token = input("Dán __RequestVerificationToken: ").strip()
        client.build_session_from_manual_input(cookie_str, token)
        log_fn("Đã dựng session từ Cookie/token dán tay.")


# ============================ CLI ============================

def main():
    if len(sys.argv) < 2:
        print("Cách dùng: python tracuu.py <duong_dan_excel.xlsx> [ten_goc] [dieuKienCapGiay]")
        print("  Bỏ trống dieuKienCapGiay nếu chỉ muốn tra cứu, không cập nhật.")
        print("  dieuKienCapGiay: " + "; ".join(f"{k}={v}" for k, v in sorted(DIEU_KIEN_CAP_GIAY_MAP.items())))
        sys.exit(1)

    input_file = sys.argv[1]
    if not os.path.isfile(input_file):
        print(f"Không thấy file: {input_file}")
        sys.exit(1)

    ten_goc = sys.argv[2].strip() if len(sys.argv) > 2 and sys.argv[2].strip() else None
    if not ten_goc:
        ten_goc = input("Nhập tên gốc cho file kết quả (vd: KetQua): ").strip()
    if not ten_goc:
        print("Tên gốc không được rỗng.")
        sys.exit(1)

    dieu_kien_cap_giay = sys.argv[3].strip() if len(sys.argv) > 3 and sys.argv[3].strip() else None

    try:
        che_do, cot_id = xac_dinh_che_do(input_file)
    except Exception as e:
        print(f"Lỗi Excel: {e}")
        sys.exit(1)

    client = MplisClient()
    try:
        lay_session_console(client)
    except ValueError as e:
        print(f"Lỗi: {e}")
        sys.exit(1)

    output_dir = os.path.dirname(os.path.abspath(input_file))

    def xac_nhan_console(dkcg):
        nhan_dkcg = DIEU_KIEN_CAP_GIAY_MAP.get(int(dkcg)) if dkcg.isdigit() else None
        print(f"\nSẮP GHI DỮ LIỆU THẬT lên MPLIS.")
        print(f"dieuKienCapGiay sẽ gán cho các bản ghi tìm được: {dkcg}" + (f" ({nhan_dkcg})" if nhan_dkcg else ""))
        print("Cập nhật sẽ chạy NGAY theo từng nhóm vừa tra/lấy chi tiết xong, không đợi hết toàn bộ file.")
        return input("Gõ 'DONG Y' để tiếp tục: ").strip() == "DONG Y"

    cap_nhat_ngay, lay_tong_ket = tao_ham_cap_nhat_tang_dan(client, dieu_kien_cap_giay, print, xac_nhan_console)

    if che_do == "id":
        print(f"Phát hiện cột '{cot_id}' → tra thẳng theo tinhHinhDangKyId, bỏ qua bước tra Số tờ/Số thửa.")
        try:
            tinh_hinh_dang_ky_ids = doc_excel_ids(input_file, cot_id)
        except Exception as e:
            print(f"Lỗi Excel: {e}")
            sys.exit(1)
        print(f"Đã đọc {len(tinh_hinh_dang_ky_ids)} tinhHinhDangKyId.")
        try:
            xu_ly_theo_id(
                tinh_hinh_dang_ky_ids, ten_goc, output_dir, client, log_fn=print, cap_nhat_ngay=cap_nhat_ngay
            )
        except Exception as e:
            print(f"Lỗi khi lấy thông tin chi tiết: {e}")
    else:
        try:
            df = doc_excel(input_file)
        except Exception as e:
            print(f"Lỗi Excel: {e}")
            sys.exit(1)
        print(f"Đã đọc {len(df)} dòng từ Excel.")
        xu_ly_theo_xa(df, ten_goc, input_file, client, log_fn=print, cap_nhat_ngay=cap_nhat_ngay)

    if dieu_kien_cap_giay:
        thanh_cong, loi = lay_tong_ket()
        print(f"TỔNG CẬP NHẬT: {thanh_cong} thành công, {loi} lỗi.")
    print("HOÀN TẤT.")


def main_cap_nhat():
    """python tracuu.py cap-nhat <duong_dan_file_ThongTinChiTiet.json> <dieuKienCapGiay>
    GHI DỮ LIỆU THẬT lên MPLIS: cập nhật hồ sơ quét (ép loaiHoSoQuet=2, laGiayToVeNguonGoc=true) +
    thông tin đăng ký (dieuKienCapGiay) cho TẤT CẢ bản ghi trong file JSON, cùng 1 dieuKienCapGiay."""
    if len(sys.argv) < 4:
        print("Cách dùng: python tracuu.py cap-nhat <duong_dan_file.json> <dieuKienCapGiay>")
        print("  dieuKienCapGiay: " + "; ".join(f"{k}={v}" for k, v in sorted(DIEU_KIEN_CAP_GIAY_MAP.items())))
        sys.exit(1)

    json_path = sys.argv[2]
    if not os.path.isfile(json_path):
        print(f"Không thấy file: {json_path}")
        sys.exit(1)
    dieu_kien_cap_giay = sys.argv[3].strip()

    try:
        so_ban_ghi = len(doc_thong_tin_chi_tiet(json_path))
    except Exception as e:
        print(f"Lỗi file JSON: {e}")
        sys.exit(1)

    nhan_dkcg = DIEU_KIEN_CAP_GIAY_MAP.get(int(dieu_kien_cap_giay)) if dieu_kien_cap_giay.isdigit() else None
    print(f"SẮP GHI DỮ LIỆU THẬT lên MPLIS cho {so_ban_ghi} bản ghi từ file: {json_path}")
    print(f"dieuKienCapGiay sẽ gán cho TẤT CẢ: {dieu_kien_cap_giay}" + (f" ({nhan_dkcg})" if nhan_dkcg else ""))
    if input("Gõ 'DONG Y' để tiếp tục: ").strip() != "DONG Y":
        print("Đã hủy.")
        sys.exit(0)

    client = MplisClient()
    try:
        lay_session_console(client)
    except ValueError as e:
        print(f"Lỗi: {e}")
        sys.exit(1)

    cap_nhat_hang_loat_tu_file(client, json_path, dieu_kien_cap_giay, log_fn=print)


# ============================ TKINTER (chọn file bằng giao diện) ============================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Tra cứu Tình hình đăng ký MPLIS")
        root.geometry("720x560")
        root.minsize(650, 480)

        self.running = False
        self.var_file = tk.StringVar()
        self.var_ten_goc = tk.StringVar()
        self.var_dieu_kien_cap_giay = tk.StringVar()

        frm = ttk.LabelFrame(root, text="Tra cứu + lấy thông tin chi tiết + cập nhật (tuỳ chọn)", padding=10)
        frm.pack(fill="x", padx=10, pady=(10, 5))

        ttk.Label(
            frm,
            text="File Excel (cột Mã xã/Số tờ/Số thửa, HOẶC cột tinhHinhDangKyId có sẵn):",
        ).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Entry(frm, textvariable=self.var_file).grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 5))
        ttk.Button(frm, text="Chọn file...", command=self.chon_file).grid(row=1, column=2)
        frm.columnconfigure(0, weight=1)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Tên gốc file kết quả:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frm, textvariable=self.var_ten_goc, width=30).grid(row=3, column=0, sticky="w")

        ttk.Label(
            frm, text="dieuKienCapGiay để CẬP NHẬT luôn (để trống nếu chỉ tra cứu, không ghi dữ liệu):",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))
        cbo_dkcg = ttk.Combobox(
            frm, textvariable=self.var_dieu_kien_cap_giay, state="readonly", width=45,
            values=[""] + [f"{k} - {v}" for k, v in sorted(DIEU_KIEN_CAP_GIAY_MAP.items())],
        )
        cbo_dkcg.grid(row=5, column=0, columnspan=2, sticky="w")

        self.btn_run = ttk.Button(frm, text="Chạy", command=self.chay)
        self.btn_run.grid(row=5, column=2, sticky="e")

        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(8, 0))
        self.lbl_status = ttk.Label(root, text="Chưa chạy")
        self.lbl_status.pack(anchor="w", padx=10)

        log_frame = ttk.Frame(root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=8)
        self.txt = tk.Text(log_frame, wrap="word", height=18)
        self.txt.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(log_frame, command=self.txt.yview)
        scroll.pack(side="right", fill="y")
        self.txt.configure(yscrollcommand=scroll.set)

    def chon_file(self):
        f = filedialog.askopenfilename(
            title="Chọn file Excel (Mã xã/Số tờ/Số thửa hoặc tinhHinhDangKyId)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tất cả", "*.*")],
        )
        if f:
            self.var_file.set(f)
            if not self.var_ten_goc.get().strip():
                self.var_ten_goc.set("KetQua")

    def hoi_xac_nhan(self, tieu_de, noi_dung):
        """Hiện messagebox.askyesno an toàn từ luồng nền (gọi qua root.after), trả về True/False."""
        ket_qua = {}
        xong = threading.Event()

        def _hien_thi():
            ket_qua["val"] = messagebox.askyesno(tieu_de, noi_dung)
            xong.set()

        self.root.after(0, _hien_thi)
        xong.wait()
        return ket_qua.get("val", False)

    def log(self, msg):
        def _append():
            self.txt.insert("end", f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n")
            self.txt.see("end")
        self.root.after(0, _append)

    def set_status(self, msg):
        self.root.after(0, lambda: self.lbl_status.config(text=msg))

    def set_progress(self, value, maximum):
        self.root.after(0, lambda: self.progress.config(maximum=max(maximum, 1), value=value))

    def hoi_cookie_token(self):
        """Hộp thoại nhỏ để dán Cookie/token tay khi không tự lấy được (chạy an toàn từ luồng nền)."""
        ket_qua = {}
        xong = threading.Event()

        def _hien_thi():
            dlg = tk.Toplevel(self.root)
            dlg.title("Dán Cookie / Token tay")
            dlg.geometry("520x320")
            dlg.transient(self.root)

            ttk.Label(
                dlg,
                text="Không tự lấy được Cookie/token. Mở DevTools (F12) > Network trên\n"
                     "trang MPLIS đã đăng nhập, copy Cookie và __RequestVerificationToken rồi dán:",
                justify="left",
            ).pack(anchor="w", padx=10, pady=(10, 5))

            ttk.Label(dlg, text="Cookie:").pack(anchor="w", padx=10)
            txt_cookie = tk.Text(dlg, height=8)
            txt_cookie.pack(fill="both", expand=True, padx=10)

            ttk.Label(dlg, text="__RequestVerificationToken:").pack(anchor="w", padx=10, pady=(8, 0))
            ent_token = ttk.Entry(dlg)
            ent_token.pack(fill="x", padx=10)

            def _ok():
                ket_qua["cookie"] = txt_cookie.get("1.0", "end").strip()
                ket_qua["token"] = ent_token.get().strip()
                dlg.destroy()
                xong.set()

            def _huy():
                ket_qua["cookie"] = None
                ket_qua["token"] = None
                dlg.destroy()
                xong.set()

            btns = ttk.Frame(dlg)
            btns.pack(pady=10)
            ttk.Button(btns, text="OK", command=_ok).pack(side="left", padx=5)
            ttk.Button(btns, text="Hủy", command=_huy).pack(side="left", padx=5)
            dlg.protocol("WM_DELETE_WINDOW", _huy)
            dlg.grab_set()

        self.root.after(0, _hien_thi)
        xong.wait()
        return ket_qua.get("cookie"), ket_qua.get("token")

    def chay(self):
        if self.running:
            return

        input_file = self.var_file.get().strip()
        ten_goc = self.var_ten_goc.get().strip()
        if not input_file or not os.path.isfile(input_file):
            messagebox.showwarning("Thiếu file", "Chọn file Excel hợp lệ.")
            return
        if not ten_goc:
            messagebox.showwarning("Thiếu tên", "Nhập tên gốc cho file kết quả.")
            return

        lua_chon = self.var_dieu_kien_cap_giay.get().strip()
        dieu_kien_cap_giay = lua_chon.split(" - ", 1)[0].strip() if lua_chon else None

        self.running = True
        self.btn_run.config(state="disabled")
        self.set_status("Đang chạy...")
        threading.Thread(
            target=self._run_batch, args=(input_file, ten_goc, dieu_kien_cap_giay), daemon=True
        ).start()

    def _lay_session(self, client):
        """Tự lấy Cookie/token từ trình duyệt; lỗi thì bật hộp thoại dán tay. Trả về True nếu có session."""
        self.log("Đang tự lấy Cookie + token từ trình duyệt (Chrome/Edge) đã đăng nhập MPLIS...")
        try:
            trinh_duyet = client.build_session_from_browser_cookies()
            self.log(f"Đã tự lấy session + token thành công từ {trinh_duyet}.")
            return True
        except Exception as e:
            self.log(f"Không tự lấy được ({e}).")
            cookie_str, token = self.hoi_cookie_token()
            if not cookie_str or not token:
                self.log("Đã hủy nhập Cookie/Token. Dừng.")
                return False
            try:
                client.build_session_from_manual_input(cookie_str, token)
            except ValueError as e2:
                self.log(f"Lỗi: {e2}")
                return False
            self.log("Đã dựng session từ Cookie/token dán tay.")
            return True

    def _run_batch(self, input_file, ten_goc, dieu_kien_cap_giay):
        try:
            try:
                che_do, cot_id = xac_dinh_che_do(input_file)
            except Exception as e:
                self.log(f"Lỗi Excel: {e}")
                return

            client = MplisClient()
            if not self._lay_session(client):
                return

            output_dir = os.path.dirname(os.path.abspath(input_file))

            def xac_nhan_gui(dkcg):
                lua_chon = self.var_dieu_kien_cap_giay.get().strip()
                return self.hoi_xac_nhan(
                    "GHI DỮ LIỆU THẬT lên MPLIS",
                    "Sẽ cập nhật NGAY theo từng nhóm vừa tra/lấy chi tiết xong "
                    "(không đợi hết toàn bộ file).\n\n"
                    "- Hồ sơ quét: ép loaiHoSoQuet=2, laGiayToVeNguonGoc=true (1 file đại diện mỗi hồ sơ quét)\n"
                    f"- Thông tin đăng ký: dieuKienCapGiay = {lua_chon}\n"
                    "- Gửi yêu cầu phân loại lại cho các thửa đất trong đơn\n\n"
                    "Thao tác này GHI TRỰC TIẾP vào dữ liệu MPLIS, không dễ hoàn tác. Tiếp tục?",
                )

            cap_nhat_ngay, lay_tong_ket = tao_ham_cap_nhat_tang_dan(
                client, dieu_kien_cap_giay, self.log, xac_nhan_gui
            )

            if che_do == "id":
                self.log(f"Phát hiện cột '{cot_id}' → tra thẳng theo tinhHinhDangKyId, bỏ qua bước tra Số tờ/Số thửa.")
                try:
                    tinh_hinh_dang_ky_ids = doc_excel_ids(input_file, cot_id)
                except Exception as e:
                    self.log(f"Lỗi Excel: {e}")
                    return
                self.log(f"Đã đọc {len(tinh_hinh_dang_ky_ids)} tinhHinhDangKyId.")
                self.set_progress(1, 1)
                try:
                    xu_ly_theo_id(
                        tinh_hinh_dang_ky_ids, ten_goc, output_dir, client,
                        log_fn=self.log, cap_nhat_ngay=cap_nhat_ngay,
                    )
                except Exception as e:
                    self.log(f"Lỗi khi lấy thông tin chi tiết: {e}")
            else:
                try:
                    df = doc_excel(input_file)
                except Exception as e:
                    self.log(f"Lỗi Excel: {e}")
                    return
                self.log(f"Đã đọc {len(df)} dòng từ Excel.")
                self.set_progress(0, len(df))
                xu_ly_theo_xa(
                    df, ten_goc, input_file, client,
                    log_fn=self.log, progress_fn=self.set_progress, cap_nhat_ngay=cap_nhat_ngay,
                )

            if dieu_kien_cap_giay:
                thanh_cong, loi = lay_tong_ket()
                self.log(f"TỔNG CẬP NHẬT: {thanh_cong} thành công, {loi} lỗi.")

            self.log("HOÀN TẤT.")
            self.set_status("Hoàn tất")
        except Exception:
            # Bắt mọi lỗi bất ngờ chưa lường tới, tránh luồng nền chết âm thầm
            # (trước đây lỗi kiểu này không hiện lên log gì cả, chỉ dừng ngang).
            self.log("LỖI KHÔNG LƯỜNG TRƯỚC:\n" + traceback.format_exc())
            self.set_status("Lỗi")
        finally:
            self.running = False
            self.root.after(0, lambda: self.btn_run.config(state="normal"))


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "cap-nhat":
        main_cap_nhat()
    elif len(sys.argv) >= 2:
        main()
    else:
        root = tk.Tk()
        App(root)
        root.mainloop()
