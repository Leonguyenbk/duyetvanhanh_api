# -*- coding: utf-8 -*-
"""
Tra cứu Giấy chứng nhận theo Số phát hành (AdvancedSearchGiayChungNhan) để lấy
tinhHinhDangKyId (tình hình đăng ký id) gắn với giấy đó.
- Khi chạy sẽ hỏi dán Cookie + __RequestVerificationToken (lấy từ trình duyệt đang
  đăng nhập MPLIS, xem hướng dẫn lấy ở dưới).

Chạy: python tracuu_giaychungnhan.py "DG 781336" [tinhId] [xaId] [huyenId]
Cài đặt: pip install requests
"""

import sys
import json
import time
from http.cookies import SimpleCookie

import requests


REFERER_URL = "https://dla.mplis.gov.vn/dc/DonDangKy/KeKhaiDangKyV2"
URL_SEARCH_GCN = "https://dla.mplis.gov.vn/dc/CungCapThongTinGiayChungNhanAjax/AdvancedSearchGiayChungNhan"


def post_voi_retry(session, url, retries=3, **kwargs):
    """Gọi session.post với retry khi gặp lỗi socket/kết nối tạm thời (vd. OSError [Errno 22]
    khi server đóng kết nối keep-alive giữa 2 request liên tiếp trong 1 vòng lặp dài)."""
    for attempt in range(1, retries + 1):
        try:
            return session.post(url, **kwargs)
        except (requests.exceptions.RequestException, OSError):
            if attempt >= retries:
                raise
            time.sleep(1.5 * attempt)

TIMEOUT = 120

DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

def build_session_from_manual_input(cookie_str, token):
    """Dựng session trực tiếp từ Cookie header và token dán tay (dùng khi không tự đọc được
    cookie trình duyệt, ví dụ Chrome lỗi giải mã cookie hoặc Edge cần quyền admin)."""
    cookie_str = (cookie_str or "").strip()
    token = (token or "").strip()
    if not cookie_str:
        raise ValueError("Cookie không được rỗng.")
    if not token:
        raise ValueError("Token không được rỗng.")

    session = requests.Session()
    session.headers.update({
        "User-Agent": DEFAULT_USER_AGENT,
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

    return session


def build_payload_giay_chung_nhan(so_phat_hanh, tinh_id=66, xa_id=0, huyen_id=0, length=10):
    """Payload form-urlencoded giống đúng request DataTables của AdvancedSearchGiayChungNhan."""
    columns = [
        ("", "", True, False),
        ("GiayChungNhan", "GiayChungNhan", True, False),
        ("ChuSoHuu", "ChuSoHuu", True, False),
        ("TaiSan", "TaiSan", True, False),
    ]

    payload = {"draw": "2"}
    for i, (data, name, searchable, orderable) in enumerate(columns):
        payload[f"columns[{i}][data]"] = data
        payload[f"columns[{i}][name]"] = name
        payload[f"columns[{i}][searchable]"] = "true" if searchable else "false"
        payload[f"columns[{i}][orderable]"] = "true" if orderable else "false"
        payload[f"columns[{i}][search][value]"] = ""
        payload[f"columns[{i}][search][regex]"] = "false"

    payload.update({
        "start": "0",
        "length": str(length),
        "search[value]": "",
        "search[regex]": "false",
        "isAdvancedSearch": "true",
        "tinhId": str(tinh_id),
        "xaId": str(xa_id),
        "huyenId": str(huyen_id),
        "timChinhXac": "true",
        "andOperator": "false",
        "loaiGiayChungNhanId": "",
        "maVach": "",
        "soPhatHanh": so_phat_hanh,
        "soVaoSo": "",
        "soHoSoGoc": "",
        "soHoSoGocCu": "",
        "soVaoSoCu": "",
        "hoTen": "",
        "namSinh": "",
        "soGiayTo": "",
        "soThuTuThua": "",
        "soHieuToBanDo": "",
        "soThuTuThuaCu": "",
        "soHieuToBanDoCu": "",
        "soNha": "",
        "diaChiChiTiet": "",
    })
    return payload


def tra_cuu_giay_chung_nhan(session, so_phat_hanh, tinh_id=66, xa_id=0, huyen_id=0):
    headers = dict(session.headers)
    headers["Content-Type"] = "application/x-www-form-urlencoded; charset=UTF-8"
    payload = build_payload_giay_chung_nhan(so_phat_hanh, tinh_id, xa_id, huyen_id)

    res = post_voi_retry(session, URL_SEARCH_GCN, data=payload, headers=headers, timeout=TIMEOUT)
    res.raise_for_status()
    ct = res.headers.get("Content-Type") or ""
    if "json" not in ct:
        raise RuntimeError(
            f"Server không trả JSON (status {res.status_code}), có thể mất session. "
            f"Nội dung: {res.text[:400]}"
        )
    return res.json()


def lay_tinh_hinh_dang_ky_id(row):
    """Lấy (các) tinhHinhDangKyId trong 1 row kết quả AdvancedSearchGiayChungNhan.
    Theo mẫu response thực tế, ID nằm ở GiayChungNhan.ListDangKyQuyen[].tinhHinhDangKyId
    (và có thể cả ListDangKyTaiSan[].tinhHinhDangKyId nếu là tài sản gắn liền với đất)."""
    if not isinstance(row, dict):
        return []
    gcn = row.get("GiayChungNhan") or row.get("giayChungNhan") or {}

    ids = []
    for key in ("ListDangKyQuyen", "ListDangKyTaiSan"):
        for item in gcn.get(key) or []:
            if isinstance(item, dict) and item.get("tinhHinhDangKyId"):
                ids.append(item["tinhHinhDangKyId"])
    return ids


def lay_to_chuc_tu_gcn_row(row):
    """Lấy object ToChuc (chủ sở hữu dạng tổ chức) trực tiếp từ 1 row kết quả
    AdvancedSearchGiayChungNhan, không cần gọi thêm GetThongTinDangKyByTinhHinhDangKyIds:
    ID nằm ở GiayChungNhan.ListDangKyQuyen[].ToChuc (và ListDangKyTaiSan[].ToChuc nếu có).
    Trả về None nếu chủ sở hữu không phải Tổ chức (Cá nhân/Hộ gia đình...)."""
    if not isinstance(row, dict):
        return None
    gcn = row.get("GiayChungNhan") or row.get("giayChungNhan") or {}

    for key in ("ListDangKyQuyen", "ListDangKyTaiSan"):
        for item in gcn.get(key) or []:
            if isinstance(item, dict) and item.get("ToChuc"):
                return item["ToChuc"]
    return None


def lay_tat_ca_to_chuc_tu_gcn_rows(rows):
    """Lấy TẤT CẢ object ToChuc (không trùng toChucId) xuất hiện trong toàn bộ rows kết quả
    AdvancedSearchGiayChungNhan — dùng khi 1 GCN bị trùng dữ liệu chủ sở hữu thành nhiều bản ghi
    ToChuc khác nhau (khác toChucId, cùng thực thể ngoài đời) cần gộp lại thành 1. Bản ghi đầu
    tiên trong list trả về được coi là bản ghi CHÍNH; các bản còn lại là bản TRÙNG cần đồng bộ
    theo bản chính. Trả về [] nếu chủ sở hữu không phải Tổ chức."""
    ket_qua = []
    da_thay = set()
    for row in rows:
        if not isinstance(row, dict):
            continue
        gcn = row.get("GiayChungNhan") or row.get("giayChungNhan") or {}
        for key in ("ListDangKyQuyen", "ListDangKyTaiSan"):
            for item in gcn.get(key) or []:
                if not isinstance(item, dict):
                    continue
                to_chuc = item.get("ToChuc")
                if not isinstance(to_chuc, dict):
                    continue
                to_chuc_id = to_chuc.get("toChucId")
                if to_chuc_id in da_thay:
                    continue
                da_thay.add(to_chuc_id)
                ket_qua.append(to_chuc)
    return ket_qua


def main():
    if len(sys.argv) < 2:
        print('Cách dùng: python tracuu_giaychungnhan.py "<soPhatHanh>" [tinhId] [xaId] [huyenId]')
        sys.exit(1)

    so_phat_hanh = sys.argv[1]
    tinh_id = int(sys.argv[2]) if len(sys.argv) > 2 else 66
    xa_id = int(sys.argv[3]) if len(sys.argv) > 3 else 0
    huyen_id = int(sys.argv[4]) if len(sys.argv) > 4 else 0

    print("Mở dla.mplis.gov.vn trên trình duyệt đang đăng nhập > F12 > tab Network > bấm 1 request")
    print("bất kỳ tới dla.mplis.gov.vn > tab Headers > copy phần Cookie và __RequestVerificationToken.")
    cookie_str = input("Dán Cookie: ").strip()
    token = input("Dán __RequestVerificationToken: ").strip()
    session = build_session_from_manual_input(cookie_str, token)

    js = tra_cuu_giay_chung_nhan(session, so_phat_hanh, tinh_id, xa_id, huyen_id)

    print(json.dumps(js, ensure_ascii=False, indent=2))

    rows = js.get("data") or []
    print(f"\nSố bản ghi tìm được: {len(rows)}")
    for i, row in enumerate(rows, start=1):
        thdk_id = lay_tinh_hinh_dang_ky_id(row)
        print(f"{i}. tinhHinhDangKyId: {thdk_id}")


if __name__ == "__main__":
    main()
