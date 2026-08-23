# -*- coding: utf-8 -*-
"""
Cập nhật thông tin Chủ sở hữu qua ChuSoHuuAjax/UpdateToChuc (Tổ chức) và
ChuSoHuuAjax/UpdateCaNhan (Cá nhân - vd người đại diện của Tổ chức).
- Payload gốc lấy nguyên object ToChuc (value[0].ChuSoHuu.ToChucs[0]) hoặc CaNhan
  (to_chuc["NguoiDaiDien"]) từ kết quả tra cứu, chỉ ghi đè các trường cần sửa rồi gửi lên
  nguyên cấu trúc.
- Người đại diện (NguoiDaiDien) là 1 CaNhan độc lập, có bản ghi riêng - sửa field lồng
  "NguoiDaiDien.xxx" trong payload UpdateToChuc chỉ đổi bản sao gắn trong Tổ chức, KHÔNG
  đổi bản ghi CaNhan gốc. Muốn đổi thật thông tin người đại diện phải gọi update_ca_nhan()
  RIÊNG (với chính object NguoiDaiDien) TRƯỚC, rồi mới gọi update_to_chuc().
- LƯU Ý: đây là API GHI DỮ LIỆU thật trên MPLIS, không có hoàn tác tự động - kiểm tra kỹ
  payload (đặc biệt toChucId/caNhanId/nguoiDaiDienId/version) trước khi chạy thật.

Chạy: python capnhat_chusohuu.py
"""

import copy
import json
import time
from http.cookies import SimpleCookie

import requests


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


REFERER_URL = "https://dla.mplis.gov.vn/dc/DonDangKy/KeKhaiDangKyV2"
URL_UPDATE_TO_CHUC = "https://dla.mplis.gov.vn/dc/ChuSoHuuAjax/UpdateToChuc"
URL_UPDATE_CA_NHAN = "https://dla.mplis.gov.vn/dc/ChuSoHuuAjax/UpdateCaNhan"

TIMEOUT = 120

DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


def build_session_from_manual_input(cookie_str, token):
    """Dựng session trực tiếp từ Cookie header và token dán tay."""
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


def _dong_bo_dia_chi(entity, dia_chi_moi, nhan=""):
    """entity: dict có key 'ListDiaChi' (ToChuc hoặc CaNhan). Field 'diaChi' cấp cao nhất chỉ
    để hiển thị - server dựng lại nó từ ListDiaChi[i].diaChiChiTiet (entry có laDiaChiChinh=true,
    không có thì lấy entry đầu), set riêng field 'diaChi' KHÔNG có tác dụng (đã kiểm chứng thực tế
    qua UpdateCaNhan: gửi diaChi mới nhưng response vẫn trả địa chỉ cũ y hệt ListDiaChi[0]).
    Đã kiểm chứng lại: khi đồng bộ cả 2 nơi thì response trả đúng địa chỉ mới.

    Nếu ListDiaChi rỗng/null (thường gặp với người đại diện chưa từng có địa chỉ cấu trúc
    trong hệ thống - vd object lấy từ AdvancedSearchGiayChungNhan) thì KHÔNG có bản ghi nào để
    ghi đè diaChiChiTiet -> trước đây hàm này âm thầm bỏ qua, khiến địa chỉ mới KHÔNG được gửi
    lên dù request vẫn trả success=true. Giờ báo lỗi rõ ràng thay vì âm thầm bỏ qua."""
    danh_sach = entity.get("ListDiaChi") or []
    muc_tieu = next((d for d in danh_sach if d.get("laDiaChiChinh")), None)
    if muc_tieu is None and danh_sach:
        muc_tieu = danh_sach[0]
    if muc_tieu is None:
        raise RuntimeError(
            f"Không có bản ghi ListDiaChi{f' ({nhan})' if nhan else ''} để đồng bộ địa chỉ mới - "
            "server chỉ đọc ListDiaChi[i].diaChiChiTiet, set riêng field 'diaChi' không có tác dụng "
            "nên địa chỉ sẽ KHÔNG được cập nhật thật. Phải thêm địa chỉ lần đầu qua giao diện MPLIS "
            "(để tạo bản ghi ListDiaChi) rồi mới dùng công cụ này sửa tiếp được."
        )
    muc_tieu["diaChiChiTiet"] = dia_chi_moi


def _dong_bo_so_giay_to_tuy_than(entity, so_giay_to_moi):
    """entity: dict có key 'ListGiayToTuyThan' (CaNhan). Giống 'diaChi', 'soGiayTo' thật của
    giấy tờ tuỳ thân (CCCD) nằm trong ListGiayToTuyThan[i].soGiayTo, không phải field
    'maSoDinhDanh' cấp cao nhất - đồng bộ luôn khi override maSoDinhDanh (ưu tiên entry loại
    CCCD, không có thì lấy entry đầu)."""
    danh_sach = entity.get("ListGiayToTuyThan") or []
    muc_tieu = next(
        (g for g in danh_sach if (g.get("LoaiGiayToTuyThan") or {}).get("maLoaiGiayTo") == "CCCD"), None
    )
    if muc_tieu is None and danh_sach:
        muc_tieu = danh_sach[0]
    if muc_tieu is not None:
        muc_tieu["soGiayTo"] = so_giay_to_moi


# Các trường định danh/hệ thống của riêng 1 bản ghi ToChuc - PHẢI giữ nguyên theo từng bản ghi
# khi gộp nhiều ToChuc trùng (khác toChucId, cùng 1 chủ thật) thành 1 thông tin duy nhất, để
# update đúng bản ghi (không lỡ đổi toChucId) và không vỡ khoá lạc quan (version)/liên kết đồ thị
# (Id/Name/Path/InId/OutId). "xaId" cũng giữ nguyên vì có thể là khoá phân vùng dữ liệu theo xã
# gắn với chính bản ghi đó trong đồ thị, không chắc an toàn khi đổi theo bản chính.
TO_CHUC_TRUONG_GIU_NGUYEN = [
    "toChucId", "Id", "Title", "Description", "Name", "Path", "ParentPath",
    "Layer", "InId", "OutId", "CreatedDate", "ModifiedDate",
    "version", "isLastest", "isNew", "isChange", "xaId",
]


def xac_dinh_ma_chu_su_dung(to_chuc_chinh):
    """Xác định 1 mã chủ sử dụng dùng chung cho tất cả ToChuc trùng của cùng 1 GCN:
    ưu tiên maChuSuDung có sẵn (không rỗng) trên bản ghi CHÍNH, không có thì dùng chính
    toChucId của bản ghi chính làm mã (ép về chuỗi để nhất quán kiểu dữ liệu)."""
    ma = to_chuc_chinh.get("maChuSuDung")
    if ma not in (None, ""):
        return ma
    return str(to_chuc_chinh.get("toChucId"))


def build_gop_to_chuc_trung_payload(to_chuc_trung, to_chuc_dich_da_resolve):
    """Dựng payload cập nhật cho 1 ToChuc TRÙNG (khác toChucId nhưng là cùng 1 chủ thật, do dữ
    liệu bị nhập lặp) để nó trở thành GIỐNG HỆT to_chuc_dich_da_resolve (bản ghi CHÍNH, đã áp
    overrides + maChuSuDung) về mọi thông tin nội dung (tên, địa chỉ, mã số, người đại diện,
    maChuSuDung...), NHƯNG giữ nguyên các trường trong TO_CHUC_TRUONG_GIU_NGUYEN của chính
    bản ghi trùng đó. to_chuc_trung: dict nguyên bản (chưa sửa) của bản ghi trùng."""
    payload = copy.deepcopy(to_chuc_dich_da_resolve)
    for truong in TO_CHUC_TRUONG_GIU_NGUYEN:
        if truong in to_chuc_trung:
            payload[truong] = to_chuc_trung[truong]
        else:
            payload.pop(truong, None)
    return payload


def build_update_to_chuc_payload(to_chuc, **overrides):
    """to_chuc: dict nguyên bản của ChuSoHuu.ToChucs[0] (từ GetThongTinDangKyByTinhHinhDangKyIds).
    overrides: field muốn sửa ở cấp ToChuc (vd tenToChuc=, diaChi=, maSoDinhDanh=), hoặc field
    lồng trong người đại diện dùng tiền tố "NguoiDaiDien." (vd "NguoiDaiDien.hoTen"=).
    Field lồng trong giấy tờ bổ sung (số GPKD/ngày cấp) sửa trực tiếp trên
    to_chuc["ListGiayToBoSung"][i] trước khi gọi hàm này.
    Override "diaChi"/"NguoiDaiDien.diaChi" sẽ tự đồng bộ vào ListDiaChi tương ứng (xem
    _dong_bo_dia_chi) vì field "diaChi" cấp cao nhất chỉ để hiển thị, không được server ghi nhận."""
    payload = copy.deepcopy(to_chuc)
    for key, value in overrides.items():
        if key.startswith("NguoiDaiDien."):
            sub_key = key.split(".", 1)[1]
            payload.setdefault("NguoiDaiDien", {})[sub_key] = value
        else:
            payload[key] = value

    if "diaChi" in overrides:
        _dong_bo_dia_chi(payload, overrides["diaChi"], nhan="tổ chức")
    if "NguoiDaiDien.diaChi" in overrides and payload.get("NguoiDaiDien"):
        _dong_bo_dia_chi(payload["NguoiDaiDien"], overrides["NguoiDaiDien.diaChi"], nhan="người đại diện")

    # Lưu ý: "maSoDinhDanh" cấp ToChuc là mã số của chính tổ chức, giấy tờ liên quan nằm ở
    # ListGiayToBoSung (soGiayTo GPKD...) chứ không phải ListGiayToTuyThan (chỉ CaNhân mới có),
    # nên không đồng bộ ở đây. NguoiDaiDien là CaNhân nên có ListGiayToTuyThan, đồng bộ như CaNhân.
    if "NguoiDaiDien.maSoDinhDanh" in overrides and payload.get("NguoiDaiDien"):
        _dong_bo_so_giay_to_tuy_than(payload["NguoiDaiDien"], overrides["NguoiDaiDien.maSoDinhDanh"])

    return payload


def update_to_chuc(session, payload):
    headers = dict(session.headers)
    headers["Content-Type"] = "application/json; charset=UTF-8"

    res = post_voi_retry(
        session,
        URL_UPDATE_TO_CHUC,
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


def build_update_ca_nhan_payload(ca_nhan, **overrides):
    """ca_nhan: dict nguyên bản của 1 Cá nhân (vd to_chuc["NguoiDaiDien"]).
    overrides: field muốn sửa trực tiếp trên CaNhan (vd hoTen=, maSoDinhDanh=, ngaySinh=, diaChi=).
    Field nào không truyền thì giữ nguyên dữ liệu gốc.

    LƯU Ý (rút ra từ test thực tế): field "diaChi" cấp cao nhất CHỈ ĐỂ HIỂN THỊ - server dựng
    lại nó từ ListDiaChi[i].diaChiChiTiet (entry có laDiaChiChinh=true, không có thì lấy entry
    đầu), không đọc trực tiếp field "diaChi" gửi lên. Set riêng "diaChi" trong payload KHÔNG
    có tác dụng - vì vậy override "diaChi" ở đây sẽ tự đồng bộ luôn vào ListDiaChi tương ứng.
    Tương tự, "soGiayTo" thật của CCCD nằm trong ListGiayToTuyThan[i].soGiayTo - override
    "maSoDinhDanh" sẽ tự đồng bộ luôn vào đó."""
    payload = copy.deepcopy(ca_nhan)
    payload.update(overrides)

    if "diaChi" in overrides:
        _dong_bo_dia_chi(payload, overrides["diaChi"], nhan="cá nhân")
    if "maSoDinhDanh" in overrides:
        _dong_bo_so_giay_to_tuy_than(payload, overrides["maSoDinhDanh"])

    return payload


def update_ca_nhan(session, payload):
    headers = dict(session.headers)
    headers["Content-Type"] = "application/json; charset=UTF-8"

    res = post_voi_retry(
        session,
        URL_UPDATE_CA_NHAN,
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


def main():
    print("Dán Cookie và __RequestVerificationToken (F12 > Network > request bất kỳ tới dla.mplis.gov.vn).")
    cookie_str = input("Dán Cookie: ").strip()
    token = input("Dán __RequestVerificationToken: ").strip()
    session = build_session_from_manual_input(cookie_str, token)

    with open("to_chuc_payload.json", encoding="utf-8") as f:
        to_chuc = json.load(f)

    # TODO: điền overrides thật rồi mới chạy - hiện để rỗng để không lỡ tay gửi update.
    overrides = {}

    payload = build_update_to_chuc_payload(to_chuc, **overrides)

    if not overrides:
        print("Chưa có overrides nào - dừng lại, không gửi update. Điền overrides trong file rồi chạy lại.")
        return

    result = update_to_chuc(session, payload)
    with open("update_to_chuc_result.json", "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print("Đã gửi update, xem kết quả trong update_to_chuc_result.json")


if __name__ == "__main__":
    main()
