# -*- coding: utf-8 -*-
"""
Luồng đầy đủ: tra cứu Giấy chứng nhận theo Số phát hành (AdvancedSearchGiayChungNhan) ->
lấy thẳng object ToChuc (chủ sở hữu tổ chức) đã có sẵn trong kết quả tra cứu đó (không cần
gọi thêm GetThongTinDangKyByTinhHinhDangKyIds) -> nếu có sửa thông tin người đại diện thì
gọi UpdateCaNhan (sửa đúng bản ghi CaNhan gốc) TRƯỚC -> rồi mới build + gửi UpdateToChuc cho
Tổ chức. Field nào truyền giá trị mới thì ghi đè, field nào không truyền thì giữ nguyên dữ
liệu gốc lấy từ MPLIS.

Mặc định KHÔNG gửi update lên MPLIS, chỉ ghi payload ra file để kiểm tra. Thêm --xac-nhan-gui
mới thực sự gửi UpdateCaNhan/UpdateToChuc.

Chạy:
  python flow_capnhat_chusohuu.py "DG 781336" 66 --ten-to-chuc "Tên mới" --dia-chi "Địa chỉ mới"
  (kiểm tra các file *_payload.json xong thì chạy lại thêm --xac-nhan-gui để gửi thật)
"""

import argparse
import copy
import json

import tracuu_giaychungnhan as gcn
import capnhat_chusohuu as cnh


DAI_DIEN_MAPPING = {
    "dai_dien_ho_ten": "hoTen",
    "dai_dien_so_dinh_danh": "maSoDinhDanh",
    "dai_dien_ngay_sinh": "ngaySinh",
    "dai_dien_dia_chi": "diaChi",
}


def build_overrides_tu_args(args):
    """Chỉ đưa field nào người dùng thực sự truyền giá trị (khác None) vào overrides -
    field không truyền sẽ giữ nguyên dữ liệu gốc lấy từ MPLIS. Dùng cho UpdateToChuc:
    field người đại diện có tiền tố "NguoiDaiDien." (chỉ đổi bản sao gắn trong Tổ chức)."""
    overrides = {}

    top_level_mapping = {
        "ten_to_chuc": "tenToChuc",
        "dia_chi": "diaChi",
        "so_dinh_danh": "maSoDinhDanh",
        "ma_doanh_nghiep": "maDoanhNghiep",
        "ma_so_thue": "maSoThue",
    }
    for arg_name, field_name in top_level_mapping.items():
        value = getattr(args, arg_name)
        if value is not None:
            overrides[field_name] = value

    for arg_name, field_name in DAI_DIEN_MAPPING.items():
        value = getattr(args, arg_name)
        if value is not None:
            overrides[f"NguoiDaiDien.{field_name}"] = value

    return overrides


def build_dai_dien_overrides_tu_args(args):
    """Overrides phẳng (không tiền tố) dùng cho UpdateCaNhan - sửa đúng bản ghi CaNhan gốc
    của người đại diện, khác với overrides có tiền tố "NguoiDaiDien." chỉ đổi bản sao gắn
    trong payload UpdateToChuc."""
    overrides = {}
    for arg_name, field_name in DAI_DIEN_MAPPING.items():
        value = getattr(args, arg_name)
        if value is not None:
            overrides[field_name] = value
    return overrides


def ghi_de_giay_to_bo_sung(to_chuc, ngay_cap=None, so_giay_to=None):
    """Sửa ngày cấp / số giấy tờ (GPKD) trong ListGiayToBoSung[0] nếu có giá trị mới truyền vào;
    field nào không truyền thì giữ nguyên. Trả về to_chuc (bản sao, không sửa object gốc)."""
    if ngay_cap is None and so_giay_to is None:
        return to_chuc
    to_chuc = copy.deepcopy(to_chuc)
    danh_sach = to_chuc.get("ListGiayToBoSung") or []
    if danh_sach:
        if ngay_cap is not None:
            danh_sach[0]["ngayCap"] = ngay_cap
        if so_giay_to is not None:
            danh_sach[0]["soGiayTo"] = so_giay_to
    return to_chuc


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("so_phat_hanh")
    parser.add_argument("tinh_id", type=int, nargs="?", default=66)

    parser.add_argument("--ten-to-chuc")
    parser.add_argument("--dia-chi")
    parser.add_argument("--so-dinh-danh", help="maSoDinhDanh cấp Tổ chức")
    parser.add_argument("--ma-doanh-nghiep")
    parser.add_argument("--ma-so-thue")
    parser.add_argument("--ngay-cap-giay-to", help="ngayCap trong ListGiayToBoSung[0], định dạng ISO vd 2024-01-15T00:00:00.000Z")
    parser.add_argument("--so-giay-to", help="soGiayTo (số GPKD) trong ListGiayToBoSung[0]")

    parser.add_argument("--dai-dien-ho-ten")
    parser.add_argument("--dai-dien-so-dinh-danh")
    parser.add_argument("--dai-dien-ngay-sinh", help="ISO vd 1990-05-20T00:00:00.000Z")
    parser.add_argument("--dai-dien-dia-chi")

    parser.add_argument("--xac-nhan-gui", action="store_true",
                         help="Gửi update thật lên MPLIS. Không truyền thì chỉ ghi payload ra file để kiểm tra.")
    args = parser.parse_args()

    print("Dán Cookie và __RequestVerificationToken (F12 > Network > request bất kỳ tới dla.mplis.gov.vn).")
    cookie_str = input("Dán Cookie: ").strip()
    token = input("Dán __RequestVerificationToken: ").strip()
    session = gcn.build_session_from_manual_input(cookie_str, token)

    # Bước 1: tra GCN theo Số phát hành -> lấy thẳng ToChuc (chủ sở hữu tổ chức) trong kết quả
    js_gcn = gcn.tra_cuu_giay_chung_nhan(session, args.so_phat_hanh, args.tinh_id)
    rows = js_gcn.get("data") or []
    if not rows:
        print("Không tìm thấy GCN nào khớp Số phát hành:", args.so_phat_hanh)
        return

    to_chuc = None
    for row in rows:
        to_chuc = gcn.lay_to_chuc_tu_gcn_row(row)
        if to_chuc:
            break
    if not to_chuc:
        print("Không tìm thấy Chủ sở hữu dạng Tổ chức trong hồ sơ này (có thể chủ là cá nhân/hộ gia đình).")
        return
    print("Đã lấy ToChuc, toChucId:", to_chuc.get("toChucId"))

    # Bước 2: nếu có sửa thông tin người đại diện -> UpdateCaNhan TRƯỚC (sửa đúng bản ghi
    # CaNhan gốc; ghi đè "NguoiDaiDien.xxx" trong payload UpdateToChuc chỉ đổi bản sao gắn
    # trong Tổ chức, không đổi bản ghi CaNhan thật).
    dai_dien_overrides = build_dai_dien_overrides_tu_args(args)
    if dai_dien_overrides:
        ca_nhan = to_chuc.get("NguoiDaiDien")
        if not ca_nhan:
            print("Có overrides người đại diện nhưng ToChuc không có NguoiDaiDien - dừng lại.")
            return

        ca_nhan_payload = cnh.build_update_ca_nhan_payload(ca_nhan, **dai_dien_overrides)
        with open("update_ca_nhan_payload.json", "w", encoding="utf-8") as f:
            json.dump(ca_nhan_payload, f, ensure_ascii=False, indent=2)
        print("Đã ghi payload người đại diện vào update_ca_nhan_payload.json - kiểm tra kỹ trước khi gửi.")

        if args.xac_nhan_gui:
            ca_nhan_result = cnh.update_ca_nhan(session, ca_nhan_payload)
            with open("update_ca_nhan_result.json", "w", encoding="utf-8") as f:
                json.dump(ca_nhan_result, f, ensure_ascii=False, indent=2)
            print("Đã gửi UpdateCaNhan, xem kết quả trong update_ca_nhan_result.json")

    # Bước 3: áp overrides Tổ chức (chỉ field có giá trị mới), giữ nguyên field khác
    to_chuc = ghi_de_giay_to_bo_sung(to_chuc, args.ngay_cap_giay_to, args.so_giay_to)
    overrides = build_overrides_tu_args(args)
    payload = cnh.build_update_to_chuc_payload(to_chuc, **overrides)

    with open("update_to_chuc_payload.json", "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print("Đã ghi payload cập nhật vào update_to_chuc_payload.json - kiểm tra kỹ trước khi gửi.")

    if not args.xac_nhan_gui:
        print("Chưa gửi update (thêm --xac-nhan-gui sau khi đã kiểm tra payload để gửi thật).")
        return

    result = cnh.update_to_chuc(session, payload)
    with open("update_to_chuc_result.json", "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print("Đã gửi update, xem kết quả trong update_to_chuc_result.json")


if __name__ == "__main__":
    main()
