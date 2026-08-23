# -*- coding: utf-8 -*-
"""
TEST 1 BẢN GHI — FULL FLOW 3 BƯỚC:
  1. Search theo Số biên nhận (AdvancedSearchHoSoTiepNhan) → lấy Id (GUID)
  2. GetHoSoTiepNhanById với Id đó → lấy duongDanFile trong danhSachVanBanChoBoSung
  3. Tải file qua FileHandler.ashx → lưu tên VBTD_{SOBIENNHAN}.pdf
Chạy: python test_full_flow.py → đăng nhập Chrome xong bấm Enter
"""

import json
import re
import time
import requests

from dotenv import load_dotenv
import os

try:
    from google import genai
    load_dotenv()
    GEMINI_CLIENT = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
except Exception as _e:
    GEMINI_CLIENT = None
    print(f"⚠ Không khởi tạo được Gemini ({_e}) — sẽ bỏ qua bước kiểm tra nội dung.")

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

MAHOSO = "H15.50-251203-2035"
TINH_ID = "66"

REFERER_URL = "https://dla.mplis.gov.vn/dc/DonDangKy/KeKhaiDangKyV2"
URL_SEARCH = "https://dla.mplis.gov.vn/dc/DangKyAjax/AdvancedSearchHoSoTiepNhan"
URL_GET_BY_ID = "https://dla.mplis.gov.vn/dc/CungCapThongTinHoSoTiepNhan/GetHoSoTiepNhanById"
URL_FILE = "https://dla.mplis.gov.vn/dc/Handlers/FileHandler.ashx"


# ---------------- payload ----------------

def build_payload_search(so_bien_nhan, tinh_id=TINH_ID):
    return [
        ("start", "0"),
        ("length", "10"),
        ("model[tinhId]", tinh_id),
        ("model[huyenId]", ""),
        ("model[xaId]", ""),
        ("model[quytrinh]", ""),
        ("model[state]", ""),
        ("model[soBienNhan]", so_bien_nhan),
        ("model[laHoSoMotCua]", "false"),
        ("model[tiepNhanTuNgay]", ""),
        ("model[tiepNhanDenNgay]", ""),
        ("model[henTraTuNgay]", ""),
        ("model[henTraDenNgay]", ""),
        ("model[trangThaiHoSo]", "0"),
        ("model[trangThaiKetISO][]", "0"),
        ("model[diaChiTaiSan]", ""),
        ("model[soThua]", ""),
        ("model[soTo]", ""),
        ("model[diaChi]", ""),
        ("model[hoTen]", ""),
        ("model[soDienThoai]", ""),
        ("model[giayChungMinh]", ""),
        ("model[daXuLy]", "-1"),
    ]


# ---------------- session (selenium) ----------------

def lay_token(driver):
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


def build_session(driver):
    token = lay_token(driver)
    if not token:
        raise RuntimeError("Không lấy được token — kiểm tra đã đăng nhập chưa.")
    s = requests.Session()
    s.headers.update({
        "User-Agent": driver.execute_script("return navigator.userAgent;"),
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "https://dla.mplis.gov.vn",
        "Referer": REFERER_URL,
        "__requestverificationtoken": token,
        "__RequestVerificationToken": token,
        "RequestVerificationToken": token,
    })
    for c in driver.get_cookies():
        s.cookies.set(name=c["name"], value=c["value"],
                      domain=c.get("domain"), path=c.get("path", "/"))
    print("✅ Đã lấy session + token.")
    return s


# ---------------- helpers ----------------

FORM_HEADERS = {"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}

# Bảng loại văn bản phát sinh (loaiVanBanPhatSinh)
LOAI_VAN_BAN = {
    0: "Chờ bổ sung hồ sơ",
    2: "Chờ công dân thực hiện NVTC",
    3: "Gia hạn hồ sơ",
    4: "Tạm dừng hồ sơ",
    5: "Chờ cơ quan thuế xác định NVTC",
    6: "Niêm yết hồ sơ",
}


def parse_dotnet_date(v):
    """'/Date(1777248998173)/' → int ms (để sort tìm văn bản mới nhất), lỗi → 0."""
    m = re.search(r"/Date\((-?\d+)\)/", str(v or ""))
    return int(m.group(1)) if m else 0


def post_form_json(session, url, data, name):
    r = session.post(url, data=data, headers=FORM_HEADERS, timeout=60, allow_redirects=False)
    print(f"[{name}] status={r.status_code}  ct={r.headers.get('Content-Type')}")
    if r.status_code == 200 and "json" in (r.headers.get("Content-Type") or ""):
        return r.json()
    print("Location:", r.headers.get("Location"))
    print("Body:", r.text[:400])
    raise RuntimeError(f"{name}: không nhận được JSON.")


def safe_filename(s):
    return re.sub(r'[\\/:*?"<>|]', "_", s)


# ---------------- 3 bước ----------------

def buoc1_search(session, so_bien_nhan):
    """Search theo SBN → trả về GUID Id của hồ sơ."""
    js = post_form_json(session, URL_SEARCH, build_payload_search(so_bien_nhan), "SEARCH")
    rows = js.get("data") or []
    if isinstance(js.get("value"), list):
        rows = js["value"]
    if not rows:
        raise RuntimeError(f"Không tìm thấy hồ sơ với SBN '{so_bien_nhan}'.")

    # Ưu tiên bản ghi khớp chính xác soBienNhan
    khop = [r for r in rows if str(r.get("soBienNhan", "")).strip().upper()
            == so_bien_nhan.strip().upper()]
    item = khop[0] if khop else rows[0]

    guid = item.get("Id")
    print(f"→ Bước 1 OK: Id={guid}, tiepNhanHoSoId={item.get('tiepNhanHoSoId')}, "
          f"soBienNhan={item.get('soBienNhan')}")
    if not guid:
        raise RuntimeError("Bản ghi search không có field 'Id'. Xem response_search.json")
    return guid


def buoc2_get_by_id(session, ho_so_guid):
    """GetHoSoTiepNhanById → trả về list duongDanFile của văn bản tạm dừng/chờ bổ sung."""
    # Gửi kèm nhiều tên tham số — MVC binder sẽ nhận đúng cái nó cần, bỏ qua phần thừa
    data = {
        "hoSoTiepNhanID": ho_so_guid,
        "hoSoTiepNhanId": ho_so_guid,
        "id": ho_so_guid,
    }
    js = post_form_json(session, URL_GET_BY_ID, data, "GET_BY_ID")
    if not js.get("success"):
        raise RuntimeError("GetHoSoTiepNhanById lỗi: " + str(js)[:400])

    with open("response_getbyid.json", "w", encoding="utf-8") as f:
        json.dump(js, f, ensure_ascii=False, indent=2)

    # Ưu tiên danhSachVanBanChoBoSung, fallback VanBanPhatSinhs
    vbs = js.get("danhSachVanBanChoBoSung") or []
    if not vbs:
        value = js.get("value") or {}
        vbs = value.get("VanBanPhatSinhs") or []
    if not vbs:
        raise RuntimeError("Hồ sơ không có văn bản phát sinh nào.")

    # Log toàn bộ văn bản + loại để soi
    print(f"→ Hồ sơ có {len(vbs)} văn bản phát sinh:")
    for vb in vbs:
        loai = vb.get("loaiVanBanPhatSinh")
        ten_loai = LOAI_VAN_BAN.get(loai, f"Không rõ ({loai})")
        co_file = "có file" if vb.get("duongDanFile") else "KHÔNG có file"
        print(f"   • vanBanPhatSinhId={vb.get('vanBanPhatSinhId')} | loại {loai} - {ten_loai} "
              f"| {co_file} | tenFile={vb.get('tenFile')}")

    # Lọc: chỉ lấy loaiVanBanPhatSinh = 0 (Chờ bổ sung hồ sơ) VÀ có duongDanFile
    ung_vien = [
        vb for vb in vbs
        if vb.get("loaiVanBanPhatSinh") == 0 and vb.get("duongDanFile")
    ]
    if not ung_vien:
        raise RuntimeError("Không có văn bản loại 0 (Chờ bổ sung hồ sơ) nào có duongDanFile.")

    # Lấy văn bản có NGÀY MỚI NHẤT: ưu tiên ngayVanBan, trùng thì so CreatedDate
    vb_moi_nhat = max(
        ung_vien,
        key=lambda vb: (parse_dotnet_date(vb.get("ngayVanBan")),
                        parse_dotnet_date(vb.get("CreatedDate"))),
    )
    print(f"→ Bước 2 OK: chọn văn bản mới nhất vanBanPhatSinhId={vb_moi_nhat.get('vanBanPhatSinhId')}, "
          f"ngayVanBan={vb_moi_nhat.get('ngayVanBan')}, "
          f"duongDanFile={vb_moi_nhat.get('duongDanFile')} ({vb_moi_nhat.get('tenFile')})")

    return [{"duongDanFile": vb_moi_nhat.get("duongDanFile"),
             "tenFile": vb_moi_nhat.get("tenFile")}]


def buoc3_tai_file(session, duong_dan_file, ten_luu):
    """Tải file qua FileHandler.ashx, thử các tên tham số phổ biến."""
    for param in ("id", "Id", "DocId", "fileId"):
        r = session.get(URL_FILE, params={param: duong_dan_file}, timeout=120)
        ct = (r.headers.get("Content-Type") or "").lower()
        print(f"[FILE ?{param}=] status={r.status_code}  ct={ct}  size={len(r.content)}")
        if r.status_code == 200 and len(r.content) > 500 and "html" not in ct:
            with open(ten_luu, "wb") as f:
                f.write(r.content)
            print(f"→ Bước 3 OK: đã lưu {ten_luu} ({len(r.content):,} bytes)")
            return ten_luu
    raise RuntimeError("Không tải được file — thử mở DevTools xem FileHandler.ashx dùng tham số gì.")


PROMPT_KIEM_TRA = """Đây là văn bản tạm dừng/chờ bổ sung hồ sơ đất đai. Hãy kiểm tra:
1. Văn bản có nêu lý do tạm dừng/bổ sung rõ ràng không? Lý do có phù hợp không?
2. Văn bản có chữ ký và con dấu không? (không có chữ ký hoặc không có con dấu = KHÔNG PHÙ HỢP)

Trả về DUY NHẤT một JSON, không markdown, không giải thích thêm, theo mẫu:
{"phu_hop": true hoặc false, "tom_tat": "tóm tắt ngắn gọn lý do tạm dừng", "ly_do_khong_phu_hop": "để trống nếu phù hợp, ngược lại nêu ngắn gọn vì sao không phù hợp"}"""


def buoc4_kiem_tra_gemini(file_path):
    """Upload PDF lên Gemini → kiểm tra sự phù hợp + tóm tắt. Trả về dict kết quả."""
    if GEMINI_CLIENT is None:
        return {"phu_hop": None, "tom_tat": "", "ly_do_khong_phu_hop": "Chưa cấu hình GEMINI_API_KEY"}

    pdf = GEMINI_CLIENT.files.upload(file=file_path)
    response = GEMINI_CLIENT.models.generate_content(
        model="gemini-3.5-flash",
        contents=[pdf, PROMPT_KIEM_TRA],
    )
    text = (response.text or "").strip()
    print(f"→ Bước 4 (Gemini) trả về: {text[:300]}")

    # Bóc JSON (phòng model bọc ```json ... ```)
    text = re.sub(r"^```(json)?|```$", "", text, flags=re.MULTILINE).strip()
    try:
        kq = json.loads(text)
    except Exception:
        kq = {"phu_hop": None, "tom_tat": text[:500], "ly_do_khong_phu_hop": "Không parse được JSON"}

    trang_thai = {True: "✅ PHÙ HỢP", False: "❌ KHÔNG PHÙ HỢP"}.get(kq.get("phu_hop"), "⚠ KHÔNG RÕ")
    print(f"→ Bước 4 OK: {trang_thai} | Tóm tắt: {kq.get('tom_tat')}")
    if kq.get("ly_do_khong_phu_hop"):
        print(f"   Lý do không phù hợp: {kq['ly_do_khong_phu_hop']}")
    return kq


# ---------------- main ----------------

def main():
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                              options=options)
    driver.get(REFERER_URL)
    time.sleep(2)

    input(">>> Đăng nhập xong trên Chrome rồi bấm Enter ở đây... ")
    session = build_session(driver)
    driver.quit()

    # Bước 1: search theo SBN → GUID Id
    guid = buoc1_search(session, MAHOSO)

    # Bước 2: GetHoSoTiepNhanById → duongDanFile
    files = buoc2_get_by_id(session, guid)

    # Bước 3: tải file, đặt tên VBTD_{SBN}.pdf
    sbn_sach = safe_filename(MAHOSO)
    ket_qua_kiem_tra = []
    for i, f_ in enumerate(files, start=1):
        ten = f"VBTD_{sbn_sach}.pdf" if len(files) == 1 else f"VBTD_{sbn_sach}_{i}.pdf"
        buoc3_tai_file(session, f_["duongDanFile"], ten)

        # Bước 4: đọc văn bản bằng Gemini → tóm tắt + kiểm tra phù hợp
        kq = buoc4_kiem_tra_gemini(ten)
        kq["file"] = ten
        ket_qua_kiem_tra.append(kq)

    with open(f"ket_qua_kiem_tra_{sbn_sach}.json", "w", encoding="utf-8") as f:
        json.dump(ket_qua_kiem_tra, f, ensure_ascii=False, indent=2)
    print(f"💾 Đã lưu kết quả kiểm tra → ket_qua_kiem_tra_{sbn_sach}.json")

    print("===== XONG =====")


if __name__ == "__main__":
    main()