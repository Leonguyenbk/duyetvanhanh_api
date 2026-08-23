# -*- coding: utf-8 -*-
"""
Tool kiểm tra VĂN BẢN TẠM DỪNG/CHỜ BỔ SUNG hàng loạt theo Số biên nhận (MPLIS)
Flow mỗi bản ghi:
  1. Search theo Số biên nhận (AdvancedSearchHoSoTiepNhan) → GUID Id
  2. GetHoSoTiepNhanById → chọn văn bản MỚI NHẤT có file, bất kể loại
     (ưu tiên ngayVanBan, trùng thì CreatedDate). Không có văn bản có file →
     ghi "Không có văn bản", bỏ qua bước 3-4
  3. Tải file qua FileHandler.ashx → VBTD_{SBN}.pdf vào folder đã chọn
  4. Gemini đọc PDF → tóm tắt + kiểm tra phù hợp (ưu tiên model chọn trên UI,
     lỗi thì tự fallback sang model kế tiếp)
Kết quả xuất Excel: Bước xử lý | Loại văn bản | Ngày văn bản | Phù hợp | Lý do...
Cài đặt: pip install selenium webdriver-manager requests pandas openpyxl google-genai python-dotenv
.env cùng thư mục: GEMINI_API_KEY=...
"""

import os
import re
import json
import time
import threading
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

try:
    from dotenv import load_dotenv
    import sys
    # Ưu tiên .env nằm CẠNH exe/script; không có thì fallback load_dotenv()
    # mặc định — cơ chế này tự dò ngược lên các thư mục cha (vd exe trong
    # dist/ mà .env ở thư mục mẹ vẫn tìm thấy).
    if getattr(sys, "frozen", False):
        _thu_muc = os.path.dirname(sys.executable)
    else:
        _thu_muc = os.path.dirname(os.path.abspath(__file__))
    _env_canh_exe = os.path.join(_thu_muc, ".env")
    if os.path.isfile(_env_canh_exe):
        load_dotenv(_env_canh_exe)
    else:
        load_dotenv()
except Exception:
    pass

try:
    from google import genai
except Exception:
    genai = None


# ============================ CẤU HÌNH ============================

REFERER_URL = "https://dla.mplis.gov.vn/dc/DonDangKy/KeKhaiDangKyV2"
URL_SEARCH = "https://dla.mplis.gov.vn/dc/DangKyAjax/AdvancedSearchHoSoTiepNhan"
URL_GET_BY_ID = "https://dla.mplis.gov.vn/dc/CungCapThongTinHoSoTiepNhan/GetHoSoTiepNhanById"
URL_FILE = "https://dla.mplis.gov.vn/dc/Handlers/FileHandler.ashx"

TIMEOUT = 120
TINH_ID_MAC_DINH = "66"

# Model AI: model chọn trên UI được ưu tiên, lỗi thì thử tiếp các model còn lại
DANH_SACH_MODEL = ["gemini-3.5-flash", "gemini-2.5-flash"]

COL_SO_BIEN_NHAN = "Số biên nhận"
REQUIRED_COLS = [COL_SO_BIEN_NHAN]

LOAI_VAN_BAN = {
    0: "Chờ bổ sung hồ sơ",
    2: "Chờ công dân thực hiện NVTC",
    3: "Gia hạn hồ sơ",
    4: "Tạm dừng hồ sơ",
    5: "Chờ cơ quan thuế xác định NVTC",
    6: "Niêm yết hồ sơ",
}

PROMPT_KIEM_TRA = """Đây là văn bản tạm dừng/chờ bổ sung hồ sơ đất đai. Hãy kiểm tra:
1. Văn bản có nêu lý do tạm dừng/bổ sung rõ ràng không? Lý do có phù hợp không?
2. Văn bản có chữ ký và con dấu không? (không có chữ ký hoặc không có con dấu = KHÔNG PHÙ HỢP)

Trả về DUY NHẤT một JSON, không markdown, không giải thích thêm, theo mẫu:
{"phu_hop": true hoặc false, "tom_tat": "tóm tắt ngắn gọn lý do tạm dừng", "ly_do_khong_phu_hop": "để trống nếu phù hợp, ngược lại nêu ngắn gọn vì sao không phù hợp"}"""


# ============================ HELPER ============================

def clean_cell(v):
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


def safe_filename(s):
    return re.sub(r'[\\/:*?"<>|]', "_", str(s))


def parse_dotnet_date(v):
    """'/Date(1777248998173)/' → int ms, lỗi → 0."""
    m = re.search(r"/Date\((-?\d+)\)/", str(v or ""))
    return int(m.group(1)) if m else 0


def dotnet_date_to_str(v):
    """'/Date(ms)/' → 'dd/mm/yyyy' theo giờ VN, không parse được → ''."""
    ms = parse_dotnet_date(v)
    if ms <= 0:
        return ""
    dt = datetime.fromtimestamp(ms / 1000, tz=timezone.utc) + timedelta(hours=7)
    return dt.strftime("%d/%m/%Y")


def build_payload_search(so_bien_nhan, tinh_id=TINH_ID_MAC_DINH):
    return [
        ("start", "0"),
        ("length", "10"),
        ("model[tinhId]", clean_cell(tinh_id)),
        ("model[huyenId]", ""),
        ("model[xaId]", ""),
        ("model[quytrinh]", ""),
        ("model[state]", ""),
        ("model[soBienNhan]", clean_cell(so_bien_nhan)),
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


# ============================ GEMINI ============================

class GeminiChecker:
    def __init__(self, log_fn, api_key=""):
        self.log = log_fn
        self.client = None
        # Ưu tiên key nhập trên UI, để trống thì lấy GEMINI_API_KEY trong .env
        api_key = (api_key or "").strip() or os.getenv("GEMINI_API_KEY")
        if genai is None:
            self.log("⚠ Chưa cài google-genai — bỏ qua bước kiểm tra AI.")
        elif not api_key:
            self.log("⚠ Chưa nhập API key (UI) và không có GEMINI_API_KEY trong .env — bỏ qua kiểm tra AI.")
        else:
            try:
                self.client = genai.Client(api_key=api_key)
            except Exception as e:
                self.log(f"⚠ Không khởi tạo được Gemini: {e}")

    def kiem_tra(self, file_path, model_uu_tien):
        """Đọc PDF, thử lần lượt model ưu tiên → fallback. Trả dict kết quả."""
        if self.client is None:
            return {"phu_hop": None, "tom_tat": "",
                    "ly_do_khong_phu_hop": "Chưa cấu hình Gemini", "model": ""}

        # Model chọn trên UI đứng đầu, các model còn lại nối sau làm fallback
        thu_tu = [model_uu_tien] + [m for m in DANH_SACH_MODEL if m != model_uu_tien]

        pdf = self.client.files.upload(file=file_path)
        loi_cuoi = ""
        for model in thu_tu:
            try:
                response = self.client.models.generate_content(
                    model=model,
                    contents=[pdf, PROMPT_KIEM_TRA],
                )
                text = (response.text or "").strip()
                text = re.sub(r"^```(json)?|```$", "", text, flags=re.MULTILINE).strip()
                try:
                    kq = json.loads(text)
                except Exception:
                    kq = {"phu_hop": None, "tom_tat": text[:500],
                          "ly_do_khong_phu_hop": "Không parse được JSON"}
                kq["model"] = model
                return kq
            except Exception as e:
                loi_cuoi = str(e)
                self.log(f"      ⚠ Model {model} lỗi: {loi_cuoi[:150]} → thử model tiếp theo...")

        return {"phu_hop": None, "tom_tat": "",
                "ly_do_khong_phu_hop": f"Tất cả model đều lỗi: {loi_cuoi[:200]}", "model": ""}


# ============================ CORE API ============================

class KhongCoVanBan(Exception):
    """Hồ sơ không có văn bản (loại 0) có file → không cần tải/đọc, không phải lỗi."""
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
            user_box = pass_box = None
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
        token = self.driver.execute_script(js)
        if not token:
            raise RuntimeError("Không lấy được token. Kiểm tra đã đăng nhập chưa.")

        session = requests.Session()
        session.headers.update({
            "User-Agent": self.driver.execute_script("return navigator.userAgent;"),
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": "https://dla.mplis.gov.vn",
            "Referer": REFERER_URL,
            "__requestverificationtoken": token,
            "__RequestVerificationToken": token,
            "RequestVerificationToken": token,
        })
        for c in self.driver.get_cookies():
            session.cookies.set(name=c["name"], value=c["value"],
                                domain=c.get("domain"), path=c.get("path", "/"))
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
    def post_form_json(self, url, data, name):
        headers = dict(self.session.headers)
        headers["Content-Type"] = "application/x-www-form-urlencoded; charset=UTF-8"
        r = self.session.post(url, data=data, headers=headers,
                              timeout=TIMEOUT, allow_redirects=False)
        if r.status_code != 200 or "json" not in (r.headers.get("Content-Type") or ""):
            raise RuntimeError(
                f"{name}: status={r.status_code}, không nhận được JSON "
                f"(có thể mất session). Body: {r.text[:300]}"
            )
        return r.json()

    # ---------- business ----------
    def buoc1_search(self, so_bien_nhan, tinh_id):
        js = self.post_form_json(URL_SEARCH, build_payload_search(so_bien_nhan, tinh_id), "SEARCH")
        rows = js.get("data") or (js.get("value") if isinstance(js.get("value"), list) else None) or []
        if not rows:
            raise RuntimeError("Không tìm thấy hồ sơ với số biên nhận này.")

        sbn = clean_cell(so_bien_nhan).upper()
        khop = [r for r in rows if clean_cell(r.get("soBienNhan")).upper() == sbn]
        item = khop[0] if khop else rows[0]

        guid = item.get("Id")
        if not guid:
            raise RuntimeError("Bản ghi search không có field 'Id'.")
        return guid, item

    def buoc2_chon_van_ban(self, ho_so_guid):
        """GetHoSoTiepNhanById → chọn văn bản loại 0 có ngày mới nhất.
        Trả về (van_ban, buoc_xu_ly). Không có văn bản có file → raise KhongCoVanBan."""
        data = {"hoSoTiepNhanID": ho_so_guid, "hoSoTiepNhanId": ho_so_guid, "id": ho_so_guid}
        js = self.post_form_json(URL_GET_BY_ID, data, "GET_BY_ID")
        if not js.get("success"):
            raise RuntimeError("GetHoSoTiepNhanById lỗi: " + str(js)[:300])

        value = js.get("value") or {}
        state = value.get("state") or {}
        buoc_xu_ly = state.get("Title") or value.get("StateName") or ""

        vbs = js.get("danhSachVanBanChoBoSung") or []
        if not vbs:
            vbs = value.get("VanBanPhatSinhs") or []

        # Không lọc theo loại nữa — lấy văn bản MỚI NHẤT có file, bất kể loại
        ung_vien = [vb for vb in vbs if vb.get("duongDanFile")]
        if not ung_vien:
            raise KhongCoVanBan(buoc_xu_ly)

        vb = max(ung_vien, key=lambda v: (parse_dotnet_date(v.get("ngayVanBan")),
                                          parse_dotnet_date(v.get("CreatedDate"))))
        return vb, buoc_xu_ly

    def buoc3_tai_file(self, duong_dan_file, ten_luu):
        for param in ("id", "Id", "DocId", "fileId"):
            r = self.session.get(URL_FILE, params={param: duong_dan_file}, timeout=TIMEOUT)
            ct = (r.headers.get("Content-Type") or "").lower()
            if r.status_code == 200 and len(r.content) > 500 and "html" not in ct:
                with open(ten_luu, "wb") as f:
                    f.write(r.content)
                return ten_luu
        raise RuntimeError("Không tải được file qua FileHandler.ashx.")


# ============================ TKINTER UI ============================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Kiểm tra văn bản tạm dừng MPLIS hàng loạt")
        root.geometry("900x680")

        self.client = MplisClient(self.log)
        self.gemini = None
        self.running = False
        self.stop_flag = False

        frm = ttk.Frame(root, padding=10)
        frm.pack(fill="x")

        # Username / password
        ttk.Label(frm, text="Username:").grid(row=0, column=0, sticky="w")
        self.ent_user = ttk.Entry(frm, width=28)
        self.ent_user.grid(row=0, column=1, sticky="w", padx=5, pady=3)

        ttk.Label(frm, text="Password:").grid(row=0, column=2, sticky="w")
        self.ent_pass = ttk.Entry(frm, width=28, show="*")
        self.ent_pass.grid(row=0, column=3, sticky="w", padx=5, pady=3)

        # Tỉnh Id + Model AI
        ttk.Label(frm, text="Tỉnh Id:").grid(row=1, column=0, sticky="w")
        self.var_tinh = tk.StringVar(value=TINH_ID_MAC_DINH)
        ttk.Entry(frm, textvariable=self.var_tinh, width=10).grid(row=1, column=1, sticky="w", padx=5)

        ttk.Label(frm, text="Model AI ưu tiên:").grid(row=1, column=2, sticky="w")
        self.var_model = tk.StringVar(value=DANH_SACH_MODEL[0])
        ttk.Combobox(frm, textvariable=self.var_model, values=DANH_SACH_MODEL,
                     width=26, state="readonly").grid(row=1, column=3, sticky="w", padx=5)

        # Gemini API key (ẩn như password, để trống thì dùng GEMINI_API_KEY trong .env)
        ttk.Label(frm, text="Gemini API key:").grid(row=2, column=0, sticky="w")
        self.var_api_key = tk.StringVar(value="")
        self.ent_api_key = ttk.Entry(frm, textvariable=self.var_api_key, width=58, show="*")
        self.ent_api_key.grid(row=2, column=1, columnspan=2, sticky="we", padx=5, pady=3)
        self.var_hien_key = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="Hiện key", variable=self.var_hien_key,
                        command=self.toggle_hien_key).grid(row=2, column=3, sticky="w")

        # Excel
        ttk.Label(frm, text="File Excel (cột 'Số biên nhận'):").grid(row=3, column=0, sticky="w")
        self.var_excel = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_excel, width=58).grid(
            row=3, column=1, columnspan=2, sticky="we", padx=5, pady=3)
        ttk.Button(frm, text="Chọn Excel...", command=self.chon_excel).grid(row=3, column=3, sticky="w")

        # Folder lưu file
        ttk.Label(frm, text="Folder lưu VBTD:").grid(row=4, column=0, sticky="w")
        self.var_folder = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_folder, width=58).grid(
            row=4, column=1, columnspan=2, sticky="we", padx=5, pady=3)
        ttk.Button(frm, text="Chọn folder...", command=self.chon_folder).grid(row=4, column=3, sticky="w")

        # Buttons
        btn_frm = ttk.Frame(root, padding=(10, 0))
        btn_frm.pack(fill="x")
        self.btn_login = ttk.Button(btn_frm, text="1. Mở Chrome đăng nhập", command=self.mo_chrome)
        self.btn_login.pack(side="left", padx=5)
        self.btn_confirm = ttk.Button(btn_frm, text="2. Đã đăng nhập → Lấy session",
                                      command=self.lay_session, state="disabled")
        self.btn_confirm.pack(side="left", padx=5)
        self.btn_run = ttk.Button(btn_frm, text="3. Chạy kiểm tra hàng loạt",
                                  command=self.chay, state="disabled")
        self.btn_run.pack(side="left", padx=5)
        self.btn_stop = ttk.Button(btn_frm, text="Dừng", command=self.dung, state="disabled")
        self.btn_stop.pack(side="left", padx=5)

        # Progress + log
        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(8, 0))
        self.lbl_status = ttk.Label(root, text="Chưa chạy")
        self.lbl_status.pack(anchor="w", padx=10)

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

    def toggle_hien_key(self):
        self.ent_api_key.config(show="" if self.var_hien_key.get() else "*")

    def chon_excel(self):
        f = filedialog.askopenfilename(title="Chọn file Excel",
                                       filetypes=[("Excel", "*.xlsx *.xls"), ("Tất cả", "*.*")])
        if f:
            self.var_excel.set(f)

    def chon_folder(self):
        d = filedialog.askdirectory(title="Chọn folder lưu văn bản tạm dừng")
        if d:
            self.var_folder.set(d)

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
                self.log("Chrome đã mở. Hoàn tất đăng nhập rồi bấm nút 2.")
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
            raise RuntimeError(f"Excel thiếu cột: {', '.join(missing)}")
        df[COL_SO_BIEN_NHAN] = df[COL_SO_BIEN_NHAN].map(clean_cell)
        df = df[df[COL_SO_BIEN_NHAN] != ""].copy()
        truoc = len(df)
        df = df.drop_duplicates(subset=[COL_SO_BIEN_NHAN], keep="first").reset_index(drop=True)
        if truoc != len(df):
            self.log(f"Đã gộp {truoc - len(df)} bản ghi trùng 'Số biên nhận'. Còn {len(df)}.")
        return df

    def chay(self):
        if self.running:
            return
        excel = self.var_excel.get().strip()
        folder = self.var_folder.get().strip()
        if not excel or not os.path.isfile(excel):
            messagebox.showwarning("Thiếu Excel", "Chọn file Excel hợp lệ.")
            return
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Thiếu folder", "Chọn folder lưu văn bản hợp lệ.")
            return
        try:
            df = self.doc_excel(excel)
        except Exception as e:
            messagebox.showerror("Lỗi Excel", str(e))
            return
        if df.empty:
            messagebox.showwarning("Excel rỗng", "Không có bản ghi nào để xử lý.")
            return
        if not messagebox.askyesno("Xác nhận",
                                   f"Sẽ kiểm tra {len(df)} số biên nhận.\n"
                                   f"Model AI ưu tiên: {self.var_model.get()}\nTiếp tục?"):
            return

        self.gemini = GeminiChecker(self.log, api_key=self.var_api_key.get())
        self.running = True
        self.stop_flag = False
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progress.config(maximum=len(df), value=0)

        tinh_id = self.var_tinh.get().strip() or TINH_ID_MAC_DINH
        model = self.var_model.get()
        threading.Thread(target=self._run_batch,
                         args=(df, excel, folder, tinh_id, model), daemon=True).start()

    def _save_results_excel(self, out_path, results):
        if not results:
            return
        columns = [
            COL_SO_BIEN_NHAN,
            "Bước xử lý",
            "Loại văn bản",
            "Ngày văn bản",
            "Tên file",
            "Phù hợp",
            "Tóm tắt / Lý do",
            "Model AI",
            "Kết quả",
            "Lỗi",
        ]
        pd.DataFrame(results, columns=columns).to_excel(out_path, index=False)

    def _run_batch(self, df, excel_path, folder, tinh_id, model):
        results = []
        total = len(df)
        out = os.path.join(
            os.path.dirname(excel_path),
            f"ket_qua_kiem_tra_vbtd_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        self.log(f"📄 Sẽ lưu kết quả sau mỗi hồ sơ vào: {out}")

        for i, row in df.iterrows():
            if self.stop_flag:
                self.log("⏹ Đã dừng theo yêu cầu.")
                break

            sbn = row[COL_SO_BIEN_NHAN]
            self.set_status(f"Đang xử lý {i + 1}/{total}: {sbn}")
            self.log(f"[{i + 1}/{total}] Số biên nhận: {sbn}")

            kq = {
                COL_SO_BIEN_NHAN: sbn,
                "Bước xử lý": "",
                "Loại văn bản": "",
                "Ngày văn bản": "",
                "Tên file": "",
                "Phù hợp": "",
                "Tóm tắt / Lý do": "",
                "Model AI": "",
                "Kết quả": "OK",
                "Lỗi": "",
            }

            try:
                # Bước 1: search → GUID
                guid, item = self.client.buoc1_search(sbn, tinh_id)
                self.log(f"   → Id={guid}")

                # Bước 2: chọn văn bản mới nhất loại 0
                try:
                    vb, buoc_xu_ly = self.client.buoc2_chon_van_ban(guid)
                except KhongCoVanBan as e:
                    kq["Bước xử lý"] = str(e)
                    kq["Kết quả"] = "KHÔNG CÓ VĂN BẢN"
                    self.log("   → ⏭ Không có văn bản (loại 0) có file — bỏ qua tải/đọc.")
                    results.append(kq)
                    self.root.after(0, lambda v=i + 1: self.progress.config(value=v))
                    continue

                loai = vb.get("loaiVanBanPhatSinh")
                kq["Bước xử lý"] = buoc_xu_ly
                kq["Loại văn bản"] = f"{loai} - {LOAI_VAN_BAN.get(loai, 'Không rõ')}"
                kq["Ngày văn bản"] = dotnet_date_to_str(vb.get("ngayVanBan"))
                self.log(f"   → Văn bản: {kq['Loại văn bản']} | ngày {kq['Ngày văn bản']} "
                         f"| duongDanFile={vb.get('duongDanFile')}")

                # Bước 3: tải file
                ten_file = os.path.join(folder, f"VBTD_{safe_filename(sbn)}.pdf")
                self.client.buoc3_tai_file(vb["duongDanFile"], ten_file)
                kq["Tên file"] = os.path.basename(ten_file)
                self.log(f"   → Đã tải {kq['Tên file']}")

                # Bước 4: Gemini kiểm tra
                ai = self.gemini.kiem_tra(ten_file, model)
                kq["Model AI"] = ai.get("model", "")
                if ai.get("phu_hop") is True:
                    kq["Phù hợp"] = "PHÙ HỢP"
                    kq["Tóm tắt / Lý do"] = ai.get("tom_tat", "")
                elif ai.get("phu_hop") is False:
                    kq["Phù hợp"] = "KHÔNG PHÙ HỢP"
                    kq["Tóm tắt / Lý do"] = (ai.get("ly_do_khong_phu_hop") or "") \
                        + (f" | Tóm tắt: {ai.get('tom_tat')}" if ai.get("tom_tat") else "")
                else:
                    kq["Phù hợp"] = "KHÔNG RÕ"
                    kq["Tóm tắt / Lý do"] = ai.get("ly_do_khong_phu_hop") or ai.get("tom_tat", "")
                self.log(f"   → ✅ {kq['Phù hợp']} | {kq['Tóm tắt / Lý do'][:120]}")

            except Exception as e:
                kq["Kết quả"] = "LỖI"
                kq["Lỗi"] = str(e)
                self.log(f"   → ❌ {e}")

            results.append(kq)
            try:
                self._save_results_excel(out, results)
                self.log(f"📄 Đã lưu kết quả sau hồ sơ {i + 1}/{total}: {out}")
            except Exception as e:
                self.log(f"⚠ Không lưu được kết quả tạm thời: {e}")
            self.root.after(0, lambda v=i + 1: self.progress.config(value=v))

        # Xuất Excel kết quả cuối cùng
        try:
            self._save_results_excel(out, results)
            self.log(f"📄 Đã xuất file kết quả cuối cùng: {out}")
        except Exception as e:
            self.log(f"⚠ Không xuất được file kết quả: {e}")

        ok = sum(1 for r in results if r["Kết quả"] == "OK")
        khong_vb = sum(1 for r in results if r["Kết quả"] == "KHÔNG CÓ VĂN BẢN")
        loi = sum(1 for r in results if r["Kết quả"] == "LỖI")
        self.log(f"===== HOÀN TẤT: {ok} OK | {khong_vb} không có văn bản | {loi} lỗi / {len(results)} =====")
        self.set_status(f"Hoàn tất: {ok} OK | {khong_vb} không có VB | {loi} lỗi / {len(results)}")

        self.running = False
        self.root.after(0, lambda: (self.btn_run.config(state="normal"),
                                    self.btn_stop.config(state="disabled")))

    def on_close(self):
        if self.running and not messagebox.askyesno("Đang chạy", "Đang xử lý, thoát luôn?"):
            return
        self.client.close_browser()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()