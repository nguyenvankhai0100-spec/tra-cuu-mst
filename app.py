#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ứng dụng web (chạy local) để tra cứu mã số thuế hàng loạt và điền vào file Excel.

Tự nhận diện cột (MST, Tên, Địa chỉ, Cơ quan thuế, Trạng thái) dựa theo tiêu đề
cột trong file, nên mỗi người dùng có thể dùng file với layout khác nhau -
nếu tự nhận diện sai, người dùng chỉnh lại bằng dropdown trên giao diện trước
khi chạy.

Chạy:
    pip install -r requirements.txt
    python app.py
Sau đó mở trình duyệt: http://127.0.0.1:5000
"""

import io
import os
import re
import secrets
import socket
import threading
import time
import unicodedata
import uuid
import webbrowser
from datetime import timedelta
from pathlib import Path

import requests
from flask import (Flask, jsonify, redirect, render_template, request,
                    send_file, session, url_for)
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

APP_DIR = Path(__file__).resolve().parent
STORE_DIR = APP_DIR / "uploads"
STORE_DIR.mkdir(exist_ok=True)

# Mật khẩu chung để vào app khi được host công khai (deploy lên Internet).
# Đặt biến môi trường APP_PASSWORD trên dịch vụ hosting để bật yêu cầu đăng
# nhập. Khi chạy local (không đặt biến này) app dùng bình thường, không hỏi
# mật khẩu - vì lúc đó chỉ có mình máy bạn truy cập được (127.0.0.1) rồi.
APP_PASSWORD = os.environ.get("APP_PASSWORD")

# LỊCH SỬ (để không ai lặp lại nhầm lẫn này lần nữa): bản trước từng cho
# rằng api.xinvoice.vn (đứng sau Cloudflare) chặn request gọi từ IP hosting/
# cloud (Render, AWS...) nên chuyển việc gọi API sang chạy ở trình duyệt
# người dùng bằng JavaScript. Nhưng thực tế kiểm tra lại (qua route chẩn
# đoán tạm /api/test_lookup, đã xoá sau khi dùng xong) cho thấy: (1) Render
# gọi api.xinvoice.vn HOÀN TOÀN BÌNH THƯỜNG, không hề bị chặn; (2) ngược lại,
# api.xinvoice.vn KHÔNG trả header access-control-allow-origin, nên cách gọi
# từ trình duyệt mới là cách chắc chắn không bao giờ chạy được (trình duyệt
# tự chặn theo CORS, không phải do api chặn) - xem lỗi Console thực tế người
# dùng gặp: "No 'Access-Control-Allow-Origin' header is present...". Vì vậy
# việc gọi API được chuyển lại về server (xem lookup_mst_serverside +
# process_job_serverside bên dưới) - đơn giản và đúng hơn hẳn.
API_TEMPLATE = "https://api.xinvoice.vn/gdt-api/tax-payer-records/{tax_code}"
NOT_FOUND_TEXT = "Không tìm thấy"
ERROR_TEXT = "Lỗi tra cứu"

MAX_PREVIEW_ROWS = 12
MAX_HEADER_SCAN_ROWS = 10

# In-memory store: upload_id -> {"path": Path, "sheet": str}
#                  result_id -> Path
#                  job_id -> job state dict (cập nhật bởi /api/report, /api/waiting, /api/finish)
_uploads = {}
_results = {}
_jobs = {}
_jobs_lock = threading.Lock()

# Bộ nhớ đệm CHUNG giữa mọi lần chạy (mọi người dùng, mọi file) trong cùng
# tiến trình server: mã số thuế đã tra rồi thì lần sau dùng lại kết quả, không
# gọi api.xinvoice.vn nữa. Chỉ lưu kết quả dứt khoát ('ok' / 'not_found'),
# không lưu lỗi. Mỗi mục hết hạn sau MST_CACHE_TTL_HOURS giờ (mặc định 6) vì
# trạng thái mã số thuế có thể thay đổi; đặt 0 để tắt hẳn bộ nhớ đệm. Nằm
# trong RAM nên mất khi server khởi động lại (Render free tự ngủ/khởi động
# lại khi lâu không dùng) - chấp nhận được, chỉ là tra lại từ đầu.
try:
    MST_CACHE_TTL_SECONDS = max(0.0, float(os.environ.get("MST_CACHE_TTL_HOURS", "6"))) * 3600
except ValueError:
    MST_CACHE_TTL_SECONDS = 6 * 3600
MST_CACHE_MAX_ENTRIES = 50000
_mst_cache = {}  # mst chuẩn hoá -> (hết_hạn_lúc, outcome, data)
_mst_cache_lock = threading.Lock()


def _cache_key(mst):
    return str(mst).strip().upper()


def cache_get(mst):
    """Trả {"outcome", "data"} nếu còn hạn, ngược lại None."""
    if MST_CACHE_TTL_SECONDS <= 0:
        return None
    key = _cache_key(mst)
    now = time.time()
    with _mst_cache_lock:
        item = _mst_cache.get(key)
        if item is None:
            return None
        expires_at, outcome, data = item
        if expires_at <= now:
            del _mst_cache[key]
            return None
        return {"outcome": outcome, "data": dict(data) if data else None}


def cache_put(mst, outcome, data):
    if MST_CACHE_TTL_SECONDS <= 0 or outcome not in ("ok", "not_found"):
        return
    now = time.time()
    with _mst_cache_lock:
        _mst_cache[_cache_key(mst)] = (now + MST_CACHE_TTL_SECONDS, outcome, dict(data) if data else None)
        if len(_mst_cache) > MST_CACHE_MAX_ENTRIES:
            for k in [k for k, v in _mst_cache.items() if v[0] <= now]:
                del _mst_cache[k]
            overflow = len(_mst_cache) - MST_CACHE_MAX_ENTRIES
            if overflow > 0:
                for k, _ in sorted(_mst_cache.items(), key=lambda kv: kv[1][0])[:overflow]:
                    del _mst_cache[k]

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25MB
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=7)
# SECRET_KEY nên đặt cố định qua biến môi trường khi deploy (nếu không, mỗi
# lần server khởi động lại mọi người sẽ bị đăng xuất). Chạy local thì không
# quan trọng, tự sinh ngẫu nhiên là đủ.
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)


@app.before_request
def require_login():
    if not APP_PASSWORD:
        return  # không đặt mật khẩu (vd: chạy local) -> khỏi cần đăng nhập
    if request.endpoint in ("login", "static"):
        return
    if session.get("authenticated"):
        return
    if request.path.startswith("/api/"):
        return jsonify({"error": "Phiên đăng nhập đã hết hạn, hãy tải lại trang."}), 401
    return redirect(url_for("login", next=request.path))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session.permanent = True
            session["authenticated"] = True
            return redirect(request.args.get("next") or url_for("index"))
        error = "Sai mật khẩu, vui lòng thử lại."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Chuẩn hóa & nhận diện cột
# ---------------------------------------------------------------------------

def normalize(s):
    if s is None:
        return ""
    s = str(s)
    # "Đ/đ" là chữ cái riêng trong tiếng Việt, NFD không tách được -> xử lý tay trước
    s = s.replace("Đ", "D").replace("đ", "d")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


# Từ khóa nhận diện. Thứ tự xử lý các field CÓ Ý NGHĨA: field nào có cụm từ
# càng đặc trưng/dài thì nên được xét (và "giữ chỗ" cột) trước, để tránh nhận
# nhầm - ví dụ cột "Trạng thái mã số thuế" chứa cả cụm "mã số thuế" nên nếu
# xét mst_col trước sẽ bị nhầm sang cột trạng thái.
FIELD_KEYWORDS = {
    "status_col": [
        "trang thai ma so thue", "trang thai mst", "trang thai",
    ],
    "dept_col": [
        "co quan quan ly thue", "co quan thue quan ly", "co quan thue", "co quan quan ly",
    ],
    "address_col": [
        "dia chi tru so", "dia chi kinh doanh", "dia chi",
    ],
    "name_col": [
        "ten nguoi nop thue", "ten nnt", "ten to chuc ca nhan nop thue", "ten don vi",
    ],
    "mst_col": [
        "ma so thue", "mst", "tax code", "taxcode", "tax id", "taxid",
    ],
}


def guess_header_row(ws):
    """Tìm dòng tiêu đề: dòng có nhiều ô khớp từ khóa nhất trong MAX_HEADER_SCAN_ROWS dòng đầu."""
    best_row, best_score = 1, -1
    max_col = min(ws.max_column, 500)
    max_row = min(ws.max_row, MAX_HEADER_SCAN_ROWS)
    for r in range(1, max_row + 1):
        score = 0
        for c in range(1, max_col + 1):
            val = normalize(ws.cell(row=r, column=c).value)
            if not val:
                continue
            for kws in FIELD_KEYWORDS.values():
                if any(kw in val for kw in kws):
                    score += 1
                    break
        if score > best_score:
            best_score = score
            best_row = r
    return best_row if best_score > 0 else 1


def guess_mapping(ws):
    header_row = guess_header_row(ws)
    max_col = min(ws.max_column, 500)

    assigned = {}
    used_cols = set()

    for field, keywords in FIELD_KEYWORDS.items():
        found_col = None
        # thử từng từ khóa theo thứ tự ưu tiên (cụ thể -> chung)
        for kw in keywords:
            for c in range(1, max_col + 1):
                if c in used_cols:
                    continue
                val = normalize(ws.cell(row=header_row, column=c).value)
                if val and kw in val:
                    found_col = c
                    break
            if found_col:
                break
        if found_col:
            assigned[field] = get_column_letter(found_col)
            used_cols.add(found_col)
        else:
            assigned[field] = None

    # Đoán dòng bắt đầu dữ liệu = dòng đầu tiên sau header_row có ô ở cột MST khác rỗng
    start_row = header_row + 1
    if assigned.get("mst_col"):
        mst_idx = column_index_from_string(assigned["mst_col"])
        for r in range(header_row + 1, min(ws.max_row, header_row + 20) + 1):
            if ws.cell(row=r, column=mst_idx).value not in (None, ""):
                start_row = r
                break

    assigned["header_row"] = header_row
    assigned["start_row"] = start_row
    return assigned


def build_preview(ws):
    """Xem trước dữ liệu: liệt kê đúng từ cột A tới cột cuối cùng THỰC SỰ có
    dữ liệu, thay vì một số cột cố định. Trước đây bị chặn cứng ở 12 cột
    (A..L) nên file nào có cột M, N... trở đi thì mất luôn, không chọn được
    trong ô chọn cột lẫn không hiện trong bảng xem trước."""
    max_row = min(ws.max_row, MAX_PREVIEW_ROWS)
    # Trần quét an toàn để không dò vô hạn nếu Excel "phồng" max_column do
    # định dạng ô để trống — nhưng đủ rộng để không chặn cột dữ liệu thật.
    scan_col_cap = min(ws.max_column, 500)

    last_data_col = 0
    for r in range(1, max_row + 1):
        for c in range(1, scan_col_cap + 1):
            if ws.cell(row=r, column=c).value not in (None, ""):
                if c > last_data_col:
                    last_data_col = c
    max_col = max(last_data_col, 1)

    rows = []
    for r in range(1, max_row + 1):
        row_vals = []
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            row_vals.append("" if v is None else str(v))
        rows.append(row_vals)
    col_letters = [get_column_letter(c) for c in range(1, max_col + 1)]
    return col_letters, rows


# ---------------------------------------------------------------------------
# Gọi API tra cứu MST
# ---------------------------------------------------------------------------

def clean_tax_code(raw):
    if raw is None:
        return ""
    s = str(raw).strip()
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        s = s[:-2]
    return s


_LOOKUP_HEADERS = {
    "Accept": "application/json",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
}


def lookup_mst_serverside(mst, on_wait=None):
    """Gọi api.xinvoice.vn thật, ngay từ server (đã xác nhận Render gọi
    được bình thường - xem ghi chú ở API_TEMPLATE). Có tự thử lại khi gặp
    lỗi tạm thời, lùi thời gian chờ dài hơn hẳn khi gặp lỗi mạng thật (khả
    năng cao là bị giới hạn tốc độ/chặn tạm thời) so với các lỗi vặt khác.
    on_wait(wait_seconds, reason), nếu có, được gọi trước mỗi lần chờ để nơi
    gọi (process_job_serverside) có thể cập nhật bảng tiến trình cho người
    dùng thấy đang chờ, không phải bị treo."""
    max_retries = 2
    base_backoff = 1.5
    last_detail = None
    saw_network_failure = False

    for attempt in range(max_retries + 1):
        try:
            resp = requests.get(API_TEMPLATE.format(tax_code=mst), headers=_LOOKUP_HEADERS, timeout=15)
        except Exception as e:
            saw_network_failure = True
            last_detail = f"Mất kết nối: {e}"
            wait_s = 8 * (attempt + 1)
            if on_wait:
                on_wait(wait_s, "Mất kết nối, đang thử lại...")
            time.sleep(wait_s)
            continue

        if resp.status_code == 200:
            try:
                payload = resp.json()
            except ValueError:
                last_detail = "Phản hồi không phải JSON hợp lệ"
                wait_s = base_backoff * (attempt + 1)
                if on_wait:
                    on_wait(wait_s, "Phản hồi lỗi, đang thử lại...")
                time.sleep(wait_s)
                continue
            if payload and payload.get("success") and payload.get("data"):
                return {"outcome": "ok", "data": payload["data"][0]}
            return {"outcome": "not_found"}

        if resp.status_code == 404:
            return {"outcome": "not_found"}

        body_snippet = ""
        if resp.text:
            body_snippet = " ".join(resp.text.split())[:160]
        last_detail = f"HTTP {resp.status_code}" + (f": {body_snippet}" if body_snippet else "")

        if resp.status_code != 429 and 400 <= resp.status_code < 500:
            return {"outcome": "error", "detail": last_detail}

        if resp.status_code == 429:
            # api.xinvoice.vn giới hạn khoảng 10 lượt/30 giây (thấy được qua
            # header ratelimit trên response) - lùi thời gian chờ nhiều hơn.
            saw_network_failure = True
            wait_s = base_backoff * (attempt + 2) * 2
            reason = "Chờ do giới hạn tốc độ..."
        else:
            wait_s = base_backoff * (attempt + 1)
            reason = "Đang thử lại..."
        if on_wait:
            on_wait(wait_s, reason)
        time.sleep(wait_s)

    return {"outcome": "error", "detail": last_detail or "Lỗi không xác định", "network_failure": saw_network_failure}


def process_job_serverside(job_id):
    """Chạy nền (background thread): lần lượt tra cứu từng mã số thuế còn
    'pending' của job, ghi kết quả vào workbook + trạng thái job qua
    apply_lookup_result/mark_waiting (dùng chung với model cũ), rồi lưu file
    kết quả bằng finish_job khi xong. Tôn trọng nút Tạm dừng/Tiếp tục qua
    job['control']. Nếu 3 mã liên tiếp đều gặp lỗi mạng (networkFailure) -
    dấu hiệu rõ ràng là đang bị giới hạn tốc độ/chặn tạm thời, không phải
    xui từng mã lẻ tẻ - tự nghỉ hẳn 60 giây trước khi thử mã tiếp theo, thay
    vì cứ gọi dồn dập làm tình trạng đó kéo dài thêm.

    Bộ nhớ đệm chung (cache_get/cache_put): mã số thuế đã tra rồi - ở dòng
    khác của lần chạy này, hoặc ở lần chạy/file khác trước đó còn trong hạn -
    thì dùng luôn kết quả cũ, không gọi API lần nữa (và cũng không phải chờ
    giãn cách). Chỉ lưu kết quả dứt khoát ('ok' / 'not_found'); mã bị lỗi thì
    KHÔNG lưu để lần sau còn được thử tra lại."""
    job = _jobs.get(job_id)
    if not job:
        return

    COOLDOWN_AFTER = 3
    COOLDOWN_SECONDS = 60
    consecutive_network_failures = 0

    todo = [r for r in job["rows"] if r["result"] == "pending"]
    for entry in todo:
        while job.get("control") == "pause":
            time.sleep(0.3)

        row = entry["row"]
        mst = entry["mst"]

        cached = cache_get(mst)
        if cached is not None:
            with _jobs_lock:
                apply_lookup_result(job, row, cached["outcome"], data=cached.get("data"), from_cache=True)
            continue

        if consecutive_network_failures >= COOLDOWN_AFTER:
            with _jobs_lock:
                mark_waiting(
                    job, row, COOLDOWN_SECONDS,
                    f"Có vẻ đang bị chặn tạm thời do gọi quá nhanh — tạm nghỉ {COOLDOWN_SECONDS}s rồi tự tiếp tục...",
                )
            time.sleep(COOLDOWN_SECONDS)
            consecutive_network_failures = 0

        def on_wait(wait_s, reason, _row=row):
            with _jobs_lock:
                mark_waiting(job, _row, wait_s, reason)

        result = lookup_mst_serverside(mst, on_wait=on_wait)

        with _jobs_lock:
            apply_lookup_result(job, row, result["outcome"], data=result.get("data"), detail=result.get("detail"))

        if result["outcome"] in ("ok", "not_found"):
            cache_put(mst, result["outcome"], result.get("data"))

        if result.get("network_failure"):
            consecutive_network_failures += 1
        else:
            consecutive_network_failures = 0

        if job["delay"] > 0:
            time.sleep(job["delay"])

    with _jobs_lock:
        finish_job(job, stopped=False)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html", show_logout=bool(APP_PASSWORD))


@app.route("/api/upload", methods=["POST"])
def api_upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "Chưa chọn file"}), 400
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify({"error": "Chỉ hỗ trợ file .xlsx / .xlsm"}), 400

    upload_id = uuid.uuid4().hex
    save_path = STORE_DIR / f"{upload_id}.xlsx"
    f.save(save_path)

    try:
        wb = load_workbook(save_path, data_only=True)
    except Exception as e:
        save_path.unlink(missing_ok=True)
        return jsonify({"error": f"Không đọc được file Excel: {e}"}), 400

    sheet_names = wb.sheetnames
    default_sheet = sheet_names[0]
    ws = wb[default_sheet]

    mapping = guess_mapping(ws)
    col_letters, preview_rows = build_preview(ws)

    _uploads[upload_id] = {"path": save_path}

    return jsonify({
        "upload_id": upload_id,
        "sheets": sheet_names,
        "active_sheet": default_sheet,
        "mapping": mapping,
        "preview": {"columns": col_letters, "rows": preview_rows},
        "max_row": ws.max_row,
        "max_col": min(ws.max_column, 40),
    })


@app.route("/api/sheet_preview", methods=["POST"])
def api_sheet_preview():
    """Khi người dùng đổi sheet, trả lại preview + mapping gợi ý cho sheet đó."""
    data = request.get_json(force=True)
    upload_id = data.get("upload_id")
    sheet_name = data.get("sheet")
    if upload_id not in _uploads:
        return jsonify({"error": "Phiên upload không tồn tại, hãy tải file lại"}), 400

    path = _uploads[upload_id]["path"]
    wb = load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        return jsonify({"error": "Sheet không tồn tại"}), 400
    ws = wb[sheet_name]

    mapping = guess_mapping(ws)
    col_letters, preview_rows = build_preview(ws)

    return jsonify({
        "mapping": mapping,
        "preview": {"columns": col_letters, "rows": preview_rows},
        "max_row": ws.max_row,
        "max_col": min(ws.max_column, 40),
    })


def _find_row_entry(job, row):
    for r in job["rows"]:
        if r["row"] == row:
            return r
    return None


def apply_lookup_result(job, row, outcome, data=None, detail=None, from_cache=False):
    """Ghi kết quả tra cứu MỘT dòng vào workbook + trạng thái job. outcome là
    'ok' / 'not_found' / 'error'. Gọi trực tiếp từ process_job_serverside
    (chạy nền trên server) sau mỗi lần lookup_mst_serverside trả về."""
    row_entry = _find_row_entry(job, row)
    if row_entry is None:
        return False

    ws = job["_ws"]
    name_col = job["name_col"]
    address_col = job["address_col"]
    dept_col = job["dept_col"]
    status_col = job["status_col"]

    row_entry.pop("wait_until", None)
    row_entry["from_cache"] = bool(from_cache)

    if outcome == "ok" and data:
        if name_col:
            ws[f"{name_col}{row}"] = data.get("name", "")
        if address_col:
            ws[f"{address_col}{row}"] = data.get("address", "")
        if dept_col:
            ws[f"{dept_col}{row}"] = data.get("taxDepartment", "")
        if status_col:
            ws[f"{status_col}{row}"] = data.get("status", "")
        row_entry["result"] = "ok"
        row_entry["name"] = data.get("name", "")
        row_entry["status_val"] = data.get("status", "")
        row_entry["detail"] = None
        job["ok"] += 1
    elif outcome == "not_found":
        for col in (name_col, address_col, dept_col, status_col):
            if col:
                ws[f"{col}{row}"] = NOT_FOUND_TEXT
        row_entry["result"] = "not_found"
        if name_col:
            row_entry["name"] = NOT_FOUND_TEXT
        if status_col:
            row_entry["status_val"] = NOT_FOUND_TEXT
        job["not_found"] += 1
    else:
        for col in (name_col, address_col, dept_col, status_col):
            if col:
                ws[f"{col}{row}"] = ERROR_TEXT
        row_entry["result"] = "error"
        row_entry["detail"] = detail or "lỗi không xác định"
        if name_col:
            row_entry["name"] = ERROR_TEXT
        if status_col:
            row_entry["status_val"] = ERROR_TEXT
        job["errors"] += 1

    job["processed"] += 1
    return True


def mark_waiting(job, row, wait_seconds, reason):
    row_entry = _find_row_entry(job, row)
    if row_entry is None:
        return False
    row_entry["result"] = "waiting"
    row_entry["wait_until"] = time.time() + max(0, wait_seconds)
    row_entry["detail"] = reason
    return True


def finish_job(job, stopped):
    """Lưu file kết quả (nếu chạy xong hoặc dừng giữa chừng) - gọi từ cuối
    process_job_serverside khi vòng lặp tra cứu trên server đã kết thúc."""
    if job.get("result_id"):
        return  # đã lưu rồi (tránh lưu 2 lần)
    if stopped:
        job["status"] = "stopped"
    else:
        wb = job["_wb"]
        result_id = uuid.uuid4().hex
        result_path = STORE_DIR / f"{result_id}_result.xlsx"
        wb.save(result_path)
        _results[result_id] = result_path
        job["result_id"] = result_id
        job["status"] = "done"


@app.route("/api/process", methods=["POST"])
def api_process():
    data = request.get_json(force=True)
    upload_id = data.get("upload_id")
    if upload_id not in _uploads:
        return jsonify({"error": "Phiên upload không tồn tại, hãy tải file lại"}), 400

    sheet_name = data.get("sheet")
    mst_col = (data.get("mst_col") or "").strip().upper()
    name_col = (data.get("name_col") or "").strip().upper()
    address_col = (data.get("address_col") or "").strip().upper()
    dept_col = (data.get("dept_col") or "").strip().upper()
    status_col = (data.get("status_col") or "").strip().upper()
    skip_filled = bool(data.get("skip_filled", True))
    try:
        start_row = int(data.get("start_row"))
    except (TypeError, ValueError):
        return jsonify({"error": "Dòng bắt đầu không hợp lệ"}), 400
    try:
        delay = float(data.get("delay", 3.0))
    except (TypeError, ValueError):
        delay = 1.0
    delay = max(0.0, min(delay, 10.0))

    if not mst_col:
        return jsonify({"error": "Bạn cần chọn cột chứa mã số thuế"}), 400
    target_cols = [c for c in (name_col, address_col, dept_col, status_col) if c]
    if not target_cols:
        return jsonify({"error": "Bạn cần chọn ít nhất 1 cột để điền kết quả"}), 400

    path = _uploads[upload_id]["path"]
    wb = load_workbook(path)
    if sheet_name not in wb.sheetnames:
        return jsonify({"error": "Sheet không tồn tại"}), 400
    ws = wb[sheet_name]

    # Nhãn tiêu đề cột (chỉ để hiển thị đẹp trên bảng tiến trình) - giả định
    # dòng tiêu đề nằm ngay phía trên dòng bắt đầu dữ liệu.
    header_probe_row = start_row - 1 if start_row > 1 else None
    name_header = (str(ws[f"{name_col}{header_probe_row}"].value).strip()
                   if name_col and header_probe_row and ws[f"{name_col}{header_probe_row}"].value else "Tên người nộp thuế")
    status_header = (str(ws[f"{status_col}{header_probe_row}"].value).strip()
                      if status_col and header_probe_row and ws[f"{status_col}{header_probe_row}"].value else "Trạng thái")

    # Quét trước (không gọi mạng) để biết tổng số dòng và đánh dấu dòng nào có
    # thể bỏ qua vì đã có đủ dữ liệu từ trước.
    rows = []
    row = start_row
    empty_streak = 0
    while row <= ws.max_row + 1 and empty_streak < 3:
        raw = ws[f"{mst_col}{row}"].value
        mst = clean_tax_code(raw)
        if mst == "":
            empty_streak += 1
            row += 1
            continue
        empty_streak = 0

        entry = {
            "row": row, "mst": mst, "name": None, "status_val": None,
            "result": "pending", "detail": None,
        }

        if skip_filled and target_cols:
            all_filled = True
            for c in target_cols:
                v = ws[f"{c}{row}"].value
                if v in (None, "") or v in (NOT_FOUND_TEXT, ERROR_TEXT):
                    all_filled = False
                    break
            if all_filled:
                entry["result"] = "skipped"
                if name_col:
                    entry["name"] = ws[f"{name_col}{row}"].value
                if status_col:
                    entry["status_val"] = ws[f"{status_col}{row}"].value

        rows.append(entry)
        row += 1

    total = len(rows)
    skipped_initial = sum(1 for r in rows if r["result"] == "skipped")

    job_id = uuid.uuid4().hex
    job = {
        "status": "running",
        "control": "run",
        "total": total,
        "ok": 0, "not_found": 0, "errors": 0, "skipped": skipped_initial,
        "processed": skipped_initial,
        "rows": rows,
        "mst_col": mst_col, "name_col": name_col, "address_col": address_col,
        "dept_col": dept_col, "status_col": status_col,
        "name_header": name_header, "status_header": status_header,
        "delay": delay,
        "result_id": None,
        "_wb": wb, "_ws": ws,
    }
    with _jobs_lock:
        _jobs[job_id] = job

    # Server tự chạy nền, tự gọi api.xinvoice.vn cho từng dòng còn "pending"
    # (xem process_job_serverside + ghi chú ở API_TEMPLATE) - trình duyệt chỉ
    # còn việc poll /api/progress để hiển thị bảng theo dõi.
    threading.Thread(target=process_job_serverside, args=(job_id,), daemon=True).start()

    return jsonify({
        "job_id": job_id,
        "total": total,
        "columns": {
            "mst_col": mst_col, "name_col": name_col, "status_col": status_col,
            "name_header": name_header, "status_header": status_header,
        },
    })


@app.route("/api/progress/<job_id>")
def api_progress(job_id):
    job = _jobs.get(job_id)
    if not job:
        return jsonify({"error": "Phiên xử lý không tồn tại"}), 404

    now = time.time()
    rows_out = []
    for r in job["rows"]:
        item = {
            "row": r["row"], "mst": r["mst"], "name": r.get("name"),
            "status_val": r.get("status_val"), "result": r["result"],
            "detail": r.get("detail"), "from_cache": bool(r.get("from_cache")),
        }
        wait_until = r.get("wait_until")
        if wait_until:
            item["wait_seconds"] = max(0, round(wait_until - now))
        rows_out.append(item)

    return jsonify({
        "status": job["status"],
        "total": job["total"],
        "processed": job["processed"],
        "ok": job["ok"], "not_found": job["not_found"],
        "errors": job["errors"], "skipped": job["skipped"],
        "rows": rows_out,
        "result_id": job.get("result_id"),
    })


@app.route("/api/pause/<job_id>", methods=["POST"])
def api_pause(job_id):
    job = _jobs.get(job_id)
    if not job:
        return jsonify({"error": "Phiên xử lý không tồn tại"}), 404
    job["control"] = "pause"
    job["status"] = "paused"
    return jsonify({"status": "paused"})


@app.route("/api/resume/<job_id>", methods=["POST"])
def api_resume(job_id):
    job = _jobs.get(job_id)
    if not job:
        return jsonify({"error": "Phiên xử lý không tồn tại"}), 404
    job["control"] = "run"
    job["status"] = "running"
    return jsonify({"status": "running"})


@app.route("/api/download/<result_id>")
def api_download(result_id):
    if result_id not in _results:
        return "File không tồn tại hoặc đã hết hạn", 404
    path = _results[result_id]
    return send_file(
        path,
        as_attachment=True,
        download_name="ket_qua_tra_cuu_mst.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _port_is_free(host, port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind((host, port))
        return True
    except OSError:
        return False
    finally:
        s.close()


def _open_browser_when_ready(url, host, port, timeout=10):
    """Chờ tới khi server thật sự nhận kết nối rồi mới mở trình duyệt, để
    tránh mở ra một trang 'không kết nối được' nếu server khởi động chậm."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection((host, port), timeout=0.5):
                break
        except OSError:
            time.sleep(0.2)
    try:
        webbrowser.open(url)
    except Exception:
        pass


if __name__ == "__main__":
    HOST = "127.0.0.1"
    PORT = 5000
    URL = f"http://{HOST}:{PORT}"

    print("=" * 64)
    print("  ĐANG KHỞI ĐỘNG ỨNG DỤNG TRA CỨU MST")
    print("=" * 64)

    if not _port_is_free(HOST, PORT):
        print()
        print(f"KHÔNG THỂ KHỞI ĐỘNG: cổng {PORT} trên máy này đang bị")
        print("chương trình khác sử dụng (có thể một cửa sổ chạy app này")
        print("từ trước vẫn còn mở, hoặc phần mềm khác đang chiếm cổng đó).")
        print()
        print("Cách khắc phục:")
        print(f"  1. Kiểm tra xem có cửa sổ nào khác đang chạy '{Path(__file__).name}' không, đóng nó lại rồi thử lại.")
        print("  2. Nếu vẫn lỗi, khởi động lại máy rồi thử lại.")
        print()
        input("Nhấn Enter để đóng cửa sổ này...")
        raise SystemExit(1)

    print(f"  Trình duyệt sẽ TỰ ĐỘNG mở tại: {URL}")
    print("  (nếu không tự mở được, bạn tự copy địa chỉ trên dán vào trình duyệt)")
    print()
    print("  !!! QUAN TRỌNG: ĐỪNG ĐÓNG cửa sổ đen này trong lúc dùng app !!!")
    print("  Đóng cửa sổ này = tắt ứng dụng ngay lập tức, trình duyệt sẽ báo")
    print("  lỗi 'không kết nối được' / ERR_CONNECTION_REFUSED.")
    print("  Dùng xong thì quay lại đây và nhấn Ctrl+C để tắt cho gọn.")
    print("=" * 64)

    threading.Thread(target=_open_browser_when_ready, args=(URL, HOST, PORT), daemon=True).start()

    try:
        app.run(host=HOST, port=PORT, debug=False, threaded=True)
    except OSError as e:
        print()
        print(f"LỖI KHI CHẠY SERVER: {e}")
        input("Nhấn Enter để đóng cửa sổ này...")
