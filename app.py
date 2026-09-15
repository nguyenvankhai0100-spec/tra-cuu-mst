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
from urllib.parse import quote

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

API_TEMPLATE = "https://api.xinvoice.vn/gdt-api/tax-payer-records/{tax_code}"
NOT_FOUND_TEXT = "Không tìm thấy"
ERROR_TEXT = "Lỗi tra cứu"
# Một số máy chủ (đặc biệt là các IP của nhà cung cấp hosting/cloud như
# Render, AWS, GCP...) bị api.xinvoice.vn (hoặc lớp tường lửa phía trước nó)
# chặn nếu request trông "giống bot" - ví dụ User-Agent mặc định của thư viện
# requests là "python-requests/x.y.z". Giả lập header của một trình duyệt
# thật giúp giảm khả năng bị chặn kiểu này.
API_REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://api.xinvoice.vn/",
    "Origin": "https://api.xinvoice.vn",
}

MAX_PREVIEW_ROWS = 12
MAX_HEADER_SCAN_ROWS = 10

# In-memory store: upload_id -> {"path": Path, "sheet": str}
#                  result_id -> Path
#                  job_id -> job state dict (xem process_job)
_uploads = {}
_results = {}
_jobs = {}
_jobs_lock = threading.Lock()

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
    max_col = min(ws.max_column, 30)
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
    max_col = min(ws.max_column, 40)

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
    max_col = min(ws.max_column, 12)
    max_row = min(ws.max_row, MAX_PREVIEW_ROWS)
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


def lookup_tax_code(tax_code, session, on_wait=None, timeout=15, retries=2, backoff=1.5):
    """
    Gọi API cho một mã số thuế.
    on_wait(wait_seconds, reason): callback được gọi NGAY TRƯỚC khi sleep để chờ
    thử lại - dùng để báo cáo trạng thái "đang chờ" ra ngoài (ví dụ cho UI biết
    dòng này đang chờ do bị giới hạn tốc độ).
    """
    url = API_TEMPLATE.format(tax_code=quote(tax_code, safe="-"))
    last_err = None
    for attempt in range(retries + 1):
        try:
            resp = session.get(url, timeout=timeout)
        except requests.RequestException as e:
            last_err = f"lỗi kết nối: {e}"
            wait_s = backoff * (attempt + 1)
            if on_wait:
                on_wait(wait_s, "Mất kết nối, đang thử lại...")
            time.sleep(wait_s)
            continue

        if resp.status_code == 200:
            try:
                payload = resp.json()
            except ValueError:
                last_err = "phản hồi không phải JSON hợp lệ"
                wait_s = backoff * (attempt + 1)
                if on_wait:
                    on_wait(wait_s, "Phản hồi lỗi, đang thử lại...")
                time.sleep(wait_s)
                continue
            if payload.get("success") and payload.get("data"):
                return payload["data"][0], None
            return None, "not_found"

        if resp.status_code == 404:
            return None, "not_found"

        last_err = f"HTTP {resp.status_code}"
        # Ghi kèm một đoạn ngắn nội dung phản hồi để biết lý do bị chặn thật sự
        # (ví dụ trang chặn của tường lửa/Cloudflare) thay vì chỉ có mã lỗi.
        try:
            body_snippet = resp.text.strip().replace("\n", " ")[:160]
        except Exception:
            body_snippet = ""
        if body_snippet:
            last_err = f"{last_err}: {body_snippet}"
        if 400 <= resp.status_code < 500 and resp.status_code != 429:
            break
        if resp.status_code == 429:
            retry_after = resp.headers.get("Retry-After")
            wait_s = float(retry_after) if retry_after and retry_after.isdigit() else backoff * (attempt + 2) * 2
            reason = "Chờ do giới hạn tốc độ..."
        else:
            wait_s = backoff * (attempt + 1)
            reason = "Đang thử lại..."
        if on_wait:
            on_wait(wait_s, reason)
        time.sleep(wait_s)

    return None, last_err or "lỗi không xác định"


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


def process_job(job_id):
    """Chạy trong thread nền: lần lượt tra cứu các dòng 'pending' và cập nhật
    trạng thái job để /api/progress đọc được theo thời gian thực."""
    job = _jobs[job_id]
    wb = job["_wb"]
    ws = job["_ws"]
    session = requests.Session()
    session.headers.update(API_REQUEST_HEADERS)

    mst_col = job["mst_col"]
    name_col = job["name_col"]
    address_col = job["address_col"]
    dept_col = job["dept_col"]
    status_col = job["status_col"]
    delay = job["delay"]

    for row_entry in job["rows"]:
        if row_entry["result"] != "pending":
            continue

        # Tạm dừng: chờ tới khi được resume hoặc bị dừng hẳn
        while job["control"] == "pause":
            job["status"] = "paused"
            time.sleep(0.3)
        if job["control"] == "stop":
            break
        job["status"] = "running"

        row = row_entry["row"]
        mst = row_entry["mst"]

        def on_wait(wait_s, reason, _entry=row_entry):
            _entry["result"] = "waiting"
            _entry["wait_until"] = time.time() + wait_s
            _entry["detail"] = reason

        data, err = lookup_tax_code(mst, session, on_wait=on_wait)
        row_entry.pop("wait_until", None)

        if data:
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
        elif err == "not_found":
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
            row_entry["detail"] = err
            if name_col:
                row_entry["name"] = ERROR_TEXT
            if status_col:
                row_entry["status_val"] = ERROR_TEXT
            job["errors"] += 1

        job["processed"] += 1

        if job["control"] == "stop":
            break
        if delay > 0:
            time.sleep(delay)

    if job["control"] == "stop":
        job["status"] = "stopped"
    else:
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
        delay = float(data.get("delay", 1.0))
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

    thread = threading.Thread(target=process_job, args=(job_id,), daemon=True)
    thread.start()

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
            "detail": r.get("detail"),
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
