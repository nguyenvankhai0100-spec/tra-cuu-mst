#!/bin/bash
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
    echo "============================================================"
    echo "  LỖI: Không tìm thấy Python 3 trên máy này."
    echo "============================================================"
    echo "  Cài Python 3 trước: https://www.python.org/downloads/"
    echo "  Cài xong thì chạy lại file này."
    echo "============================================================"
    read -p "Nhấn Enter để đóng cửa sổ này..."
    exit 1
fi

echo "Đang cài đặt thư viện cần thiết (chỉ mất thời gian ở lần chạy đầu)..."
if ! python3 -m pip install -r requirements.txt; then
    echo
    echo "============================================================"
    echo "  LỖI: Cài đặt thư viện thất bại. Kiểm tra kết nối mạng rồi thử lại."
    echo "============================================================"
    read -p "Nhấn Enter để đóng cửa sổ này..."
    exit 1
fi

echo
echo "============================================================"
echo "  ĐỪNG ĐÓNG cửa sổ này khi đang dùng ứng dụng!"
echo "  Đóng cửa sổ này = tắt ứng dụng ngay lập tức."
echo "============================================================"
python3 app.py

echo
read -p "Ứng dụng đã dừng. Nhấn Enter để đóng cửa sổ này..."
