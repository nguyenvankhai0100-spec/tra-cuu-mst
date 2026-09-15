# Đưa app lên Internet thật (một link cho tất cả mọi người, không cần cài gì)

Làm theo đúng thứ tự bên dưới, mất khoảng 15-20 phút cho lần đầu. Sau khi xong,
bạn có một link dạng `https://vntax-tra-cuu-mst.onrender.com` — gửi link đó +
mật khẩu cho đồng nghiệp là họ dùng được ngay trên trình duyệt, không cần cài
Python, không có cửa sổ đen nào cả.

Bạn chỉ cần làm phần này **một lần**. Từ lần sau, mỗi khi tôi sửa code, tôi sẽ
gửi bạn file mới, bạn chỉ cần upload đè lên GitHub là Render tự cập nhật.

---

## Bước 1 — Đưa code lên GitHub (không cần biết lệnh git)

1. Đăng nhập https://github.com.
2. Bấm dấu **+** ở góc trên bên phải → **New repository**.
3. Đặt tên, ví dụ `vntax-tra-cuu-mst`. Chọn **Private** (chỉ mình bạn/đồng
   nghiệp thấy code) hoặc Public đều được — không ảnh hưởng người dùng cuối.
   Không cần tick "Add a README file". Bấm **Create repository**.
4. Ở trang repo vừa tạo, bấm dòng chữ **"uploading an existing file"**
   (nằm trong đoạn hướng dẫn "...or uploading an existing file").
5. Kéo **toàn bộ nội dung bên trong thư mục `tax_web_app`** (không kéo cả thư
   mục `tax_web_app`, mà kéo các file/thư mục *bên trong* nó: `app.py`,
   `requirements.txt`, `Procfile`, `render.yaml`, thư mục `templates`, các
   file `.bat`/`.sh`, `HUONG_DAN.md`...) thả vào khung upload của GitHub.
   Trình duyệt Chrome hỗ trợ kéo thả cả thư mục `templates` và giữ đúng cấu
   trúc — nếu dùng trình duyệt khác không kéo được thư mục, vào thư mục
   `templates` kéo từng file `index.html`, `login.html` vào một thư mục con
   `templates` mà GitHub tự tạo khi bạn gõ `templates/index.html` vào ô tên
   file lúc kéo thả.
6. Cuộn xuống cuối trang, bấm **Commit changes**.
7. Kiểm tra lại: mở repo, phải thấy đủ các file `app.py`, `requirements.txt`,
   `Procfile`, và một thư mục `templates` chứa `index.html` + `login.html`.
   Thiếu thư mục `templates` là nguyên nhân phổ biến nhất khiến bước sau lỗi.

## Bước 2 — Tạo tài khoản Render và kết nối GitHub

1. Vào https://render.com, bấm **Get Started** (hoặc Sign Up), chọn
   **Sign up with GitHub** cho nhanh (Render sẽ xin quyền truy cập GitHub
   của bạn — đồng ý).
2. Sau khi vào Dashboard, bấm **New +** → **Web Service**.
3. Chọn repo `vntax-tra-cuu-mst` bạn vừa tạo ở Bước 1 (nếu chưa thấy, bấm
   "Configure account" để cấp quyền cho Render thấy repo đó).

## Bước 3 — Cấu hình Web Service

Render sẽ tự nhận đây là app Python. Điền/kiểm tra các mục sau:

- **Name**: đặt tên tùy ý, ví dụ `vntax-tra-cuu-mst` (tên này quyết định
  link cuối cùng: `https://<name>.onrender.com`).
- **Region**: chọn Singapore nếu có (gần Việt Nam nhất, tải nhanh hơn).
- **Build Command**: `pip install -r requirements.txt`
- **Start Command**: `gunicorn --worker-class gthread --workers 1 --threads 8 --timeout 120 -b 0.0.0.0:$PORT app:app`
- **Instance Type**: chọn **Free**.

Kéo xuống mục **Environment Variables**, bấm **Add Environment Variable**,
thêm 2 dòng:

| Key | Value |
|---|---|
| `APP_PASSWORD` | mật khẩu bạn muốn đặt cho cả nhóm dùng chung, ví dụ `Vntax2026!` |
| `SECRET_KEY` | một chuỗi ngẫu nhiên bất kỳ, càng dài càng tốt, ví dụ `f3a9c1e7b2d4...` (gõ bừa tay cũng được, miễn dài ~32 ký tự trở lên) |

Bấm **Deploy Web Service**.

## Bước 4 — Chờ deploy xong

Render sẽ hiện log build trực tiếp trên màn hình (khoảng 2-4 phút cho lần
đầu). Khi thấy dòng kiểu `Your service is live 🎉` là xong. Link app nằm ở
đầu trang, dạng `https://vntax-tra-cuu-mst.onrender.com`.

Mở thử link đó — sẽ hiện màn hình đăng nhập, gõ đúng `APP_PASSWORD` bạn đặt
ở Bước 3 để vào dùng.

## Bước 5 — Gửi cho đồng nghiệp

Gửi họ 2 thứ: **link** và **mật khẩu**. Vậy là xong — họ mở trình duyệt bất
kỳ (điện thoại cũng được), vào link, nhập mật khẩu, dùng luôn.

---

## Lưu ý quan trọng về gói miễn phí của Render

- **App sẽ "ngủ" sau 15 phút không ai dùng.** Lần mở tiếp theo sau khi ngủ
  sẽ mất khoảng 30-60 giây để "thức dậy" (Render tự hiện trang loading trong
  lúc chờ) — không phải app bị lỗi, cứ đợi chút rồi trang sẽ lên. Nếu công ty
  bạn dùng thường xuyên và không muốn chờ, Render có gói trả phí giữ app luôn
  "thức" — nói tôi nếu bạn muốn tìm hiểu thêm.
- Nếu Render báo hạn chế do gọi API bên ngoài quá nhiều trong thời gian ngắn,
  đó là do gói free giới hạn lưu lượng — với quy mô vài chục/vài trăm MST mỗi
  lần tra thì không đáng lo.

## Cập nhật code sau này

Mỗi khi tôi gửi bạn bản code mới:
1. Vào lại repo trên GitHub, bấm **Add file → Upload files**, kéo thả các
   file mới đè lên (GitHub tự hỏi có muốn thay thế file trùng tên không, bấm
   đồng ý), **Commit changes**.
2. Render tự động phát hiện thay đổi và deploy lại sau vài phút — không cần
   làm gì thêm bên Render.
