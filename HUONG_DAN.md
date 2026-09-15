# Web App tra cứu mã số thuế hàng loạt

> **Chia sẻ cho nhiều người dùng (khuyên dùng)?** Đọc file
> `HUONG_DAN_DUA_LEN_MANG.md` — đưa app lên một link web thật, ai cũng mở
> được bằng trình duyệt, không ai cần cài Python hay đụng tới cửa sổ cmd nào
> cả. Cách chạy local dưới đây chỉ phù hợp khi bạn tự dùng một mình trên máy
> mình, vì cách này dễ gặp lỗi trên máy tính văn phòng bị khóa quyền cài đặt.

Đây là bản nâng cấp thành **web app chạy local**, có giao diện, tự nhận diện cột
theo tiêu đề file Excel — mỗi người dùng có thể có file với layout khác nhau
(MST ở cột khác, cột cần điền khác, dòng bắt đầu khác), app sẽ tự đoán và cho
phép chỉnh tay nếu đoán sai, trước khi chạy.

Sau khi bấm "Bắt đầu tra cứu", bạn sẽ thấy một bảng theo dõi tiến trình trực
tiếp: thanh tiến độ (%), 4 ô số liệu (Tìm thấy / Không tìm thấy / Lỗi kết nối /
Bỏ qua vì đã có dữ liệu), và một bảng liệt kê từng dòng đang được xử lý, cập
nhật theo thời gian thực — kể cả trạng thái "đang chờ do giới hạn tốc độ API"
khi cần. Có nút **Tạm dừng / Tiếp tục** để dừng giữa chừng và chạy tiếp sau.

Mặc định app sẽ **bỏ qua các dòng đã có đủ dữ liệu** (từ lần chạy trước), chỉ
tra lại những dòng còn trống hoặc trước đó bị "Không tìm thấy"/"Lỗi tra cứu" —
bạn có thể tắt tùy chọn này nếu muốn tra lại toàn bộ từ đầu.

## Vì sao không phải là "artifact" chạy trên Claude?

Artifact xuất bản trên Claude chỉ được phép gọi tới một số ít domain đã duyệt
sẵn (CDN, Google Fonts...). Domain `api.xinvoice.vn` không nằm trong danh sách
đó, nên **bất kỳ ai** mở một artifact như vậy — không riêng bạn — đều sẽ bị
trình duyệt chặn khi cố gọi API. Đây là giới hạn cố định của nền tảng, không
thể tắt/bật.

Giải pháp thật sự dùng được: một ứng dụng web nhỏ **chạy trên máy của người
dùng** (bạn, hoặc đồng nghiệp bạn chia sẻ file này cho). Nó vẫn là giao diện
web, chạy trong trình duyệt, chỉ khác là "server" đứng sau nó chạy ngay trên
máy đó — nên gọi API bình thường, không bị chặn.

## Cài đặt (một lần)

1. Cài **Python 3** nếu máy chưa có: https://www.python.org/downloads/
   (khi cài trên Windows, nhớ tick chọn "Add Python to PATH").
2. Giải nén toàn bộ thư mục này ra một chỗ cố định trên máy.

## Cách chạy

**Windows:** nhấp đúp vào file `chay_ung_dung.bat`.

**macOS / Linux:** mở Terminal tại thư mục này rồi chạy:
```bash
./chay_ung_dung.sh
```

Lần đầu chạy sẽ tự cài các thư viện cần thiết (Flask, openpyxl, requests),
có thể mất khoảng 30 giây - 1 phút.

**Trình duyệt sẽ TỰ ĐỘNG mở ra** khi server sẵn sàng — bạn không cần tự gõ
hay copy link gì cả. Nếu vì lý do gì đó trình duyệt không tự mở, tự mở
trình duyệt và vào `http://127.0.0.1:5000`.

## ⚠️ QUAN TRỌNG — lỗi hay gặp nhất: "This site can't be reached" / ERR_CONNECTION_REFUSED

Đây **luôn luôn** là vì cửa sổ đen (cmd/terminal) đang chạy app đã bị đóng
hoặc chưa từng chạy thành công. App này không phải một website thật trên
Internet — nó chỉ tồn tại khi cửa sổ đó đang mở. Cụ thể:

- **Cửa sổ chạy app phải luôn mở** trong suốt lúc bạn dùng. Đóng cửa sổ đó
  (kể cả vô tình bấm nhầm dấu X) = tắt app ngay lập tức = trình duyệt báo lỗi
  "refused to connect" dù bạn có reload bao nhiêu lần.
- Đây **không phải một cái link chia sẻ được** như link Google Drive hay
  Zalo. Địa chỉ `http://127.0.0.1:5000` chỉ có nghĩa trên đúng máy đang chạy
  app — copy gửi cho người khác, hoặc mở trên điện thoại/máy khác, sẽ luôn
  báo lỗi này vì trên máy đó không có app nào đang chạy ở địa chỉ đó cả. Mỗi
  người muốn dùng phải tự chạy `chay_ung_dung.bat`/`.sh` trên máy của họ.
- Nếu cửa sổ hiện lỗi đỏ kiểu "python is not recognized" / "không tìm thấy
  Python" → máy đó chưa cài Python 3 đúng cách (xem lại mục Cài đặt ở trên).
- **Nếu cửa sổ hiện đúng câu:** *"Python was not found; run without
  arguments to install from the Microsoft Store, or disable this shortcut
  from Settings > Apps > Advanced app settings > App execution aliases."*
  → Đây là lỗi rất hay gặp trên Windows: máy có sẵn một "python" **giả**
  (Windows gọi là App execution alias) — gõ lệnh `python` chỉ hiện thông báo
  này chứ không chạy Python thật, **kể cả khi máy đã cài Python thật rồi**.
  Cách sửa, làm đúng theo thứ tự:
  1. Mở **Settings → Apps → Advanced app settings → App execution aliases**.
  2. **Tắt (gạt OFF)** hai mục "App Installer python.exe" và
     "App Installer python3.exe".
  3. Nếu máy **chưa** cài Python thật, cài tại
     https://www.python.org/downloads/ — ở màn hình đầu tiên lúc cài, phải
     **tick chọn ô "Add python.exe to PATH"** trước khi bấm Install Now.
  4. Đóng hết các cửa sổ cmd đang mở (hoặc khởi động lại máy cho chắc),
     rồi chạy lại `chay_ung_dung.bat`.
- Nếu cửa sổ báo cổng 5000 đang bị dùng → có một cửa sổ app khác vẫn đang
  chạy ở đâu đó (kể cả từ hôm trước), đóng hết các cửa sổ cũ rồi chạy lại.

Để tắt ứng dụng đúng cách: quay lại cửa sổ đó và nhấn `Ctrl + C`, hoặc đóng
cửa sổ khi đã dùng xong (không phải trước hoặc trong khi dùng).

## Cách dùng trên giao diện web

1. **Kéo thả hoặc chọn file Excel (.xlsx)** cần tra cứu.
2. App sẽ tự động:
   - Đoán sheet, dòng tiêu đề, và cột nào là MST / Tên / Địa chỉ / Cơ quan
     thuế / Trạng thái, dựa vào chữ trong tiêu đề cột.
   - Hiển thị bảng xem trước 12 dòng đầu để bạn đối chiếu.
3. **Kiểm tra lại các dropdown** — nếu app đoán sai (ví dụ file không có cột
   "Trạng thái" hoặc đặt tên khác), bạn tự chọn lại đúng cột trong dropdown.
   Có thể chọn "-- không dùng --" cho cột nào bạn không cần điền.
4. Kiểm tra **"Dòng dữ liệu bắt đầu từ"** — mặc định app đoán dựa vào dòng đầu
   tiên có MST, nhưng bạn nên xác nhận lại.
5. Bấm **"Bắt đầu tra cứu"**. App sẽ gọi API cho từng dòng, có nghỉ giữa các
   lần gọi (mặc định 1 giây) để tránh bị giới hạn tần suất.
6. Khi xong, xem **bảng tổng kết** (số dòng thành công / không tìm thấy / lỗi)
   và **nhật ký từng dòng**, rồi bấm **"Tải file kết quả (.xlsx)"**.

File gốc bạn tải lên không bị thay đổi — file kết quả là một file mới.

## Chia sẻ cho đồng nghiệp

Gửi nguyên thư mục này (nén thành file .zip) cho người khác. Họ chỉ cần:
1. Cài Python 3 (nếu chưa có).
2. Giải nén và chạy `chay_ung_dung.bat` (Windows) hoặc `chay_ung_dung.sh` (Mac/Linux).
3. Mở trình duyệt vào `http://127.0.0.1:5000`, dùng như bình thường — kể cả
   nếu file Excel của họ có layout khác hoàn toàn (MST ở cột khác, thứ tự cột
   điền khác), app vẫn tự nhận diện hoặc để họ tự chọn tay.

Mỗi người chạy một bản độc lập trên máy mình — không ai cần "server chung",
không cần đăng ký tài khoản, không giới hạn số lượng người dùng.

## Lưu ý về dữ liệu MST

Nếu một mã số thuế không tra ra được, app sẽ ghi **"Không tìm thấy"**. Trước
khi kết luận API sai, hãy kiểm tra:
- MST có bị mất số 0 ở đầu không (thường do cột định dạng Number thay vì Text
  trong Excel trước đó).
- MST có đúng định dạng của Tổng cục Thuế không (10 số cho cá nhân/tổ chức,
  hoặc 10 số + `-XXX` cho đơn vị phụ thuộc).

## Giới hạn hiện tại

- Ứng dụng chỉ chạy khi bạn (hoặc người dùng) đang mở cửa sổ terminal/cmd đó
  — đóng cửa sổ là tắt server. Đây là app chạy local, không phải dịch vụ trên
  Internet.
- Không có xác thực đăng nhập — chỉ máy đang chạy app mới truy cập được
  (địa chỉ `127.0.0.1` = chỉ máy này), phù hợp dùng cá nhân/nội bộ.
- Muốn biến thành một dịch vụ web thật sự (một link cho nhiều người dùng cùng
  lúc, không cần cài gì), cần triển khai (deploy) `app.py` lên một nơi lưu trữ
  ứng dụng web (ví dụ Render, Railway, một VPS...). Nếu bạn muốn, tôi có thể
  hướng dẫn thêm bước đó.
