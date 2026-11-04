# Chạm tiếp thị thay modem — TTVT × VNPT

Website tĩnh (HTML/CSS/JS) đưa lên **GitHub Pages**, gọi **API Google Apps Script** để:

- Đăng nhập theo tài khoản nhân viên TTVT/VNPT; mỗi tài khoản chỉ thấy đúng MA_KV được phân.
- Đọc danh sách hộ trực tiếp từ Google Sheet.
- Ghi từng lượt chạm ngược lại sheet `CHAM_TIEP_THI` (tự điền người nhập/đơn vị).
- Ghi khách hẹn thay modem vào `HEN_THAY_MODEM`, có bộ lọc ngày và theo dõi đã thay/hủy.
- Cập nhật cột trạng thái trên `Sheet1` (R→AG).
- Thống kê, báo cáo và tải Excel ngay trên trình duyệt.

```
Người dùng ──► GitHub Pages (HTML/JS) ──► Apps Script API ──► Google Sheet
                     ▲                        │
                     └──── JSONP response ────┘
```

> Apps Script không cho đặt header CORS, nên frontend gọi API bằng **JSONP** (`?callback=...`).
> API vẫn là một endpoint JSON duy nhất và **hỗ trợ cả POST JSON** cho curl/Postman/ứng dụng khác.

---

## 1. Cấu trúc thư mục

| File | Vai trò |
|---|---|
| `index.html` | Trang chính |
| `assets/style.css` | Giao diện |
| `assets/app.js` | Toàn bộ logic (lọc, phiếu chạm, thống kê, xuất Excel) |
| `config.js` | Cấu hình API cố định dùng chung (bắt buộc điền `API_URL`) |
| `config.example.js` | Mẫu cấu hình |
| `.nojekyll` | Giữ nguyên file bắt đầu bằng `_` khi deploy |
| `404.html` | Chuyển về trang chính |

Repo không cần build, không cần `node_modules`, không dùng CDN.

---

## 2. Tài khoản & phân quyền

- 42 tài khoản từ file `Phân ô địa bàn VNPT TAH.xlsx`: 01 Giám đốc, 02 Phó Giám đốc, 01 HTKT, 05 Tổ trưởng, 33 Nhân viên.
- **Giám đốc, Phó Giám đốc, HTKT:** toàn quyền 32 khu vực.
- **Tổ trưởng, Nhân viên:** chỉ thấy hộ thuộc MA_KV được phân; server kiểm tra phạm vi khi đọc và khi ghi.
- Mật khẩu mặc định **123456**; đổi tại menu tài khoản. Quản trị đặt lại mật khẩu bằng menu trong Google Sheet.
- Một tài khoản VNPT có thể phụ trách nhiều khu vực theo file phân ô.

## 3. Chuẩn bị Google Sheet

1. Tải file `cham thay modem - TACH DIA CHI.xlsx` lên Google Drive.
2. Mở bằng Google Trang tính → **File → Lưu dưới dạng Google Trang tính**.
3. Giữ nguyên 17 cột gốc, không chèn/xóa/đổi thứ tự.

---

## 4. Triển khai API (Apps Script)

1. Trong Google Sheet: **Tiện ích mở rộng → Apps Script**.
2. Dán toàn bộ `Code.gs` (trong bộ `Cham_tiep_thi_modem_AppsScript.zip`).
3. (Tùy chọn) **Cài đặt dự án** → bật `appsscript.json` → dán nội dung manifest.
4. Chạy hàm `setupSheets` một lần → cấp quyền.
   - Tạo sheet `CHAM_TIEP_THI`.
   - Thêm 16 cột trạng thái (R→AG): 10 cột chạm + 6 cột hẹn thay.
5. **Triển khai → Triển khai mới → Ứng dụng web**:
   - Thực thi dưới dạng: **Tôi**
   - Ai có quyền truy cập: **Bất kỳ ai**
6. Sao chép URL kết thúc bằng **`/exec`**.

`setupSheets` tạo thêm:
- `TAI_KHOAN`: 42 tài khoản từ file phân ô, mật khẩu mặc định **123456**;
- `PHAN_O`: 32 khu vực TAH001–TAH032;
- `HEN_THAY_MODEM`: lịch sử hẹn thay;
- 6 cột hẹn thay trên `Sheet1` (sau 10 cột trạng thái chạm).

Kiểm tra nhanh bằng trình duyệt:

```
https://script.google.com/macros/s/AKfy…/exec?action=ping
```

Kết quả mong đợi:

```json
{"ok":true,"rows":3542,"sheetName":"Sheet1","historySheet":"CHAM_TIEP_THI",…}
```

> Nếu trình duyệt hiện trang đăng nhập Google thay vì JSON → web app chưa để quyền **Bất kỳ ai**.

---

## 5. Đưa lên GitHub Pages

### Cách A — tải trực tiếp trên web GitHub

1. Tạo repository mới (nên để **Public**; GitHub Pages cho repo Private tùy gói tài khoản).
2. Bấm **Add file → Upload files**, kéo **toàn bộ nội dung** thư mục này vào (không kéo cả thư mục cha):
   ```
   index.html
   config.js
   config.example.js
   404.html
   .nojekyll
   README.md
   assets/style.css
   assets/app.js
   ```
3. Commit.
4. **Settings → Pages** → Source: **Deploy from a branch** → Branch: `main` / `(root)` → **Save**.
5. Chờ 1–2 phút, mở `https://<tên-tài-khoản>.github.io/<tên-repo>/`.

### Cách B — dùng git

```bash
cd github_pages
git init
git add .
git commit -m "Cham tiep thi modem - TTVT VNPT"
git branch -M main
git remote add origin https://github.com/<tài-khoản>/<repo>.git
git push -u origin main
```

Sau đó bật Pages như bước 4.

### Cách C — đặt trong thư mục `docs/`

Nếu muốn giữ repo chung với mã Apps Script, đưa các file vào `docs/` rồi chọn
**Settings → Pages → Branch: main / docs**.

---

## 6. Cấu hình API cố định (chỉ một lần)

Mở `config.js` trong repo và điền URL `/exec`:

```js
window.APP_CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfy…/exec",
  ACCESS_CODE: ""   // để trống nếu Code.gs không đặt mã
};
```

Commit lại. **Tất cả nhân viên dùng chung một API**, không ai phải nhập URL hay mã trên giao diện; mỗi người chỉ đăng nhập tài khoản.

Nếu `API_URL` để trống, trang hiện thông báo **“Chưa cấu hình API”** và chỉ rõ cần sửa `config.js` — không còn hộp thoại nhập API.

> `config.js` là file công khai trên GitHub Pages. Vì vậy **không nên đặt `ACCESS_CODE`** ở đây; hãy dùng đăng nhập tài khoản làm lớp bảo vệ chính. Nếu bắt buộc đặt mã, dùng repo Private hoặc chấp nhận mã hiển thị công khai.

---

## 7. Lọc nhanh trên danh sách

Ngay dưới bộ lọc chi tiết có hai nhóm nút luôn hiển thị (kể cả trên điện thoại khi bộ lọc phụ đang thu gọn):

- **Thay modem:** `Chưa thay` · `Đã thay` · `Tất cả` — kèm số hộ từng nhóm (2.159 / 1.383 / 3.542).
- **Nhân viên kỹ thuật:** bấm để mở danh sách nhân viên, gõ tên hoặc mã để tìm, mỗi người có số hộ đang phụ trách theo bộ lọc hiện tại. Chọn một người để lọc ngay; chọn “Tất cả nhân viên” để bỏ lọc.

Hai bộ lọc này kết hợp được với Tổ, Xã, Ấp, Đường, nhu cầu dịch vụ và trạng thái chạm.

## 8. API reference

Base URL: `https://script.google.com/macros/s/<ID>/exec`

| Action | Tham số | Việc làm |
|---|---|---|
| `ping` | `code` | Kiểm tra kết nối, trả số hộ |
| `login` | `username`, `password` | Trả token phiên + hồ sơ tài khoản |
| `me` | `token` | Kiểm tra phiên đăng nhập |
| `changePassword` | `token`, `old_password`, `new_password` | Đổi mật khẩu |
| `logout` | `token` | Kết thúc phiên |
| `bootstrap` | `token` | Hộ + lịch sử chạm + hẹn thay **trong phạm vi MA_KV** |
| `saveTouch` | `token`, `ma`, `ngay`, `gio`, `ket_qua`, `tu_van[]`, `dong_y[]`, `ghi_chu`, `client_id` | Ghi 1 lượt chạm; tự điền người nhập/đơn vị |
| `deleteTouch` | `token`, `id` | Xóa 1 lượt chạm |
| `saveAppointment` | `token`, `ma`, `ngay_hen`, `gio_hen`, `dich_vu[]`, `ghi_chu` | Lưu/cập nhật hẹn thay |
| `updateAppointment` | `token`, `id`, `status`, `ngay_hen`… | Đã thay / Hủy / Dời ngày |

### GET (JSON hoặc JSONP)

```
# 1. Đăng nhập → token
/exec?action=login&payload={"username":"nmtien.hcm","password":"123456"}&callback=cb

# 2. Các lệnh cần token
/exec?action=bootstrap&token=<TOKEN>
/exec?action=bootstrap&token=<TOKEN>&callback=cb
/exec?action=saveTouch&token=<TOKEN>&payload=<JSON đã encode>&callback=cb
/exec?action=saveAppointment&token=<TOKEN>&payload=<JSON đã encode>&callback=cb
```

### POST JSON

```bash
curl -X POST "https://script.google.com/macros/s/<ID>/exec" \
  -H "Content-Type: text/plain;charset=utf-8" \
  -d '{"action":"login","username":"nmtien.hcm","password":"123456"}'
```

### Chống ghi trùng

Mỗi lần lưu, frontend sinh `client_id`. Nếu mạng chập chờn và request bị gửi lại,
server trả lại kết quả cũ thay vì tạo thêm dòng. Cache 10 phút.

### Bảo mật

- **Bắt buộc đổi mật khẩu mặc định `123456`** cho từng tài khoản.
- Phân quyền theo MA_KV được kiểm tra ở server; không chỉ ẩn trên giao diện.
- `ACCESS_CODE` trong `Code.gs` (mặc định rỗng) là lớp bảo vệ thêm nếu website công khai.
- Chọn **Bất kỳ ai** để nhân viên mở không cần đăng nhập Google; tài khoản ứng dụng vẫn là lớp chính.
- Không đưa URL `/exec` kèm mã/token lên nơi công khai.
- Token phiên hết hạn sau 6 giờ; đăng xuất sẽ hủy token.

---

## 9. Luồng dữ liệu khi lưu

1. Người dùng bấm **Lưu phiếu chạm**.
2. Frontend gửi `saveTouch` qua JSONP kèm token đăng nhập.
3. Apps Script:
   - Xác thực token, kiểm tra hộ thuộc MA_KV của tài khoản.
   - Tự điền người nhập/đơn vị từ tài khoản.
   - Ghi 1 dòng vào `CHAM_TIEP_THI`.
   - Tính lại 10 cột trạng thái chạm của hộ.
   - Trả về bản ghi vừa lưu.
4. Frontend cập nhật danh sách, KPI, tab TTVT/VNPT, báo cáo.

Với **Gọi hẹn thay**: lưu vào `HEN_THAY_MODEM`, cập nhật 6 cột hẹn thay trên `Sheet1`, không tạo phiếu chạm; khách chuyển sang tab **KH hẹn thay**.

Xóa lượt chạm cũng xóa đúng dòng lịch sử và tính lại trạng thái.

---

## 10. Cập nhật giao diện

- Sửa trực tiếp `index.html`, `assets/app.js`, `assets/style.css` trên GitHub → commit.
- GitHub Pages tự build lại sau ~1 phút.
- Nếu đổi API URL: sửa `API_URL` trong `config.js` → commit. Không cần sửa giao diện.

---

## 11. Xử lý sự cố

| Hiện tượng | Nguyên nhân / cách xử lý |
|---|---|
| Trang báo “Chưa cấu hình API” | Điền `API_URL` trong `config.js` rồi commit |
| `Không tải được API` | URL sai, web app chưa deploy, hoặc quyền không phải “Bất kỳ ai” |
| Trình duyệt hiện trang đăng nhập Google | Web app đang để “Chỉ mình tôi” hoặc giới hạn domain |
| `Mã truy cập không đúng` | Kiểm tra `ACCESS_CODE` trong `Code.gs` và `config.js` phải giống nhau |
| `Missing required parameter` | Dùng sai action; xem bảng API |
| 404 trên GitHub Pages | Chưa bật Pages, sai branch/thư mục, hoặc chưa có `index.html` ở root |
| Trang cũ vẫn hiện | Ctrl/Cmd + Shift + R (cache), hoặc chờ GitHub build |
| Lưu báo lỗi | Mất mạng, hoặc Sheet đang khóa — thử lại |
| Tải Excel bị chặn | Dùng Chrome/Edge, cho phép tải xuống |

---

## 12. Ghi chú kỹ thuật

- Không cần framework, không CDN, không build step.
- Dữ liệu 3.542 hộ + lịch sử được tải một lần khi mở trang; lọc/thống kê chạy phía trình duyệt.
- Excel `.xlsx` được tạo bằng JavaScript thuần (ZIP + OOXML), không cần server.
- Điện thoại: giao diện thẻ, bộ lọc thu gọn, phiếu chạm toàn màn hình.
- Múi giờ: `Asia/Ho_Chi_Minh`.
- API dùng `LockService` khi ghi để an toàn nhiều người nhập cùng lúc.

---

## 13. Liên kết

- Hướng dẫn đầy đủ về cấu trúc sheet, cột trạng thái, kích bản tiếp thị:
  `HUONG_DAN_TRIEN_KHAI.md` trong bộ Apps Script.
- Bản xem trước giao diện (dữ liệu mẫu, không ghi Sheet):
  `XEM_TRUOC_giao_dien.html`.
