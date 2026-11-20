# Auto Click Recorder V4 — Windows

## Cấu hình mặc định

- Ngày bắt đầu: **2026-11-20**.
- Số vòng: **5** (20, 21, 22, 23, 24/11/2026).
- Tự tăng **1 ngày sau khi hoàn tất toàn bộ các bước và các dòng Excel trong một vòng**.
- Không bỏ qua cuối tuần.
- Bước `set_datetime` có thể đặt ở vị trí bất kỳ trong kịch bản, mặc định `mode=auto`, `value={run_date}`, `format=%d/%m/%Y`.

## Chọn ngày thông minh

`mode=auto`: nhập trực tiếp trước. Nếu có `verify_image`, ứng dụng kiểm tra ảnh trạng thái sau nhập; nếu không thấy, chuyển sang chọn lịch qua `calendar_image` và `day_image`. **Nếu không có verify_image, ứng dụng không thể tự biết việc nhập ngày có thành công hay không và sẽ không tự fallback**. `mode=calendar` luôn chọn bằng ảnh lịch. Các ảnh phải phù hợp giao diện ứng dụng và **ngày thực tế**; một ảnh ngày cố định không thể đại diện cho mọi ngày. Nếu không thấy ảnh sẽ dừng có báo lỗi, không đoán tọa độ.

Cách chỉnh bước: chọn bước → nhấp đúp → sửa `kind`, `mode`, `format`, `value`, ảnh nhận diện, thời gian chờ → lưu. `select_all=true` chọn toàn bộ nội dung ô trước khi nhập; cần click vào ô ngày ở bước trước.

## Excel, ghi thao tác, hẹn giờ

- Nạp Excel `.xlsx`, chọn chạy từng dòng, thêm bước `excel_text` và nhập đúng tên cột Excel.
- Ghi click/phím, sửa từng bước, lưu/mở kịch bản JSON, hẹn giờ thực thi, dừng F8 hoặc rê chuột lên góc trên trái.
- Kịch bản V3 vẫn mở được; nếu có cấu hình V3 sẽ giữ các giá trị đó.
- Hẹn giờ cần ứng dụng mở, máy đang hoạt động và phiên Windows chưa khóa.

## Build EXE bằng GitHub Actions

Đưa **nội dung thư mục `AutoClickRecorder/`** lên thư mục gốc repository, vào Actions → Build Windows EXE → Run workflow → Artifacts → `AutoClickRecorder-Windows`. Workflow dùng Windows và PyInstaller; máy sử dụng EXE không cần cài Python. Chưa kiểm chứng build thực tế trên Windows.

## Giới hạn và an toàn

- OCR/nhận diện ảnh chỉ nhận biết giao diện, không hiểu mọi kiểu lịch. Lịch tháng/năm khác nhau cần ảnh mẫu hoặc tích hợp theo ứng dụng cụ thể.
- Không ghi mật khẩu; kiểm thử với dữ liệu giả trước khi bật lặp tự động. Phím F8 dừng ở lần kiểm tra tiếp theo, không đảo ngược thao tác đã thực hiện.
- Gõ bằng `pyautogui.write` có thể không nhập được tiếng Việt Unicode; nên kiểm thử dữ liệu Excel thực tế.
