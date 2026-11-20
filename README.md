# Auto Click Recorder 2.0 – Windows

Ứng dụng Windows để ghi thao tác chuột/bàn phím, sửa từng bước, nhận diện nút bằng hình ảnh, lập lịch chạy và nhập dữ liệu Excel.

## Sử dụng

1. Nhấn **Ghi**, thực hiện thao tác, nhấn **Dừng ghi** hoặc **F8**.
2. Chọn một dòng trong bảng rồi nhấn **Sửa bước** hoặc nhấp đúp; chỉnh các trường và **Lưu thay đổi**. Dùng ↑ ↓ để đổi thứ tự.
3. Thêm bước `wait_image` để chờ nút hiển thị, hoặc `click_image` để tìm và bấm nút dựa vào PNG đã chụp.
4. Chọn **Nạp Excel** để mở `.xlsx`/`.xlsm`. Dòng đầu là tên cột, mỗi dòng còn lại là một bộ dữ liệu. Bật **Chạy từng dòng Excel**: ứng dụng thực thi *toàn bộ kịch bản cho từng dòng*, sau đó chuyển dòng tiếp theo. Dùng bước `excel_text` với `column` bằng đúng tên cột (ví dụ `Mã SAP`). Hoặc trong bước `text` dùng `{Mã SAP}`, `{Số lượng}`.
5. Thêm bước **Chọn ngày/giờ** (`set_datetime`) và chỉnh `value` thành `{today}`, `{now}`, `{date:%d/%m/%Y}`, hoặc chuỗi cố định `20/11/2026`. Trường `format` điều khiển cách hiển thị khi `value={today}`. Các bước nhập văn bản gửi vào ô đang được chọn.
6. Điền **Hẹn giờ** theo `YYYY-MM-DD HH:MM:SS`, nhấn **Hẹn chạy**. Giữ app đang mở và Windows chưa khóa màn hình; hẹn giờ chỉ hoạt động khi app còn chạy, không phải Windows Task Scheduler.
7. Số vòng `0` = chạy liên tục. Nếu có Excel, mỗi vòng duyệt lại tất cả dòng Excel. F8 để dừng; di chuyển chuột về góc trái trên cũng là fail-safe.
8. **Lưu** / **Mở** kịch bản JSON. File Excel phải nạp lại sau khi mở kịch bản.

## Giới hạn quan trọng

- Chỉ dùng cho thao tác được cho phép. Không ghi mật khẩu, mã OTP, thông tin nhạy cảm.
- Ứng dụng ghi/phát lại thao tác, không phải AI tự suy luận giao diện. Nhận diện hình ảnh dựa vào ảnh mẫu và độ tương đồng.
- Các ứng dụng chạy quyền Administrator có thể cần chạy công cụ cùng quyền; Windows UAC/secure desktop không được hỗ trợ.
- Ngày giờ trong ứng dụng chỉ *nhập vào ô hiện đang có focus*, không tự điều khiển lịch popup tùy biến; với lịch popup nên ghi các thao tác chọn ngày hoặc dùng ảnh mẫu.
- `pyautogui.write` có thể không nhập đúng ký tự Unicode tiếng Việt trên một số ứng dụng. Nên thử trên môi trường an toàn trước.
- Đặt hiển thị Windows cố định (độ phân giải/tỉ lệ), vì bước click theo tọa độ nhạy với thay đổi vị trí.
- Không chạy kịch bản khi máy khóa, ngủ hoặc cửa sổ mục tiêu không khả dụng.
- Nếu ô Excel là số, Excel có thể đã làm mất số 0 đầu; định dạng cột mã thành Text trước khi nhập.
- Khi file Excel có công thức, chỉ đọc giá trị đã lưu trong file (`data_only=True`).

## Build EXE trên GitHub Actions

Upload cả thư mục vào repo GitHub, vào **Actions → Build Windows EXE → Run workflow**. Sau khi chạy thành công, vào **Artifacts → AutoClickRecorder-Windows**, tải và giải nén `AutoClickRecorder.exe`. Máy đích không cần cài Python. Chưa xác thực build trên Windows thực tế.

## Phát triển local

```bash
pip install -r requirements.txt
python app.py
```


## V3 - Ngày kịch bản tự tăng
- Đặt **Ngày bắt đầu** theo YYYY-MM-DD.
- Thêm bước **Chọn ngày/giờ** ở bất kỳ vị trí nào, nhập `{run_date}` và chọn `format` như `%d/%m/%Y`, `%m/%d/%Y` hoặc `%Y-%m-%d`.
- Bật **Tăng ngày sau mỗi vòng hoàn tất**, chọn số ngày tăng (mặc định 1), tùy chọn bỏ qua thứ 7 và chủ nhật.
- Ngày chỉ tăng sau khi **toàn bộ các bước và mọi dòng Excel trong một vòng** hoàn tất; nếu lỗi hoặc nhấn F8 thì không tăng tiếp.
- Có thể dùng `{run_date:%Y-%m-%d}` trong bước văn bản hoặc `{run_time}` để kết hợp ngày kịch bản với giờ thực tế.
- `0` vòng = chạy liên tục; cần giới hạn số vòng để tránh chạy không ngừng.
- Chọn ngày qua popup riêng của ứng dụng bên thứ ba vẫn cần ghi thao tác hoặc nhận diện hình ảnh, không tự động hiểu mọi lịch popup.
