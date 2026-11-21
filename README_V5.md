# Auto Click Recorder V5 — Hướng dẫn chạy theo khung giờ

## Mặc định
- Ngày bắt đầu: **2026-11-21**
- Ngày kết thúc: **2026-11-30** (bao gồm ngày cuối)
- Khung 1: **06:00–14:30**
- Khung 2: **14:30–23:30**
- Không bỏ qua thứ Bảy, Chủ nhật.
- **20 lượt tải dự kiến** nếu mỗi khung giờ chỉ tải một file và không dùng chế độ Excel.

## Cách thiết lập
1. Bật **V5: Chạy từng khung giờ, hết ngày mới tăng ngày**.
2. Nhấn **＋ Mẫu 4 ô ngày/giờ + tải** (chỉ dùng khi chưa có các bước tương ứng; nếu đã ghi kịch bản thì sửa các bước hiện có).
3. Sửa 4 bước click để đặt đúng tọa độ ô **giờ bắt đầu, ngày bắt đầu, giờ kết thúc, ngày kết thúc**. Không để tọa độ 0,0.
4. Bốn bước `set_datetime` đi kèm dùng lần lượt `{window_start_time}`, `{window_start_date}`, `{window_end_time}`, `{window_end_date}`. Có thể chỉnh `mode=text` nếu ứng dụng chỉ hỗ trợ nhập trực tiếp.
5. Chụp ảnh riêng của **nút tải danh sách**, **dấu hiệu danh sách đã sẵn sàng** và **nút tải file**. Mở sửa từng bước `click_image` / `wait_image`, nhập đường dẫn PNG chính xác và chỉnh timeout.
6. Lưu kịch bản JSON và thử một ngày trên dữ liệu không quan trọng trước khi chạy cả tháng.

## Thứ tự chạy
Với từng ngày: thực hiện **toàn bộ các bước** cho khung 06:00–14:30, chờ ảnh sẵn sàng rồi tải; sau đó thực hiện cho khung 14:30–23:30. Chỉ sau khi hoàn thành cả hai khung mới tăng ngày. Dừng sau ngày 30/11/2026. F8 dừng khẩn cấp.

## Lưu ý an toàn
- V5 không tự đoán vị trí các ô hoặc trạng thái đã tải xong. Người dùng phải ghi tọa độ và cung cấp ảnh mẫu từ chính phần mềm.
- Bước `wait_image` xác minh **ảnh xuất hiện**, không bảo đảm dữ liệu đã mới hoặc file đã lưu xong. Hãy chọn ảnh chỉ xuất hiện khi kết quả thực sự sẵn sàng, và bổ sung bước chờ nếu cần.
- Nếu phần mềm hiển thị cùng ảnh sẵn sàng từ lượt trước, cần ghi thêm thao tác làm mới/đợi dấu hiệu đang tải biến mất để tránh tải nhầm dữ liệu cũ.
- Nếu dùng Excel, mỗi dòng Excel sẽ chạy toàn bộ kịch bản **trong mỗi khung giờ**.
- Chế độ V4 không chọn V5 vẫn giữ cách chạy theo số vòng. V5 chạy đến hết ngày kết thúc, không giới hạn bởi ô số vòng V4.
- Hẹn giờ cần giữ ứng dụng mở. Nhấn F8 để dừng.
- EXE được build qua workflow `.github/workflows/Build_exe.yml` trên Windows và upload từ `dist/Autoclick.exe`.
