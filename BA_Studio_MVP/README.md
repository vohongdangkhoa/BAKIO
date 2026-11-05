# BA Requirements Studio – Human-in-the-loop (phiên bản 2.0)

Công cụ web chạy cục bộ cho nhân viên BA thu thập yêu cầu, chuẩn hóa và tạo hồ sơ Word cho SA, UX/UI và QA. **Không cần database, không tích hợp AI.** Nhân viên chịu trách nhiệm xác minh.

## Chạy
1. Mở thư mục trong VS Code.
2. Chạy `py -m pip install -r requirements.txt` (hoặc `python -m pip install -r requirements.txt`).
3. Chạy `py server.py` (hoặc `python server.py`).
4. Mở `http://localhost:8000`.
5. Bấm **Nạp mẫu Kio** hoặc nhập hồ sơ mới; xem ví dụ ngay dưới mỗi trường.
6. Lưu JSON để chỉnh sửa sau, xuất Word để trình ký và bàn giao.

## Phân hệ thu thập
- Tổng quan và phạm vi
- Stakeholder, Actor, User Journey và Business Rules
- Yêu cầu chức năng và dữ liệu
- Danh sách màn hình và yêu cầu UX/UI
- Khảo sát, AS-IS, TO-BE, ngoại lệ
- User Story, Use Case, tiêu chí chấp nhận
- Phối hợp Dev/Tester, SIT, UAT, đào tạo và nhật ký thay đổi
- Ràng buộc, câu hỏi mở, xác nhận

## Lưu trữ và bảo mật
Dữ liệu được lưu bằng localStorage trên trình duyệt hiện tại; khi chạy local, server chỉ nhận dữ liệu để tạo Word và không ghi xuống database. Nên định kỳ tải JSON dự phòng; không nhập dữ liệu khách hàng nhạy cảm nếu máy tính không được bảo vệ.

## Trạng thái tài liệu
Không đánh dấu xác nhận: Word ghi **BẢN NHÁP**. Tích xác nhận: Word ghi **BA/PO đã xác nhận trên ứng dụng** (chỉ là tự khai báo, không phải chữ ký số).

## Giới hạn
Không có đăng nhập, cộng tác đa người, API AI, phê duyệt điện tử hoặc vẽ UML tự động. Hồ sơ được tạo dựa trên thông tin BA nhập; các quyết định còn thiếu được đánh dấu để rà soát.
