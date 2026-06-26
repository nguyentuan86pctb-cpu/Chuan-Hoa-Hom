Tên ý tưởng:
Web App Chuẩn Hóa Mã Hòm Công Tơ Hiện Trường

Mục tiêu dự án:
Xây dựng một web app/PWA dùng trên điện thoại để cán bộ hiện trường chuẩn hóa mã hòm công tơ và mã công tơ theo quy tắc ngành, giảm tối đa nhập tay, tự sinh mã theo quy tắc Z, lưu dữ liệu an toàn trong quá trình làm việc và xuất file Excel gửi về điện lực.

Người dùng mục tiêu:
Cán bộ hiện trường đi kiểm tra/cập nhật dữ liệu công tơ, không yêu cầu biết kỹ thuật, thao tác chủ yếu trên điện thoại Android.

Vấn đề cần giải quyết:
Hiện nay việc chuẩn hóa mã hòm công tơ/công tơ dễ bị nhập tay nhiều, sai quy tắc, trùng mã, mất thời gian, khó kiểm soát tiến độ. Dữ liệu đầu vào là file Excel có danh sách số cột và khách hàng trên từng cột. Cán bộ hiện trường cần quan sát thực tế khách hàng nằm trong hòm/tủ nào rồi chuẩn hóa mã theo quy tắc.

Giải pháp đã chọn:
Làm web app MVP dạng PWA chạy trên điện thoại, có thể thêm vào màn hình chính như app Android. App import Excel, cho cán bộ chọn cột, chọn cấu trúc hòm/tủ bằng giao diện lưới Z-Rule, tự sinh mã, tự kiểm tra lỗi, tự lưu liên tục trên máy và xuất Excel kết quả gửi về điện lực. MVP không bắt buộc backend/server lưu dữ liệu khách hàng; dữ liệu xử lý trên thiết bị và lưu bằng IndexedDB.

Phạm vi MVP:
1. Import file Excel đầu vào.
2. Map các cột dữ liệu: số cột, mã khách hàng, tên khách hàng, loại khách hàng, lộ hạ thế nếu có.
3. Hiển thị danh sách cột và số lượng khách hàng trên từng cột.
4. Cho cán bộ chọn từng cột để chuẩn hóa.
5. Hỗ trợ 2 loại vị trí:
   - Hòm công tơ tại cột, mã 7 ký tự.
   - Tủ điện thanh cái hỗn hợp, mã 8 ký tự.
6. Giao diện lưới Z-Rule để gán khách hàng vào vị trí công tơ.
7. Tự gợi ý chia hòm theo số khách hàng, ví dụ 8 khách = 2 hòm H4.
8. Tự tăng STT hòm trên cùng cột.
9. Tự tăng STT khách hàng trong hòm/tủ theo chiều trái sang phải, trên xuống dưới.
10. Tự sinh mã chuẩn hóa.
11. Kiểm tra lỗi trước khi xuất.
12. Autosave từng thao tác.
13. Lưu nháp offline bằng IndexedDB.
14. Backup định kỳ.
15. Cho khôi phục phiên làm việc khi mở lại app.
16. Cho xuất bản nháp bất kỳ lúc nào.
17. Cho xuất file Excel hoàn thành gửi về điện lực.

Quy tắc mã hóa hòm công tơ tại cột:
Định dạng 7 ký tự:
[Pha][Lộ hạ thế 2 số][STT hòm 2 số][Loại hòm 1 số][STT khách hàng 1 số]

Ký tự 1:
A, B, C cho khách hàng 1 pha; D cho khách hàng 3 pha.

Ký tự 2,3:
Lộ hạ thế, định dạng 2 chữ số: 01, 02, 03...

Ký tự 4,5:
STT hòm trên cùng một cột, định dạng 2 chữ số: 01, 02, 03...

Ký tự 6:
Loại hòm:
H1 = 1
H2 = 2
H4 = 4
H6 = 6
Hòm 3 pha = 3

Ký tự 7:
STT khách hàng trong hòm: 1, 2, 3...

Ví dụ:
A010141
A = pha A
01 = lộ 01
01 = hòm thứ 01
4 = hòm H4
1 = khách hàng thứ 1 trong hòm

Quy tắc mã hóa công tơ tại tủ điện thanh cái hỗn hợp:
Định dạng 8 ký tự:
[Pha][Lộ hạ thế 2 số][Số mặt tủ 1 số][Loại tủ 2 số][STT khách hàng 2 số]

Ký tự 1:
A, B, C cho khách hàng 1 pha; D cho khách hàng 3 pha.

Ký tự 2,3:
Lộ hạ thế, định dạng 2 chữ số: 01, 02, 03...

Ký tự 4:
Số mặt tủ: 1 hoặc 2.

Ký tự 5,6:
Loại tủ theo số lượng công tơ thiết kế:
06, 08, 10, 12, 16, 20...

Ký tự 7,8:
STT khách hàng trong tủ, định dạng 2 chữ số: 01, 02, 03...

Ví dụ:
A0111201
A = pha A
01 = lộ 01
1 = tủ 1 mặt
12 = tủ 12 công tơ
01 = khách hàng thứ 01 trong tủ

Quy tắc Z-Rule:
Tự động đánh số theo chiều từ trái sang phải, từ trên xuống dưới.
Áp dụng cho:
1. STT hòm trên cùng một cột.
2. STT khách hàng trong cùng một hòm.
3. STT khách hàng trong cùng một tủ.

Ví dụ lưới H4:
[1] [2]
[3] [4]

Ví dụ tủ 12:
[01] [02] [03] [04]
[05] [06] [07] [08]
[09] [10] [11] [12]

Tính năng nội suy/gợi ý:
1. Dựa vào số khách hàng trên cột để gợi ý cách chia hòm.
Ví dụ:
8 khách: gợi ý 2 hòm H4, hoặc 1 H6 + 1 H2.
6 khách: gợi ý 1 H6, hoặc 1 H4 + 1 H2.
4 khách: gợi ý 1 H4.
2 khách: gợi ý 1 H2.
2. Sau khi chọn cấu trúc hòm/tủ, app có nút “Tự gán khách theo thứ tự Excel”.
3. Nếu thực tế khác, cán bộ có thể đổi vị trí khách hàng bằng thao tác chọn/kéo-thả.
4. Khi gán khách đầu tiên, các vị trí tiếp theo tự sinh mã theo thứ tự.

Yêu cầu giao diện:
Giao diện phải tối ưu cho điện thoại, nút lớn, ít nhập tay, ưu tiên chọn bằng dropdown/nút bấm.
Các màn hình chính:
1. Màn hình bắt đầu:
   - Nhập tên cán bộ.
   - Nhập đơn vị/đội.
   - Hiển thị/cho sửa tên trạm đang làm để tránh làm lẫn trạm.
   - Tải file Excel.
   - Tiếp tục phiên đang làm nếu có.

2. Màn hình danh sách cột:
   - Hiển thị số cột.
   - Số khách hàng trên cột.
   - Trạng thái: chưa làm, đang làm, có lỗi, đã xong.
   - Có tìm kiếm số cột.

3. Màn hình chuẩn hóa một cột:
   - Hiển thị số cột và tổng khách.
   - Chọn lộ hạ thế.
   - Chọn pha.
   - Thêm hòm hoặc tủ.
   - Gợi ý chia hòm.
   - Giao diện lưới Z-Rule.
   - Nút tự gán khách.
   - Nút kiểm tra lỗi.
   - Nút xuất bản nháp.

4. Màn hình kiểm tra lỗi:
   - Tổng khách.
   - Đã gán.
   - Chưa gán.
   - Thiếu mã/chưa gán.
   - Khách bị gán nhiều vị trí nếu có.

5. Màn hình xuất file:
   - Xuất bản nháp.
   - Xuất bản nháp và chia sẻ Zalo, tên file ghi rõ trạm.
   - Xuất file hoàn thành.
   - Xuất file hoàn thành và chia sẻ Zalo, tên file ghi rõ trạm.
   - Tải file backup khôi phục.

Yêu cầu chống mất dữ liệu:
1. Autosave từng thao tác.
2. Lưu dữ liệu cục bộ bằng IndexedDB.
3. Hiển thị trạng thái lưu: Đang lưu / Đã tự lưu lúc ... / Lưu lỗi.
4. Backup định kỳ, giữ 5–10 bản gần nhất.
5. Khi mở lại app, nếu có phiên chưa hoàn thành thì hỏi “Tiếp tục phiên đang làm?”.
6. Có nút xuất bản nháp bất kỳ lúc nào.
7. Có nút tải file backup JSON để khôi phục.
8. Có chức năng import file backup JSON để tiếp tục làm.
9. Cảnh báo trước khi thoát/reload nếu có dữ liệu chưa xuất.
10. Nhắc xuất bản dự phòng sau mỗi 10–15 cột hoàn thành.

File Excel đầu vào chuẩn:
- Dùng file theo mẫu `File mau chuan.xlsx`.
- Sheet chuẩn: `Sheet1`.
- Các cột bắt buộc: `MA_KHANG`, `TEN_KHANG`, `SO_TBI`, `SO_COT`, `Tram`, `CHUAN_HOA`.
- Cột `CHUAN_HOA` có thể để trống ở file đầu vào.

File Excel đầu ra:
Xuất file .xlsx có sheet đầu tiên theo đúng mẫu chuẩn:
Tên file xuất phải có tên trạm và trạng thái `BAN_NHAP` hoặc `HOAN_THANH`.

Sheet 1: Sheet1
Các cột gồm:
- MA_KHANG
- TEN_KHANG
- SO_TBI
- SO_COT
- Tram
- CHUAN_HOA
- GHI_CHU

Quy tắc ghi chú Sheet1:
- Khách thêm thực tế: thêm vào đúng `SO_COT` đang làm và ghi `GHI_CHU = Thêm mới`.
- Khách bị xóa: vẫn giữ trong `Sheet1`, tô đỏ cả dòng và ghi rõ lý do xóa trong `GHI_CHU`.

Các sheet đối soát phía sau:

Sheet 2: DU_LIEU_CHUAN_HOA
Các cột gồm:
- Số cột
- Mã khách hàng
- Tên khách hàng
- Địa chỉ nếu có
- Pha
- Lộ hạ thế
- Loại vị trí: hòm cột hoặc tủ thanh cái
- STT hòm
- Loại hòm
- Số mặt tủ
- Loại tủ
- Vị trí khách hàng
- Mã chuẩn hóa
- Trạng thái

Sheet 3: TONG_HOP_THEO_COT
Các cột gồm:
- Số cột
- Tổng khách hàng
- Đã gán
- Chưa gán
- Số hòm
- Số tủ
- Trạng thái
- Có lỗi hay không

Sheet 4: DANH_SACH_LOI
Các cột gồm:
- Số cột
- Mã khách hàng
- Loại lỗi
- Mô tả lỗi
- Gợi ý xử lý

Sheet 5: NHAT_KY_THAO_TAC
Các cột gồm:
- Thời gian
- Tên cán bộ
- Số cột
- Hành động

Các lỗi cần kiểm tra:
1. Khách hàng chưa được gán mã chuẩn hóa.
2. Một khách hàng bị gán nhiều hơn một vị trí nếu có.
8. Một khách hàng bị gán nhiều vị trí.
9. Hòm bị gán quá số lượng công tơ thiết kế.
10. Tủ bị gán quá số lượng công tơ thiết kế.
11. Khách hàng chưa được gán.
12. Khách hàng 3 pha nhưng mã không bắt đầu bằng D.
13. Khách hàng 1 pha nhưng mã bắt đầu bằng D.

Stack kỹ thuật đề xuất:
Frontend:
- React hoặc Next.js
- Tailwind CSS
- PWA
- IndexedDB để lưu offline
- Zustand hoặc Redux Toolkit để quản lý trạng thái
- SheetJS/xlsx để đọc và xuất Excel

Backend:
- MVP không bắt buộc backend.
- Dữ liệu khách hàng xử lý cục bộ trên điện thoại.
- Server chỉ dùng để host web app nếu cần.

Triển khai:
- Có thể triển khai static web app lên Vercel, Netlify, Cloudflare Pages hoặc hosting nội bộ.
- App phải dùng được trên trình duyệt điện thoại Android.
- Có thể thêm vào màn hình chính như app.

Tính năng loại khỏi giai đoạn đầu:
1. Đăng nhập/phân quyền phức tạp.
2. Đồng bộ cloud nhiều người.
3. Bản đồ GIS.
4. GPS.
5. Chụp ảnh hiện trường.
6. OCR/AI nhận diện hòm từ ảnh.
7. Tích hợp hệ thống nội bộ.
8. Android native APK.
9. Dashboard quản lý cấp điện lực.

Tiêu chí thành công MVP:
1. Cán bộ hiện trường có thể import Excel và chuẩn hóa dữ liệu bằng điện thoại.
2. Giảm tối đa nhập mã thủ công.
3. App tự sinh đúng mã 7 ký tự và 8 ký tự.
4. App áp dụng đúng quy tắc Z.
5. Không mất dữ liệu khi reload, tắt trình duyệt hoặc mất mạng.
6. Xuất được file Excel bản nháp và bản hoàn thành.
7. File xuất có sheet dữ liệu, tổng hợp, lỗi và nhật ký.
8. Người dùng không biết kỹ thuật vẫn dùng được sau hướng dẫn ngắn.

Prompt yêu cầu trợ lý tiếp theo cần thực hiện:
Hãy thiết kế và triển khai MVP cho “Web App Chuẩn Hóa Mã Hòm Công Tơ Hiện Trường” theo brief trên. Ưu tiên tạo bản chạy được trên điện thoại, giao diện đơn giản, import/export Excel, giao diện lưới Z-Rule, tự sinh mã theo quy tắc, autosave bằng IndexedDB, có backup/restore, có xuất bản nháp và xuất file hoàn thành. Không cần backend trong MVP đầu. Hãy bắt đầu bằng thiết kế kiến trúc, cấu trúc dữ liệu, luồng màn hình, sau đó triển khai code hoàn chỉnh.
