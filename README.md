# Web App Chuẩn Hóa Mã Hòm Công Tơ

App dùng để nhập file Excel theo mẫu `File mau chuan.xlsx`, lấy dữ liệu từ sheet `Sheet1`, chuẩn hóa mã hòm/tủ theo quy tắc Z và xuất file Excel kết quả.

## Cách chạy trên Windows

1. Bấm đúp file `run.bat`.
2. Trình duyệt sẽ tự mở địa chỉ `http://localhost:5173`.
3. Nếu trình duyệt không tự mở, copy địa chỉ trên màn hình đen và dán vào Chrome/Edge.
4. Bấm **Chọn file Excel** và chọn file dữ liệu.
5. Kiểm tra ô **Tên trạm đang làm**. App sẽ tự lấy từ cột `Tram`, nhưng có thể sửa tay nếu cần.

Máy cần có Python để chạy server local. Nếu lệnh `py` không chạy, cài Python bản mới hoặc mở app bằng lệnh:

```bat
cd app
python -m http.server 5173
```

## Quy tắc lấy dữ liệu

- App đọc dữ liệu theo mẫu `File mau chuan.xlsx`.
- Sheet chuẩn là `Sheet1`.
- Các cột cần có: `MA_KHANG`, `TEN_KHANG`, `SO_TBI`, `SO_COT`, `Tram`, `CHUAN_HOA`.
- Cột `CHUAN_HOA` có thể để trống ở file đầu vào; app sẽ điền mã chuẩn hóa khi xuất.

## Xuất dữ liệu

File xuất có sheet đầu tiên theo đúng mẫu chuẩn và thêm cột ghi chú:

- `Sheet1`: gồm `MA_KHANG`, `TEN_KHANG`, `SO_TBI`, `SO_COT`, `Tram`, `CHUAN_HOA`, `GHI_CHU`
- `DU_LIEU_CHUAN_HOA`
- `TONG_HOP_THEO_COT`
- `DANH_SACH_LOI`
- `NHAT_KY_THAO_TAC`

Tên file xuất có kèm tên trạm và trạng thái, ví dụ:

```text
chuan_hoa_Tay_Giang_12_BAN_NHAP_20260626_1050.xlsx
chuan_hoa_Tay_Giang_12_HOAN_THANH_20260626_1055.xlsx
```

Trong màn **Chuẩn hóa**, có thể thêm khách thực tế nếu ngoài hiện trường phát sinh khác file Excel:

1. Mở cột cần sửa.
2. Nhập **Số thiết bị** và **Tên khách hàng** nếu có ở phần **Thêm khách thực tế**.
3. Bấm **Thêm khách**.
4. Chọn khách mới trong ô gán hòm/tủ như khách bình thường.

Khi bấm **Xóa khách**, app sẽ không xóa mất dữ liệu mà chuyển khách sang mục **Khách hàng khác**. Sang cột đúng, bấm **Thêm vào cột này** để đưa khách đó vào cột hiện tại.

Nếu nhập **Số thiết bị** ở phần **Thêm khách thực tế** mà số đó đang nằm trong **Khách hàng khác**, app sẽ tự đưa khách đó vào cột hiện tại để tránh tạo trùng. Khi xuất file, các khách còn nằm trong mục này được ghi thêm ở sheet `KHACH_HANG_KHAC`.

Ở màn **Danh sách cột**, cuối danh sách có nút **Thêm cột**. Dùng nút này khi ngoài hiện trường có cột phát sinh chưa có trong file Excel, sau đó nhập khách ở phần **Thêm khách thực tế**.

Nút **Xuất nháp & chia sẻ Zalo** dùng để gửi bản nháp, tên file sẽ ghi rõ trạm nào. Nút **Xuất & chia sẻ Zalo** dùng để gửi bản hoàn thành. Nếu điện thoại/trình duyệt hỗ trợ chia sẻ file, app sẽ mở bảng chia sẻ để sếp chọn Zalo và gửi file Excel trực tiếp.

Người ở hiện trường khi nhận file qua Zalo có thể mở ngay trong đoạn chat Zalo. Nếu muốn lưu lại, mới cần bấm tải/lưu file trên máy hoặc điện thoại.

## Đóng gói bản web tĩnh

Bấm đúp `build.bat`. Kết quả nằm trong thư mục `dist`.
## CÃ¡ch chá»n file Excel trÃªn Ä‘iá»‡n thoáº¡i

Náº¿u file Excel Ä‘Æ°á»£c gá»­i qua Zalo, Messenger hoáº·c email, Ä‘iá»‡n thoáº¡i thÆ°á»ng khÃ´ng Ä‘Æ°a file tháº³ng ra mÃ n hÃ¬nh chá»n file. LÃ m theo cÃ¡ch nÃ y:

1. Má»Ÿ file Excel trong Zalo/email.
2. Báº¥m **LÆ°u vá» mÃ¡y**, **Táº£i xuá»‘ng** hoáº·c **Chia sáº»**.
3. Quay láº¡i app, báº¥m **Chá»n file Excel**.
4. TrÃªn mÃ n chá»n file, tÃ¬m trong **Gáº§n Ä‘Ã¢y**, **Táº£i xuá»‘ng/Download** hoáº·c thÆ° má»¥c **Zalo**.

Náº¿u dÃ¹ng Android vÃ  Ä‘Ã£ cÃ i app lÃªn mÃ n hÃ¬nh chÃ­nh:

1. Má»Ÿ file Excel trong Zalo.
2. Báº¥m **Chia sáº»**.
3. Chá»n app **Chuáº©n hÃ³a CT**.
4. App sáº½ tá»± má»Ÿ vÃ  Ä‘á»c file Excel, khÃ´ng cáº§n tÃ¬m file thá»§ cÃ´ng.
