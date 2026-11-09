# Kết quả kiểm thử

**Ngày 07/11/2026 · Windows · Python 3.12 · Tesseract 5.5.0 · `vie` và `eng`**

Ba PDF tổng hợp được gửi qua HTTP đến ứng dụng Flask đang chạy trên máy. Kiểm thử kiểm tra số trang đã xử lý, loại trang, dữ liệu xem trước, ảnh gốc, ảnh minh họa, tải Word và dọn phiên. Các tệp `.docx` trong `mau/` là kết quả tải xuống từ ứng dụng.

| Mẫu | Chế độ | Trang PDF → trang kết xuất | Bảng Word | Ô gộp ngang | Ảnh Word | Thời gian lượt kiểm thử |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `01-van-ban` | Văn bản | 2 → 2 | 2 | 2 | 1 | 0,61 giây |
| `02-scan-tieng-viet` | Scan | 1 → 1 | 1 | 1 | 1 | 2,11 giây |
| `03-hon-hop` | Hỗn hợp | 1 → 1 | 1 | 1 | 2 | 0,83 giây |

Thời gian gồm gửi tệp, chờ chuyển đổi, đọc các trang xem trước và tải kết quả. Đây là số đo trên các tệp mẫu nhỏ.

## Đối chiếu từng trang

| Mẫu và trang | Kết quả quan sát | Chi tiết cần chỉnh thủ công |
| --- | --- | --- |
| Văn bản, trang 1 | Hai cột đúng vị trí. Thứ tự khối xuất từ cột trái sang cột phải. Giữ chữ đậm, nghiêng, gạch chân, ảnh ở tọa độ gốc, đầu trang, chân trang và số trang. Ô tiêu đề gộp 3 cột, có viền và màu nền. | Đường cơ sở của một số dòng và mép dưới bảng dịch nhẹ. |
| Văn bản, trang 2 | Giữ khổ ngang và ngắt trước trang 2. Có hai mục đánh số, một dấu đầu dòng, một bảng có ô gộp và trường số trang 2. | Dấu đầu dòng và phần chữ ở chân trang có thể lệch nhẹ so với PDF. |
| Scan tiếng Việt, trang 1 | Hai câu thân bài mẫu đúng dấu. Dựng lại bảng có ô gộp và màu nền ước lượng. Ảnh phong cảnh được cắt ra riêng. | OCR đọc “HỒ SƠ KIỂM THỬ” thành “HÒ SƠ KIÊM THỬ”, “TỔNG HỢP ĐỐI CHIẾU” thành “TÓNG HỢP ĐÓI CHIẾU”, “Hoàn tất” thành “Hoàn tắt” và “Trang 1” thành “Trang †1”. Màu và kiểu chữ OCR, độ đậm viền, kích thước ảnh cắt và gạch chân dưới tiêu đề phụ còn khác. Các vùng dưới 75% được đánh dấu. Một số lỗi OCR có độ tin cậy cao hơn ngưỡng. |
| Hỗn hợp, trang 1 | Chữ gốc và hai dòng OCR xuất hiện đúng một lần. Hai ảnh ở tọa độ gốc. Giữ bảng gộp, màu nền, đầu và chân trang. | Hai dòng chữ OCR trong khung xanh có cỡ, màu và khoảng cách khác nhẹ. Nền ảnh được tô lại tại vùng chữ để tránh trùng. |

## Kiểm thử khác

- **Đạt qua HTTP.** Tệp lỗi trả thông báo lỗi, tệp có mật khẩu trả thông báo mở khóa, tệp trên 35 MB bị từ chối, trang trắng có cảnh báo thiếu chữ. Đã dùng cả ba lựa chọn OCR. Sau khi tải Word, phiên xem trước và lượt tải lại trả 404.
- **Trình duyệt.** Đã mở giao diện trên máy, kiểm tra bố cục ở chiều rộng hẹp và chọn PDF mẫu qua hộp chọn tệp.
- **Đối chiếu trực quan.** Đã kết xuất và xem đủ bốn trang bằng Aspose.Words 26.10.0 trên máy. Dấu bản dùng thử chỉ xuất hiện trong bản kết xuất kiểm tra. Tệp bàn giao không có dấu này.
- **Sửa và mở lại.** Đã sửa đoạn “Đoạn này có thể sửa trong Word.” và ô “Hoàn tất” trong một bản sao bằng Aspose.Words. Hai thay đổi còn nguyên sau khi lưu và mở lại. Kiểm tra lại xác nhận ô đã sửa vẫn thuộc bảng.
- **Microsoft Word 2013.** Đã mở Word và hộp thoại chọn tệp. Phiên điều khiển dừng khi người dùng nhấn Escape trước khi mở tệp mẫu. Việc đối chiếu, chỉnh sửa, lưu và mở lại trong Microsoft Word vẫn chưa được xác nhận.

## Lặp lại bước nghiệm thu trong Microsoft Word

1. Mở từng PDF và DOCX cùng tên trong `mau/`. Đối chiếu theo bảng trên ở chế độ Print Layout.
2. Lưu một bản sao của `01-van-ban.docx`. Trên trang 1, bấm vào đoạn ở cột trái và thay câu “Đoạn này có thể sửa trong Word.”. Sửa ô “Hoàn tất” trong bảng cùng trang.
3. Lưu, đóng rồi mở lại bản sao. Kiểm tra cả hai thay đổi, ô gộp, vị trí ảnh, số trang và trang ngang.
