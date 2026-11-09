# PDF sang Word

Ứng dụng Flask chạy trên Windows tại `http://127.0.0.1/` (cổng 80 mặc định). Giao diện tiếng Việt có ba bước **Chọn PDF → Chuyển đổi → Tải Word**. PDF, ảnh xem trước, nội dung OCR và tệp Word được xử lý trên máy chạy ứng dụng.

## Cài đặt trên Windows

1. Cài Python 3.12 64-bit và Tesseract OCR 5.x. Có thể dùng bộ cài Windows được liên kết trong [hướng dẫn cài Tesseract](https://tesseract-ocr.github.io/tessdoc/Installation.html).
2. Kiểm tra thư mục `tessdata` có `vie.traineddata` và `eng.traineddata`. Với tài liệu scan nhiều chữ, dùng [tessdata_best](https://github.com/tesseract-ocr/tessdata_best) để ưu tiên độ chính xác. Có thể dùng [tessdata_fast](https://github.com/tesseract-ocr/tessdata_fast) khi ưu tiên tốc độ. Đặt hai tệp cùng loại trong một thư mục.
3. Mở PowerShell trong thư mục chứa `app.py`. Tạo môi trường Python và cài thư viện.

   ```powershell
   py -3.12 -m venv .venv
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

4. Cấu hình Tesseract rồi kiểm tra danh sách ngôn ngữ. Sửa hai đường dẫn theo vị trí thực tế.

   ```powershell
   $env:TESSERACT_CMD = 'C:\Program Files\Tesseract-OCR\tesseract.exe'
   $env:TESSDATA_PREFIX = 'C:\Program Files\Tesseract-OCR\tessdata'
   & $env:TESSERACT_CMD --list-langs --tessdata-dir $env:TESSDATA_PREFIX
   ```

5. Chạy ứng dụng rồi mở `http://127.0.0.1/` trong trình duyệt. Giữ PowerShell mở khi sử dụng. Nhấn **Ctrl+C** để dừng.

   ```powershell
   .\.venv\Scripts\python.exe app.py
   ```

   Có thể dùng `Chay-ung-dung.ps1` nếu PowerShell cho phép chạy tập lệnh. Nếu cổng 80 đang được dùng, chạy `app.py --port 8766` hoặc `Khoi-dong-PDF-sang-Word.cmd --port 8766`, rồi mở `http://127.0.0.1:8766/`.

## Sử dụng và đối chiếu

1. Chọn hoặc kéo thả một PDF tối đa **100 MB**, **900 trang**.
2. Chọn tiếng Việt, tiếng Anh hoặc cả hai cho OCR. Chữ có sẵn trong PDF được ưu tiên khi trùng vùng.
3. Nếu có lượt xử lý trước, xem vị trí chờ được cập nhật mỗi 0,5 giây. Số 1 là tệp chờ kế tiếp. Khi đến lượt, ứng dụng hiển thị tiến độ theo trang.
4. Đối chiếu từng trang và các vùng cảnh báo, sau đó tải `.docx`. Kết quả sẵn sàng giữ trong bộ nhớ tối đa 15 phút và được dọn khi máy chủ bắt đầu gửi tệp. Để tải lại sau khi đã dọn, chuyển đổi lại PDF.

Bấm **Hủy** bên cạnh **Tải tệp Word** nếu muốn bỏ kết quả đang xem và giải phóng phiên ngay. Đồng hồ **Tự hủy sau** hiển thị thời gian còn lại của phiên.

Ứng dụng nhận tối đa **10 phiên** cùng lúc, gồm tệp đang xử lý, tệp đang chờ và kết quả đang giữ để tải xuống. Tệp được chuyển đổi lần lượt.

Trong Word, bấm vào chữ trong khối để đặt con trỏ và sửa trực tiếp. Bảng dùng ô Word, gồm ô gộp khi xác định được. Có thể thay đổi nội dung, thêm hàng và chỉnh đường viền. Với chỉnh sửa dài làm thay đổi bố cục, có thể cần kéo giãn khung văn bản hoặc dời khối kế tiếp. Số trang nhận diện được ở chân trang dùng trường `PAGE` khi giá trị khớp số trang PDF.

## Giới hạn cần biết khi kiểm tra bản chuyển đổi

- Vị trí được dựng theo từng trang và thứ tự đọc được suy ra từ khoảng cách giữa các khối. Khối ở nhiều cột được đặt tại tọa độ gốc. Khi sửa dài, Word có thể xuống dòng trong khung hiện tại và cần chỉnh vị trí các khối xung quanh.
- Phông, cỡ, màu, đậm, nghiêng và gạch chân lấy từ lớp chữ và nét PDF khi xác định được. Trang scan dùng ước lượng cỡ chữ và vị trí từ OCR. Ngôn ngữ, dấu, độ nghiêng, ảnh mờ, kiểu chữ và kích thước nhỏ có thể làm giảm độ chính xác. Văn bản OCR giữ nguyên kết quả trả về.
- Trang scan nhiều chữ không có lưới bảng dùng đoạn văn Word theo luồng. Các dòng chữ in được ưu tiên khi làm sạch nền scan. Cần đối chiếu thủ công chữ viết tay, chữ trên con dấu và vùng bị dấu che.
- Trên trang scan nhiều chữ, ứng dụng có thể giữ dòng ngắn có mực viết tay, chữ ký và con dấu dưới dạng ảnh cắt từ PDF tại vị trí gốc. Các vùng này được đánh dấu trong bản đối chiếu. Văn bản ngoài vùng ảnh vẫn có thể chỉnh sửa trong Word.
- Bảng đường kẻ rõ có khả năng dựng lại hàng, cột, ô gộp và viền. Bảng không kẻ, đường chéo, viền trang trí, màu nền ảnh scan hoặc ô quá phức tạp có thể phải chỉnh tay. Bố cục không có lưới rõ có thể được giữ dưới dạng các khối chữ.
- Ảnh minh họa được giữ theo vùng. Trên ảnh chứa chữ, nền tại vùng chữ có thể có vết sau khi tách. Các hình vector phức tạp, dấu, đồ họa chồng lấp và ảnh scan xám có thể cần phục hồi thủ công.
- Bản xem trước dùng mô hình bố cục của bộ chuyển đổi. Để kiểm tra ngắt trang và cách Word dùng phông thay thế, mở tệp đã tải trong Microsoft Word.
- Tệp trên 100 trang có ảnh PDF xem trước nhỏ hơn để giảm bộ nhớ. Giới hạn 15 phút chỉ áp dụng sau khi đã chuyển đổi xong. Nếu dữ liệu xuất và ảnh xem trước vượt tổng giới hạn 512 MB, ứng dụng yêu cầu chia nhỏ tệp.
- Cảnh báo về chữ quá nhỏ hoặc dòng OCR không đọc được xuất hiện ở trang tương ứng trong mục đối chiếu. Lỗi làm dừng chuyển đổi xuất hiện dưới ô chọn PDF.
- Waitress chỉ lắng nghe trên loopback. Dữ liệu chuyển đổi nằm trong RAM, Tesseract nhận ảnh và trả dữ liệu qua luồng chuẩn. Gói cài đặt và gói OCR có thể cần Internet khi tải lần đầu. Quy trình chuyển đổi không gọi dịch vụ ngoài.

## Mã nguồn và kiểm thử

| Thành phần | Vai trò |
| --- | --- |
| `app.py` | Máy chủ cục bộ, hàng đợi một bộ xử lý, tiến độ, xem trước, tải xuống và dọn phiên |
| `layout.py` | Trích xuất PyMuPDF, OCR Tesseract, bố cục trang, bảng và ảnh |
| `word.py` | Xuất đoạn văn, danh sách, hình, bảng và ngắt trang bằng python-docx và OOXML |
| `templates/`, `static/` | Giao diện và bản xem trước cục bộ |
| `kiem_thu.py` | Kiểm thử HTTP và kiểm tra cấu trúc Word |
| `tao_mau.py`, `mau/` | Tạo và lưu bộ mẫu tổng hợp |

Chạy kiểm thử khi ứng dụng đang mở. Dùng thêm `--write-samples` để xuất lại ba tệp mẫu. Cài thêm `reportlab` nếu muốn tạo lại PDF bằng `tao_mau.py`.

```powershell
.\.venv\Scripts\python.exe kiem_thu.py
```

Xem `KET-QUA-KIEM-THU.md` để biết kết quả trên bộ mẫu bàn giao và các chi tiết cần đối chiếu thủ công.
