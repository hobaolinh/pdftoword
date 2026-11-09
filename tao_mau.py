"""Generate the three synthetic PDFs used for local acceptance testing."""

from io import BytesIO
from pathlib import Path
import os

from PIL import Image, ImageDraw, ImageFont
import pymupdf
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas

ROOT = Path(__file__).resolve().parent / "mau"
FONTS = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
A4 = (595.28, 841.89)
INK = "#18365B"
BLUE = "#265E99"


def text(c, x, y, value, size=12, font="Arial", color="#18263B"):
    c.setFont(font, size)
    c.setFillColor(HexColor(color))
    c.drawString(x, c._pagesize[1]-y, value)


def head(c, title, subtitle):
    text(c, 54, 37, "HỒ SƠ KIỂM THỬ  /  STUDIO TÀI LIỆU", 9, "ArialB", BLUE)
    text(c, 54, 93, title, 24, "ArialB", INK)
    text(c, 54, 116, subtitle, 11, "ArialI", BLUE)
    c.setStrokeColor(HexColor(BLUE))
    c.line(54, c._pagesize[1]-120, 54+pdfmetrics.stringWidth(subtitle, "ArialI", 11), c._pagesize[1]-120)


def foot(c, page):
    text(c, 54, c._pagesize[1]-32, "LINHHB  •  Tài liệu kiểm thử", 9, "Arial", BLUE)
    text(c, c._pagesize[0]-86, c._pagesize[1]-32, f"Trang {page}", 9, "Arial", BLUE)


def table(c, x, y):
    h = c._pagesize[1]
    c.setStrokeColor(HexColor("#456489"))
    c.setLineWidth(.8)
    c.setFillColor(HexColor("#DBE9FA"))
    c.rect(x, h-y-31, 480, 31, fill=1, stroke=0)
    c.setFillColor(HexColor("#EFF5FC"))
    c.rect(x, h-y-62, 480, 31, fill=1, stroke=0)
    for yy in (0, 31, 62, 93, 124):
        c.line(x, h-y-yy, x+480, h-y-yy)
    for xx, top in ((0,0),(164,31),(318,31),(480,0)):
        c.line(x+xx, h-y-top, x+xx, h-y-124)
    text(c, x+12, y+21, "TỔNG HỢP ĐỐI CHIẾU", 12, "ArialB", INK)
    for r, values in enumerate((("Hạng mục","Trạng thái","Ghi chú"),
                                ("Văn bản","Hoàn tất","Có dấu"),
                                ("Bảng biểu","Đã kiểm tra","Ô gộp"))):
        for col, value in enumerate(values):
            text(c, x+(12,176,330)[col], y+52+r*31, value, 11, "ArialB" if not r else "Arial")


def illustration():
    im = Image.new("RGB", (780, 300), "#DDEBFA")
    draw = ImageDraw.Draw(im)
    draw.rounded_rectangle((16,16,764,284), 27, fill="#BDD6F2")
    draw.ellipse((584,38,674,128), fill="#F9CC6F")
    draw.polygon(((22,248),(237,78),(390,215),(490,120),(752,264)), fill="#3975A6")
    draw.polygon(((30,275),(280,194),(426,269),(541,200),(754,275)), fill="#315884")
    return im


def image_text():
    im = Image.new("RGB", (1100, 235), "#EDF5FE")
    draw = ImageDraw.Draw(im)
    draw.rounded_rectangle((5,5,1094,228), 20, outline="#BDD3EE", width=5)
    font = ImageFont.truetype(str(FONTS/"arial.ttf"), 52)
    draw.text((35,35), "Điểm tiếp nhận hồ sơ", fill="#18365B", font=font)
    draw.text((35,123), "Opening hours: 08:00 - 17:00", fill="#18365B", font=font)
    return im


def picture(c, image, x, y, width, height):
    c.drawImage(ImageReader(image), x, c._pagesize[1]-y-height, width, height)


def pdf_text():
    c = Canvas(str(ROOT/"01-van-ban.pdf"), pagesize=A4)
    head(c, "Báo cáo đối chiếu", "Văn bản tiếng Việt, bố cục hai cột và bảng có ô gộp")
    text(c, 54, 158, "Cột trái", 16, "ArialB", BLUE)
    for i, line in enumerate(("Kiểm tra dữ liệu nguồn theo ngày.", "Ghi nhận sai lệch và đối soát.",
                              "Đoạn này có thể sửa trong Word.")):
        text(c, 54, 182+i*17, line)
    text(c, 321, 158, "Cột phải", 16, "ArialB", BLUE)
    for i, line in enumerate(("Xác nhận số liệu đã tổng hợp.", "Chuyển hồ sơ sang bộ phận xử lý.",
                              "Giữ nguyên thứ tự các mục.")):
        text(c, 321, 182+i*17, line)
    table(c, 54, 264)
    picture(c, illustration(), 54, 434, 255, 98)
    text(c, 54, 550, "Hình 1. Minh họa phong cảnh", 10, "ArialI", BLUE)
    foot(c, 1)
    c.showPage()
    c.setPageSize(A4[::-1])
    head(c, "Phụ lục trang ngang", "Kiểm tra kích thước giấy và ngắt trang")
    text(c, 54, 157, "1. Đọc nội dung gốc.")
    text(c, 54, 180, "2. Đối chiếu tệp Word.")
    text(c, 54, 203, "• Kiểm tra định dạng tiếng Việt.")
    table(c, 54, 246)
    foot(c, 2)
    c.save()


def pdf_scan():
    tmp = BytesIO()
    c = Canvas(tmp, pagesize=A4)
    head(c, "Phiếu kiểm thử bản quét", "Nhận dạng tiếng Việt có dấu")
    text(c, 54, 164, "Hồ sơ đã được kiểm tra và xác nhận đầy đủ.", 14)
    text(c, 54, 188, "Vui lòng đối chiếu nội dung trước khi lưu.", 14)
    table(c, 54, 250)
    picture(c, illustration(), 54, 434, 255, 98)
    foot(c, 1)
    c.save()
    with pymupdf.open(stream=tmp.getvalue(), filetype="pdf") as original, pymupdf.open() as output:
        pix = original[0].get_pixmap(dpi=230)
        page = output.new_page(width=A4[0], height=A4[1])
        page.insert_image(page.rect, stream=pix.tobytes("jpeg", jpg_quality=89))
        output.save(ROOT/"02-scan-tieng-viet.pdf", deflate=True)


def pdf_mixed():
    c = Canvas(str(ROOT/"03-hon-hop.pdf"), pagesize=A4)
    head(c, "Thông báo tiếp nhận", "Văn bản gốc kết hợp ảnh chứa chữ")
    text(c, 54, 161, "Dòng này thuộc lớp chữ có sẵn trong PDF.", 13)
    picture(c, image_text(), 54, 204, 480, 103)
    text(c, 54, 341, "Thông tin phía trên được trích xuất bằng OCR.", 12)
    table(c, 54, 386)
    picture(c, illustration(), 54, 561, 255, 98)
    foot(c, 1)
    c.save()


if __name__ == "__main__":
    ROOT.mkdir(exist_ok=True)
    for name, file in (("Arial","arial.ttf"),("ArialB","arialbd.ttf"),("ArialI","ariali.ttf")):
        pdfmetrics.registerFont(TTFont(name, str(FONTS/file)))
    pdf_text()
    pdf_scan()
    pdf_mixed()
