"""Run the embedded Tesseract engine with isolated standard streams."""

import sys

import pymupdf as pdf


def main():
    language, dpi, tessdata = sys.argv[1:]
    pixmap = pdf.Pixmap(sys.stdin.buffer.read())
    pixmap.set_dpi(int(dpi), int(dpi))
    sys.stdout.buffer.write(pixmap.pdfocr_tobytes(language=language, tessdata=tessdata))


if __name__ == "__main__":
    main()
