"""Run the acceptance cases over HTTP against the local Windows server."""

from argparse import ArgumentParser
from io import BytesIO
import json
from pathlib import Path
from time import monotonic, sleep
from urllib.error import HTTPError
from urllib.parse import quote
from urllib.request import Request, urlopen
from zipfile import ZipFile

from lxml import etree
import pymupdf

DIR = Path(__file__).resolve().parent / "mau"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "v": "urn:schemas-microsoft-com:vml", "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}
CASES = [
    ("01-van-ban", "vie+eng", ["Văn bản","Văn bản"], "Đoạn này có thể sửa trong Word."),
    ("02-scan-tieng-viet", "vie", ["Scan"], "Hồ sơ đã được kiểm tra và xác nhận đầy đủ."),
    ("03-hon-hop", "vie+eng", ["Hỗn hợp"], "Điểm tiếp nhận hồ sơ"),
]


def call(base, path, data=None, headers=None, expected=200):
    try:
        response = urlopen(Request(base + path, data=data, headers=headers or {}), timeout=20)
    except HTTPError as err:
        response = err
    with response:
        content = response.read()
        assert response.status == expected, (path, response.status, content[:500])
        if response.headers.get_content_type() == "application/json":
            return json.loads(content)
        return content


def submit(base, data, name, lang, expected=202):
    return call(base, "/api/jobs?lang=" + quote(lang), data,
                {"X-Local-Converter":"1", "Content-Type":"application/pdf", "X-PDF-Name":quote(name)},
                expected)


def wait(base, path):
    deadline = monotonic()+150
    while monotonic() < deadline:
        state = call(base, path)
        if state["state"] in {"ready","failed"}:
            return state
        sleep(.2)
    raise AssertionError("Job did not finish")


def inspect_docx(data, pages, expected_text, name):
    with ZipFile(BytesIO(data)) as archive:
        assert archive.testzip() is None
        root = etree.fromstring(archive.read("word/document.xml"))
        text = " ".join(root.xpath("//w:txbxContent//w:t/text()", namespaces=NS))
        assert expected_text in text, text
        section_count = len(root.xpath("//w:sectPr", namespaces=NS))
        assert section_count == pages, section_count
        tables = len(root.xpath("//w:tbl", namespaces=NS))
        merged = root.xpath("//w:gridSpan[@w:val='3']", namespaces=NS)
        assert tables >= 1 and merged, (tables, len(merged), text)
        pictures = root.xpath("//wp:anchor/wp:extent", namespaces=NS)
        assert pictures and all(int(p.get("cx")) < 7_000_000 and int(p.get("cy")) < 7_000_000
                                for p in pictures), "Missing illustration or full-page image"
        if name.startswith("01"):
            assert root.xpath("//w:pgSz[@w:orient='landscape']", namespaces=NS)
            assert root.xpath("//w:numPr", namespaces=NS)
            assert root.xpath("//w:shd[@w:fill='dbe9fa']", namespaces=NS)
            assert root.xpath("//w:rPr/w:u[@w:val='single']", namespaces=NS)
            assert text.index("Cột trái") < text.index("Đoạn này") < text.index("Cột phải"), text
        if name.startswith("03"):
            assert text.count(expected_text) == 1 and "Opening hours" in text, text
            assert text.count("Dòng này thuộc lớp chữ có sẵn trong PDF.") == 1, text
        return {"sections":section_count, "tables":tables, "merged_cells":len(merged),
                "pictures":len(pictures)}


def main():
    args = ArgumentParser()
    args.add_argument("--url", default="http://127.0.0.1")
    args.add_argument("--write-samples", action="store_true")
    args = args.parse_args()
    base = args.url
    assert b"VikkiBank" in call(base,"/")
    for name, lang, modes, sample_text in CASES:
        start = monotonic()
        key = submit(base, (DIR/(name+".pdf")).read_bytes(), name+".pdf", lang)["id"]
        path = "/api/jobs/" + key
        state = wait(base, path)
        assert state["state"] == "ready", state
        assert state["complete"] == state["total"] == len(modes)
        assert [item["mode"] for item in state["pages"]] == modes, state
        warnings = []
        for i in range(len(modes)):
            url = path + "/pages/" + str(i)
            model = call(base, url)
            assert call(base, url + "/original").startswith(b"\xff\xd8")
            for image in [b for b in model["items"] if b["kind"] == "image"]:
                assert call(base, url + "/assets/" + str(image["asset"])).startswith(b"\xff\xd8")
            warnings.append(model["warnings"])
        result = call(base, path + "/download", b"", {"X-Local-Converter":"1"})
        checks = inspect_docx(result, len(modes), sample_text, name)
        call(base, path, expected=404)
        call(base, path + "/download", b"", {"X-Local-Converter":"1"}, expected=404)
        if args.write_samples:
            (DIR/(name+".docx")).write_bytes(result)
        print(json.dumps({"file":name, "seconds":round(monotonic()-start,2), **checks,
                          "warnings":warnings}, ensure_ascii=False))

    invalid = submit(base, b"not a pdf", "hong.pdf", "vie", 400)
    assert "bị hỏng" in invalid["error"]
    too_big = submit(base, b"%"*(100*1024*1024+1), "qua-lon.pdf", "eng", 413)
    assert "100 MB" in too_big["error"]
    with pymupdf.open() as doc:
        doc.new_page().insert_text((50,50), "Protected sample")
        protected = doc.tobytes(encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="sample-pass",
                               owner_pw="sample-owner-pass")
    state = wait(base, "/api/jobs/" + submit(base, protected, "co-mat-khau.pdf", "eng")["id"])
    assert state["state"] == "failed" and "mật khẩu" in state["error"], state
    with pymupdf.open() as doc:
        doc.new_page()
        empty = doc.tobytes()
    empty_path = "/api/jobs/" + submit(base, empty, "trang-trang.pdf", "eng")["id"]
    assert wait(base, empty_path)["state"] == "ready"
    assert any(w["code"] == "empty" for w in call(base, empty_path + "/pages/0")["warnings"])
    call(base, empty_path + "/download", b"", {"X-Local-Converter":"1"})
    print("PASS: invalid, size limit, password, no text, download cleanup")


if __name__ == "__main__":
    main()
