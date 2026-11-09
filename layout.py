"""Extract a page model shared by the Word exporter and browser preview."""

from collections import defaultdict
from functools import lru_cache
from io import BytesIO
from pathlib import Path
import csv
import math
import os
import re
import shutil
import subprocess
import sys

import pymupdf as pdf
from PIL import Image, ImageChops, ImageDraw, ImageOps, ImageStat
from scan_marks import preserve_ink

LOW_CONFIDENCE = 75
OCR_SCALE = 300 / 72
OCR_PIXELS = 8_000_000
OCR_WORD = re.compile(rb"<((?:[0-9a-fA-F]{4})+)>\s*Tj")


class ConversionError(Exception):
    """An actionable, user-facing conversion failure."""


def rect(box):
    return pdf.Rect(box)


def bounds(items):
    box = rect(items[0]["box"])
    for item in items[1:]:
        box |= rect(item["box"])
    return tuple(box)


def inside(item, box):
    region = rect(item["box"])
    return box.contains((region.tl + region.br) / 2)


def coverage(a, b):
    return (a & b).get_area() / max(0.01, a.get_area())


def raster(page, clip, scale):
    scale = min(scale, math.sqrt(OCR_PIXELS / max(1, clip.get_area())))
    pix = page.get_pixmap(matrix=pdf.Matrix(scale, scale), clip=clip, alpha=False)
    image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    return image, (pix.x / scale, pix.y / scale), scale


def encoded(image, kind, **options):
    result = BytesIO()
    image.save(result, format=kind, **options)
    return result.getvalue()


def ocr_notices(stderr, report):
    if report is None:
        return
    ignored = {
        "Warning: Parameter not found: language_model_ngram_on",
        "Warning: Parameter not found: classify_misfit_junk_penalty",
    }
    lines = {line.strip() for line in stderr.decode("utf-8", "replace").splitlines()}
    if any(line.startswith(("Image too small to scale", "Line cannot be recognized")) for line in lines):
        report("ocr-small", "OCR bỏ qua ít nhất một nét hoặc dòng quá nhỏ. Hãy đối chiếu chữ nhỏ với PDF gốc.")
    if any(line.startswith(("Warning:", "Error:")) for line in lines - ignored):
        report("ocr-notice", "OCR gặp cảnh báo khi nhận dạng trang này. Hãy đối chiếu với PDF gốc.")


def ocr_command(lang):
    configured = os.environ.get("TESSERACT_CMD")
    default = Path(os.environ.get("ProgramFiles", "C:/Program Files")) / "Tesseract-OCR/tesseract.exe"
    executable = configured or shutil.which("tesseract") or str(default)
    command = [executable]
    tessdata = os.environ.get("TESSDATA_PREFIX")
    if tessdata:
        command.extend(["--tessdata-dir", tessdata])
    embedded_data = os.environ.get("PDF_WORD_EMBEDDED_TESSDATA") or tessdata
    fallback = embedded_data and all((Path(embedded_data) / f"{name}.traineddata").is_file()
                                     for name in lang.split("+"))
    if fallback and os.environ.get("PDF_WORD_OCR_ENGINE") == "embedded":
        return None
    try:
        result = subprocess.run(command + ["--list-langs"], capture_output=True,
                                timeout=10, **ocr_process_options(command))
    except (OSError, subprocess.TimeoutExpired) as exc:
        if fallback:
            return None
        raise ConversionError("Không chạy được Tesseract. Kiểm tra TESSERACT_CMD trong hướng dẫn.") from exc
    output = result.stdout + b"\n" + result.stderr
    available = {line.strip().lstrip("\ufeff")
                 for line in output.decode("utf-8", "replace").splitlines()}
    missing = [name for name in lang.split("+") if name not in available]
    if result.returncode or missing:
        if fallback:
            return None
        raise ConversionError("Thiếu dữ liệu OCR: " + ", ".join(missing or [lang]) +
                              ". Kiểm tra thư mục TESSDATA_PREFIX.")
    return command


def ocr_process_options(command):
    options = {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}
    executable = Path(command[0])
    if executable.is_absolute() and executable.is_file():
        env = dict(os.environ)
        env["PATH"] = str(executable.parent) + os.pathsep + env.get("PATH", "")
        env.setdefault("OMP_THREAD_LIMIT", "1")
        options.update(cwd=str(executable.parent), env=env)
    return options


def recognize_embedded(image, origin, scale, lang, report=None):
    try:
        dpi = max(1, round(72 * scale))
        tessdata = os.environ.get("PDF_WORD_EMBEDDED_TESSDATA") or os.environ["TESSDATA_PREFIX"]
        worker = str(Path(__file__).with_name("ocr_worker.py"))
        result = subprocess.run([sys.executable, worker, lang, str(dpi), tessdata],
                                input=encoded(image, "PNG"), capture_output=True, timeout=90,
                                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        if result.returncode or not result.stdout.startswith(b"%PDF-"):
            raise ConversionError("Tesseract tích hợp không xử lý được vùng ảnh. Kiểm tra cấu hình và dữ liệu OCR.")
        ocr_notices(result.stderr, report)
        with pdf.open(stream=result.stdout, filetype="pdf") as document:
            page = document[0]
            # Tesseract writes each recognized word as a positioned UTF-16 Tj.
            # PDF text extraction otherwise joins neighboring words when their
            # invisible glyph boxes touch. Preserve those source word boundaries.
            for stream_id in page.get_contents():
                stream = document.xref_stream(stream_id)
                separated = OCR_WORD.sub(rb"<\g<1>0020>Tj", stream)
                if separated != stream:
                    document.update_stream(stream_id, separated)
            sx, sy = image.width / page.rect.width / scale, image.height / page.rect.height / scale
            groups = defaultdict(list)
            for x0, y0, x1, y1, text, block, line, _ in page.get_text("words"):
                box = (x0*sx+origin[0], y0*sy+origin[1], x1*sx+origin[0], y1*sy+origin[1])
                groups[(str(block), "0", str(line))].append({"box":box, "text":text, "conf":None})
            return groups
    except subprocess.TimeoutExpired as exc:
        raise ConversionError("OCR vượt quá 90 giây ở một vùng ảnh. Hãy thử chia nhỏ PDF.") from exc
    except (KeyError, OSError, ValueError, RuntimeError) as exc:
        raise ConversionError("Tesseract tích hợp không đọc được dữ liệu ngôn ngữ OCR.") from exc


def recognize(image, origin, scale, command, lang, psm, report=None):
    if command is None:
        return recognize_embedded(image, origin, scale, lang, report)
    try:
        result = subprocess.run(
            command + ["stdin", "stdout", "-l", lang, "--psm", str(psm),
                       "--dpi", str(round(72 * scale)), "-c", "tessedit_create_tsv=1"],
            input=encoded(image, "PNG"), capture_output=True, timeout=90,
            **ocr_process_options(command))
    except subprocess.TimeoutExpired as exc:
        raise ConversionError("OCR vượt quá 90 giây ở một vùng ảnh. Hãy thử chia nhỏ PDF.") from exc
    if result.returncode:
        raise ConversionError("Tesseract gặp lỗi khi đọc ảnh. Kiểm tra cấu hình và gói ngôn ngữ OCR.")
    ocr_notices(result.stderr, report)
    rows = result.stdout.decode("utf-8", "replace").splitlines()
    if not rows or not rows[0].startswith("level\t"):
        raise ConversionError("Tesseract không trả về dữ liệu vị trí và độ tin cậy của OCR.")
    groups = defaultdict(list)
    for word in csv.DictReader(rows,
                               delimiter="\t", quoting=csv.QUOTE_NONE):
        content = (word.get("text") or "").strip()
        if not content or word["level"] != "5":
            continue
        x, y, w, h = [int(word[k]) / scale for k in ("left", "top", "width", "height")]
        x, y = x + origin[0], y + origin[1]
        key = tuple(word[k] for k in ("block_num", "par_num", "line_num"))
        groups[key].append({"box": (x, y, x+w, y+h), "text": content, "conf": float(word["conf"])})
    return groups


@lru_cache(maxsize=1)
def installed_fonts():
    result = {name.lower(): name for name in ("Arial", "Times New Roman", "Courier New", "Calibri")}
    try:
        import winreg
    except ImportError:
        return result
    path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, path) as key:
                for i in range(winreg.QueryInfoKey(key)[1]):
                    name = winreg.EnumValue(key, i)[0].split(" (")[0]
                    result[re.sub(r"[\s_-]", "", name).lower()] = name
        except OSError:
            continue
    return result


def font_name(name, flags):
    original = name.split("+")[-1]
    base = re.sub(r"[-,]?(bolditalic|boldoblique|bold|italic|oblique|regular|roman|psmt|mt)$",
                  "", original, flags=re.I)
    base = re.sub(r"[-,]?(bolditalic|boldoblique|bold|italic|oblique|regular|roman|psmt|mt)$",
                  "", base, flags=re.I).strip(" ,-")
    key = re.sub(r"[\s_-]", "", base).lower()
    aliases = {"helvetica": "Arial", "times": "Times New Roman", "timesnew": "Times New Roman",
               "timesnewroman": "Times New Roman", "courier": "Courier New"}
    known = installed_fonts().get(key)
    if known or key in aliases:
        return known or aliases[key], None if known else original
    fallback = "Courier New" if flags & 8 else "Times New Roman" if flags & 4 else "Arial"
    return fallback, original


def native_blocks(page, drawings, warn):
    blocks = []
    substitutions = set()
    source = page.get_text("rawdict", flags=pdf.TEXTFLAGS_RAWDICT & ~pdf.TEXT_PRESERVE_IMAGES)
    underlines = []
    for path in drawings:
        for part in path["items"]:
            if part[0] == "l" and abs(part[1].y - part[2].y) < 0.5:
                a, b = sorted((part[1], part[2]), key=lambda p: p.x)
                underlines.append((a.x, b.x, a.y))
    for block in source["blocks"]:
        lines = []
        for line in block["lines"]:
            spans = []
            angle = -math.degrees(math.atan2(line["dir"][1], line["dir"][0]))
            if abs(angle) > 3:
                warn("rotation", "Có chữ xoay. Cần đối chiếu góc và vị trí.", line["bbox"])
            for span in line["spans"]:
                if span.get("alpha", 255) == 0 or not span.get("char_flags", 24) & 24:
                    continue
                chars = [{"text": c["c"], "box": c["bbox"]} for c in span["chars"]
                         if valid_xml_char(c["c"])]
                if not any(c["text"].strip() for c in chars):
                    continue
                x0, _, x1, _ = span["bbox"]
                baseline = span["origin"][1]
                flags = span["flags"]
                font, replaced = font_name(span["font"], flags)
                if replaced:
                    substitutions.add((replaced, font))
                spans.append({
                    "text": "".join(c["text"] for c in chars), "chars": chars, "box": bounds(chars),
                    "baseline": baseline, "size": round(span["size"], 2),
                    "font": font, "color": f'{span["color"]:06x}', "bold": bool(flags & 16),
                    "italic": bool(flags & 2),
                    "underline": bool(span.get("char_flags", 0) & 2) or any(
                        baseline <= y <= baseline + max(4.5, span["size"]*.28) and
                        min(b, x1) - max(a, x0) >= (x1 - x0) * .8 for a, b, y in underlines),
                })
            if spans:
                groups = [[]]
                for span in spans:
                    if groups[-1] and abs(angle) < 3 and span["box"][0] - groups[-1][-1]["box"][2] > 28:
                        groups.append([])
                    groups[-1].append(span)
                lines.extend({"box": bounds(group), "spans": group, "angle": angle} for group in groups)
        groups = []
        for line in sorted(lines, key=lambda item: (item["box"][1], item["box"][0])):
            left, top, right, _ = line["box"]
            prior = next((g for g in reversed(groups)
                          if 0 < top-g[-1]["box"][1] <= 2.1*max(s["size"] for s in line["spans"])
                          and min(right,g[-1]["box"][2])-max(left,g[-1]["box"][0]) > 2), None)
            if prior is None:
                groups.append([line])
            else:
                prior.append(line)
        blocks.extend({"kind": "text", "lines": group, "box": bounds(group)} for group in groups)
    if substitutions:
        fonts = ", ".join(f"{a} → {b}" for a, b in sorted(substitutions)[:8])
        warn("font", "Thay phông chữ: " + fonts)
    return blocks


def valid_xml_char(char):
    return all(c in "\t\n\r" or 32 <= ord(c) <= 0xD7FF or
               0xE000 <= ord(c) <= 0xFFFD or ord(c) >= 0x10000 for c in char)


def long_runs(image, axis, minimum):
    if axis == "v":
        image = image.transpose(Image.Transpose.TRANSPOSE)
    width, height = image.size
    pixels = image.tobytes()
    pattern = re.compile(rb"\x00{" + str(minimum).encode() + rb",}")
    active, finished = [], []
    for y in range(height):
        next_active = []
        for match in pattern.finditer(pixels, y * width, (y + 1) * width):
            start, end = match.start() - y * width, match.end() - y * width
            prior = next((r for r in active if abs(r[0]-start) <= 4 and abs(r[1]-end) <= 4), None)
            if prior:
                active.remove(prior)
                next_active.append((prior[0], prior[1], prior[2], y))
            else:
                next_active.append((start, end, y, y))
        finished.extend(active)
        active = next_active
    finished.extend(active)
    for start, end, first, last in finished:
        if last - first <= 18:
            mid = (first + last) / 2
            yield (start, mid, end, mid) if axis == "h" else (mid, start, mid, end)


def scan_grid(image, origin, scale):
    gray = image.convert("L").point([0] * 130 + [255] * 126)
    horizontal = list(long_runs(gray, "h", max(20, round(scale*28))))
    vertical = list(long_runs(gray, "v", max(20, round(scale*18))))
    if len(horizontal) > 200 or len(vertical) > 200:
        return []
    def crosses(h, v):
        return h[0]-6 <= v[0] <= h[2]+6 and v[1]-6 <= h[1] <= v[3]+6

    def supported(line, cuts, start, end):
        return len(cuts) >= 2 and any(min(abs(pos-line[start]), abs(pos-line[end])) <= 7
                                     for pos in cuts)

    crossing_h = {h for h in horizontal if supported(h, [v[0] for v in vertical if crosses(h,v)], 0, 2)}
    crossing_v = {v for v in vertical if supported(v, [h[1] for h in crossing_h if crosses(h,v)], 1, 3)}
    if len(crossing_h) < 3 or len(crossing_v) < 3:
        return []
    return [tuple(value / scale + origin[i % 2] for i, value in enumerate(line))
            for line in sorted(crossing_h | crossing_v)]


def make_tables(page, drawings, scan_lines, detections, warn):
    with pdf.open() as synthetic:
        candidates = []
        try:
            if drawings:
                candidates.extend(page.find_tables(strategy="lines_strict", paths=drawings).tables)
            if scan_lines:
                target = synthetic.new_page(width=page.rect.width, height=page.rect.height)
                for x0, y0, x1, y1 in scan_lines:
                    target.draw_line((x0, y0), (x1, y1), width=0.5)
                # The grid finder ignores empty grids. Mark recognized positions
                # in the in-memory probe. These marks never go into the result.
                for _, groups in detections:
                    for words in groups.values():
                        for word in words:
                            x0, y0, x1, y1 = word["box"]
                            target.insert_text(((x0+x1)/2, (y0+y1)/2), "x", fontsize=2)
                candidates.extend(target.find_tables(strategy="lines", snap_tolerance=3,
                                                     join_tolerance=5, intersection_tolerance=4).tables)
            return table_models(candidates, drawings, warn)
        except (ValueError, RuntimeError, IndexError):
            warn("table", "Không xác định được lưới bảng. Nội dung nhận dạng được giữ ở dạng văn bản.")
            return []


def grid_axis(values):
    axis = []
    for value in sorted(values):
        if not axis or value - axis[-1] > 2:
            axis.append(value)
    return axis


def table_models(candidates, drawings, warn):
    tables = []
    for candidate in candidates:
        box = rect(candidate.bbox)
        if any(coverage(box, rect(t["box"])) > .8 for t in tables):
            continue
        cells = [rect(cell) for cell in candidate.cells]
        xs = grid_axis([v for c in cells for v in (c.x0, c.x1)])
        ys = grid_axis([v for c in cells for v in (c.y0, c.y1)])
        if not 2 <= len(xs) <= 33 or not 2 <= len(ys) <= 151 or len(xs)*len(ys) > 1800:
            continue
        items, occupied = [], set()
        for cell in sorted(cells, key=lambda c: -c.get_area()):
            col, end_col = [min(range(len(xs)), key=lambda i: abs(xs[i]-x)) for x in (cell.x0, cell.x1)]
            row, end_row = [min(range(len(ys)), key=lambda i: abs(ys[i]-y)) for y in (cell.y0, cell.y1)]
            grid = {(r, c) for r in range(row, end_row) for c in range(col, end_col)}
            if not grid or grid & occupied:
                continue
            occupied.update(grid)
            fill = next((p["fill"] for p in reversed(drawings) if p.get("fill") and
                         coverage(cell, p["rect"]) > .94 and p["rect"].get_area() <= box.get_area()*1.1), None)
            color = "".join(f"{min(255, max(0, round(x*255))):02x}" for x in fill) if fill else None
            stroke = next((p for p in reversed(drawings) if p.get("color") and
                           ((p["rect"] & cell).get_area() >= cell.get_area()*.8 or
                            (p["rect"].width >= cell.width*.8 and
                             min(abs(p["rect"].y0-cell.y0), abs(p["rect"].y0-cell.y1)) < 2 and
                             p["rect"].x0 <= cell.x0+2 and p["rect"].x1 >= cell.x1-2))), None)
            border = "".join(f"{round(x*255):02x}" for x in stroke["color"]) if stroke else "64748b"
            items.append({"row": row, "col": col, "rows": end_row-row, "cols": end_col-col,
                          "box": tuple(cell), "lines": [], "fill": color, "border": border,
                          "weight": (stroke.get("width") or .5) if stroke else .5})
        if items and len(occupied) >= (len(xs)-1)*(len(ys)-1)*.75:
            tables.append({"kind": "table", "box": tuple(box), "xs": xs, "ys": ys, "cells": items})
        else:
            warn("table", "Lưới bảng không rõ. Các ô chưa được tách tự động.", tuple(box))
    return tables


def assign_table_text(blocks, tables):
    cells = [cell for table in tables for cell in table["cells"]]
    if not cells:
        return blocks
    kept = []
    for block in blocks:
        rest = []
        for line in block["lines"]:
            if not any(rect(line["box"]).intersects(rect(cell["box"])) for cell in cells):
                rest.append(line)
                continue
            groups = defaultdict(list)
            for span in line["spans"]:
                chars = span.get("chars", [span])
                runs = []
                for char in chars:
                    idx = next((i for i, cell in enumerate(cells) if inside(char, rect(cell["box"]))), -1)
                    if not runs or runs[-1][0] != idx:
                        runs.append((idx, []))
                    runs[-1][1].append(char)
                for idx, run in runs:
                    if any(c["text"].strip() for c in run):
                        groups[idx].append({**span, "text": "".join(c["text"] for c in run),
                                            "box": bounds(run)})
            for idx, spans in groups.items():
                piece = {**line, "spans": spans, "box": bounds(spans)}
                if idx == -1:
                    rest.append(piece)
                else:
                    cells[idx]["lines"].append(piece)
        if rest:
            kept.append({**block, "lines": rest, "box": bounds(rest)})
    for cell in cells:
        lines = sorted(cell["lines"], key=lambda line: (line["box"][1], line["box"][0]))
        merged = []
        for line in lines:
            if merged and abs(line["box"][1] - merged[-1]["box"][1]) < 3:
                merged[-1]["spans"].extend(line["spans"])
                merged[-1]["spans"].sort(key=lambda s: s["box"][0])
                merged[-1]["box"] = bounds(merged[-1]["spans"])
            else:
                merged.append(line)
        cell["lines"] = merged
    return kept


def ocr_blocks(groups, native, table_cells, warn, printed_scan_height=0):
    blocks = defaultdict(list)
    for key, words in groups.items():
        kept = [w for w in words if valid_xml_char(w["text"]) and not any(
            coverage(rect(w["box"]), rect(n["box"])) > .6 or inside(w, rect(n["box"])) for n in native)]
        if not kept:
            continue
        splits = [[]]
        for word in kept:
            if splits[-1]:
                last = splits[-1][-1]
                gap = word["box"][0] - last["box"][2]
                outer_band = printed_scan_height and (word["box"][1] < printed_scan_height*.25 or
                                                       word["box"][1] > printed_scan_height*.80)
                if gap > 15 and (gap > (20 if outer_band else 45) or any(
                    inside(word, rect(c["box"])) != inside(last, rect(c["box"])) for c in table_cells)):
                    splits.append([])
            splits[-1].append(word)
        for i, segment in enumerate(splits):
            box = bounds(segment)
            height = sorted(rect(w["box"]).height for w in segment)[len(segment)//2]
            # The integrated OCR's invisible PDF font box includes extra
            # leading, unlike the CLI's tight word rectangles.
            embedded = all(w["conf"] is None for w in segment)
            size = min(72, max(5, height * (.85 if embedded else 1.12)))
            if printed_scan_height and not any(any(c.isalnum() for c in w["text"]) for w in segment):
                continue
            spans = [{"text": " ".join(w["text"] for w in segment), "box": box,
                      "chars": [{**w, "text": w["text"] + (" " if j+1 < len(segment) else "")}
                                for j, w in enumerate(segment)],
                      "baseline": box[3] - .18*height, "size": round(size, 2), "font": "Arial",
                      "color": "000000", "bold": False, "italic": False, "underline": False}]
            confidence = min((w["conf"] for w in segment if w["conf"] is not None), default=100)
            if confidence < LOW_CONFIDENCE:
                warn("low", f"OCR có từ dưới {LOW_CONFIDENCE}% (thấp nhất {max(0, confidence):.1f}%).", box)
            blocks[(*key[:2], i)].append({"box": box, "spans": spans, "angle": 0})
    return [{"kind": "text", "box": bounds(lines), "lines": lines, "ocr": True}
            for lines in blocks.values()]


def scan_text_image(image):
    """Make neutral print easier to separate from tinted paper and ink."""
    red, green, blue = image.split()
    neutral = ImageChops.lighter(ImageChops.lighter(red, green), blue)
    gray = ImageOps.autocontrast(neutral, cutoff=(1, 2))
    return gray.point([min(255, max(0, round((x-62)*255/145))) for x in range(256)]).convert("RGB")


def separate_colored_marks(groups, image, origin, scale, page):
    """Reject colored seal fragments misread as isolated print near a page edge."""
    red, green, blue = image.split()
    high = ImageChops.lighter(ImageChops.lighter(red, green), blue)
    low = ImageChops.darker(ImageChops.darker(red, green), blue)
    saturation = ImageChops.subtract(high, low)
    colored = ImageChops.multiply(saturation.point([0]*50+[255]*206), low.point([255]*175+[0]*81))
    neutral = ImageChops.multiply(saturation.point([255]*34+[0]*222), high.point([255]*167+[0]*89))
    red_ink = ImageChops.multiply(ImageChops.subtract(red,green).point([0]*42+[255]*214),
                                  ImageChops.subtract(red,blue).point([0]*28+[255]*228))
    result = {}
    for key, words in groups.items():
        kept = []
        for word in words:
            x0,y0,x1,y1 = word["box"]
            box = tuple(round((v-origin[i%2])*scale) for i,v in enumerate(word["box"]))
            near_stamp = False
            if y0 > page.height*.45 and x0 > page.width*.57:
                padding = round(12*scale)
                neighborhood = (box[0]-padding, box[1]-padding, box[2]+padding, box[3]+padding)
                patch = red_ink.crop(neighborhood)
                near_stamp = patch.histogram()[255] > patch.width*patch.height*.028
            if near_stamp or y0 > page.height*.76 or x1 > page.width*.93 or x0 < page.width*.04:
                color_count = colored.crop(box).histogram()[255]
                dark_count = neutral.crop(box).histogram()[255]
                area = max(1, (box[2]-box[0])*(box[3]-box[1]))
                if color_count > area*.035 and color_count > dark_count*2.5:
                    continue
            kept.append(word)
        if kept:
            result[key] = kept
    return result


def scan_illustrations(image, origin, scale, groups, tables):
    """Keep separate color/dense raster areas outside text and grid regions."""
    small = image.copy()
    small.thumbnail((256, 256))
    ratio_x, ratio_y = image.width / small.width, image.height / small.height
    pixels = small.load()
    occupied = [rect(w["box"]) for words in groups.values() for w in words]
    occupied.extend(rect(t["box"]) for t in tables)
    tiles = set()
    step = 4
    for y in range(0, small.height-step+1, step):
        for x in range(0, small.width-step+1, step):
            px = [pixels[a,b] for a in range(x,x+step) for b in range(y,y+step)]
            if sum(max(p)-min(p)>26 or max(p)<150 for p in px) < 10:
                continue
            point = (origin[0]+(x+2)*ratio_x/scale, origin[1]+(y+2)*ratio_y/scale)
            if not any(r.contains(point) for r in occupied):
                tiles.add((x, y))
    results = []
    while tiles:
        component = [tiles.pop()]
        for x, y in component:
            for p in ((x+step,y),(x-step,y),(x,y+step),(x,y-step)):
                if p in tiles:
                    tiles.remove(p)
                    component.append(p)
        if len(component) < 10:
            continue
        x0, y0 = min(x for x,y in component), min(y for x,y in component)
        x1, y1 = max(x for x,y in component)+step, max(y for x,y in component)+step
        if min(x1-x0,y1-y0) < 12 or (x1-x0)*(y1-y0) > small.width*small.height*.36:
            continue
        crop = (round(x0*ratio_x),round(y0*ratio_y),round(x1*ratio_x),round(y1*ratio_y))
        box = tuple(value/scale+origin[i%2] for i,value in enumerate(crop))
        if any(rect(box).intersects(r) and (rect(box)&r).get_area() > 1 for r in occupied):
            continue
        results.append((box, image.crop(crop)))
    return results[:20]


def likely_seal_crop(box, image, page):
    if box[1] < page.height*.77 or box[2]-box[0] > page.width*.25:
        return False
    sample = image.copy()
    sample.thumbnail((120, 120))
    red = sum(r > g+36 and r > b+24 for r,g,b in sample.get_flattened_data())
    return red > sample.width*sample.height*.055


def mask_text(image, groups, origin, scale):
    draw = ImageDraw.Draw(image)
    for word in [w for words in groups.values() for w in words]:
        x0, y0, x1, y1 = [round((v-origin[i%2])*scale) for i,v in enumerate(word["box"])]
        above = image.crop((max(0,x0-2), max(0,y0-5), min(image.width,x1+2), max(1,y0-1)))
        color = tuple(round(c) for c in ImageStat.Stat(above).median)
        draw.rectangle((x0-2,y0-2,x1+2,y1+2), fill=color)


def scan_table_color(tables, image, origin, scale):
    for table in tables:
        for cell in table["cells"]:
            if cell["fill"]:
                continue
            x0,y0,x1,y1 = cell["box"]
            margin = min(3, (x1-x0)/5, (y1-y0)/5)
            pixels = tuple(round((v-origin[i%2])*scale) for i,v in enumerate(
                (x0+margin,y0+margin,x1-margin,y1-margin)))
            color = ImageStat.Stat(image.crop(pixels)).median
            if min(color) < 244 and min(color) >= 150:
                cell["fill"] = "".join(f"{round(c):02x}" for c in color)


def reading_order(items, width, depth=0):
    """Split at real gutters, reading each column before moving to the next."""
    if len(items) < 2 or depth >= 6:
        return sorted(items, key=lambda b: (b["box"][1], b["box"][0]))
    intervals = sorted((i["box"][0], i["box"][2]) for i in items)
    end, gaps = intervals[0][1], []
    for start, right in intervals[1:]:
        if start-end >= max(18, width*.025):
            gaps.append((start-end, (start+end)/2))
        end = max(end, right)
    for _, mid in sorted(gaps, reverse=True):
        left = [i for i in items if i["box"][2] <= mid]
        right = [i for i in items if i["box"][0] >= mid]
        if left and right and min(bounds(left)[3],bounds(right)[3]) > max(bounds(left)[1],bounds(right)[1]):
            return reading_order(left,width,depth+1) + reading_order(right,width,depth+1)
    barriers = sorted([i for i in items if i["box"][2]-i["box"][0] >= width*.67],
                      key=lambda i: i["box"][1])
    ordered, rest = [], items[:]
    for item in barriers:
        before = [i for i in rest if i is not item and i["box"][3] <= item["box"][1]+2]
        ordered.extend(reading_order(before, width, depth+1))
        ordered.append(item)
        rest = [i for i in rest if i is not item and i not in before]
    if barriers:
        return ordered + reading_order(rest, width, depth+1)
    vertical = sorted((i["box"][1], i["box"][3]) for i in items)
    end, breaks = vertical[0][1], []
    for start, bottom in vertical[1:]:
        if start-end >= 13:
            breaks.append((start-end, (start+end)/2))
        end = max(end, bottom)
    if breaks:
        _, cut = max(breaks)
        return (reading_order([i for i in items if i["box"][3] < cut], width, depth+1) +
                reading_order([i for i in items if i["box"][1] > cut], width, depth+1))
    return sorted(items, key=lambda b: (b["box"][1], b["box"][0]))


def extract_page(page, get_ocr, lang):
    warnings = []
    ocr_warning_codes = set()

    def warn(code, message, box=None):
        if len(warnings) < 100:
            warnings.append({"code": code, "text": message, "box": list(box) if box else None})

    def report_ocr(code, message):
        if code not in ocr_warning_codes:
            ocr_warning_codes.add(code)
            warn(code, message)

    if page.rotation:
        page.remove_rotation()
    page_box = page.rect
    drawings = page.get_drawings()
    native = native_blocks(page, drawings, warn)
    direct = [s for b in native for l in b["lines"] for s in l["spans"]]
    if any("\ufffd" in span["text"] for span in direct):
        warn("encoding", "Lớp chữ PDF chứa ký tự không giải mã được (�).")
    regions = []
    for info in page.get_image_info():
        box = rect(info["bbox"]) & page_box
        if min(box.width, box.height) >= 8 and not any(coverage(box, old) > .98 for old in regions):
            regions.append(box)
    scan = any(coverage(page_box, b) > .60 for b in regions)
    candidates = [page_box] if scan or not native else regions
    if len(candidates) > 32:
        warn("layout", "Trang có nhiều vùng ảnh. OCR được gộp thành một lượt trên toàn trang.")
        candidates = [page_box]
    scan_lines, detections, illustration_data, ink_art = [], [], [], []
    for region in candidates:
        if not scan and native and min(region.width, region.height) < 12:
            continue
        image, origin, scale = raster(page, region, OCR_SCALE)
        full = coverage(page_box, region) > .95
        command = get_ocr()
        if command is None and not any(w["code"] == "ocr-fallback" for w in warnings):
            warn("ocr-fallback", "Dùng Tesseract tích hợp. Cần kiểm tra chữ OCR vì chế độ này không chấm độ tin cậy từng từ.")
        groups = recognize(image, origin, scale, command, lang, 3 if full else 11, report_ocr)
        if full and scan and len(groups) >= 20:
            groups = separate_colored_marks(groups, image, origin, scale, page_box)
        elif full and scan and not groups:
            groups = recognize(scan_text_image(image), origin, scale, command, lang, 3, report_ocr)
        if full and scan and len(groups) < 20:
            for edge, band in enumerate((rect((0,0,page_box.width,page_box.height*.065)),
                                         rect((0,page_box.height*.9,page_box.width,page_box.height)))):
                detail, offset, factor = raster(page, band, OCR_SCALE)
                extra = recognize(detail, offset, factor, command, lang, 11, report_ocr)
                old = [rect(w["box"]) for words in groups.values() for w in words]
                for key, words in extra.items():
                    fresh = [w for w in words if not any(coverage(rect(w["box"]),
                             rect((box.x0-2,box.y0-2,box.x1+2,box.y1+2))) > .45 for box in old)]
                    if fresh:
                        groups[(f"edge{edge}-"+key[0], *key[1:])] = fresh
        if scan or not native or any(groups.values()):
            scan_lines.extend(scan_grid(image, origin, scale))
        if scan and not native and full and len(groups) >= 15:
            crops, groups = preserve_ink(image, origin, scale, page_box, groups)
            ink_art.extend(crops)
        detections.append((region, groups))
        if full and (scan or not native):
            illustration_data.append((image, origin, scale, groups))

    tables = make_tables(page, drawings, scan_lines, detections, warn)
    for image, origin, scale, _ in illustration_data:
        scan_table_color(tables, image, origin, scale)
    cells = [c for table in tables for c in table["cells"]]
    recognized = []
    dense_scan = scan and not native and sum(len(groups) for _, groups in detections) >= 15
    for _, groups in detections:
        recognized.extend(ocr_blocks(groups, direct + [
            s for b in recognized for l in b["lines"] for s in l["spans"]], cells, warn,
            page_box.height if dense_scan else 0))
    if recognized:
        warn("ocr", "Văn bản OCR giữ nguyên kết quả nhận dạng. Phông, kiểu chữ và ngắt đoạn là ước lượng.")
    if scan:
        warn("scan", "Ảnh nền scan đã được tách khỏi nội dung. Kiểm tra minh họa, đường kẻ và dấu tiếng Việt.")
    if dense_scan:
        warn("ink", "Cần đối chiếu chữ viết tay, chữ trên dấu và vùng chữ bị dấu che với PDF gốc.")
    if (scan or recognized) and scan_lines and tables:
        warn("table", "Bảng từ ảnh được dựng theo lưới phát hiện. Hãy kiểm tra ô gộp và đường viền.")
    pictures = []
    for i, region in enumerate(regions):
        if scan and coverage(page_box, region) > .60:
            continue
        matching = [groups for box, groups in detections if coverage(region, box) > .98]
        words = {key: [w for w in group if inside(w, region)]
                 for groups in matching for key, group in groups.items()}
        if not any(words.values()) and not any(rect(s["box"]).intersects(region) for s in direct):
            if (scan or not native) and coverage(page_box, region) > .35:
                continue
        image, origin, scale = raster(page, region, 2.2)
        if any(words.values()):
            mask_text(image, words, origin, scale)
            warn("image", "Đã tách chữ OCR khỏi ảnh. Kiểm tra nền và vị trí chữ tại vùng này.", tuple(region))
        pictures.append((tuple(region), image))
    for image, origin, scale, groups in illustration_data:
        pictures.extend((box, crop) for box, crop in scan_illustrations(image, origin, scale, groups, tables)
                        if (not dense_scan or not likely_seal_crop(box, crop, page_box)) and
                        not any(coverage(rect(box), rect(area)) > .2 for area, _, _ in ink_art))
    ink_labels = {}
    for box, crop, kind in ink_art:
        label = ("Dòng có chữ viết tay từ PDF gốc" if kind == "handwriting"
                 else "Con dấu và chữ ký từ PDF gốc" if kind == "seal" else "Chữ ký từ PDF gốc")
        warn("ink-preserved", label + " được giữ dưới dạng ảnh cắt để đối chiếu.", box)
        ink_labels[box] = label
        pictures.append((box, crop))
    if not native and not recognized:
        warn("empty", "Không tìm thấy chữ. Kiểm tra ngôn ngữ OCR và chất lượng ảnh.")
    if any(p.get("fill") is not None or any(i[0] in {"c", "qu"} for i in p["items"]) for p in drawings):
        warn("vector", "Kiểm tra hình vector, nền và đường kẻ ngoài bảng.")
    text = assign_table_text(native + recognized, tables)
    media = {}
    image_blocks = []
    for i, (box, image) in enumerate(pictures):
        media[i] = encoded(image, "JPEG", quality=90, optimize=True)
        image_blocks.append({"kind": "image", "box": box, "asset": i,
                             "label": ink_labels.get(box, "Ảnh từ PDF"), "preserved_ink": box in ink_labels})
    items = reading_order(text + tables, page_box.width) + image_blocks
    margin_box = bounds(items) if items else (36, 36, page_box.width-36, page_box.height-36)
    return {"width": page_box.width, "height": page_box.height, "items": items,
            "margins": (max(0,margin_box[0]), max(0,margin_box[1]),
                        max(0,page_box.width-margin_box[2]), max(0,page_box.height-margin_box[3])),
            "warnings": warnings, "flow_scan": dense_scan and not tables,
            "mode": "Scan" if scan or not native else "Hỗn hợp" if recognized else "Văn bản",
            "media": media}
