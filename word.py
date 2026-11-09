"""Write editable Word objects positioned in page coordinates."""

from functools import lru_cache
from io import BytesIO
from pathlib import Path
from statistics import median
import os
import re

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
from lxml import etree
from PIL import ImageFont

VML = "urn:schemas-microsoft-com:vml"
OFFICE = "urn:schemas-microsoft-com:office:office"
PAGE_NUMBER = re.compile(r"^(?:(Trang|Page)\s+)?(\d+)$", re.I)
LIST_MARKER = re.compile(r"^\s*([•◦▪‣]|\d+[.)])\s+")
MIN_PAGE_PT = 36
MAX_PAGE_PT = 1584


def element(tag, attrs=None):
    node = OxmlElement(tag)
    for key, value in (attrs or {}).items():
        node.set(qn(key), str(value))
    return node


def blank(paragraph, height=1):
    f = paragraph.paragraph_format
    f.space_before = Pt(0)
    f.space_after = Pt(0)
    f.line_spacing = Pt(max(1, height))
    f.keep_with_next = False
    f.keep_together = False
    f.widow_control = False


def line_text(line):
    return "".join(span["text"] for span in line["spans"])


def classify(block, page, page_index):
    if block["kind"] != "text":
        return
    top, bottom = block["box"][1], block["box"][3]
    text = line_text(block["lines"][0]).strip()
    match = PAGE_NUMBER.fullmatch(text)
    if len(block["lines"]) == 1 and bottom >= page["height"]*.90 and match:
        if int(match[2]) == page_index + 1:
            block["page_field"] = match[1] or ""
    if bottom < page["height"]*.066 and max(s["size"] for s in block["lines"][0]["spans"]) <= 12:
        block["part"] = "header"
    elif top > page["height"]*.91 or "page_field" in block:
        block["part"] = "footer"


def line_position(line):
    largest = max(line["spans"], key=lambda s: s["size"])
    return largest["baseline"] - .92 * largest["size"]


def add_numbering(doc, paragraph, marker, span):
    numbering = doc.part.numbering_part.element
    existing = numbering.findall(qn("w:abstractNum"))
    aid = 1 + max((int(x.get(qn("w:abstractNumId"))) for x in existing), default=0)
    abstract = element("w:abstractNum", {"w:abstractNumId": aid})
    lvl = element("w:lvl", {"w:ilvl": 0})
    numeric = marker[0].isdigit()
    lvl.append(element("w:start", {"w:val": marker[:-1] if numeric else 1}))
    lvl.append(element("w:numFmt", {"w:val": "decimal" if numeric else "bullet"}))
    lvl.append(element("w:suff", {"w:val": "space"}))
    lvl.append(element("w:lvlText", {"w:val": "%1" + marker[-1] if numeric else marker}))
    lvl.append(element("w:lvlJc", {"w:val": "left"}))
    properties = element("w:rPr")
    properties.append(element("w:rFonts", {"w:ascii": span["font"], "w:hAnsi": span["font"]}))
    properties.append(element("w:sz", {"w:val": round(span["size"]*2)}))
    lvl.append(properties)
    abstract.append(lvl)
    first_num = numbering.find(qn("w:num"))
    if first_num is None:
        numbering.append(abstract)
    else:
        first_num.addprevious(abstract)
    number = numbering.add_num(aid)
    pr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    pr.get_or_add_ilvl().val = 0
    pr.get_or_add_numId().val = number.numId


@lru_cache(maxsize=48)
def measure_font(name, bold, italic, size):
    family = {"Arial":"arial", "Times New Roman":"times", "Courier New":"cour",
              "Calibri":"calibri"}.get(name)
    if not family:
        return None
    if family in {"times", "cour"}:
        suffix = "bi" if bold and italic else "bd" if bold else "i" if italic else ""
    elif family == "calibri":
        suffix = "z" if bold and italic else "b" if bold else "i" if italic else ""
    else:
        suffix = "bi" if bold and italic else "bd" if bold else "i" if italic else ""
    folder = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    try:
        return ImageFont.truetype(str(folder / (family + suffix + ".ttf")), max(1, round(size*4)))
    except OSError:
        return None


def add_run(paragraph, span, content, fit=True, scale_limits=(70, 150)):
    run = paragraph.add_run(content)
    font = run.font
    font.name = span["font"]
    font.size = Pt(max(1, min(600, span["size"])))
    font.color.rgb = RGBColor.from_string(span["color"].upper())
    font.bold, font.italic, font.underline = span["bold"], span["italic"], span["underline"]
    pr = run._r.get_or_add_rPr()
    pr.rFonts.set(qn("w:eastAsia"), span["font"])
    pr.rFonts.set(qn("w:cs"), span["font"])
    measured = measure_font(span["font"], span["bold"], span["italic"], span["size"])
    if fit and measured and content.strip():
        actual = measured.getlength(content) / 4
        intended = span["box"][2] - span["box"][0]
        ratio = round(100 * intended / max(1, actual))
        if scale_limits[0] <= ratio <= scale_limits[1] and ratio != 100:
            scale = element("w:w", {"w:val": ratio})
            size = pr.find(qn("w:sz"))
            size.addprevious(scale)
            span["stretch"] = ratio
    return run


def add_page_field(paragraph, field_prefix, span, index):
    if field_prefix:
        add_run(paragraph, span, field_prefix + " ", fit=False)
    run = add_run(paragraph, span, "", fit=False)
    run._r.append(element("w:fldChar", {"w:fldCharType": "begin"}))
    instruction = element("w:instrText")
    instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instruction.text = " PAGE "
    run._r.append(instruction)
    run._r.append(element("w:fldChar", {"w:fldCharType": "separate"}))
    add_run(paragraph, span, str(index+1), fit=False)
    run = paragraph.add_run()
    run._r.append(element("w:fldChar", {"w:fldCharType": "end"}))


def add_lines(doc, container, lines, left, top, width, first=None, page_field=None, index=0):
    if not lines:
        if first is not None:
            blank(first)
        return
    lines = sorted(lines, key=lambda line: line["box"][1])
    centered = len(lines) > 1 and all(abs((l["box"][0]+l["box"][2])/2 -
                                          (lines[0]["box"][0]+lines[0]["box"][2])/2) < 3
                                      for l in lines) and max(l["box"][0] for l in lines) - left > 5
    right_aligned = len(lines) > 1 and max(l["box"][2] for l in lines) - min(l["box"][2] for l in lines) < 3
    common_left = min(l["box"][0] for l in lines)
    initial_indent = max(0, lines[0]["box"][0] - common_left)
    paragraph = first if first is not None else container.add_paragraph()
    deltas = [line_position(b)-line_position(a) for a,b in zip(lines, lines[1:])]
    single_paragraph = (bool(deltas) and min(deltas) > 4 and max(deltas)-min(deltas) < 3
                        and all(abs(l["angle"]) < 3 and not LIST_MARKER.match(line_text(l)) for l in lines)
                        and (centered or right_aligned or
                             all(abs(l["box"][0]-common_left) < 3 for l in lines[1:])))
    for i, line in enumerate(lines):
        spans = line["spans"]
        size = max(s["size"] for s in spans)
        height = size*1.23
        text = line_text(line)
        match = LIST_MARKER.match(text)
        heading = size >= 17 and len(lines) == 1 and len(text) <= 140 and first is None
        if i and single_paragraph:
            paragraph.add_run().add_break()
            for span in spans:
                add_run(paragraph, span, span["text"])
            continue
        if i:
            paragraph = container.add_paragraph()
        if heading:
            paragraph.style = "Heading 1" if size >= 23 else "Heading 2"
        blank(paragraph, sum(deltas)/len(deltas) if single_paragraph else height)
        fmt = paragraph.paragraph_format
        pos = line_position(line)
        previous = line_position(lines[i-1]) + max(s["size"] for s in lines[i-1]["spans"])*1.23 if i else top
        fmt.space_before = Pt(max(0, pos-previous))
        if centered:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif right_aligned:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            fmt.right_indent = Pt(max(0, left+width-max(l["box"][2] for l in lines)-2))
        else:
            indent = max(0, line["box"][0]-left)
            if i == 0 and len(lines) > 1 and initial_indent > 2:
                fmt.left_indent = Pt(max(0, common_left-left))
                fmt.first_line_indent = Pt(initial_indent)
            else:
                fmt.left_indent = Pt(indent)
        if page_field is not None:
            add_page_field(paragraph, page_field, spans[0], index)
            continue
        skip = 0
        if match:
            add_numbering(doc, paragraph, match[1], spans[0])
            skip = match.end()
        for span in spans:
            content = span["text"]
            removed = min(len(content), skip)
            skip -= removed
            if content[removed:]:
                add_run(paragraph, span, content[removed:], fit=not removed and abs(line["angle"]) < 3)


class TextBox:
    def __init__(self, doc, anchor, shape_id, box, layer, angle=0):
        self.doc = doc
        x0, y0, x1, y1 = box
        run = anchor.add_run()
        pict = element("w:pict")
        run._r.append(pict)
        shape = etree.SubElement(pict, f"{{{VML}}}rect", nsmap={"v": VML, "o": OFFICE})
        shape.set("id", f"pdfObject{shape_id}")
        shape.set(f"{{{OFFICE}}}allowoverlap", "t")
        shape.set("filled", "f")
        shape.set("stroked", "f")
        style = (f"position:absolute;margin-left:{x0:.2f}pt;margin-top:{y0:.2f}pt;"
                 f"width:{max(2,x1-x0):.2f}pt;height:{max(2,y1-y0):.2f}pt;z-index:{layer};"
                 "mso-position-horizontal-relative:page;mso-position-vertical-relative:page;"
                 "mso-wrap-style:none")
        if abs(angle) >= 3:
            style += f";rotation:{-angle:.1f}"
        shape.set("style", style)
        textbox = etree.SubElement(shape, f"{{{VML}}}textbox", inset="0,0,0,0")
        textbox.set("style", "mso-fit-shape-to-text:f")
        self.content = element("w:txbxContent")
        textbox.append(self.content)

    def add_paragraph(self):
        node = element("w:p")
        self.content.append(node)
        return Paragraph(node, self.doc._body)

    def add_table(self, rows, cols):
        table = self.doc.add_table(rows=rows, cols=cols)
        self.content.append(table._tbl)
        return table


def add_table(doc, box, model):
    xs, ys = model["xs"], model["ys"]
    table = box.add_table(len(ys)-1, len(xs)-1)
    table.autofit = False
    tblpr = table._tbl.tblPr
    size = tblpr.find(qn("w:tblW"))
    size.set(qn("w:type"), "dxa")
    size.set(qn("w:w"), str(round((xs[-1]-xs[0])*20)))
    margins = element("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        margins.append(element("w:" + side, {"w:w": 0, "w:type": "dxa"}))
    tblpr.append(margins)
    for j, column in enumerate(table.columns):
        column.width = Pt(xs[j+1]-xs[j])
    for i, row in enumerate(table.rows):
        row.height = Pt(ys[i+1]-ys[i])
        row._tr.get_or_add_trPr().append(element("w:cantSplit"))
        for j, cell in enumerate(row.cells):
            cell.width = Pt(xs[j+1]-xs[j])
            blank(cell.paragraphs[0])
    for item in model["cells"]:
        row, col = item["row"], item["col"]
        cell = table.cell(row, col)
        if item["rows"] > 1 or item["cols"] > 1:
            cell = cell.merge(table.cell(row+item["rows"]-1, col+item["cols"]-1))
        p = cell._tc.get_or_add_tcPr()
        border = element("w:tcBorders")
        for edge in ("top", "left", "bottom", "right"):
            border.append(element("w:"+edge, {"w:val":"single", "w:color":item["border"],
                                               "w:sz": max(2, min(96, round(item["weight"]*8)))}))
        p.append(border)
        if item["fill"]:
            p.append(element("w:shd", {"w:val":"clear", "w:fill":item["fill"]}))
        left, top, right, _ = item["box"]
        add_lines(doc, cell, item["lines"], left, top, right-left, first=cell.paragraphs[0])
    blank(box.add_paragraph())


def add_picture(anchor, data, box, layer, label="Ảnh từ PDF", foreground=False):
    left, top, right, bottom = box
    run = anchor.add_run()
    picture = run.add_picture(BytesIO(data), width=Pt(right-left), height=Pt(bottom-top))
    inline = picture._inline
    inline.find(qn("wp:docPr")).set("descr", label)
    floating = element("wp:anchor")
    for name, value in {"distT": 0, "distB": 0, "distL": 0, "distR": 0, "simplePos": 0,
                        "relativeHeight": layer, "behindDoc": int(not foreground), "locked": 0,
                        "layoutInCell": 1, "allowOverlap": 1}.items():
        floating.set(name, str(value))
    simple = element("wp:simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    floating.append(simple)
    for axis, value in (("H",left), ("V",top)):
        pos = element("wp:position" + axis)
        pos.set("relativeFrom", "page")
        offset = element("wp:posOffset")
        offset.text = str(Pt(value))
        pos.append(offset)
        floating.append(pos)
    floating.append(inline.find(qn("wp:extent")))
    floating.append(element("wp:wrapNone"))
    for child in list(inline):
        floating.append(child)
    inline.getparent().replace(inline, floating)


def scan_rows(page):
    lines = [line for block in page["items"] if block["kind"] == "text" for line in block["lines"]]
    if not lines:
        return []
    body = [s["size"] for l in lines for s in l["spans"]
            if len(s["text"]) >= 32 and l["box"][2]-l["box"][0] >= page["width"]*.25]
    usual_size = min(15, max(8, median(body))) if body else 11
    rows = []
    for line in sorted(lines, key=lambda l: (l["spans"][0]["baseline"], l["box"][0])):
        left, y0, right, y1 = line["box"]
        # An isolated tiny component at the scan edge is usually paper or a seal fragment.
        if right-left < 26 and (right > page["width"]*.96 or left < page["width"]*.04):
            continue
        for span in line["spans"]:
            span["font"] = "Times New Roman"
            short = len(span["text"]) < 78 and span["text"].isupper()
            span["bold"] = short or bool(re.match(r"^(?:[IVX]+\.|\d+\.)\s", span["text"]))
            ceiling = usual_size*1.42 if short else usual_size*1.16
            span["size"] = round(max(usual_size*.75, min(ceiling, span["size"]*1.12)), 2)
        base = median([s["baseline"] for s in line["spans"]])
        if rows and base - rows[-1]["base"] <= 6 and all(
                (right + 1 < old["box"][0] or left - 1 > old["box"][2])
                and min(y1, old["box"][3]) - max(y0, old["box"][1]) >
                .28 * min(y1-y0, old["box"][3]-old["box"][1])
                for old in rows[-1]["lines"]):
            rows[-1]["lines"].append(line)
        else:
            rows.append({"base":base, "lines":[line]})
    for row in rows:
        row["lines"].sort(key=lambda line: line["box"][0])
        row["size"] = max(s["size"] for l in row["lines"] for s in l["spans"])
        row["top"] = max(0, row["base"] - .96*row["size"])
    rows.sort(key=lambda row: row["top"])
    return rows


def same_scan_paragraph(previous, row, left):
    if len(previous["lines"]) != 1 or len(row["lines"]) != 1:
        return False
    old, line = previous["lines"][0], row["lines"][0]
    text = line_text(line)
    return (abs(line["box"][0]-left) < 5 and not any(s["bold"] for s in line["spans"])
            and row["size"]*.78 <= row["top"]-previous["top"] <= row["size"]*1.65
            and old["box"][2]-old["box"][0] > 170 and not LIST_MARKER.match(text)
            and not re.match(r"^(?:[-+]|Căn cứ\b|Điều \d)", text))


def add_scan_page(doc, section, page):
    rows = scan_rows(page)
    if not rows:
        blank(doc.add_paragraph())
        return
    width, height = page["width"], page["height"]
    left = max(0, min(line["box"][0] for row in rows for line in row["lines"])-2)
    long_line_starts = sorted(line["box"][0] for row in rows for line in row["lines"]
                              if line["box"][2]-line["box"][0] > width*.45)
    body_left = long_line_starts[len(long_line_starts)//6] if long_line_starts else left
    first_top = rows[0]["top"]
    section.top_margin = Pt(first_top)
    section.bottom_margin = Pt(8)
    section.left_margin = Pt(left)
    section.right_margin = Pt(4)
    groups = []
    for row in rows:
        if groups and same_scan_paragraph(groups[-1][-1], row, body_left):
            groups[-1].append(row)
        else:
            groups.append([row])
    previous_end = first_top
    for group_index, group in enumerate(groups):
        row = group[0]
        paragraph = doc.add_paragraph()
        top = row["top"]
        last = group[-1]
        if len(group) > 1:
            leading = (last["top"]-top)/(len(group)-1)
        else:
            next_top = groups[group_index+1][0]["top"] if group_index+1 < len(groups) else height-12
            leading = min(row["size"]*1.15, max(row["size"]*.94, next_top-top))
        leading = max(4, leading)
        blank(paragraph, leading)
        fmt = paragraph.paragraph_format
        fmt.space_before = Pt(max(0, top-previous_end))
        start = row["lines"][0]["box"][0]
        if len(group) > 1:
            x = min(line["box"][0] for r in group for line in r["lines"])
            fmt.left_indent = Pt(max(0, x-left))
            fmt.first_line_indent = Pt(max(0, start-x))
        else:
            fmt.left_indent = Pt(max(0, start-left))
        for j, current in enumerate(group):
            if j:
                paragraph.add_run().add_break()
            for i, line in enumerate(current["lines"]):
                if i:
                    fmt.tab_stops.add_tab_stop(Pt(line["box"][0]-left))
                    paragraph.add_run("\t")
                for span in line["spans"]:
                    add_run(paragraph, span, span["text"], scale_limits=(45, 190))
        previous_end = max(top, previous_end) + leading*len(group)
    for block in page["items"]:
        if block["kind"] == "image":
            add_picture(paragraph, page["media"][block["asset"]], block["box"], 1,
                        block.get("label", "Ảnh từ PDF"), block.get("preserved_ink", False))


def fit_word_page(page):
    width, height = page["width"], page["height"]
    if MIN_PAGE_PT <= min(width, height) and max(width, height) <= MAX_PAGE_PT:
        return

    # Preserve the aspect ratio. Add space only if the short edge still falls
    # below the minimum after the long edge has reached the maximum.
    scale = min(max(1, MIN_PAGE_PT / min(width, height)), MAX_PAGE_PT / max(width, height))
    scaled_width, scaled_height = width * scale, height * scale
    padded = min(scaled_width, scaled_height) < MIN_PAGE_PT - 1e-6
    page["source_size"] = {"width": width, "height": height}
    page["word_scale"] = scale
    page["width"] = min(MAX_PAGE_PT, max(MIN_PAGE_PT, scaled_width))
    page["height"] = min(MAX_PAGE_PT, max(MIN_PAGE_PT, scaled_height))
    page["flow_scan"] = False

    def scaled(values):
        return tuple(value * scale for value in values)

    page["margins"] = scaled(page["margins"])
    for block in page["items"]:
        block["box"] = scaled(block["box"])
        lines = block.get("lines", [])
        if block["kind"] == "table":
            block["xs"], block["ys"] = scaled(block["xs"]), scaled(block["ys"])
            for cell in block["cells"]:
                cell["box"] = scaled(cell["box"])
                cell["weight"] *= scale
                lines.extend(cell["lines"])
        for line in lines:
            line["box"] = scaled(line["box"])
            for span in line["spans"]:
                span["box"] = scaled(span["box"])
                span["baseline"] *= scale
                span["size"] *= scale

    message = (f"Kích thước trang và nội dung trong Word bằng {scale:.3%} kích thước PDF gốc "
               "để vừa khổ giấy Word.") if scale != 1 else ""
    if padded:
        message += " Đã thêm khoảng trắng ở cạnh ngắn để vừa khổ giấy Word."
    page["warnings"].insert(0, {"code": "word-page-size", "text": message.strip(), "box": None})


def export_word(pages):
    doc = Document()
    doc.core_properties.title = "Bản chuyển đổi PDF"
    doc.core_properties.author = ""
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)
    shape_id = 0
    for index, page in enumerate(pages):
        fit_word_page(page)
        section = doc.sections[0] if index == 0 else doc.add_section(WD_SECTION_START.NEW_PAGE)
        if index:
            blank(doc.paragraphs[-1])
        width, height = page["width"], page["height"]
        section.orientation = WD_ORIENT.LANDSCAPE if width > height else WD_ORIENT.PORTRAIT
        section.page_width, section.page_height = Pt(width), Pt(height)
        left, top, right, bottom = page["margins"]
        section.left_margin, section.right_margin = Pt(min(left, width*.4)), Pt(min(right, width*.4))
        section.top_margin, section.bottom_margin = Pt(min(top, height*.4)), Pt(min(bottom, height*.4))
        section.header_distance = Pt(0)
        section.footer_distance = Pt(0)
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        if page.get("flow_scan"):
            add_scan_page(doc, section, page)
            continue
        anchor = doc.add_paragraph()
        blank(anchor)
        for block in page["items"]:
            classify(block, page, index)
            x0, y0, x1, y1 = block["box"]
            shape_id += 1
            target = getattr(section, block["part"]).paragraphs[0] if "part" in block else anchor
            blank(target)
            if block["kind"] == "text":
                lines = block["lines"]
                top = max(0, min(line_position(line) for line in lines))
                angle = lines[0]["angle"]
                extra = 7
                if abs(angle) >= 3:
                    box_width = max(x1-x0, y1-y0) + extra
                    box_height = max(s["size"] for l in lines for s in l["spans"])*1.7
                    xmid, ymid = (x0+x1)/2, (y0+y1)/2
                    bounds = (xmid-box_width/2, ymid-box_height/2,
                              xmid+box_width/2, ymid+box_height/2)
                    x0, top = bounds[:2]
                else:
                    bounds = (x0, top, min(width, x1+extra), min(height, max(y1+9, top+16)))
                box = TextBox(doc, target, shape_id, bounds, 20+shape_id, angle)
                add_lines(doc, box, lines, x0, top, bounds[2]-bounds[0],
                          page_field=block.get("page_field"), index=index)
            elif block["kind"] == "table":
                box = TextBox(doc, target, shape_id, (x0,y0,x1+2,min(height,y1+10)), 20+shape_id)
                add_table(doc, box, block)
            else:
                add_picture(target, page["media"][block["asset"]], block["box"], shape_id,
                            block.get("label", "Ảnh từ PDF"), block.get("preserved_ink", False))
        if index < len(pages)-1:
            # Word stores a section break on its own paragraph, so keep it short.
            anchor.paragraph_format.space_after = Pt(0)
    result = BytesIO()
    doc.save(result)
    return result.getvalue()
