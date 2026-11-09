"""Locate small regions of colored ink in scans of printed documents."""

from math import ceil
import re

import pymupdf as pdf
from PIL import Image, ImageChops


def ink_masks(image):
    red, green, blue = image.split()
    red_ink = ImageChops.multiply(
        ImageChops.subtract(red, green).point([0]*42+[255]*214),
        ImageChops.subtract(red, blue).point([0]*29+[255]*227),
    )
    blue_ink = ImageChops.multiply(
        ImageChops.subtract(blue, red).point([0]*34+[255]*222),
        ImageChops.subtract(green, red).point([0]*12+[255]*244),
    )
    return red_ink, blue_ink


def nearby(box, horizontal, vertical):
    return pdf.Rect(box.x0-horizontal, box.y0-vertical, box.x1+horizontal, box.y1+vertical)


def pixel_box(box, origin, scale):
    return tuple(round((value-origin[i%2])*scale) for i, value in enumerate(box))


def ink_components(mask, origin, scale):
    """Group nearby saturated strokes without grouping the background."""
    width, height = ceil(mask.width/(scale*3)), ceil(mask.height/(scale*3))
    small = mask.resize((width, height), Image.Resampling.BOX)
    dx, dy = mask.width/width, mask.height/height
    active = {(i%width, i//width) for i, level in enumerate(small.get_flattened_data()) if level >= 8}
    boxes = []
    while active:
        component = [active.pop()]
        for x, y in component:
            for a in range(x-1, x+2):
                for b in range(y-1, y+2):
                    if (a,b) in active:
                        active.remove((a,b))
                        component.append((a,b))
        if len(component) < 3:
            continue
        x0, x1 = min(x for x,y in component), max(x for x,y in component)+1
        y0, y1 = min(y for x,y in component), max(y for x,y in component)+1
        crop = (round(x0*dx),round(y0*dy),round(x1*dx),round(y1*dy))
        ink = mask.crop(crop).getbbox()
        if ink is not None:
            boxes.append(pdf.Rect(*(value/scale+origin[i%2] for i,value in enumerate(
                (crop[0]+ink[0],crop[1]+ink[1],crop[0]+ink[2],crop[1]+ink[3])))))
    return sorted(boxes, key=lambda box: (box.y0,box.x0))


def detected_ink(image, origin, scale, page):
    red, blue = ink_masks(image)
    red_parts = ink_components(red, origin, scale)
    seals = [box for box in red_parts if 26 <= box.height <= 150 and (
             box.y0 > page.height*.32 and 28 <= box.width <= 145 or
             box.y0 > page.height*.20 and box.x1 > page.width*.92 and 7 <= box.width <= 65)]
    blue_parts = ink_components(blue, origin, scale)
    graphics, handwriting = [], []
    while seals:
        box = seals.pop(0)
        # Keep signature strokes and stamped names close to the seal together.
        region = nearby(box, 28, 21)
        combined = pdf.Rect(box)
        for red_box in red_parts:
            if region.intersects(red_box):
                combined |= red_box
                if red_box in seals:
                    seals.remove(red_box)
        for blue_box in blue_parts[:]:
            if region.intersects(blue_box):
                combined |= blue_box
                blue_parts.remove(blue_box)
        graphics.append((nearby(combined, 5, 8) & page, "seal"))
    for box in blue_parts:
        if box.y0 > page.height*.60 and box.width >= 32 and box.width >= 1.45*box.height and box.height <= 64:
            # Leave a role label above the signature editable, while keeping
            # the printed name immediately below the handwritten strokes.
            signature = pdf.Rect(box.x0-4, box.y0-1, box.x1+4, box.y1+21) & page
            for seal, kind in graphics:
                if kind == "seal" and abs(seal.y0-signature.y0) <= 14:
                    seal.y0 += 4
            graphics.append((signature, "signature"))
        elif 4 <= box.width <= 52 and 5 <= box.height <= 25 and box.x1 < page.width*.93:
            handwriting.append(box)
    return graphics, handwriting


def handwriting_line(ink, words, page, dark, origin, scale):
    """Limit the fallback to the printed line containing the handwritten ink."""
    mid = (ink.y0+ink.y1)/2
    row = [pdf.Rect(word["box"]) for word in words]
    row = [box for box in row if abs((box.y0+box.y1)/2-mid) <= 7 and box.height <= 24]
    line = pdf.Rect(ink)
    remaining = row[:]
    while remaining:
        linked = [box for box in remaining if box.x1 >= line.x0-16 and box.x0 <= line.x1+16]
        if not linked:
            break
        for box in linked:
            line |= box
            remaining.remove(box)
    line.y0 = max(line.y0, mid-12)
    line.y1 = min(line.y1, mid+12)
    # OCR may put the year or the last few words in a separate line.
    # Recover their horizontal extent from the printed ink on the same baseline.
    band = pdf.Rect(page.x0, mid-4.5, page.x1, mid+4.5) & page
    pixels = pixel_box(band, origin, scale)
    strip = dark.crop(pixels).resize((max(1, round(band.width*2)), 1), Image.Resampling.BOX)
    columns = [band.x0 + i*band.width/strip.width
               for i, value in enumerate(strip.get_flattened_data()) if value > 25]
    for x in reversed(columns):
        if x < line.x0 and line.x0-x < 17:
            line.x0 = x
    for x in columns:
        if x > line.x1 and x-line.x1 < 17:
            line.x1 = x
    line = nearby(line, 2, 1.4) & page
    if line.y0 < page.height*.24 and line.x1 < page.width*.47:
        # Some form headers place an underlined printed line just above the
        # handwritten reference number. Keep it whole when the boxes touch.
        search = pdf.Rect(line.x0-10, line.y0-10, line.x1+10, line.y0-1) & page
        x0,y0,x1,y1 = pixel_box(search, origin, scale)
        strip = dark.crop((x0,y0,x1,y1)).tobytes()
        width = x1-x0
        rule = re.compile(rb"\xff{" + str(round(28*scale)).encode() + rb",}")
        matches = [(len(hit.group()),y,hit.start()-y*width,hit.end()-y*width)
                   for y in range(y1-y0)
                   for hit in rule.finditer(strip,y*width,(y+1)*width)]
        if matches:
            _, row, start, end = max(matches)
            line.x0 = min(line.x0, (x0+start)/scale+origin[0]-3)
            line.x1 = max(line.x1, (x0+end)/scale+origin[0]+3)
            line.y0 = (y0+row)/scale+origin[1]-12
    return line & page


def preserve_ink(image, origin, scale, page, groups):
    """Return bounded scan crops and words outside those regions."""
    graphics, handwriting = detected_ink(image, origin, scale, page)
    words = [word for group in groups.values() for word in group]
    dark = image.convert("L").point([255]*150+[0]*106) if handwriting else None
    # Keep the original strokes for a line only when a printed/OCR context exists.
    ink_lines = []
    for ink in handwriting:
        if any(nearby(box, 7, 8).intersects(ink) for box, _ in graphics):
            continue
        region = handwriting_line(ink, words, page, dark, origin, scale)
        if region.width < 24:
            continue
        related = [line for line in ink_lines if line.intersects(region) and
                   abs((line.y0+line.y1-region.y0-region.y1)/2) <= 6]
        for line in related:
            region |= line
            ink_lines.remove(line)
        ink_lines.append(region)
    graphics.extend((box, "handwriting") for box in ink_lines)
    crops = [(tuple(box), image.crop(pixel_box(box, origin, scale)), kind)
             for box, kind in graphics[:16] if box.get_area() <= page.get_area()*.10]
    areas = [pdf.Rect(box) for box, _, _ in crops]
    filtered = {}
    for key, group in groups.items():
        kept = []
        for word in group:
            box = pdf.Rect(word["box"])
            center = (box.tl+box.br)/2
            if not any(area.contains(center) or (box&area).get_area() > box.get_area()*.58
                       for area in areas):
                kept.append(word)
        if kept:
            filtered[key] = kept
    return crops, filtered
