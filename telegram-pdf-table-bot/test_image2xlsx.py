"""Checks for image2xlsx.py.

Images are rendered from PDFs and compared, cell by cell, with what pdf2xlsx.py
reads from the same page (the PDF text is exact, so it is the answer key).
Spaces are ignored in the comparison, because OCR on photos can drop them.
"""
import sys
import tempfile
from pathlib import Path

import fitz
from openpyxl import load_workbook
from PIL import Image, ImageEnhance, ImageFilter

import image2xlsx
import pdf2xlsx
from test_pdf2xlsx import SAMPLES, make_pdf


def check(condition, label):
    if not condition:
        raise AssertionError(label)
    print("passed:", label)


def render(pdf, page_no, png, dpi):
    fitz.open(pdf)[page_no].get_pixmap(dpi=dpi).save(png)


def photo_like(src, dst, angle=2.5, long_side=None, quality=70):
    """Simulate a phone photo: background, perspective, rotation, warm tint, blur, JPEG."""
    im = Image.open(src).convert("RGB")
    w, h = im.size
    canvas = Image.new("RGB", (int(w * 1.15), int(h * 1.12)), (120, 110, 95))
    canvas.paste(im, (int(w * 0.07), int(h * 0.05)))
    W, H = canvas.size
    im = canvas.transform((W, H), Image.PERSPECTIVE, (1.0, 0.035, -20, -0.02, 1.0, 15, 0.00001, 0.00002),
                          Image.BICUBIC, fillcolor=(120, 110, 95))
    im = im.rotate(angle, resample=Image.BICUBIC, fillcolor=(120, 110, 95))
    im = ImageEnhance.Brightness(im).enhance(0.9)
    r, g, b = im.split()
    im = Image.merge("RGB", (r, g.point(lambda v: v * 0.97), b.point(lambda v: v * 0.88)))
    im = im.filter(ImageFilter.GaussianBlur(0.8))
    if long_side:
        im.thumbnail((long_side, long_side))
    im.save(dst, quality=quality)


def norm(v):
    return None if v is None else str(v).replace(" ", "").replace("：", ":").replace("，", ",").upper()


def answer_key(pdf, page_no):
    doc = fitz.open(pdf)
    return [row for kind, rows in pdf2xlsx.page_items(doc[page_no]) if kind == "table" for row in rows]


def compare(xlsx, key, label, min_cells, qty_cols=()):
    rows = [list(r) for r in load_workbook(xlsx).active.iter_rows(values_only=True)]
    # Title rows have one value; compare only table rows, filtered the same way on both sides.
    table = [r for r in rows if sum(v is not None for v in r) >= 2]
    key = [r for r in key if sum(v is not None for v in r) >= 2]
    check(len(table) == len(key), f"{label}: {len(key)} table rows (got {len(table)})")
    total = same = 0
    wrong = []
    for got, want in zip(table, key):
        for c, w in enumerate(want):
            g = got[c] if c < len(got) else None
            total += 1
            if norm(g) == norm(w):
                same += 1
            else:
                wrong.append((w, g))
    share = same / total
    print(f"   {label}: {same}/{total} cells match ({share:.1%}); differences: {wrong[:12]}")
    check(share >= min_cells, f"{label}: at least {min_cells:.0%} of cells match")
    for c in qty_cols:
        bad = [(want[c], got[c]) for got, want in zip(table[1:], key[1:]) if norm(got[c]) != norm(want[c])]
        check(not bad, f"{label}: every quantity in column {c + 1} is exact {bad}")
    return rows


def main():
    tmp = Path(tempfile.mkdtemp(prefix="image2xlsx-test-"))

    pdf = tmp / "synthetic.pdf"
    make_pdf(pdf)
    png = tmp / "synthetic.png"
    render(pdf, 0, png, 150)
    out = tmp / "synthetic.xlsx"
    info = image2xlsx.convert(str(png), str(out))
    rows = compare(out, answer_key(pdf, 0), "synthetic screenshot", 1.0, qty_cols=(3,))
    check(rows[0][0] == "PO.TEST", "title above the table is kept")
    check(info["tables"] == 1, "one table found")
    check(not any("Page" in str(v) for r in rows for v in r), "page footer is left out")

    blank = tmp / "blank.png"
    Image.new("RGB", (800, 600), "white").save(blank)
    try:
        image2xlsx.convert(str(blank), str(tmp / "blank.xlsx"))
        check(False, "blank photo is refused")
    except pdf2xlsx.NoTables as e:
        check("No text" in str(e), "blank photo is refused with a message")

    broken = tmp / "broken.jpg"
    broken.write_bytes(b"not an image")
    try:
        image2xlsx.convert(str(broken), str(tmp / "broken.xlsx"))
        check(False, "unreadable image is refused")
    except pdf2xlsx.NoTables as e:
        check("could not be read" in str(e), "unreadable image is refused with a message")

    if not SAMPLES.exists():
        print("skipped: CPVC PO sample PDFs not found on this PC")
        print("ALL IMAGE CHECKS PASSED")
        return
    p20469, p20288 = SAMPLES / "PO.20469  LIST.pdf", SAMPLES / "PO.20288  LIST.pdf"
    qty = (4, 5, 6)  # CPVC Qty, DC CPVC Qty, Total Qty

    clean = tmp / "20469_clean.png"
    render(p20469, 0, clean, 150)
    compare_out = tmp / "20469_clean.xlsx"
    image2xlsx.convert(str(clean), str(compare_out))
    compare(compare_out, answer_key(p20469, 0), "PO.20469 screenshot", 0.97, qty)

    base = tmp / "20288.png"
    render(p20288, 0, base, 130)
    photo = tmp / "20288_photo.jpg"
    photo_like(base, photo, angle=2.5)
    photo_out = tmp / "20288_photo.xlsx"
    image2xlsx.convert(str(photo), str(photo_out))
    compare(photo_out, answer_key(p20288, 0), "PO.20288 skewed phone photo", 0.95, qty)

    squeezed = tmp / "20288_telegram.jpg"
    photo_like(base, squeezed, angle=-1.5, long_side=1280, quality=80)
    squeezed_out = tmp / "20288_telegram.xlsx"
    image2xlsx.convert(str(squeezed), str(squeezed_out))
    compare(squeezed_out, answer_key(p20288, 0), "PO.20288 photo compressed by Telegram", 0.95, qty)

    page3 = tmp / "20288_p3.png"
    render(p20288, 2, page3, 150)
    page3_out = tmp / "20288_p3.xlsx"
    info = image2xlsx.convert(str(page3), str(page3_out))
    compare(page3_out, answer_key(p20288, 2), "PO.20288 page 3 with two tables", 0.95, qty_cols=(0,))
    print("ALL IMAGE CHECKS PASSED")


if __name__ == "__main__":
    sys.exit(main())
