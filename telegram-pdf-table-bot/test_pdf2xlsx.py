"""Checks for pdf2xlsx.py: a synthetic two-page PDF, plus the user's CPVC PO lists when present on this PC."""
import sys
import tempfile
from pathlib import Path

import fitz
from openpyxl import load_workbook

import pdf2xlsx

SAMPLES = Path(r"D:\C3D1\Piping\BQ\CPVC\CPVC Pipe & Fitting\PO")


def check(condition, label):
    if not condition:
        raise AssertionError(label)
    print("passed:", label)


def draw_table(page, top, rows, xs=(20, 60, 160, 260, 360), height=30):
    ys = [top + i * height for i in range(len(rows) + 1)]
    for x in xs:
        page.draw_line((x, ys[0]), (x, ys[-1]))
    for y in ys:
        page.draw_line((xs[0], y), (xs[-1], y))
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            page.insert_text((xs[c] + 4, ys[r] + 20), value, fontsize=10)


def make_pdf(path):
    doc = fitz.open()
    header = ["No", "Size", "Name", "Qty"]
    page = doc.new_page(width=595, height=842)
    page.insert_text((20, 40), "PO.TEST", fontsize=12)
    draw_table(page, 60, [header, ["1", "DN15", "PIPE", "10 LGH"], ["2", "DN25", "SOCKET", "1,260"]])
    page.insert_text((250, 820), "Page 1 of 2", fontsize=9)
    page = doc.new_page(width=595, height=842)
    draw_table(page, 40, [header, ["3", "DN32", "CAP", "5"]])
    page.insert_text((250, 820), "第 2 页，共 2 页", fontsize=9, fontname="china-s")
    doc.save(path)


def rows_of(xlsx):
    return [tuple(r) for r in load_workbook(xlsx).active.iter_rows(values_only=True)]


def synthetic(tmp):
    pdf, xlsx = tmp / "synthetic.pdf", tmp / "synthetic.xlsx"
    make_pdf(pdf)
    info = pdf2xlsx.convert(str(pdf), str(xlsx))
    rows = rows_of(xlsx)
    check(rows[0][0] == "PO.TEST", "text above a table is kept as a title row")
    check(rows[1] == ("No", "Size", "Name", "Qty"), "header row is read")
    check(rows[2] == (1, "DN15", "PIPE", "10 LGH"), "numbers become numeric, units stay text")
    check(rows[3][3] == 1260, "thousand separators are read as numbers")
    check(rows[4] == (3, "DN32", "CAP", 5) and len(rows) == 5, "header repeated on page 2 is dropped")
    check(not any("Page" in str(v) or "页" in str(v) for r in rows for v in r), "page footers are left out")
    check(info == {"pages": 2, "tables": 2, "rows": 5}, "conversion summary counts tables and rows")

    blank = tmp / "blank.pdf"
    d = fitz.open()
    d.new_page()
    d.save(blank)
    try:
        pdf2xlsx.convert(str(blank), str(tmp / "blank.xlsx"))
        check(False, "empty PDF is refused")
    except pdf2xlsx.NoTables as e:
        check("Scanned" in str(e), "empty/scanned PDF is refused with a message")


def samples(tmp):
    if not SAMPLES.exists():
        print("skipped: CPVC PO sample PDFs not found on this PC")
        return
    expected = {
        "PO.20288  LIST": {"rows": 115, "cells": {"B3": "尺寸Size", "A3": "项次", "G65": "1,260PCS", "C77": "TEE REDUCING BRANCH",
                                                    "E90": 5667, "G90": 8623, "A91": "DC: CLEAR PVC", "F107": 'middle 2-1/2" pipe',
                                                    "A108": "SUPPORT", "F115": "640LGH*3PCS/LGH"}},
        "PO.20469  LIST": {"rows": 42, "cells": {"E31": 669, "G31": 753, "A40": "SUPPORT", "B42": "DN80*40"}},
        "PO.20472  LIST": {"rows": 86, "cells": {"A32": 29, "E32": "15 PCS", "D32": "SCH80 CPVC", "E81": 3789,
                                                   "A82": "SUPPORT", "E86": "3960 个", "F86": "1320LGH*3PCS/LGH", "G86": None}},
    }
    for name, exp in expected.items():
        out = tmp / f"{name}.xlsx"
        pdf2xlsx.convert(str(SAMPLES / f"{name}.pdf"), str(out))
        ws = load_workbook(out).active
        check(ws.max_row == exp["rows"], f"{name}: {exp['rows']} rows")
        for ref, value in exp["cells"].items():
            check(ws[ref].value == value, f"{name}: {ref} = {value!r} (got {ws[ref].value!r})")
        values = [str(v) for r in ws.iter_rows(values_only=True) for v in r if v is not None]
        check(not any("页" in v for v in values), f"{name}: no page footer text")


def main():
    tmp = Path(tempfile.mkdtemp(prefix="pdf2xlsx-test-"))
    synthetic(tmp)
    samples(tmp)
    print("ALL CONVERTER CHECKS PASSED")


if __name__ == "__main__":
    sys.exit(main())
