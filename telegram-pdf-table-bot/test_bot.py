"""Offline checks of the bot's access control and PDF-to-Excel delivery, with a fake Telegram API."""
import io
import json
import shutil
import sys
import tempfile
from pathlib import Path

import fitz
from openpyxl import load_workbook

import bot
from test_pdf2xlsx import make_pdf


class FakeApi:
    def __init__(self, pdf_bytes):
        self.pdf = pdf_bytes
        self.calls = []

    def call(self, method, data=None, files=None, timeout=65):
        self.calls.append((method, data, files))
        if method == "getFile":
            return {"file_path": "documents/file_1.pdf"}
        return {"message_id": 1000 + len(self.calls)}

    def download(self, file_path, timeout=120):
        return self.pdf

    def texts(self):
        return [d["text"] for m, d, f in self.calls if m == "sendMessage"]

    def documents(self):
        return [(d, f) for m, d, f in self.calls if m == "sendDocument"]


def msg(uid, mid, text=None, document=None, chat_type="private"):
    m = {"message_id": mid, "date": 1789000000, "chat": {"id": uid, "type": chat_type},
         "from": {"id": uid, "is_bot": False, "first_name": "Test"}}
    if text is not None:
        m["text"] = text
    if document is not None:
        m["document"] = document
    return m


def check(condition, label):
    if not condition:
        raise AssertionError(label)
    print("passed:", label)


def main():
    root = Path(tempfile.mkdtemp(prefix="pdfbot-test-"))
    pdf_path = root / "sample.pdf"
    make_pdf(pdf_path)
    api = FakeApi(pdf_path.read_bytes())
    state = {"offset": 0, "owner": None, "users": [], "pairFailures": 0}
    b = bot.Bot(api, root, state, {"pairCode": "abc123"})
    pdf_doc = {"file_id": "F1", "file_name": "PO.1 LIST.pdf", "mime_type": "application/pdf", "file_size": 5000}

    b.handle(msg(11, 1, "/pair abc123", chat_type="group"))
    check(not api.calls, "group chats are ignored")

    b.handle(msg(11, 2, document=pdf_doc))
    check("contact the administrator" in api.texts()[-1] and not api.documents(), "unknown users get no conversion")

    b.handle(msg(11, 3, "/pair wrong"))
    check(state["pairFailures"] == 1 and state["owner"] is None, "wrong pairing code is counted")

    b.handle(msg(11, 4, "/pair abc123"))
    check(state["owner"] == 11 and json.loads((root / "state.json").read_text())["owner"] == 11, "administrator pairs")

    b.handle(msg(11, 5, document=pdf_doc))
    docs = api.documents()
    check(len(docs) == 1, "PDF is converted and one Excel file is sent")
    data, files = docs[0]
    filename, content, ctype = files["document"]
    check(filename == "PO.1 LIST.xlsx" and ctype == bot.XLSX, "Excel file keeps the PDF name")
    ws = load_workbook(io.BytesIO(content)).active
    check(ws["A1"].value == "PO.TEST" and ws["B3"].value == "DN15", "sent workbook holds the table")
    check("table(s)" in data["caption"], "caption reports tables and rows")
    delivery = json.loads((root / "outputs" / "11_5" / "delivery.json").read_text())
    check(delivery["status"] == "sent", "delivery is recorded as sent")

    b.handle(msg(11, 5, document=pdf_doc))
    check(len(api.documents()) == 1, "a re-delivered update does not send a second file")

    b.handle(msg(11, 6, document={"file_id": "F2", "file_name": "list.xlsx", "file_size": 10}))
    check("Only PDF files and photos" in api.texts()[-1], "other file types are refused")

    png = root / "table.png"
    fitz.open(pdf_path)[0].get_pixmap(dpi=150).save(png)
    api.pdf = png.read_bytes()
    photo = [{"file_id": "P_SMALL", "file_size": 900, "width": 90, "height": 120},
             {"file_id": "P_BIG", "file_size": 90000, "width": 1240, "height": 1754}]
    m = msg(11, 20)
    m["photo"] = photo
    b.handle(m)
    got_file = [d for meth, d, f in api.calls if meth == "getFile"][-1]
    check(got_file["file_id"] == "P_BIG", "the largest photo size is downloaded")
    data, files = api.documents()[-1]
    check(files["document"][0].startswith("photo_") and files["document"][0].endswith(".xlsx"), "photo gives photo_<time>.xlsx")
    ws = load_workbook(io.BytesIO(files["document"][1])).active
    check(ws["A1"].value == "PO.TEST" and ws["B3"].value == "DN15", "photo workbook holds the table")
    check("read from the photo" in data["caption"], "photo caption asks to check against the photo")
    check(ws.title.startswith("photo_"), "sheet is named after the photo, not the saved source file")

    b.handle(msg(11, 21, document={"file_id": "P3", "file_name": "Scan 01.png", "mime_type": "image/png",
                                   "file_size": 90000}))
    check(api.documents()[-1][1]["document"][0] == "Scan 01.xlsx", "image sent as a file keeps its name")
    documents_before = len(api.documents())
    api.pdf = pdf_path.read_bytes()

    b.handle(msg(11, 7, document={**pdf_doc, "file_size": bot.MAX_BYTES + 1}))
    check("20 MB" in api.texts()[-1], "files over 20 MB are refused")

    b.handle(msg(22, 8, "/allow 33"))
    check("contact the administrator" in api.texts()[-1], "non-users cannot allow others")
    b.handle(msg(11, 9, "/allow 22"))
    check(22 in state["users"], "administrator allows a colleague")
    b.handle(msg(22, 10, "/allow 33"))
    check(33 not in state["users"] and "Only the administrator" in api.texts()[-1], "colleagues cannot allow others")
    b.handle(msg(22, 11, document=pdf_doc))
    check(len(api.documents()) == documents_before + 1, "allowed colleague receives a conversion")
    b.handle(msg(11, 12, "/remove 22"))
    check(22 not in state["users"], "administrator removes a colleague")

    blank = root / "blank.pdf"
    d = fitz.open()
    d.new_page().insert_text((50, 50), "No table here")
    d.save(blank)
    api.pdf = blank.read_bytes()
    b.handle(msg(11, 13, document=pdf_doc))
    check("No table" in api.texts()[-1] and len(api.documents()) == documents_before + 1,
          "PDF without tables gets a clear message")

    api.pdf = b"not a pdf"
    b.handle(msg(11, 14, document=pdf_doc))
    check("could not be read" in api.texts()[-1], "broken PDF gets a clear message")

    shutil.rmtree(root, ignore_errors=True)
    print("ALL BOT CHECKS PASSED")


if __name__ == "__main__":
    sys.exit(main())
