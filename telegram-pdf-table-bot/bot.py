"""Telegram bot: send a PDF or a photo of a table, receive an Excel file of its tables.

A separate bot (its own BotFather token, chat, allowlist and folder) from the
fastener-order bot. Private chats only; the first administrator pairs with
/pair CODE from Setup, colleagues send /id and the administrator uses /allow ID.
The token comes from the TELEGRAM_BOT_TOKEN environment variable set by
Start.ps1 and is never logged. Photos are read offline on this PC.
"""
import datetime as dt
import json
import os
import re
import socket
import sys
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path

import image2xlsx
import pdf2xlsx

ROOT = Path(__file__).resolve().parent
LOCK_PORT = 43872  # the fastener bot uses 43871
MAX_BYTES = 20 * 1024 * 1024  # Telegram Bot API download limit
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff")

HELP = (
    "PDF / Photo Table to Excel Bot\n\n"
    "Send a PDF file or a photo of a table. The bot replies with an Excel file of the tables in it.\n\n"
    "PDF files:\n"
    "- Works with PDFs that contain real text and table borders, e.g. exported from Excel, Word or a system.\n"
    "- Scanned PDFs are not supported. Send a photo or screenshot of the page instead.\n\n"
    "Photos and screenshots:\n"
    "- Take the photo straight on, in good light, with the whole table in the frame and in focus.\n"
    "- For the best accuracy, send it as a file (paperclip > File) so Telegram does not compress it.\n"
    "- Printed English and Chinese text is supported. Handwriting is not.\n"
    "- Text recognition can misread characters or drop spaces. Always check the Excel against the photo.\n"
    "- Reading a photo takes about 20-30 seconds. Each photo gives its own Excel file.\n\n"
    "For both:\n"
    "- All tables go into one sheet in order. Repeated headers and page numbers are left out.\n"
    "- Plain numbers become Excel numbers. Values with units, such as 10 LGH, stay as text.\n"
    "- Maximum file size: 20 MB.\n\n"
    "/id - Show your Telegram ID\n"
    "/help - Show these instructions"
)


def log(message):
    print(f"{dt.datetime.now().isoformat(timespec='seconds')} {message}", flush=True)


def read_json(path, default=None):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        return default


def write_json(path, data):
    path = Path(path)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def safe_stem(name):
    stem = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", Path(name).stem).strip(" .")
    return stem[:100] or "converted"


class TelegramError(Exception):
    def __init__(self, method, code=None, retry_after=None, description=""):
        super().__init__(f"Telegram {method} failed ({code})")
        self.code = code
        self.retry_after = retry_after
        self.description = description or ""


class Api:
    def __init__(self, token):
        self._base = f"https://api.telegram.org/bot{token}/"
        self._files = f"https://api.telegram.org/file/bot{token}/"

    def call(self, method, data=None, files=None, timeout=65):
        if files:
            body, content_type = multipart(data or {}, files)
        else:
            body, content_type = json.dumps(data or {}).encode("utf-8"), "application/json"
        request = urllib.request.Request(self._base + method, data=body, headers={"Content-Type": content_type})
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                payload = json.load(response)
        except urllib.error.HTTPError as e:
            try:
                payload = json.load(e)
            except ValueError:
                raise TelegramError(method, e.code) from None
        if not payload.get("ok"):
            raise TelegramError(method, payload.get("error_code"),
                                (payload.get("parameters") or {}).get("retry_after"), payload.get("description"))
        return payload["result"]

    def download(self, file_path, timeout=120):
        url = self._files + urllib.parse.quote(file_path)
        try:
            with urllib.request.urlopen(url, timeout=timeout) as response:
                return response.read()
        except urllib.error.HTTPError as e:
            raise TelegramError("download", e.code) from None


def multipart(fields, files):
    boundary = uuid.uuid4().hex
    out = bytearray()
    for key, value in fields.items():
        out += f'--{boundary}\r\nContent-Disposition: form-data; name="{key}"\r\n\r\n'.encode("utf-8")
        out += str(value).encode("utf-8") + b"\r\n"
    for key, (filename, content, content_type) in files.items():
        filename = filename.replace('"', "'")
        out += (f'--{boundary}\r\nContent-Disposition: form-data; name="{key}"; filename="{filename}"\r\n'
                f"Content-Type: {content_type}\r\n\r\n").encode("utf-8")
        out += content + b"\r\n"
    out += f"--{boundary}--\r\n".encode("utf-8")
    return bytes(out), f"multipart/form-data; boundary={boundary}"


class Bot:
    def __init__(self, api, root, state, config):
        self.api = api
        self.root = Path(root)
        self.state = state
        self.config = config

    def save(self):
        write_json(self.root / "state.json", self.state)

    def say(self, chat, text):
        return self.api.call("sendMessage", {"chat_id": chat, "text": text})

    def handle(self, m):
        chat, sender = m.get("chat") or {}, m.get("from") or {}
        if chat.get("type") != "private" or not sender or sender.get("is_bot"):
            return
        uid, cid = sender["id"], chat["id"]
        text = (m.get("text") or "").strip()
        state = self.state

        if re.fullmatch(r"/id(?:@\w+)?", text):
            self.say(cid, f"Your Telegram ID: {uid}\nSend this ID to the administrator. "
                          f"They can grant access using /allow {uid}.")
            return
        if not state.get("owner") and re.match(r"/pair\s+", text):
            if state.get("pairFailures", 0) >= 10:
                self.say(cid, "Pairing is locked. Please ask the administrator to run Setup.cmd again.")
                return
            if text[5:].strip() != self.config.get("pairCode"):
                state["pairFailures"] = state.get("pairFailures", 0) + 1
                self.save()
                self.say(cid, "Incorrect pairing code. Please use the code shown in the latest Setup window.")
                return
            state["owner"], state["users"] = uid, [uid]
            self.save()
            self.say(cid, "Administrator pairing completed. Colleagues can send /id and share their ID with you. "
                          "Use /allow ID to grant access and /remove ID to revoke it.\n\n" + HELP)
            return
        if uid not in state.get("users", []):
            self.say(cid, "Welcome to the PDF Table to Excel Bot. Please contact the administrator for access. "
                          "Send /id to get your Telegram ID.")
            return
        allow = re.fullmatch(r"/(allow|remove)\s+(\d+)", text)
        if allow:
            if uid != state.get("owner"):
                self.say(cid, "Only the administrator can change access.")
                return
            target = int(allow.group(2))
            if target == state["owner"]:
                self.say(cid, "Administrator access remains enabled.")
                return
            users = [u for u in state["users"] if u != target]
            if allow.group(1) == "allow":
                users.append(target)
            state["users"] = users
            self.save()
            self.say(cid, f"{allow.group(1)} {target}: OK")
            return
        if re.fullmatch(r"/(?:start|help)(?:@\w+)?", text):
            self.say(cid, HELP)
            return
        if m.get("document"):
            doc = m["document"]
            name = (doc.get("file_name") or "").lower()
            mime = doc.get("mime_type") or ""
            if name.endswith(".pdf") or mime == "application/pdf":
                self.convert(m, doc, "pdf")
            elif name.endswith(IMAGE_EXTENSIONS) or mime.startswith("image/"):
                self.convert(m, doc, "image")
            else:
                self.say(cid, "Only PDF files and photos are supported. Please send a PDF, JPG or PNG.")
            return
        if m.get("photo"):
            # Telegram sends several sizes of the same photo; use the largest.
            largest = max(m["photo"], key=lambda p: (p.get("width", 0) * p.get("height", 0), p.get("file_size", 0)))
            self.convert(m, largest, "image")
            return
        self.say(cid, "Please send a PDF file or a photo of a table. Send /help for details.")

    def convert(self, m, file, kind):
        cid = m["chat"]["id"]
        if kind == "pdf":
            name = file.get("file_name") or "document.pdf"
        else:
            stamp = dt.datetime.fromtimestamp(m.get("date", time.time())).strftime("%Y%m%d_%H%M%S")
            name = file.get("file_name") or f"photo_{stamp}.jpg"
        if (file.get("file_size") or 0) > MAX_BYTES:
            self.say(cid, "This file is larger than 20 MB, which Telegram bots cannot download. "
                          "Please split or compress it and send it again.")
            return
        job_id = f"{cid}_{m['message_id']}"
        job = self.root / "outputs" / job_id
        job.mkdir(parents=True, exist_ok=True)
        delivery_path = job / "delivery.json"
        prior = read_json(delivery_path) or {}
        if prior.get("status") in ("sent", "no_tables", "failed"):
            return
        if prior.get("status") in ("sending", "uncertain"):
            self.say(cid, "The previous file delivery could not be confirmed. Please check whether you received it. "
                          "If not, send the file again.")
            return
        try:
            if kind == "image":
                self.say(cid, "Reading the photo. This takes about 20-30 seconds.")
            self.api.call("sendChatAction", {"chat_id": cid, "action": "upload_document"})
        except TelegramError:
            pass
        write_json(job / "request.json", {
            "telegramUserId": m["from"]["id"],
            "receivedAt": dt.datetime.fromtimestamp(m.get("date", time.time()), dt.timezone.utc).isoformat(),
            "kind": kind,
            "fileName": name,
            "fileSize": file.get("file_size"),
        })
        suffix = ".pdf" if kind == "pdf" else (Path(name).suffix.lower() or ".jpg")
        source = job / f"source{suffix}"
        try:
            info = self.api.call("getFile", {"file_id": file["file_id"]})
            source.write_bytes(self.api.download(info["file_path"]))
        except TelegramError as e:
            log(f"Download failed for {job_id} ({e.code})")
            if "too big" in e.description.lower():
                self.say(cid, "This file is too large for Telegram bots to download (20 MB). Please split or compress it.")
            else:
                self.say(cid, "The file could not be downloaded from Telegram. Please send it again.")
            return
        output = job / f"{safe_stem(name)}.xlsx"
        converter = pdf2xlsx.convert if kind == "pdf" else image2xlsx.convert
        try:
            result = converter(str(source), str(output), title=pdf2xlsx.sheet_title(name))
        except pdf2xlsx.NoTables as e:
            write_json(delivery_path, {"status": "no_tables"})
            hint = ("The bot converts tables with visible borders in text-based PDFs. "
                    "For a scanned PDF, send a photo or screenshot of the page." if kind == "pdf" else
                    "Please send a clear photo of a printed table, taken straight on in good light.")
            self.say(cid, f"{e}\n{hint}")
            return
        except Exception:
            log(f"Conversion failed for {job_id}\n{traceback.format_exc()}")
            write_json(delivery_path, {"status": "failed"})
            self.say(cid, "The file could not be read or converted. Please check that the file opens normally, "
                          "or contact the administrator.")
            return
        write_json(delivery_path, {"status": "sending"})
        if kind == "pdf":
            caption = (f"{output.name}\n{result['tables']} table(s), {result['rows']} row(s) from "
                       f"{result['pages']} page(s). Please check against the PDF before use.")
        else:
            caption = (f"{output.name}\n{result['tables']} table(s), {result['rows']} row(s) read from the photo. "
                       "Text recognition can misread characters or drop spaces. Please check against the photo "
                       "before use.")
        try:
            sent = self.api.call("sendDocument", {"chat_id": cid, "caption": caption,
                                                  "reply_to_message_id": m["message_id"]},
                                 files={"document": (output.name, output.read_bytes(), XLSX)}, timeout=120)
            write_json(delivery_path, {"status": "sent", "messageId": sent.get("message_id"), **result})
            log(f"Converted {kind} and delivered {job_id}: {result}")
        except Exception as e:
            write_json(delivery_path, {"status": "uncertain", **result})
            log(f"Delivery unconfirmed for {job_id} ({getattr(e, 'code', type(e).__name__)})")
            self.say(cid, "The Excel file was created, but Telegram delivery could not be confirmed. "
                          "Please check whether you received it. If not, send the file again.")


def connect(api):
    delay = 2
    while True:
        try:
            return api.call("getMe"), api.call("getWebhookInfo")
        except TelegramError as e:
            if e.code in (401, 404, 409):
                raise SystemExit(f"Telegram startup rejected ({e.code}). Check the token or other bot instances.")
            wait = max(delay, e.retry_after or 0)
        except (OSError, ValueError) as e:
            wait = delay
            log(f"Telegram startup connection failed ({type(e).__name__}).")
        log(f"Retrying Telegram connection in {wait} seconds.")
        time.sleep(wait)
        delay = min(delay * 2, 60)


def main():
    token = os.environ.pop("TELEGRAM_BOT_TOKEN", None)
    if not token:
        print("Run Setup.cmd first, then Start.cmd.", file=sys.stderr)
        return 1
    lock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        lock.bind(("127.0.0.1", LOCK_PORT))
        lock.listen(1)
    except OSError:
        print(f"Bot already running, or local port {LOCK_PORT} is in use.", file=sys.stderr)
        return 1
    config = read_json(ROOT / "config.json", {})
    state = read_json(ROOT / "state.json") or {"offset": 0, "owner": None, "users": [], "pairFailures": 0}
    api = Api(token)
    del token
    me, webhook = connect(api)
    if webhook.get("url"):
        print("This bot already has a webhook. Use a new BotFather bot.", file=sys.stderr)
        return 1
    log(f"Online: @{me['username']}. Private chats only.")
    write_json(ROOT / "status.json", {"username": me["username"], "started": dt.datetime.now().isoformat(),
                                      "pid": os.getpid()})
    bot = Bot(api, ROOT, state, config)
    backoff = 2
    while True:
        try:
            updates = api.call("getUpdates", {"offset": state.get("offset", 0), "timeout": 45,
                                              "allowed_updates": ["message"]})
            for update in updates:
                try:
                    if update.get("message"):
                        bot.handle(update["message"])
                except TelegramError as e:
                    log(f"Update {update['update_id']} failed ({e.code})")
                except Exception:
                    log(f"Update {update['update_id']} failed\n{traceback.format_exc()}")
                # Always move past an update, so one bad message cannot block the bot.
                state["offset"] = update["update_id"] + 1
                bot.save()
            backoff = 2
        except TelegramError as e:
            log(f"Polling paused ({e.code})")
            if e.code in (401, 409):
                return 1
            time.sleep(max(backoff, e.retry_after or 0))
            backoff = min(backoff * 2, 60)
        except (OSError, ValueError) as e:
            log(f"Polling paused ({type(e).__name__})")
            time.sleep(backoff)
            backoff = min(backoff * 2, 60)


if __name__ == "__main__":
    sys.exit(main())
