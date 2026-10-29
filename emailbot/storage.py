"""Saves everything as plain files in the data/ folder next to the app."""

import csv
import json
import os
import re
import sys
import threading
import uuid
from datetime import datetime
from pathlib import Path

if getattr(sys, "frozen", False):
    # Running as EmailBot.exe: the code is unpacked to a temporary folder that's
    # deleted on exit, so keep data in a folder next to the .exe instead.
    _DEFAULT_DATA = Path(sys.executable).resolve().parent / "data"
else:
    _DEFAULT_DATA = Path(__file__).resolve().parent.parent / "data"
DATA_DIR = Path(os.environ.get("EMAILBOT_DATA", _DEFAULT_DATA))
_lock = threading.Lock()

EXAMPLE_LETTER = {
    "name": "Example letter",
    "subject": "An update for {First Name}",
    "body": (
        "Dear {First Name},\n\n"
        "Thank you for being part of {Organization}. We wanted to let you know "
        "about a few changes coming next month.\n\n"
        "{Addendums}\n\n"
        "If you have any questions, just reply to this email."
    ),
    "signature": "Best regards,\nYour Name",
    "addendums": [
        {"name": "Late payment reminder",
         "text": "Our records show an outstanding balance on your account. Please get in touch so we can sort it out."},
        {"name": "Welcome (new members)",
         "text": "Since you joined recently, we've attached our welcome guide for you."},
    ],
}


def new_id():
    return uuid.uuid4().hex[:10]


def _path(name):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    return DATA_DIR / name


def read_json(name, default):
    p = _path(name)
    if not p.exists():
        return default
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def write_json(name, value):
    p = _path(name)
    tmp = p.with_suffix(".tmp")
    with _lock:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(value, f, indent=2, ensure_ascii=False)
        os.replace(tmp, p)


def _fresh_letter(template=None):
    letter = json.loads(json.dumps(template or {"name": "New letter", "subject": "", "body": "",
                                               "signature": "", "addendums": []}))
    letter["id"] = new_id()
    letter.setdefault("attachments", [])
    for a in letter["addendums"]:
        a["id"] = new_id()
    return letter


def load_workspace():
    ws = read_json("workspace.json", None)
    if ws is None:
        letter = _fresh_letter(EXAMPLE_LETTER)
        ws = {
            "letters": [letter],
            "current_letter_id": letter["id"],
            "recipients": {
                "columns": ["First Name", "Organization"],
                "rows": [{"id": new_id(), "include": True, "email": "", "cc": "",
                          "fields": {"First Name": "", "Organization": ""},
                          "addendums": [], "note": ""}],
            },
        }
        write_json("workspace.json", ws)
    return ws


def save_workspace(ws):
    write_json("workspace.json", ws)


def new_letter(copy_from=None):
    letter = _fresh_letter(copy_from)
    if copy_from:
        letter["name"] = f"{copy_from.get('name', 'Letter')} (copy)"
        letter["attachments"] = []
    return letter


# ---- attachments ---------------------------------------------------------

def attachment_dir(letter_id):
    d = _path("attachments") / re.sub(r"[^a-zA-Z0-9]", "", letter_id)
    d.mkdir(parents=True, exist_ok=True)
    return d


def safe_filename(name):
    name = os.path.basename(name or "").strip().replace("\x00", "")
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return name or "attachment"


# ---- sent tracking & log -------------------------------------------------

def load_sent():
    return read_json("sent.json", {})


def mark_sent(letter_id, email):
    sent = load_sent()
    sent.setdefault(letter_id, {})[email.lower()] = datetime.now().isoformat(timespec="seconds")
    write_json("sent.json", sent)


def log_send(letter_name, to, subject, ok, error=""):
    p = _path("send_log.csv")
    new = not p.exists()
    # A new file starts with a byte-order mark so Excel reads accented names correctly.
    with _lock, open(p, "a", newline="", encoding="utf-8-sig" if new else "utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["time", "letter", "to", "subject", "result", "error"])
        w.writerow([datetime.now().isoformat(sep=" ", timespec="seconds"), letter_name, to,
                    subject, "sent" if ok else "failed", error])


def read_log(limit=300):
    p = _path("send_log.csv")
    if not p.exists():
        return []
    with open(p, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    return rows[-limit:][::-1]
