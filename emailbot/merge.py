"""Turns a letter + one recipient into a finished email.

Placeholders are written as {Field Name} and are matched against the
recipient's columns without caring about upper/lower case or extra spaces.
The special placeholder {Addendums} marks where add-on paragraphs and the
personal note go; without it they are placed after the main text.
"""

import html
import re

FIELD_RE = re.compile(r"\{\s*([^{}\n]+?)\s*\}")
EMAIL_RE = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$")
ADDENDUM_MARKER = "addendums"
_MARKER_TOKEN = "\x00ADDENDUMS\x00"


def norm(name):
    return " ".join(str(name).split()).lower()


def split_addresses(value):
    return [a.strip() for a in re.split(r"[;,]", value or "") if a.strip()]


def is_email(value):
    return bool(EMAIL_RE.match(value or ""))


def _fill(text, fields, problems):
    def repl(m):
        key = norm(m.group(1))
        if key == ADDENDUM_MARKER:
            return _MARKER_TOKEN
        if key not in fields:
            problems["unknown"].add(m.group(1).strip())
            return m.group(0)
        value = fields[key]
        if not value.strip():
            problems["blank"].add(m.group(1).strip())
        return value

    return FIELD_RE.sub(repl, text or "")


def _join_paragraphs(parts):
    return "\n\n".join(p.strip("\n") for p in parts if p and p.strip())


def to_html(text):
    escaped = html.escape(text).replace("\n", "<br>\n")
    return (
        '<div style="font-family: Calibri, Arial, sans-serif; font-size: 11pt;">'
        f"{escaped}</div>"
    )


def render(letter, recipient):
    """Return the finished email for one recipient, plus any problems found."""
    fields = {norm(k): str(v or "") for k, v in (recipient.get("fields") or {}).items()}
    email = (recipient.get("email") or "").strip()
    fields.setdefault("email", email)
    problems = {"unknown": set(), "blank": set()}

    chosen = set(recipient.get("addendums") or [])
    extras = [
        _fill(a.get("text", ""), fields, problems)
        for a in letter.get("addendums", [])
        if a.get("id") in chosen
    ]
    extras.append(_fill(recipient.get("note", ""), fields, problems))
    extras_text = _join_paragraphs(extras)

    body = _fill(letter.get("body", ""), fields, problems)
    if _MARKER_TOKEN in body:
        body = body.replace(_MARKER_TOKEN, extras_text)
        body = re.sub(r"\n{3,}", "\n\n", body)
        main = body
    else:
        main = _join_paragraphs([body, extras_text])
    text = _join_paragraphs([main, _fill(letter.get("signature", ""), fields, problems)])
    subject = _fill(letter.get("subject", ""), fields, problems).replace(_MARKER_TOKEN, "")

    cc = split_addresses(recipient.get("cc"))
    warnings = []
    if not email:
        warnings.append({"level": "error", "msg": "No email address."})
    elif not is_email(email):
        warnings.append({"level": "error", "msg": f"“{email}” doesn't look like an email address."})
    for addr in cc:
        if not is_email(addr):
            warnings.append({"level": "error", "msg": f"CC “{addr}” doesn't look like an email address."})
    if not subject.strip():
        warnings.append({"level": "error", "msg": "The subject line is empty."})
    for name in sorted(problems["unknown"]):
        warnings.append({
            "level": "error",
            "msg": f"The letter uses {{{name}}}, but there is no column called “{name}”.",
        })
    for name in sorted(problems["blank"]):
        warnings.append({"level": "warning", "msg": f"“{name}” is blank for this person."})

    return {
        "row_id": recipient.get("id"),
        "to": email,
        "cc": cc,
        "subject": subject,
        "text": text,
        "html": to_html(text),
        "warnings": warnings,
        "has_errors": any(w["level"] == "error" for w in warnings),
    }
