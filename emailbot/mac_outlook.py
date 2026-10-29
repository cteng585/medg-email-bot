"""Sending through the Outlook app on a Mac, by scripting it with AppleScript.

No Microsoft/Azure setup is needed: mail goes out from whichever account the
Outlook app uses for new messages. Works with "new Outlook" for Mac, which
can't list or pick accounts from a script.
"""

import subprocess
import sys
import tempfile
from pathlib import Path

APP_PATH = Path("/Applications/Microsoft Outlook.app")

# Everything variable is passed in as arguments, never pasted into the script,
# so quotes or odd characters in a letter can't break it.
SEND_SCRIPT = """
on run argv
  set theSubject to item 1 of argv
  set theBody to item 2 of argv
  set toList to paragraphs of (item 3 of argv)
  set ccList to paragraphs of (item 4 of argv)
  set fileList to {}
  repeat with p in paragraphs of (item 5 of argv)
    if length of p > 0 then set end of fileList to (POSIX file (p as text))
  end repeat
  tell application "Microsoft Outlook"
    set m to make new outgoing message with properties {subject:theSubject, content:theBody}
    repeat with a in toList
      if length of a > 0 then make new to recipient at m with properties {email address:{address:(a as text)}}
    end repeat
    repeat with a in ccList
      if length of a > 0 then make new cc recipient at m with properties {email address:{address:(a as text)}}
    end repeat
    repeat with f in fileList
      make new attachment at m with properties {file:f}
    end repeat
    send m
  end tell
end run
"""

CHECK_SCRIPT = 'tell application "Microsoft Outlook" to get version'

# The step 1 check builds a complete email exactly as sending does, then deletes it unsent.
TEST_SCRIPT = SEND_SCRIPT.replace("    send m\n", "    delete m\n")

# AppleScript error numbers worth explaining in plain words.
_FRIENDLY = {
    "-1743": ("Your Mac blocked Email Bot from controlling Outlook. Open System Settings → "
              "Privacy & Security → Automation, find Terminal (or Python), and switch on "
              "“Microsoft Outlook”. Then try again."),
    "-600": "Outlook isn't responding. Open the Outlook app, wait for it to finish loading, and try again.",
    "-1712": "Outlook took too long to respond. Make sure it isn't showing a pop-up, then try again.",
}


class MacOutlookError(Exception):
    pass


def available():
    return sys.platform == "darwin" and APP_PATH.exists()


def _run(script, args=(), timeout=120):
    try:
        result = subprocess.run(
            ["osascript", "-", *args], input=script, capture_output=True, text=True, timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        raise MacOutlookError(_FRIENDLY["-1712"]) from None
    if result.returncode != 0:
        err = result.stderr.strip()
        for code, msg in _FRIENDLY.items():
            if f"({code})" in err:
                raise MacOutlookError(msg)
        raise MacOutlookError(f"Outlook reported a problem: {err or 'unknown error'}")
    return result.stdout.strip()


def uses_new_outlook():
    """True when Outlook's "New Outlook" switch is on (Outlook's own setting)."""
    try:
        result = subprocess.run(
            ["defaults", "read", "com.microsoft.Outlook", "IsRunningNewOutlook"],
            capture_output=True, text=True, timeout=10,
        )
    except (OSError, subprocess.TimeoutExpired):
        return False
    return result.stdout.strip() == "1"


def check():
    """Make sure Outlook answers; the first call triggers macOS's permission pop-up."""
    if not available():
        raise MacOutlookError("Microsoft Outlook isn't installed in this Mac's Applications folder.")
    return _run(CHECK_SCRIPT, timeout=90)


def send(to, cc, subject, html, text, attachments=()):
    if not available():
        raise MacOutlookError("Microsoft Outlook isn't installed in this Mac's Applications folder.")
    # New Outlook keeps a scripted HTML body as a draft but sends it as plain
    # text, so the tags show up in the recipient's inbox. Give it plain text.
    body = text if uses_new_outlook() else html
    _run(SEND_SCRIPT, [
        subject,
        body,
        to,
        "\n".join(cc),
        "\n".join(str(Path(p).resolve()) for p in attachments),
    ])


def full_check(test_address=None):
    """Check everything sending needs, stopping at the first problem.

    Returns {"steps": [{name, ok, detail}], "version"}.
    """
    result = {"steps": []}

    def record(name, ok, detail=""):
        result["steps"].append({"name": name, "ok": bool(ok), "detail": detail})
        return ok

    if not record("Microsoft Outlook is installed", available(),
                  "" if available() else "Microsoft Outlook isn't in this Mac's Applications folder."):
        return result
    try:
        result["version"] = check()
    except MacOutlookError as e:
        record("Outlook responds and Email Bot is allowed to control it", False, str(e))
        return result
    record("Outlook responds and Email Bot is allowed to control it", True, f"Version {result['version']}")

    attachment = Path(tempfile.gettempdir()) / "emailbot-check.txt"
    attachment.write_text("Email Bot check\n", encoding="utf-8")
    address = test_address or "emailbot-check@example.com"
    try:
        _run(TEST_SCRIPT, ["Email Bot check (not sent)", "Email Bot check", address, address, str(attachment)])
        record("Outlook can create an email (a test email was built and deleted, not sent)", True)
    except MacOutlookError as e:
        record("Outlook can create an email", False, str(e))
    return result
