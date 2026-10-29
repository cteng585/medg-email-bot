"""Sending settings, and sending through the Outlook app on this computer.

On a Mac this scripts Outlook with AppleScript (mac_outlook); on Windows it
automates classic Outlook over COM (win_outlook).
"""

import sys

from . import mac_outlook, storage, win_outlook

if sys.platform == "darwin":
    PLATFORM = "mac"
elif sys.platform == "win32":
    PLATFORM = "windows"
else:
    PLATFORM = "other"


class OutlookError(Exception):
    pass


def _available():
    if PLATFORM == "mac":
        return mac_outlook.available()
    if PLATFORM == "windows":
        return win_outlook.available()
    return False


def _unavailable_reason():
    if PLATFORM == "mac":
        return "Microsoft Outlook isn't installed on this Mac."
    if PLATFORM == "windows":
        return win_outlook.NOT_INSTALLED
    return "Email Bot only works on a Mac or a Windows PC with Outlook installed."


class Outlook:
    def config(self):
        cfg = storage.read_json("config.json", {})
        cfg.setdefault("delay_seconds", 3)
        cfg.setdefault("my_email", "")
        cfg.setdefault("account", "")  # Windows only: send from this account instead of Outlook's default
        return cfg

    def save_config(self, **changes):
        cfg = self.config()
        cfg.update(changes)
        storage.write_json("config.json", cfg)

    def status(self):
        cfg = self.config()
        available = _available()
        return {
            "platform": PLATFORM,
            "available": available,
            "ready": available,
            "unavailable_reason": "" if available else _unavailable_reason(),
            "app_label": "classic Outlook on this PC" if PLATFORM == "windows" else "the Outlook app on this Mac",
            "my_email": cfg["my_email"],
            # Where "Send a test to myself" goes.
            "test_address": cfg["my_email"],
            "account": cfg["account"] if PLATFORM == "windows" else "",
            "delay_seconds": cfg["delay_seconds"],
        }

    def check(self):
        """Check everything sending needs (step 1's button). On Windows this also lists
        Outlook's accounts. "ok" is false if any check failed."""
        cfg = self.config()
        if PLATFORM == "windows":
            result = win_outlook.full_check(account=cfg["account"] or None, test_address=cfg["my_email"] or None)
        elif PLATFORM == "mac":
            result = mac_outlook.full_check(test_address=cfg["my_email"] or None)
        else:
            result = {"steps": [{"name": "This computer can run Email Bot", "ok": False,
                                 "detail": _unavailable_reason()}]}
        result["ok"] = all(s["ok"] for s in result["steps"])
        return result

    def send(self, to, cc, subject, html, text, attachments=()):
        try:
            if PLATFORM == "windows":
                win_outlook.send(to, cc, subject, html, text, attachments,
                                 account=self.config()["account"] or None)
            else:
                mac_outlook.send(to, cc, subject, html, text, attachments)
        except (mac_outlook.MacOutlookError, win_outlook.WinOutlookError) as e:
            raise OutlookError(str(e)) from None
