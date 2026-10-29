"""Step-by-step check of sending through classic Outlook on Windows.

Run from the Email Bot folder, after start.bat has set things up once:

    .venv\\Scripts\\python -m emailbot.diagnose_windows
    .venv\\Scripts\\python -m emailbot.diagnose_windows --account you@school.edu
    .venv\\Scripts\\python -m emailbot.diagnose_windows --send-to you@school.edu

or, with the Windows app:   EmailBot.exe --diagnose

By default nothing is sent: the test emails are saved to Outlook's Drafts
folder (delete them afterwards). --send-to also sends one real email there.
"""

import argparse
import logging
import platform
import sys
import tempfile
from pathlib import Path

from . import win_outlook as W
from .merge import to_html

TEXT = "Email Bot diagnostic.\n\nLine two, with an ampersand & and an accent: café."


def run(title, fn):
    print(f"\n== {title}")
    try:
        result = fn()
        print("   [ok]")
        return result
    except W.WinOutlookError as e:
        print(f"   [FAILED] {e}")
    except Exception as e:  # noqa: BLE001 - this script exists to show whatever goes wrong
        print(f"   [FAILED] unexpected {type(e).__name__}: {e}")
    return None


def _yes_no(value):
    return "unknown" if value is None else ("yes" if value else "no")


def _is_admin():
    try:
        import ctypes
        return "yes" if ctypes.windll.shell32.IsUserAnAdmin() else "no"
    except Exception:  # noqa: BLE001
        return "unknown"


def print_state():
    """What's running before this script touches Outlook, then what Outlook reports."""
    print(f"Classic Outlook (OUTLOOK.EXE) running before this check: {_yes_no(W.process_running('OUTLOOK.EXE'))}")
    print(f"New Outlook (olk.exe) running: {_yes_no(W.process_running('olk.exe'))}")
    print(f"\"New Outlook\" switch in classic Outlook: {'on' if W.new_outlook_switch_on() else 'off'}")
    print(f"This window running as administrator: {_is_admin()}")

    def ask(app, step):
        found = {}
        for label, get in [
            ("Outlook windows open", lambda: app.Explorers.Count),
            ("Profile", lambda: app.GetNamespace("MAPI").CurrentProfileName),
            ("Default mailbox", lambda: app.GetNamespace("MAPI").DefaultStore.DisplayName),
            ("Exchange connection mode", lambda: app.GetNamespace("MAPI").ExchangeConnectionMode),
        ]:
            try:
                found[label] = get()
            except Exception as e:  # noqa: BLE001 - report, don't stop
                found[label] = f"couldn't read ({W.describe(e) if hasattr(e, 'args') else e})"
        return found

    state = run("Ask Outlook about its setup", lambda: W._with_outlook(ask))
    for label, value in (state or {}).items():
        print(f"   {label}: {value}")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--account", help="also test sending from this Outlook account")
    parser.add_argument("--send-to", help="also send one real email to this address")
    args = parser.parse_args()

    # Show each Outlook call as it happens, and the full details of any error.
    logging.basicConfig(level=logging.DEBUG, format="   %(message)s", stream=sys.stdout)

    print(f"Windows: {platform.platform()}  Python: {sys.version.split()[0]}")
    if sys.platform != "win32":
        print("This check is for Windows.")
        return 1
    try:
        import win32com  # noqa: F401
        print("pywin32: installed")
    except ImportError:
        print("pywin32: MISSING - run start.bat while connected to the internet, then try again.")
        return 1
    print(f"Classic Outlook registered: {'yes' if W.available() else 'NO'}")
    print_state()

    info = run("Talk to Outlook and list its accounts", W.check)
    if info:
        print(f"   Outlook version {info['version']}")
        for a in info["accounts"]:
            print(f"   account: {a['name']} <{a['address']}>")

    attachment = Path(tempfile.gettempdir()) / "emailbot-diagnostic.txt"
    attachment.write_text("Email Bot diagnostic attachment\n", encoding="utf-8")
    html = to_html(TEXT)
    to = args.send_to or "emailbot-test@example.com"

    def draft(**kw):
        return lambda: W.send(to, [], "Email Bot diagnostic (draft)", html, TEXT, mode="draft", **kw)

    run("Save a plain draft (no attachment, default account)", draft())
    run("Save a draft with an attachment", draft(attachments=[attachment]))
    if args.account:
        run(f"Save a draft from the account {args.account}", draft(account=args.account))
    if args.send_to:
        run(f"Send one real email to {args.send_to}",
            lambda: W.send(args.send_to, [], "Email Bot diagnostic", html, TEXT,
                           attachments=[attachment], account=args.account))

    print("\nDone. Drafts made by this check are in Outlook's Drafts folder; you can delete them.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
