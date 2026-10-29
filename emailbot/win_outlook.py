"""Sending through classic Outlook for Windows, by automating it over COM (pywin32).

Classic Outlook (the version that comes with Microsoft 365 / Office) can be
controlled by other programs; the new Outlook for Windows can't, so it isn't
supported. Unlike new Outlook on a Mac, classic Outlook can list its accounts
and send from a chosen one.

pywin32 is imported only when needed, so this module loads on any platform.
"""

import logging
import subprocess
import sys
import tempfile
from pathlib import Path

log = logging.getLogger("emailbot.outlook")

OL_MAIL_ITEM = 0
OL_FOLDER_DRAFTS = 16
OL_DISCARD = 1
DISPID_SEND_USING_ACCOUNT = 64209
DISPATCH_PROPERTYPUTREF = 8

NOT_INSTALLED = (
    "Classic Outlook isn't installed on this PC. Email Bot works with classic Outlook "
    "(the one that comes with Microsoft 365 / Office), but not with the new Outlook for Windows."
)
SETUP_MISSING = (
    "Part of Email Bot's setup is missing. Close Email Bot and double-click start.bat again "
    "while connected to the internet."
)

# COM error codes worth explaining in plain words.
_FRIENDLY = {
    -2147221005: NOT_INSTALLED,  # CO_E_CLASSSTRING: "Outlook.Application" isn't registered
    -2147221164: NOT_INSTALLED,  # REGDB_E_CLASSNOTREG
    -2146959355: ("Outlook couldn't be reached. If Outlook is showing a pop-up, close it. If Outlook "
                  "was started with “Run as administrator”, close it and open it normally. Then try again."),
    # E_ABORT: Outlook refused to create an email. Seen in practice when Office wasn't
    # activated (reduced functionality mode); also what programmatic-access security returns.
    -2147467260: ("Outlook refused to let Email Bot create an email. Usually this means Microsoft Office "
                  "isn't activated on this PC (a sign is that File → Options is greyed out in Outlook): "
                  "open File → Office Account in Outlook and sign in to activate it, then restart Outlook. "
                  "If Office is activated, check File → Options → Trust Center → Trust Center Settings → "
                  "Programmatic Access. If Outlook asked whether to allow a program to send email, click "
                  "Allow next time."),
}


class WinOutlookError(Exception):
    pass


def _registered():
    try:
        import winreg
    except ImportError:
        return False
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Outlook.Application\CLSID"))
        return True
    except OSError:
        return False


def available():
    return sys.platform == "win32" and _registered()


def _friendly(e):
    code = e.args[0] if e.args else None
    if code in _FRIENDLY:
        return _FRIENDLY[code]
    detail = ""
    if len(e.args) > 2 and e.args[2]:
        detail = e.args[2][2] or ""  # excepinfo description
    return f"Outlook reported a problem: {detail or (e.args[1] if len(e.args) > 1 else e)}"


def describe(e):
    """Everything a COM error carries, for the log: codes, source and message."""
    args = list(e.args) + [None] * 4
    code, text, excepinfo = args[0], args[1], args[2]
    parts = [f"HRESULT {_hex(code)} ({text})"]
    if excepinfo:
        info = list(excepinfo) + [None] * 6
        parts.append(f"source={info[1]!r} message={info[2]!r} scode={_hex(info[5])}")
    return "; ".join(parts)


def _hex(code):
    return f"0x{code & 0xFFFFFFFF:08X}" if isinstance(code, int) else repr(code)


def _with_outlook(work, on_step=None):
    """Run work(outlook_app, step) on this thread, with COM set up and torn down around it.

    work calls step("doing something") before each Outlook call, so an error can
    say which call failed. on_step(name) is told about each step as well.
    """
    try:
        import pythoncom
        import pywintypes
        import win32com.client
    except ImportError:
        raise WinOutlookError(SETUP_MISSING) from None

    current = ["starting Outlook"]

    def step(name):
        current[0] = name
        log.debug("Outlook: %s", name)
        if on_step:
            on_step(name)

    pythoncom.CoInitialize()  # sending runs on a background thread, which needs its own COM setup
    app = None
    try:
        step("starting Outlook")
        app = win32com.client.Dispatch("Outlook.Application")
        return work(app, step)
    except pywintypes.com_error as e:
        log.error("Outlook failed while %s: %s", current[0], describe(e))
        raise WinOutlookError(f"{_friendly(e)} (This happened while {current[0]}.){_hint(app)}") from None
    finally:
        app = None  # release Outlook before shutting COM down
        pythoncom.CoUninitialize()


def _hint(app):
    """Extra advice when Outlook was started in the background rather than open on screen."""
    try:
        if app is None or app.Explorers.Count > 0:
            return ""
    except Exception:  # noqa: BLE001 - only advice; never hide the real error
        return ""
    return (" Outlook doesn't seem to be open. Open classic Outlook, wait until your inbox shows, "
            "then try again. If you use the new Outlook, switch “New Outlook” off in the top-right "
            "corner of classic Outlook; Email Bot can't send through the new one.")


def _new_mail(app, step, account=None):
    """Create a blank email, making sure Outlook's mailbox is signed in first.

    If Outlook's usual way of creating an email fails (which happens when it was
    started in the background or its default mailbox isn't ready), create it
    straight in a Drafts folder instead.
    """
    import pywintypes

    step("connecting to Outlook's mailbox")
    session = app.GetNamespace("MAPI")
    session.Logon("", "", False, False)  # does nothing if Outlook is already signed in
    step("creating a new email")
    try:
        return app.CreateItem(OL_MAIL_ITEM)
    except pywintypes.com_error as e:
        log.warning("Creating the email the usual way failed (%s); trying the Drafts folder instead.",
                    describe(e))
    step("creating a new email in the Drafts folder")
    store = account.DeliveryStore if account is not None else session.DefaultStore
    return store.GetDefaultFolder(OL_FOLDER_DRAFTS).Items.Add(OL_MAIL_ITEM)


def _accounts(app, step):
    step("listing Outlook's accounts")
    accounts = app.Session.Accounts
    found = []
    for i in range(1, accounts.Count + 1):
        a = accounts.Item(i)
        found.append({"name": a.DisplayName, "address": a.SmtpAddress or a.DisplayName})
    return found


def _find_account(app, step, wanted):
    step(f"looking up the account {wanted}")
    accounts = app.Session.Accounts
    for i in range(1, accounts.Count + 1):
        a = accounts.Item(i)
        if wanted.lower() in ((a.SmtpAddress or "").lower(), (a.DisplayName or "").lower()):
            return a
    raise WinOutlookError(
        f"The account “{wanted}” isn't set up in Outlook anymore. Choose another one in step 1.")


def check(on_step=None):
    """Make sure Outlook answers, and list the accounts it can send from."""
    def work(app, step):
        step("reading Outlook's version")
        version = app.Version
        return {"version": version, "accounts": _accounts(app, step)}

    return _with_outlook(work, on_step)


def send(to, cc, subject, html, text, attachments=(), account=None, mode="send", on_step=None):
    """Send one email. mode "draft" saves it to Drafts instead (diagnose_windows);
    "discard" builds it completely and then throws it away (the step 1 check)."""
    def work(app, step):
        found = _find_account(app, step, account) if account else None
        mail = _new_mail(app, step, found)
        step("setting the subject")
        mail.Subject = subject
        step("setting the To address")
        mail.To = to
        step("setting the CC addresses")
        mail.CC = "; ".join(cc)
        step("setting the email text")
        mail.HTMLBody = html  # classic Outlook sends HTML properly, unlike new Outlook on a Mac
        for p in attachments:
            step(f"attaching {Path(p).name}")
            mail.Attachments.Add(str(Path(p).resolve()))
        if found is not None:
            step(f"choosing {account} as the sending account")
            # A plain "mail.SendUsingAccount = ..." doesn't stick through pywin32's late
            # binding, so set the property by reference.
            mail._oleobj_.Invoke(DISPID_SEND_USING_ACCOUNT, 0, DISPATCH_PROPERTYPUTREF, 0, found)
        if mode == "draft":
            step("saving the email to Drafts")
            mail.Save()
        elif mode == "discard":
            step("discarding the test email")
            mail.Close(OL_DISCARD)
        else:
            step("sending")
            mail.Send()

    _with_outlook(work, on_step)


# ---- the full check behind step 1's "Check that Outlook responds" ----

def process_running(exe):
    """True/False, or None if Windows wouldn't say."""
    try:
        out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {exe}", "/NH"],
                             capture_output=True, text=True, timeout=20).stdout
    except (OSError, subprocess.TimeoutExpired):
        return None
    return exe.lower() in out.lower()


def new_outlook_switch_on():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Office\16.0\Outlook\Preferences") as k:
            return bool(winreg.QueryValueEx(k, "UseNewOutlook")[0])
    except OSError:
        return False


def full_check(account=None, test_address=None):
    """Check everything sending needs, stopping at the first thing that would block it.

    Ends by building a complete test email (To, CC, attachment, chosen account)
    and discarding it unsent. Returns {"steps": [{name, ok, detail}], "version", "accounts"}.
    """
    result = {"steps": []}

    def record(name, ok, detail=""):
        result["steps"].append({"name": name, "ok": bool(ok), "detail": detail})
        return ok

    try:
        import win32com.client  # noqa: F401
    except ImportError:
        record("Email Bot's Outlook add-on is installed", False, SETUP_MISSING)
        return result
    if not record("Classic Outlook is installed", available(), "" if available() else NOT_INSTALLED):
        return result
    switched = new_outlook_switch_on()
    if not record("Classic Outlook is in use (not the new Outlook)", not switched,
                  "The “New Outlook” switch in the top-right corner of classic Outlook is on. Switch it "
                  "off; Email Bot can't send through the new Outlook." if switched else ""):
        return result
    running = process_running("OUTLOOK.EXE")
    if not record("Classic Outlook is open", running is not False,
                  "Open classic Outlook, wait until your inbox shows, then check again."
                  if running is False else ""):
        return result

    try:
        info = check()
    except WinOutlookError as e:
        record("Outlook responds", False, str(e))
        return result
    result.update(info)
    record("Outlook responds", True, f"Version {info['version']}")
    if not record("Outlook has an email account", info["accounts"],
                  "" if info["accounts"] else "Add your email account to classic Outlook first."):
        return result
    if account:
        known = any(account.lower() in (a["address"].lower(), a["name"].lower()) for a in info["accounts"])
        if not record("The chosen “Send from” account is in Outlook", known,
                      "" if known else f"“{account}” isn't in Outlook anymore. Choose another account below."):
            return result

    attachment = Path(tempfile.gettempdir()) / "emailbot-check.txt"
    attachment.write_text("Email Bot check\n", encoding="utf-8")
    address = test_address or "emailbot-check@example.com"
    try:
        send(address, [address], "Email Bot check (not sent)", "<div>Email Bot check</div>", "Email Bot check",
             attachments=[attachment], account=account or None, mode="discard")
        record("Outlook can create an email (a test email was built and thrown away, not sent)", True)
    except WinOutlookError as e:
        record("Outlook can create an email", False, str(e))
    return result
