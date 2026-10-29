"""Signing in to Microsoft and sending mail through the Microsoft Graph API."""

import base64
import mimetypes
import threading
import time

import msal
import requests

from . import storage

SCOPES = ["Mail.Send", "User.Read"]
GRAPH = "https://graph.microsoft.com/v1.0"


class OutlookError(Exception):
    pass


class Outlook:
    def __init__(self):
        self._cache = msal.SerializableTokenCache()
        cache_path = storage._path("token_cache.json")
        if cache_path.exists():
            self._cache.deserialize(cache_path.read_text(encoding="utf-8"))
        self._client = None
        self._login_thread = None
        self.login_error = ""

    # ---- configuration ----

    def config(self):
        cfg = storage.read_json("config.json", {})
        cfg.setdefault("client_id", "")
        cfg.setdefault("tenant", "common")
        cfg.setdefault("delay_seconds", 3)
        return cfg

    def save_config(self, **changes):
        cfg = self.config()
        cfg.update(changes)
        storage.write_json("config.json", cfg)
        self._client = None

    def _app(self):
        cfg = self.config()
        if not cfg["client_id"]:
            raise OutlookError("Outlook isn't set up yet. Add the Application (client) ID in step 1.")
        if self._client is None:
            self._client = msal.PublicClientApplication(
                cfg["client_id"],
                authority=f"https://login.microsoftonline.com/{cfg['tenant'] or 'common'}",
                token_cache=self._cache,
            )
        return self._client

    def _save_cache(self):
        if self._cache.has_state_changed:
            storage.write_json_text("token_cache.json", self._cache.serialize())

    # ---- sign in / out ----

    def account(self):
        try:
            accounts = self._app().get_accounts()
        except OutlookError:
            return None
        return accounts[0] if accounts else None

    def status(self):
        cfg = self.config()
        acct = self.account() if cfg["client_id"] else None
        return {
            "configured": bool(cfg["client_id"]),
            "client_id": cfg["client_id"],
            "tenant": cfg["tenant"],
            "delay_seconds": cfg["delay_seconds"],
            "signed_in": acct is not None,
            "account": acct["username"] if acct else "",
            "login_in_progress": bool(self._login_thread and self._login_thread.is_alive()),
            "login_error": self.login_error,
        }

    def start_login(self):
        if self._login_thread and self._login_thread.is_alive():
            return
        app = self._app()
        self.login_error = ""

        def run():
            try:
                result = app.acquire_token_interactive(SCOPES, prompt="select_account", timeout=300)
                if "access_token" not in result:
                    self.login_error = result.get("error_description") or result.get("error") or "Sign-in failed."
            except Exception as e:  # noqa: BLE001 - surfaced to the user as text
                self.login_error = str(e)
            finally:
                self._save_cache()

        self._login_thread = threading.Thread(target=run, daemon=True)
        self._login_thread.start()

    def logout(self):
        try:
            app = self._app()
            for acct in app.get_accounts():
                app.remove_account(acct)
        except OutlookError:
            pass
        self._save_cache()

    def _token(self):
        app = self._app()
        acct = self.account()
        if not acct:
            raise OutlookError("You're not signed in to Outlook. Go to step 1 and sign in.")
        result = app.acquire_token_silent(SCOPES, account=acct)
        self._save_cache()
        if not result or "access_token" not in result:
            raise OutlookError("Your Outlook sign-in has expired. Go to step 1 and sign in again.")
        return result["access_token"]

    # ---- sending ----

    def send(self, to, cc, subject, html, attachments=()):
        message = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": to}}],
            "ccRecipients": [{"emailAddress": {"address": a}} for a in cc],
        }
        if attachments:
            message["attachments"] = [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": path.name,
                    "contentType": mimetypes.guess_type(path.name)[0] or "application/octet-stream",
                    "contentBytes": base64.b64encode(path.read_bytes()).decode("ascii"),
                }
                for path in attachments
            ]

        for attempt in range(4):
            resp = requests.post(
                f"{GRAPH}/me/sendMail",
                headers={"Authorization": f"Bearer {self._token()}"},
                json={"message": message, "saveToSentItems": True},
                timeout=60,
            )
            if resp.status_code == 202:
                return
            if resp.status_code in (429, 503, 504) and attempt < 3:
                time.sleep(min(int(resp.headers.get("Retry-After", 10)), 120))
                continue
            try:
                detail = resp.json()["error"]["message"]
            except Exception:  # noqa: BLE001
                detail = resp.text[:300]
            raise OutlookError(f"Outlook refused the message ({resp.status_code}): {detail}")
