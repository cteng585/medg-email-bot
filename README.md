# Email Bot

Send a form letter to a list of people from your Outlook account. Each person's details (name, amount, date…) are filled in for them, and you can add optional extra paragraphs or a personal note for specific people.

Everything runs on your own computer. The app opens in your web browser, but it isn't on the internet: only this computer can see it, and your letters and lists are saved in the `data` folder here.

---

## Starting the app

**Mac:** double-click `start.command`.
(The first time, macOS may say it can't be opened. If so, right-click the file → **Open** → **Open**.)

**Windows:** double-click `start.bat`.

A black terminal window appears and your browser opens Email Bot. **Keep the terminal window open while you use the app.** Close it when you're done.

The first start takes about a minute while it installs what it needs. If you're told Python is missing, install it from <https://www.python.org/downloads/> (on Windows, tick **"Add Python to PATH"** during install) and try again.

If the browser doesn't open by itself, go to <http://127.0.0.1:5050>.

---

## Using it: the 4 steps

1. **Connect Outlook.** Click **Sign in with Microsoft** and sign in with the account you want to send from. You stay signed in between uses.
2. **Write the letter.**
   - Type the subject, letter, and sign-off.
   - To fill in a personal detail, click where it should go, then click one of the `{…}` buttons, e.g. `{First Name}`.
   - **Extra paragraphs (addendums)** are for text only *some* people should get, like "Late payment reminder". Put `{Addendums}` in the letter where they should appear. If you leave it out, they're added at the end.
   - You can save several letters and switch between them.
3. **Add recipients.**
   - Easiest way: in Excel, select your rows **including the title row**, copy, then click **Paste from Excel** here. You can also open a `.xlsx` or `.csv` file.
   - You need one column titled **Email**. A column titled **CC** is optional. Every other column (e.g. *First Name*, *Organization*) becomes a `{…}` detail you can use in the letter.
   - In the **Extras** column, tick which extra paragraphs each person gets, or add a personal note just for them.
   - Untick **Send?** to leave someone out without deleting them.
4. **Review & send.**
   - Click through every email to see exactly what each person will get.
   - Anything marked **Needs fixing** (red) must be fixed before you can send, e.g. a missing email address or a `{Detail}` with no matching column.
   - Use **Send a test to myself** first, then **Send emails**.
   - The app remembers who already got each letter and skips them, so you won't send twice by accident. Tick "Also send again…" if you really do want to resend.
   - Sent emails also appear in your Outlook **Sent Items** folder. **Sending history** at the bottom of the page shows every email sent from the app.

---

## One-time Outlook setup

Microsoft requires every app that sends from Outlook to be registered once. This takes about 5 minutes. Whoever handles IT can do it once and share the resulting ID with everyone using the app.

1. Go to <https://entra.microsoft.com> (or <https://portal.azure.com>) and sign in.
   - **Work/school account:** use your work login. If you're not allowed to register apps, ask your IT admin to follow these steps.
   - **Personal Outlook.com/Hotmail account:** you need a free Azure account to register an app. Create one at <https://azure.microsoft.com/free>.
2. Go to **App registrations** (under *Identity → Applications*, or search for it at the top) → **New registration**.
3. Fill in:
   - **Name:** `Email Bot`
   - **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts*. If you only use work accounts, choose *this organizational directory only*.
   - **Redirect URI:** choose platform **Public client/native (mobile & desktop)** and enter `http://localhost`
   - Click **Register**.
4. On the page that opens, copy the **Application (client) ID**.
5. Click **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → search for and tick **Mail.Send** → **Add permissions**.
   (*User.Read* is already there; leave it.)
   Some workplaces require an admin to click **Grant admin consent**. If sign-in later says you need admin approval, ask your IT admin to do this.
6. In Email Bot, go to **Step 1 → Connection settings**, paste the client ID, choose the matching **Account type**, and click **Save settings**. Then click **Sign in with Microsoft**.

The app can only send mail as the person who signed in. It can't read your inbox.

---

## Good to know

- **Sending limits:** Outlook limits how many emails you can send. Personal Outlook.com accounts allow roughly 300 per day; work accounts usually allow far more, but no more than about 30 per minute. The app waits 3 seconds between emails by default (you can change this in Connection settings). For very large lists, send in batches.
- **Attachments** go to everyone and are limited to 3 MB in total.
- **Where things are saved:** in the `data` folder: letters and the recipient list (`workspace.json`), the sending history (`send_log.csv`, opens in Excel), and your Outlook sign-in. Don't share the `data` folder, since it contains your sign-in. Delete it to start completely fresh.
- **Formatting:** letters are sent as plain text in a standard Outlook font. Line breaks are kept exactly as you type them.

---

## For developers

- Python 3.10+, Flask, MSAL, openpyxl. `pip install -r requirements.txt && python app.py`
- Sends through Microsoft Graph `POST /me/sendMail` with delegated `Mail.Send`; sign-in uses MSAL's interactive flow with a localhost redirect, and tokens are cached in `data/token_cache.json`.
- The server listens on `127.0.0.1` only, rejects other `Host` headers (DNS rebinding), and requires an `X-EmailBot` header on API writes, so other websites can't trigger a send.
- Environment variables: `EMAILBOT_PORT` (default 5050), `EMAILBOT_DATA` (data folder), `EMAILBOT_NO_BROWSER=1`.
- Code: `app.py` (routes, send job), `emailbot/merge.py` (placeholders → email), `emailbot/importer.py` (CSV/Excel/paste), `emailbot/outlook.py` (auth + Graph), `static/app.js` (UI).
