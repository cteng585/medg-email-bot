# Email Bot

Send a form letter to a list of people from your Outlook account. Each person's details (name, amount, date…) are filled in for them, and you can add optional extra paragraphs or a personal note for specific people.

Everything runs on your own computer. The app opens in your web browser, but it isn't on the internet: only this computer can see it, and your letters and lists are saved in the `data` folder here.

---

## Starting the app

**Mac:** double-click `start.command`.

A Terminal window appears and your browser opens Email Bot. **Keep the Terminal window open while you use the app.** Close it when you're done.

**The first time** takes a minute or two. Email Bot sets itself up inside a `.venv` folder in the Email Bot folder and downloads the add-ons it needs. This needs an internet connection but no administrator password, and nothing is installed anywhere else on the Mac. After that, it starts in a few seconds.

- **"start.command can't be opened" / "unidentified developer":** open **System Settings → Privacy & Security**, scroll down, and click **Open Anyway** next to the message about start.command. Then double-click it again. This only happens once.
- **"Email Bot needs Python":** most Macs already have it. If yours doesn't, install Python from <https://www.python.org/downloads/macos/>, or run `xcode-select --install` in Terminal. Both may need an administrator password, so ask whoever manages your Mac if you don't have one.
- **An add-on is "missed":** Email Bot still starts, but that one feature is unavailable (for example, opening `.xlsx` files, where you can still paste from Excel instead). It tries again each time you start Email Bot.
- **"Setup couldn't finish":** check your internet connection and double-click `start.command` again. If it keeps failing, send whoever helps you the `setup.log` file named in the message.
- **Something seems stuck or broken:** close Email Bot, delete the `.venv` folder inside the Email Bot folder, and double-click `start.command` to set up from scratch. Your letters and lists in `data` are kept.

**Windows:** double-click `start.bat`. If you're told Python is missing, install it from <https://www.python.org/downloads/> (tick **"Add Python to PATH"**) and try again.

If the browser doesn't open by itself, go to <http://127.0.0.1:5050>.

---

## Using it: the 4 steps

1. **Connect Outlook.** Choose how emails are sent:
   - **Outlook app on this Mac:** no setup. Emails go out through the Outlook app, from the account it uses for new emails. See [Sending with the Outlook app on a Mac](#sending-with-the-outlook-app-on-a-mac).
   - **Sign in with Microsoft:** works on any computer, but needs the [one-time setup](#one-time-microsoft-sign-in-setup) first. You stay signed in between uses.
2. **Write the letter.**
   - Type the subject, letter, and sign-off.
   - To fill in a personal detail, click where it should go, then click one of the `{…}` buttons, e.g. `{First Name}`.
   - **Extra paragraphs (addendums)** are for text only *some* people should get, like "Late payment reminder". Put `{Addendums}` in the letter where they should appear. If you leave it out, they're added at the end.
   - You can save several letters and switch between them.
3. **Add recipients.**
   - Easiest way: in Excel, select your rows **including the title row**, copy, then click **Paste from Excel** here. You can also open an Excel (`.xlsx`), `.csv`, or `.tsv` file.
   - Before anything is added, a window shows every column with a few example values. Check which column has the email addresses (and the CC column, if any), untick any columns you don't need, and choose whether to replace the current list or add to it.
   - You need one column titled **Email**. A column titled **CC** is optional. Every other column (e.g. *First Name*, *Organization*) becomes a `{…}` detail you can use in the letter.
   - In the **Extras** column, tick which extra paragraphs each person gets, or add a personal note just for them.
   - Untick **Send?** to leave someone out without deleting them.
   - To get rid of columns you don't need, click **Remove columns**, tick them, and click **Remove**. Changed your mind? Click **Undo** in the message that appears.
4. **Review & send.**
   - Click through every email to see exactly what each person will get.
   - Anything marked **Needs fixing** (red) must be fixed before you can send, e.g. a missing email address or a `{Detail}` with no matching column.
   - Use **Send a test to myself** first, then **Send emails**.
   - The app remembers who already got each letter and skips them, so you won't send twice by accident. Tick "Also send again…" if you really do want to resend.
   - Sent emails also appear in your Outlook **Sent Items** folder. **Sending history** at the bottom of the page shows every email sent from the app.

---

## Sending with the Outlook app on a Mac

This is the easiest option on a Mac. It needs no Microsoft or Azure setup, only the Outlook app, signed in to the account you want to send from.

1. Open Outlook and leave it open while you use Email Bot.
2. In Email Bot step 1, choose **Outlook app on this Mac**, type your own email address (for test emails), and click **Save address**.
3. Click **Check that Outlook responds**. The first time, your Mac asks whether Terminal may control "Microsoft Outlook". Click **OK**.
   If you clicked "Don't Allow" by mistake: open **System Settings → Privacy & Security → Automation**, find **Terminal**, and switch on **Microsoft Outlook**.
4. In step 4, use **Send a test to myself** and check that it arrives from the right account.

Good to know:
- Emails are sent from the account Outlook uses for **new** emails. If Outlook has several accounts, make that one the default in Outlook (**Settings → Accounts**). Email Bot can't choose the account itself, because new Outlook doesn't let other apps see its accounts. If you need to pick the sending account, use **Sign in with Microsoft** instead.
- Email Bot hands each email to Outlook, and Outlook does the sending. "Sent" in Email Bot means Outlook accepted it. If something looks off, check Outlook's **Outbox** and **Sent** folders.
- This option isn't available on Windows yet. Use **Sign in with Microsoft** there.

---

## One-time Microsoft sign-in setup

Only needed for the **Sign in with Microsoft** option. Microsoft requires every app that sends from Outlook this way to be registered once. This takes about 5 minutes. Whoever handles IT can do it once and share the resulting ID with everyone using the app.

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
6. In Email Bot, go to **Step 1**, choose **Sign in with Microsoft**, open **Microsoft sign-in settings**, paste the client ID, choose the matching **Account type**, and click **Save settings**. Then click **Sign in with Microsoft**.

The app can only send mail as the person who signed in. It can't read your inbox.

---

## Good to know

- **Sending limits:** Outlook limits how many emails you can send. Personal Outlook.com accounts allow roughly 300 per day; work accounts usually allow far more, but no more than about 30 per minute. The app waits 3 seconds between emails by default (you can change this under **Sending speed** in step 1). For very large lists, send in batches.
- **Attachments** go to everyone and are limited to 3 MB in total.
- **Where things are saved:** in the `data` folder: letters and the recipient list (`workspace.json`), the sending history (`send_log.csv`, opens in Excel), and your Outlook sign-in. Don't share the `data` folder, since it contains your sign-in. Delete it to start completely fresh.
- **Formatting:** letters are plain text, and line breaks are kept exactly as you type them. With **Sign in with Microsoft** they use a standard Outlook font (Calibri); with the **Outlook app on a Mac** (new Outlook) they arrive as a plain-text email in the recipient's usual font.

---

## For developers

- Python 3.9+ (Apple's built-in Python works). Flask is required; openpyxl, requests and MSAL are optional, and the app disables the matching feature when one is missing. `pip install -r requirements.txt && python app.py`
- `start.command` installs `requirements.txt` one line at a time into `.venv`. Lines commented `# optional: <feature>` may fail without stopping startup. Setup counts as finished (and is skipped on later starts) once every line installs; it's redone when `requirements.txt` changes.
- Two senders, chosen in step 1 (`method` in `data/config.json`): `mac_app` scripts the Outlook app via `osascript` (`emailbot/mac_outlook.py`); everything varying is passed as script arguments, not interpolated. `graph` sends through Microsoft Graph `POST /me/sendMail` with delegated `Mail.Send`; sign-in uses MSAL's interactive flow with a localhost redirect, and tokens are cached in `data/token_cache.json`.
- The server listens on `127.0.0.1` only, rejects other `Host` headers (DNS rebinding), and requires an `X-EmailBot` header on API writes, so other websites can't trigger a send.
- Environment variables: `EMAILBOT_PORT` (default 5050), `EMAILBOT_DATA` (data folder), `EMAILBOT_NO_BROWSER=1`.
- Code: `app.py` (routes, send job), `emailbot/merge.py` (placeholders → email), `emailbot/importer.py` (CSV/Excel/paste), `emailbot/outlook.py` (settings, Graph auth + send, picks the sender), `emailbot/mac_outlook.py` (Mac Outlook app), `static/app.js` (UI).
