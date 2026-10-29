# Email Bot

Send a form letter to a list of people through the Outlook app on your computer. Each person's details (name, amount, date…) are filled in for them, and you can add optional extra paragraphs or a personal note for specific people.

Everything runs on your own computer. The app opens in your web browser, but it isn't on the internet: only this computer can see it. Your letters and lists are saved in a `data` folder next to the app.

---

## Before you start

Email Bot doesn't send email itself: it hands each email to the Outlook app on your computer, and Outlook sends it. So Outlook has to be set up and working first.

**On a Mac**
- [ ] Microsoft Outlook is installed and signed in to the account you want to send from. New or legacy Outlook both work.

**On Windows**
- [ ] **Classic Outlook** is installed and signed in to the account you want to send from. This is the Outlook that comes with Microsoft 365 / Office. It has a **File** menu in the top-left corner.
- [ ] **The "New Outlook" switch is off.** It's in the top-right corner of classic Outlook. The new Outlook for Windows doesn't let other programs send email through it, so Email Bot can't use it.
- [ ] **Microsoft Office is activated.** In Outlook, open **File → Office Account**: it should say *Product Activated*. If it doesn't, click **Sign in** there with your work or school Microsoft account, then restart Outlook. Signing in to your mailbox isn't the same as activating Office: an unactivated Outlook still shows your email, but nothing can create new ones. Telltale signs are a greyed-out **File → Options** and a **New Email** button that doesn't work.
- [ ] You can write and send an email by hand in Outlook.

---

## Installing and starting Email Bot

### Windows: EmailBot.exe (easiest, no Python needed)

1. Download **EmailBot.exe** from the **Releases** section of this project's GitHub page.
2. Make a folder for it, e.g. `Documents\Email Bot`, and move `EmailBot.exe` there. Email Bot keeps its `data` folder next to the `.exe`, so don't run it straight from Downloads.
3. Open classic Outlook, then double-click `EmailBot.exe`.
   - **"Windows protected your PC"**: click **More info → Run anyway**. This happens because the app isn't code-signed.
   - If your antivirus flags it, that's a known false alarm for apps built this way. Ask whoever manages your PC to allow it.
4. A black window opens and your browser shows Email Bot. **Keep the black window open while you use Email Bot.** Closing it stops Email Bot.

### Windows: start.bat (if you have the full Email Bot folder)

1. If you downloaded a zip, right-click it and choose **Extract All** first.
2. Double-click `start.bat`.
3. The first time, it sets itself up inside a `.venv` folder and downloads what it needs, which takes a minute or two and needs internet. It needs **Python 3.9 or newer**. If it says Python is missing, install it from <https://www.python.org/downloads/windows/>, tick **"Add python.exe to PATH"** during installation, and double-click `start.bat` again.

### Mac: start.command

1. Double-click `start.command`.
   - **"start.command can't be opened" / "unidentified developer"**: open **System Settings → Privacy & Security**, scroll down, click **Open Anyway** next to the message about start.command, then double-click it again. This only happens once.
2. The first time, it sets itself up inside a `.venv` folder and downloads what it needs, which takes a minute or two and needs internet but no administrator password. It needs Python, which most Macs already have. If it says Python is missing, install it from <https://www.python.org/downloads/macos/> or run `xcode-select --install` in Terminal.
3. A Terminal window opens and your browser shows Email Bot. **Keep the Terminal window open while you use Email Bot.**

If the browser doesn't open by itself, go to <http://127.0.0.1:5050>.

---

## Using Email Bot: the 4 steps

### 1. Connect Outlook

1. Type your own email address (where test emails go) and click **Save address**.
2. Click **Check that Outlook responds**. It runs a full check and shows a ✓ or ✗ for each part. If anything fails, the ✗ line says what to fix. Fix it, then click the button again.
   - **Mac:** the first time, your Mac asks whether Terminal may control "Microsoft Outlook". Click **OK**.
   - **Windows:** once the check passes, a **Send from** list appears. Pick the account to send from, or leave it on "Outlook's default account".

What the check covers:

| Check | Mac | Windows |
|---|---|---|
| Outlook is installed | ✓ | ✓ (classic Outlook) |
| The "New Outlook" switch is off | | ✓ |
| Classic Outlook is open | | ✓ |
| Outlook responds, and Email Bot is allowed to control it | ✓ | ✓ |
| Outlook has an email account, and the chosen **Send from** account still exists | | ✓ |
| Outlook can create an email: a complete test email (To, CC, attachment) is built and thrown away, **not sent** | ✓ | ✓ |

On a Mac, the deleted test email may show up in Outlook's **Deleted Items**; you can delete it from there.

On a Mac, emails are sent from the account Outlook uses for **new** emails. If Outlook has several accounts, make the right one the default in Outlook (**Settings → Accounts**). Email Bot can't choose the account on a Mac, because new Outlook doesn't let other apps see its accounts.

### 2. Write the letter

- Type the subject, letter, and sign-off.
- To fill in a personal detail, click where it should go, then click one of the `{…}` buttons, e.g. `{First Name}`.
- **Extra paragraphs (addendums)** are for text only *some* people should get, like "Late payment reminder". Put `{Addendums}` in the letter where they should appear. If you leave it out, they're added at the end.
- You can save several letters and switch between them.

### 3. Add recipients

- Easiest way: in Excel, select your rows **including the title row**, copy, then click **Paste from Excel**. You can also open an Excel (`.xlsx`), `.csv`, or `.tsv` file.
- Before anything is added, a window shows every column with a few example values. Check which column has the email addresses (and the CC column, if any), untick any columns you don't need, and choose whether to replace the current list or add to it.
- Every column other than email and CC (e.g. *First Name*, *Organization*) becomes a `{…}` detail you can use in the letter.
- In the **Extras** column, tick which extra paragraphs each person gets, or add a personal note just for them.
- Untick **Send?** to leave someone out without deleting them.
- To get rid of columns you don't need, click **Remove columns**, tick them, and click **Remove**. Changed your mind? Click **Undo** in the message that appears.

### 4. Review & send

- Click through every email to see exactly what each person will get.
- Anything marked **Needs fixing** (red) must be fixed before you can send, e.g. a missing email address or a `{Detail}` with no matching column.
- Use **Send a test to myself** first, and check it arrives from the right account. Then click **Send emails**.
- Email Bot remembers who already got each letter and skips them, so you won't send twice by accident. Tick "Also send again…" if you really do want to resend.
- Sent emails appear in Outlook's **Sent** folder. **Sending history** at the bottom of the page lists every email sent from Email Bot. "Sent" means Outlook accepted the email; if something looks off, check Outlook's **Outbox** and **Sent** folders.

---

## Troubleshooting

Start with **Check that Outlook responds** in step 1: its ✗ line usually says exactly what to do.

| What you see | What to do |
|---|---|
| **"Outlook refused to let Email Bot create an email"**, or "The operation failed" (Windows) | Most often Office isn't activated. In Outlook, open **File → Office Account** and sign in to activate, then restart Outlook. If **File → Options** is greyed out, this is almost certainly the cause. If Office *is* activated, check **File → Options → Trust Center → Trust Center Settings → Programmatic Access**. |
| **"Classic Outlook isn't installed on this PC"** | You only have the new Outlook for Windows. Email Bot needs classic Outlook; ask whoever manages your PC to install it. |
| **"The New Outlook switch … is on"** | Switch off **New Outlook** in the top-right corner of classic Outlook, then check again. |
| **"Classic Outlook is open" ✗**, or "Outlook doesn't seem to be open" | Open classic Outlook, wait until your inbox appears, then try again. |
| **"Outlook couldn't be reached"** (Windows) | Close any pop-up in Outlook. If Outlook was started with "Run as administrator", close it and open it normally. |
| **"a program is trying to send an e-mail message on your behalf"** (Windows Outlook pop-up) | Click **Allow**. Outlook shows this when it can't confirm your antivirus is up to date. |
| **"Your Mac blocked Email Bot from controlling Outlook"** | Open **System Settings → Privacy & Security → Automation**, find **Terminal**, and switch on **Microsoft Outlook**. |
| **"The account … isn't in Outlook anymore"** (Windows) | Click **Check that Outlook responds** and pick another account under **Send from**. |
| An add-on was **"missed"** during setup | Email Bot still starts, but that feature (e.g. opening `.xlsx` files) is unavailable until the add-on installs. It retries each time Email Bot starts. You can paste from Excel instead. |
| **"Setup couldn't finish"** | Check your internet connection and start Email Bot again. If it keeps failing, send whoever helps you the `setup.log` file named in the message. |
| Something seems stuck or broken after an update | Close Email Bot, delete the `.venv` folder in the Email Bot folder, and start it again to set up from scratch. Your letters and lists in `data` are kept. |

### Getting details for whoever helps you

- **The log:** problems are printed in the black/Terminal window and saved to `data/emailbot.log`.
- **More detail:** start Email Bot with the environment variable `EMAILBOT_DEBUG=1` to also see every Outlook step as it happens. On Windows, in Command Prompt in the Email Bot folder: `set EMAILBOT_DEBUG=1` then `start.bat`.
- **Step-by-step Windows check:** it reports which Outlook is running, whether the "New Outlook" switch is on, the Outlook profile and mailbox, and then tries each step separately, saving drafts rather than sending. Close Email Bot, keep classic Outlook open, open Command Prompt in the Email Bot folder, and run one of these:
  - With EmailBot.exe: `EmailBot.exe --diagnose`
  - With start.bat: `.venv\Scripts\python -m emailbot.diagnose_windows`

  Add `--account you@example.com` to also test sending from a specific account, or `--send-to you@example.com` to send one real test email. Delete the drafts it creates afterwards.

---

## Good to know

- **Sending limits:** Outlook limits how many emails you can send. Personal Outlook.com accounts allow roughly 300 per day; work and school accounts usually allow far more, but no more than about 30 per minute. Email Bot waits 3 seconds between emails by default (change it under **Sending speed** in step 1). For very large lists, send in batches.
- **Attachments** go to everyone and are limited to 3 MB in total.
- **Where things are saved:** in the `data` folder: letters and the recipient list (`workspace.json`), the sending history (`send_log.csv`, opens in Excel), settings, and the log. It contains your recipients' details, so don't share it. Delete it to start completely fresh.
- **Formatting:** letters are plain text, and line breaks are kept exactly as you type them. From new Outlook on a Mac they arrive as a plain-text email in the recipient's usual font; from legacy Mac Outlook and classic Windows Outlook they use Calibri.

---

## For developers

- Python 3.9+ (Apple's built-in Python works). Flask is required; openpyxl is optional (Excel import is disabled without it); pywin32 is required on Windows only. `pip install -r requirements.txt && python app.py`
- `start.command` / `start.bat` find a Python, create `.venv`, then run `emailbot/setup_deps.py`, which installs `requirements.txt` one line at a time. Lines commented `# optional: <feature>` may fail without stopping startup, and lines whose environment marker doesn't match (e.g. `sys_platform == 'win32'`) are skipped. Setup counts as finished once every line installs, and is redone when `requirements.txt` changes.
- **Mac:** sending scripts the Outlook app via `osascript` (`emailbot/mac_outlook.py`); everything varying is passed as script arguments, not interpolated. New Outlook gets the plain-text body, because it delivers scripted HTML bodies with the tags showing.
- **Windows:** sending automates classic Outlook over COM with pywin32 (`emailbot/win_outlook.py`): `GetNamespace("MAPI").Logon`, `CreateItem` (falling back to `Items.Add` on the Drafts folder), `HTMLBody`, `Attachments.Add`, optional `SendUsingAccount` (set by reference), `Send`. COM is initialized per call because sending runs on a background thread. Errors are logged with the step that failed and the full HRESULT/excepinfo. The new Outlook for Windows has no COM interface.
- **Step 1's check** (`full_check` in each sender) ends by building a complete email and discarding it (`Close(olDiscard)` on Windows, `delete` on a Mac); `diagnose_windows.py` runs the Windows steps individually.
- The server listens on `127.0.0.1` only, rejects other `Host` headers (DNS rebinding), and requires an `X-EmailBot` header on API writes, so other websites can't trigger a send.
- Command line: `python app.py --diagnose [...]` runs the Windows diagnostic; `python app.py --check-bundle` checks the page files and add-ons load (used by the Windows build).
- Environment variables: `EMAILBOT_PORT` (default 5050), `EMAILBOT_DATA` (data folder), `EMAILBOT_NO_BROWSER=1`, `EMAILBOT_DEBUG=1`.
- `start.bat` must keep Windows (CRLF) line endings; `.gitattributes` enforces this.
- Code: `app.py` (routes, send job, command line), `emailbot/merge.py` (placeholders → email), `emailbot/importer.py` (CSV/TSV/Excel/paste), `emailbot/outlook.py` (settings, picks the Mac or Windows sender), `emailbot/mac_outlook.py`, `emailbot/win_outlook.py`, `emailbot/diagnose_windows.py`, `emailbot/setup_deps.py` (first-run installs), `static/app.js` (UI).

### Building EmailBot.exe

GitHub builds it automatically on a Windows machine (`.github/workflows/build-windows.yml`), using the PyInstaller recipe in `EmailBot.spec`:

- **Every push to `main`:** the workflow builds `EmailBot.exe`, runs `EmailBot.exe --check-bundle`, and attaches the `.exe` to the run. Open the repo's **Actions** tab → the latest "Build Windows app" run → **Artifacts** → `EmailBot-windows`.
- **A version tag** (`git tag v1.0.0` then `git push origin v1.0.0`): the same build, plus a GitHub **Release** with `EmailBot.exe` attached, which is where users download it.
- **By hand on Windows:** `pip install -r requirements.txt pyinstaller` then `pyinstaller EmailBot.spec`; the result is `dist\EmailBot.exe`.

When running as `EmailBot.exe`, the page files are read from the unpacked bundle (`sys._MEIPASS`) and data is kept in `data\` next to the `.exe`.
