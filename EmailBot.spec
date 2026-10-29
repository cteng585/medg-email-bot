# PyInstaller recipe for EmailBot.exe. On Windows, from this folder:
#     pip install -r requirements.txt pyinstaller
#     pyinstaller EmailBot.spec
# The result is dist\EmailBot.exe. GitHub builds it automatically
# (see .github/workflows/build-windows.yml).

a = Analysis(
    ["app.py"],
    # The web page itself; app.py finds these inside the unpacked bundle.
    datas=[("templates", "templates"), ("static", "static")],
    # Imported only when needed, so list them explicitly to be safe.
    hiddenimports=["win32com.client", "pythoncom", "pywintypes", "openpyxl", "emailbot.diagnose_windows"],
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    name="EmailBot",
    console=True,  # the window that says "keep this open"; closing it stops Email Bot
    upx=False,  # compressed executables trip antivirus false alarms more often
)
