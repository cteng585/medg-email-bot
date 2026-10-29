"""Installs requirements.txt one line at a time into the Python running this.

start.command (Mac) and start.bat (Windows) run it inside .venv on every start:
    python -m emailbot.setup_deps

A line commented "# optional: <feature>" may fail without stopping Email Bot;
any other line is required. Once everything installs, a copy of
requirements.txt is saved in .venv and later starts skip straight to the app.
Only uses the standard library, since nothing else is installed yet.

Exit codes: 0 all installed, 1 a required package failed, 2 only optional ones failed.
"""

import subprocess
import sys
from pathlib import Path

REQUIREMENTS = Path(__file__).resolve().parent.parent / "requirements.txt"
STAMP = Path(sys.prefix) / ".installed"
LOG = Path(sys.prefix) / "setup.log"


def _applies_here(spec):
    """False for lines limited to another platform, e.g. "pywin32; sys_platform == 'win32'"."""
    if ";" not in spec:
        return True
    try:
        from pip._vendor.packaging.markers import Marker  # pip is always in a fresh venv
        return Marker(spec.split(";", 1)[1]).evaluate()
    except Exception:  # noqa: BLE001 - if unsure, let pip decide
        return True


def main():
    wanted = REQUIREMENTS.read_bytes()
    if STAMP.exists() and STAMP.read_bytes() == wanted:
        return 0

    print("Installing the add-ons Email Bot uses (needs an internet connection)...")
    failed_required = failed_optional = False
    with open(LOG, "a", encoding="utf-8") as log:
        for line in REQUIREMENTS.read_text(encoding="utf-8").splitlines():
            spec, _, note = line.partition("#")
            spec, note = spec.strip(), note.strip()
            if not spec or not _applies_here(spec):
                continue
            log.write(f"\n=== {spec}\n")
            log.flush()
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "--no-input",
                 "--retries", "2", "--timeout", "20", spec],
                stdin=subprocess.DEVNULL, stdout=log, stderr=subprocess.STDOUT,
            )
            name = spec.split(";")[0].strip()
            if result.returncode == 0:
                print(f"  [ok]     {name}")
            elif note.startswith("optional"):
                feature = note.partition(":")[2].strip() or "one feature"
                print(f"  [missed] {name} - Email Bot will work without it, except for: {feature}")
                failed_optional = True
            else:
                print(f"  [FAILED] {name} - Email Bot can't start without this.")
                failed_required = True

    if failed_required:
        print()
        print("Setup couldn't finish. Check that this computer is connected to the internet,")
        print("then start Email Bot again to retry.")
        print(f"(Details for whoever helps you: {LOG})")
        return 1
    if failed_optional:
        print()
        print("Some extras couldn't be installed. Email Bot will try again next time it starts.")
        print()
        return 2
    STAMP.write_bytes(wanted)
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
