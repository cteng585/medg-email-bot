#!/bin/bash
# Double-click this file to start Email Bot.
#
# The first time, it sets up everything Email Bot needs inside a ".venv" folder
# next to this file. Nothing is installed anywhere else on the Mac, and no
# administrator password is needed. Later starts skip straight to the app.

cd "$(dirname "$0")" || exit 1
VENV_PY=".venv/bin/python"
STAMP=".venv/.installed"   # copy of requirements.txt from the last complete setup
LOG=".venv/setup.log"

close_window() {
  echo
  read -r -p "Press Enter to close this window."
  exit "${1:-1}"
}

# Finds a Python 3.9+ that can create a virtual environment.
find_python() {
  local c p
  for c in python3.13 python3.12 python3.11 python3.10 python3.9 \
           /Library/Frameworks/Python.framework/Versions/Current/bin/python3 \
           /opt/homebrew/bin/python3 /usr/local/bin/python3 python3; do
    p=$(command -v "$c" 2>/dev/null) || continue
    # Until Apple's Command Line Tools are installed, /usr/bin/python3 is only a
    # placeholder that shows an install pop-up, so don't run it.
    if [ "$p" = /usr/bin/python3 ] && ! xcode-select -p >/dev/null 2>&1; then
      continue
    fi
    if "$p" -c 'import sys, venv, ensurepip; sys.exit(sys.version_info < (3, 9))' >/dev/null 2>&1; then
      echo "$p"
      return 0
    fi
  done
  return 1
}

# A setup left behind by a removed or upgraded Python can't run; start it over.
if [ -e "$VENV_PY" ] && ! "$VENV_PY" -c 'import sys' >/dev/null 2>&1; then
  echo "The previous setup no longer works (Python may have been updated). Setting up again..."
  rm -rf .venv
fi

if [ ! -x "$VENV_PY" ]; then
  echo "First-time setup: getting Email Bot ready. This takes a minute or two."
  echo
  if ! PY=$(find_python); then
    echo "Email Bot needs Python (version 3.9 or newer), and this Mac doesn't have it yet."
    echo
    echo "Either:"
    echo "  - Install Python from https://www.python.org/downloads/macos/ , or"
    echo "  - Install Apple's Command Line Tools: run   xcode-select --install   in Terminal."
    echo
    echo "Both may ask for an administrator password. If you don't have one, ask"
    echo "whoever manages this Mac. Then double-click start.command again."
    close_window
  fi
  echo "Using Python at $PY"
  if ! "$PY" -m venv .venv >/dev/null 2>&1; then
    rm -rf .venv
    echo "Couldn't set up Email Bot's folder. Make sure the Email Bot folder isn't"
    echo "read-only (for example, still inside a zip file or a shared drive), then try again."
    close_window
  fi
fi

if ! cmp -s requirements.txt "$STAMP"; then
  echo "Installing the add-ons Email Bot uses (needs an internet connection)..."
  failed_required=0
  failed_optional=0
  while IFS= read -r line || [ -n "$line" ]; do
    spec=$(echo "${line%%#*}" | xargs)
    [ -z "$spec" ] && continue
    note=""
    [[ "$line" == *"#"* ]] && note=$(echo "${line#*#}" | xargs)
    if "$VENV_PY" -m pip install --disable-pip-version-check --no-input --retries 2 --timeout 20 \
         "$spec" </dev/null >>"$LOG" 2>&1; then
      echo "  [ok]     $spec"
    elif [[ "$note" == optional* ]]; then
      echo "  [missed] $spec - Email Bot will work without it, except for: ${note#optional: }"
      failed_optional=1
    else
      echo "  [FAILED] $spec - Email Bot can't start without this."
      failed_required=1
    fi
  done < requirements.txt

  if [ "$failed_required" = 1 ]; then
    echo
    echo "Setup couldn't finish. Check that this Mac is connected to the internet,"
    echo "then double-click start.command again to retry."
    echo "(Details for whoever helps you: $(pwd)/$LOG)"
    close_window
  fi
  if [ "$failed_optional" = 1 ]; then
    echo
    echo "Some extras couldn't be installed. Email Bot will try again next time it starts."
  else
    cp requirements.txt "$STAMP"
  fi
  echo
fi

"$VENV_PY" app.py
close_window $?
