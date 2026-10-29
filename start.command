#!/bin/bash
# Double-click this file (Mac) to start Email Bot.
cd "$(dirname "$0")" || exit 1
if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 is needed. Download it from https://www.python.org/downloads/ and try again."
  read -r -p "Press Enter to close."
  exit 1
fi
if [ ! -x .venv/bin/python ] || [ requirements.txt -nt .venv/.installed ]; then
  echo "First-time setup: installing what Email Bot needs (this takes a minute)..."
  python3 -m venv .venv && .venv/bin/pip install -q --upgrade pip && .venv/bin/pip install -q -r requirements.txt && touch .venv/.installed || {
    read -r -p "Setup failed. Press Enter to close."; exit 1; }
fi
.venv/bin/python app.py
