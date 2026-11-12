#!/bin/sh
cd "$(dirname "$0")"
[ -f .env ] || cp .env.example .env
if [ ! -d .venv ]; then
  python3 -m venv .venv
  .venv/bin/pip install -r requirements.txt
fi
exec .venv/bin/python app.py
