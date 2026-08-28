#!/bin/bash
# Dubbelklicka på den här filen för att starta Bad Debt-dashboarden.
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  echo "Skapar Python-miljö (sker bara första gången)..."
  python3 -m venv .venv
  source .venv/bin/activate
  pip install --quiet --upgrade pip
  pip install --quiet -r requirements.txt
else
  source .venv/bin/activate
fi

streamlit run app.py
