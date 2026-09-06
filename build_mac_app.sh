#!/bin/zsh
set -e
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
pyinstaller --clean --noconfirm --windowed --name "Electricity Invoice Reader" app.py
printf '\nBuilt: dist/Electricity Invoice Reader.app\n'
