# Electricity Invoice Reader

A macOS-friendly desktop app for reviewing electricity invoices.

## Features

- Browse to a folder containing PDF invoices or an existing `.xlsx` summary.
- Extract service periods, meter numbers, and explicit `KWH Used` values.
- Use OCR for image-only PDFs when Tesseract and Pillow are installed.
- Show prorated calendar-month usage grouped by meter.
- Export the displayed summary to `electricity_invoice_summary.xlsx`.

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

On macOS, install the Tesseract executable for scanned-document OCR:

```bash
brew install tesseract
```

## Build a macOS app

```bash
source .venv/bin/activate
pip install pyinstaller
pyinstaller --windowed --name "Electricity Invoice Reader" app.py
```

The app will be in `dist/Electricity Invoice Reader.app`.
