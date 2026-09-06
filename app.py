from __future__ import annotations

import re
import threading
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import END, BOTH, LEFT, RIGHT, TOP, X, Y, Button, Entry, Frame, Label, StringVar, Tk, filedialog, messagebox
from tkinter import ttk

import fitz
from openpyxl import Workbook, load_workbook


DATE_PATTERN = r"([A-Z][a-z]{2,8}\s+\d{1,2},\s+\d{4})"
SERVICE_PERIOD_PATTERN = re.compile(
    rf"Service Period\s+{DATE_PATTERN}\s*-\s*{DATE_PATTERN}", re.IGNORECASE
)
METER_PATTERN = re.compile(r"Meter Number\s+([A-Za-z0-9 -]+?)(?=\s+(?:Meter Reading|KWH Used|Days Billed|$))", re.IGNORECASE)
USAGE_PATTERN = re.compile(r"KWH Used\s+([\d,]+(?:\.\d+)?)", re.IGNORECASE)


@dataclass
class Invoice:
    filename: str
    start: date
    end: date
    meter: str
    usage: float
    source: str
    confidence: str = "High"
    ambiguity: str = ""


@dataclass
class MonthlyUsage:
    meter: str
    month: date
    usage: float
    billed_days: int
    invoice_total: float
    invoice_file: str


def parse_date(value: str) -> date:
    return datetime.strptime(value.strip(), "%b %d, %Y").date()


def extract_invoice(path: Path) -> Invoice:
    document = fitz.open(path)
    text = "\n".join(page.get_text() for page in document)
    if not text.strip():
        try:
            import pytesseract
            from PIL import Image

            rendered = []
            for page in document:
                pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
                rendered.append(pytesseract.image_to_string(image))
            text = "\n".join(rendered)
        except ImportError as error:
            raise ValueError("This invoice is scanned. Install OCR support with: pip install pytesseract pillow") from error

    period = SERVICE_PERIOD_PATTERN.search(text)
    meter = METER_PATTERN.search(text)
    usage_matches = USAGE_PATTERN.findall(text)
    if not period or not meter or not usage_matches:
        raise ValueError("Could not identify service period, meter number, or KWH Used")

    usage_values = [float(value.replace(",", "")) for value in usage_matches]
    usage = usage_values[0]
    ambiguity = ""
    if len(set(usage_values)) > 1:
        usage = usage_values[-1]
        ambiguity = "Multiple KWH Used values found; latest billing-table value selected."

    return Invoice(
        filename=path.name,
        start=parse_date(period.group(1)),
        end=parse_date(period.group(2)),
        meter=meter.group(1).strip(),
        usage=usage,
        source="Page 2, Current Billing Information",
        ambiguity=ambiguity,
    )


def prorate(invoices: list[Invoice]) -> list[MonthlyUsage]:
    rows = []
    for invoice in invoices:
        total_days = (invoice.end - invoice.start).days + 1
        current = invoice.start.replace(day=1)
        while current <= invoice.end.replace(day=1):
            next_month = date(current.year + (current.month == 12), 1 if current.month == 12 else current.month + 1, 1)
            month_end = next_month - timedelta(days=1)
            overlap_start = max(invoice.start, current)
            overlap_end = min(invoice.end, month_end)
            billed_days = (overlap_end - overlap_start).days + 1
            if billed_days > 0:
                rows.append(MonthlyUsage(
                    meter=invoice.meter,
                    month=current,
                    usage=invoice.usage * billed_days / total_days,
                    billed_days=billed_days,
                    invoice_total=invoice.usage,
                    invoice_file=invoice.filename,
                ))
            current = next_month
    return sorted(rows, key=lambda row: (row.meter, row.month, row.invoice_file))


def load_workbook_rows(path: Path) -> list[MonthlyUsage]:
    workbook = load_workbook(path, data_only=True)
    sheet = workbook["By meter"] if "By meter" in workbook.sheetnames else workbook.active
    headers = [cell.value for cell in sheet[1]]
    month_index = headers.index("Billing month") if "Billing month" in headers else None
    usage_index = headers.index("Prorated usage (kWh)") if "Prorated usage (kWh)" in headers else None
    meter_index = headers.index("Meter number") if "Meter number" in headers else None
    if None in (month_index, usage_index, meter_index):
        raise ValueError("The workbook must contain Meter number, Billing month, and Prorated usage (kWh) columns")

    rows = []
    for values in sheet.iter_rows(min_row=2, values_only=True):
        if not values[meter_index] or values[usage_index] is None:
            continue
        month_value = values[month_index]
        if isinstance(month_value, datetime):
            month = month_value.date().replace(day=1)
        else:
            month = datetime.strptime(str(month_value), "%B %Y").date().replace(day=1)
        rows.append(MonthlyUsage(str(values[meter_index]), month, float(values[usage_index]), 0, 0, "Existing workbook"))
    return sorted(rows, key=lambda row: (row.meter, row.month))


def export_workbook(folder: Path, invoices: list[Invoice], monthly_rows: list[MonthlyUsage]) -> Path:
    output = folder / "electricity_invoice_summary.xlsx"
    workbook = Workbook()
    readme = workbook.active
    readme.title = "Read me"
    readme["A1"] = "Electricity invoice summary"
    readme["A3"] = "Monthly usage is prorated by inclusive billed days across each invoice service period."

    source = workbook.create_sheet("Electricity invoices")
    source.append(["Invoice file", "Billing period start", "Billing period end", "Meter number", "Total usage (kWh)", "Usage source/page or section", "Confidence", "Ambiguities"])
    for invoice in invoices:
        source.append([invoice.filename, invoice.start, invoice.end, invoice.meter, invoice.usage, invoice.source, invoice.confidence, invoice.ambiguity])

    monthly = workbook.create_sheet("By meter")
    monthly.append(["Meter number", "Billing month", "Billed days allocated", "Prorated usage (kWh)", "Invoice total usage (kWh)", "Invoice file"])
    for row in monthly_rows:
        monthly.append([row.meter, row.month.strftime("%B %Y"), row.billed_days, row.usage, row.invoice_total, row.invoice_file])

    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = cell.font.copy(bold=True)
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = min(max(max(len(str(cell.value or "")) for cell in column) + 2, 14), 42)
    workbook.save(output)
    return output


class InvoiceApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("Electricity Invoice Reader")
        self.root.geometry("1180x680")
        self.root.minsize(900, 520)
        self.folder = Path()
        self.invoices: list[Invoice] = []
        self.monthly_rows: list[MonthlyUsage] = []
        self.status = StringVar(value="Choose a folder containing PDF invoices or an existing Excel summary.")
        self.path_text = StringVar()
        self.build_ui()

    def build_ui(self):
        top = Frame(self.root, padx=22, pady=18)
        top.pack(fill=X)
        Label(top, text="Electricity Invoice Reader", font=("Helvetica", 22, "bold")).pack(anchor="w")
        Label(top, text="Browse a folder to extract invoices and review prorated monthly usage.", font=("Helvetica", 11)).pack(anchor="w", pady=(4, 14))

        controls = Frame(top)
        controls.pack(fill=X)
        Entry(controls, textvariable=self.path_text, state="readonly", width=80).pack(side=LEFT, fill=X, expand=True, padx=(0, 8))
        Button(controls, text="Browse folder", command=self.choose_folder, padx=14).pack(side=LEFT)
        Button(controls, text="Export Excel", command=self.export, padx=14).pack(side=LEFT, padx=(8, 0))

        summary = Frame(self.root, padx=22)
        summary.pack(fill=X)
        self.metric = Label(summary, text="No data loaded", font=("Helvetica", 11, "bold"))
        self.metric.pack(anchor="w", pady=(0, 10))

        table_frame = Frame(self.root, padx=22, pady=0)
        table_frame.pack(fill=BOTH, expand=True, pady=(0, 14))
        columns = ("meter", "month", "usage", "days", "invoice")
        self.table = ttk.Treeview(table_frame, columns=columns, show="headings")
        headings = {"meter": "Meter number", "month": "Billing month", "usage": "Prorated usage (kWh)", "days": "Billed days", "invoice": "Invoice file"}
        widths = {"meter": 150, "month": 150, "usage": 190, "days": 110, "invoice": 390}
        for column in columns:
            self.table.heading(column, text=headings[column])
            self.table.column(column, width=widths[column], anchor="w")
        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=scroll.set)
        self.table.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.pack(side=RIGHT, fill=Y)

        Label(self.root, textvariable=self.status, relief="sunken", anchor="w", padx=10).pack(side="bottom", fill=X)

    def choose_folder(self):
        selected = filedialog.askdirectory(title="Select invoice folder")
        if selected:
            self.folder = Path(selected)
            self.path_text.set(str(self.folder))
            self.load_folder()

    def load_folder(self):
        self.status.set("Reading invoices...")
        self.root.update_idletasks()
        threading.Thread(target=self._load_folder_worker, daemon=True).start()

    def _load_folder_worker(self):
        try:
            workbook_files = sorted(self.folder.glob("*.xlsx"))
            if workbook_files:
                rows = load_workbook_rows(workbook_files[0])
                invoices = []
                message = f"Loaded {workbook_files[0].name}."
            else:
                pdf_files = sorted(self.folder.glob("*.pdf"))
                if not pdf_files:
                    raise ValueError("No PDF invoices or .xlsx summary found in this folder.")
                invoices = [extract_invoice(path) for path in pdf_files]
                rows = prorate(invoices)
                message = f"Processed {len(invoices)} invoice(s)."
            self.root.after(0, lambda: self.show_rows(invoices, rows, message))
        except Exception as error:
            self.root.after(0, lambda: messagebox.showerror("Could not load folder", str(error)))
            self.root.after(0, lambda: self.status.set("Load failed. Check the folder and invoice format."))

    def show_rows(self, invoices, rows, message):
        self.invoices = invoices
        self.monthly_rows = rows
        for item in self.table.get_children():
            self.table.delete(item)
        for row in rows:
            self.table.insert("", END, values=(row.meter, row.month.strftime("%B %Y"), f"{row.usage:,.2f}", row.billed_days or "-", row.invoice_file))
        total = sum(row.usage for row in rows)
        meters = len({row.meter for row in rows})
        months = len({row.month for row in rows})
        self.metric.config(text=f"{meters} meter(s)  |  {months} billing month(s)  |  {total:,.2f} kWh total")
        self.status.set(message)

    def export(self):
        if not self.monthly_rows:
            messagebox.showinfo("Nothing to export", "Browse to an invoice folder first.")
            return
        output = export_workbook(self.folder, self.invoices, self.monthly_rows)
        self.status.set(f"Saved {output.name}")
        messagebox.showinfo("Export complete", f"Saved:\n{output}")


if __name__ == "__main__":
    app = Tk()
    InvoiceApp(app)
    app.mainloop()
