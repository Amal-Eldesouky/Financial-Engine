"""Build an end-to-end financial model from Amazon reports and manual inputs."""
from __future__ import annotations

import csv
import math
import re
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

BASE_DIR = Path(__file__).resolve().parent
SUMMARY_FILE = BASE_DIR / "amazon_summary.csv"
SUMMARY_PDF_FILE = BASE_DIR / "amazon_summary.csv.pdf"
TRANSACTION_FILES = (BASE_DIR / "amazon_transactions.csv", BASE_DIR / "amazon_transactions.csv.csv")
OUTPUT_FILE = BASE_DIR / "Amazon_Pure_Financial_Model.xlsx"
MONEY_FORMAT = '#,##0.00;[Red]-#,##0.00'


def ask_number(label: str) -> float:
    while True:
        raw = input(f"{label}: ").strip()
        try:
            value = float(raw.replace(",", "")) if raw else 0.0
            if value < 0:
                raise ValueError
            return value
        except ValueError:
            print("Please enter a non-negative number, or press Enter for 0.00.")


def normalize(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def find_header_row(path: Path, required: Iterable[str]) -> int:
    required_normalized = {normalize(value) for value in required}
    minimum_matches = min(2, len(required_normalized))
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        for row_number, row in enumerate(csv.reader(handle)):
            row_normalized = {normalize(item) for item in row}
            if len(required_normalized.intersection(row_normalized)) >= minimum_matches:
                return row_number
    return 0


def read_report(path: Path, required_headers: Iterable[str]) -> pd.DataFrame:
    frame = pd.read_csv(path, skiprows=find_header_row(path, required_headers), dtype=str, on_bad_lines="skip")
    frame.columns = [str(column).strip() for column in frame.columns]
    frame = frame.loc[:, ~frame.columns.str.match(r"^Unnamed")]
    return frame.replace(r"^\s*$", 0, regex=True).fillna(0)


def find_column(frame: pd.DataFrame, aliases: Iterable[str]) -> str | None:
    columns = {normalize(column): column for column in frame.columns}
    for alias in aliases:
        alias_normalized = normalize(alias)
        if alias_normalized in columns:
            return columns[alias_normalized]
    return next(
        (column for normalized, column in columns.items() if any(normalize(alias) in normalized for alias in aliases)),
        None,
    )


def numeric_series(frame: pd.DataFrame, column: str) -> pd.Series:
    cleaned = frame[column].astype(str).str.replace(r"[$,]", "", regex=True)
    cleaned = cleaned.str.replace("(", "-", regex=False).str.replace(")", "", regex=False)
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def numeric_value(value: object, label: str) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError) as error:
        raise ValueError(f"Amazon report value for {label!r} is not numeric: {value!r}") from error
    if not pd.api.types.is_number(number) or not math.isfinite(number):
        raise ValueError(f"Amazon report value for {label!r} is not numeric: {value!r}")
    return number


def read_summary(path: Path) -> dict[str, float | str]:
    frame = read_report(path, ("gross sales", "sales", "refunds"))

    def total(aliases: Iterable[str]) -> float:
        column = find_column(frame, aliases)
        return float(numeric_series(frame, column).sum()) if column else 0.0

    period_column = find_column(frame, ("Month", "Date", "Period"))
    values = {
        "Gross Sales": total(("Gross Sales", "Sales", "Product Sales")),
        "Refunds": -abs(total(("Refunds", "Refund", "Refunded Amount", "Promotions", "Promotional Rebates"))),
        "Amazon Selling Fees": abs(total(("Amazon Selling Fees", "Selling Fees", "Amazon Fees", "Referral Fees"))),
        "Period": str(frame[period_column].iloc[0]) if period_column and not frame.empty else "Amazon report period",
    }
    return {label: numeric_value(value, label) if label != "Period" else value for label, value in values.items()}


def read_summary_pdf(path: Path) -> dict[str, float | str]:
    if PdfReader is None:
        raise ImportError("Reading amazon_summary.csv.pdf requires pypdf. Install it with: uv add pypdf")
    text = "\n".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)

    def labelled_total(label: str) -> float:
        match = re.search(rf"{re.escape(label)}\s+(-?[\d,]+\.\d{{2}})", text, re.IGNORECASE)
        return float(match.group(1).replace(",", "")) if match else 0.0

    return {
        "Gross Sales": labelled_total("Product sales (non-FBA)") + labelled_total("FBA product sales"),
        "Refunds": -abs(labelled_total("Product sale refunds (non-FBA)" ) + labelled_total("FBA product sale refunds")),
        "Amazon Selling Fees": abs(labelled_total("FBA selling fees") + labelled_total("Seller fulfilled selling fees")),
        "Period": "Amazon Summary Report",
    }


def read_transactions(path: Path) -> dict[str, float]:
    frame = read_report(path, ("date/time", "total", "type"))
    total_column = find_column(frame, ("total",))
    type_column = find_column(frame, ("type",))
    description_column = find_column(frame, ("description",))
    if not total_column or not type_column:
        raise ValueError("Amazon transaction report must contain Type and Total columns.")
    amounts = numeric_series(frame, total_column)
    product_sales_column = find_column(frame, ("product sales", "sales", "gross"))
    promotions_column = find_column(frame, ("promotions", "promotional rebates"))
    selling_fees_column = find_column(frame, ("selling fees", "amazon selling fees", "referral fees"))
    product_sales = numeric_series(frame, product_sales_column) if product_sales_column else pd.Series(0.0, index=frame.index)
    gross_sales = product_sales[product_sales > 0].sum()
    refunds = product_sales[product_sales < 0].sum()
    if promotions_column:
        refunds += numeric_series(frame, promotions_column).sum()
    selling_fees = numeric_series(frame, selling_fees_column).abs().sum() if selling_fees_column else 0.0
    text = frame[type_column].astype(str).str.lower()
    if description_column:
        text = text + " " + frame[description_column].astype(str).str.lower()
    values = {
        "Gross Sales": float(gross_sales),
        "Refunds": float(refunds),
        "Amazon Selling Fees": float(selling_fees),
        "PPC Spend": float(amounts[text.str.contains(r"advertis|sponsored product|sponsored brands|ppc", regex=True, na=False)].abs().sum()),
        "Storage Fees": float(amounts[text.str.contains("storage", regex=False, na=False)].abs().sum()),
        "Amazon Settlements / Payouts": float(amounts[text.str.contains(r"transfer|payout|settlement|disbursement", regex=True, na=False)].sum()),
    }
    return {label: numeric_value(value, label) for label, value in values.items()}


def set_value(sheet, cell: str, value: float, number_format: str = MONEY_FORMAT) -> None:
    sheet[cell] = numeric_value(value, f"{sheet.title}!{cell}")
    sheet[cell].number_format = number_format


def style_sheet(sheet, title: str, subtitle: str) -> None:
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "B5"
    sheet["A1"] = title
    sheet["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    sheet["A1"].fill = PatternFill("solid", fgColor="17324D")
    sheet.merge_cells("A1:B1")
    sheet["A2"] = subtitle
    sheet["A2"].font = Font(italic=True, color="5B6770")
    sheet.merge_cells("A2:B2")
    sheet.column_dimensions["A"].width = 38
    sheet.column_dimensions["B"].width = 22
    thin = Side(style="thin", color="D9E1E8")
    for row in sheet.iter_rows():
        for cell in row:
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="center")


def build_workbook(amazon: dict[str, float | str], manual: dict[str, float], output_path: Path) -> None:
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "INPUTS_DATA"
    income = workbook.create_sheet("INCOME_STATEMENT")
    balance = workbook.create_sheet("BALANCE_SHEET")
    cash_flow = workbook.create_sheet("CASH_FLOW_STATEMENT")
    dashboard = workbook.create_sheet("KPI_DASHBOARD")
    period = str(amazon["Period"])
    rows: dict[str, int] = {}

    style_sheet(inputs, "Automated Financial Engine | Inputs & Amazon Data", period)
    inputs["A4"], inputs["B4"] = "Metric", "Value"
    imported = {
        **amazon,
        "COGS": manual["COGS per Unit"] * manual["Units Sold Qty"],
        "COGS per Unit": manual["COGS per Unit"],
        "Units Sold Qty": manual["Units Sold Qty"],
    }
    imported_labels = ["Gross Sales", "Refunds", "Amazon Selling Fees", "PPC Spend", "Storage Fees", "Amazon Settlements / Payouts", "COGS", "COGS per Unit", "Units Sold Qty"]
    for label in imported_labels:
        imported[label] = numeric_value(imported.get(label, 0.0), label)
    row = 5
    for label in imported_labels:
        rows[label] = row
        inputs.cell(row, 1, label)
        inputs.cell(row, 2, imported.get(label, manual.get(label, 0.0)))
        row += 1
    inputs.cell(row, 1, "Manual non-Amazon inputs").font = Font(bold=True, color="17324D")
    row += 1
    for label in ("Salaries", "External Supplier Cash Outflows", "Owner Drawings", "Asset Purchases", "Opening Bank Balance"):
        rows[label] = row
        inputs.cell(row, 1, label)
        inputs.cell(row, 2, numeric_value(manual[label], label))
        inputs.cell(row, 2).fill = PatternFill("solid", fgColor="FFF2CC")
        row += 1
    for cell in inputs[4]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2C6E8F")
    for row_number in range(5, row):
        inputs.cell(row_number, 2).number_format = '#,##0' if row_number == rows["Units Sold Qty"] else MONEY_FORMAT

    for sheet, title in ((income, "Accrual Income Statement"), (balance, "Dynamic Balance Sheet"), (cash_flow, "Cash Flow Statement"), (dashboard, "KPI Dashboard")):
        style_sheet(sheet, title, period)
        sheet["A4"], sheet["B4"] = "Line Item", "Value"
        for cell in sheet[4]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="2C6E8F")

    gross_sales = imported["Gross Sales"]
    refunds = imported["Refunds"]
    net_revenue = gross_sales + refunds
    cogs = imported["COGS"]
    gross_profit = net_revenue - cogs
    selling_fees = imported["Amazon Selling Fees"]
    ppc = imported["PPC Spend"]
    storage = imported["Storage Fees"]
    salaries = numeric_value(manual["Salaries"], "Salaries")
    supplier_outflows = numeric_value(manual["External Supplier Cash Outflows"], "External Supplier Cash Outflows")
    asset_purchases = numeric_value(manual["Asset Purchases"], "Asset Purchases")
    owner_drawings = numeric_value(manual["Owner Drawings"], "Owner Drawings")
    opening_balance = numeric_value(manual["Opening Bank Balance"], "Opening Bank Balance")
    operating_cash_flow = imported["Amazon Settlements / Payouts"] - supplier_outflows - salaries
    ending_cash = opening_balance + operating_cash_flow - asset_purchases - owner_drawings
    inventory = max(supplier_outflows - cogs, 0)
    total_assets = ending_cash + inventory

    income_rows = [("Gross Sales", gross_sales), ("Refunds", refunds), ("Net Revenue", net_revenue), ("COGS", -cogs), ("Gross Profit", gross_profit), ("Amazon Selling Fees", -selling_fees), ("PPC Spend", -ppc), ("Storage Fees", -storage), ("Salaries / OPEX", -salaries), ("Net Income / (Loss)", gross_profit - selling_fees - ppc - storage - salaries)]
    for row_number, (label, value) in enumerate(income_rows, 5):
        income.cell(row_number, 1, label)
        set_value(income, f"B{row_number}", value)
    for row_number in (7, 9, 14):
        income[f"A{row_number}"].font = income[f"B{row_number}"].font = Font(bold=True)

    cash_rows = [("Amazon Settlements / Payouts", imported["Amazon Settlements / Payouts"]), ("External Supplier Payments", -supplier_outflows), ("Salaries Paid", -salaries), ("Operating Cash Flow", operating_cash_flow), ("Asset Purchases", -asset_purchases), ("Investing Cash Flow", -asset_purchases), ("Owner Drawings", -owner_drawings), ("Financing Cash Flow", -owner_drawings), ("Opening Bank Balance", opening_balance), ("Ending Cash", ending_cash)]
    for row_number, (label, value) in enumerate(cash_rows, 5):
        cash_flow.cell(row_number, 1, label)
        set_value(cash_flow, f"B{row_number}", value)
    for row_number in (8, 10, 12, 14):
        cash_flow[f"A{row_number}"].font = cash_flow[f"B{row_number}"].font = Font(bold=True)

    balance_rows = [("Cash Asset", ending_cash), ("Inventory Asset", supplier_outflows - cogs), ("Total Assets", total_assets), ("Accounts Payable", 0), ("Owner's Equity", total_assets), ("Total Liabilities + Equity", total_assets), ("Balance Check", 0)]
    for row_number, (label, value) in enumerate(balance_rows, 5):
        balance.cell(row_number, 1, label)
        set_value(balance, f"B{row_number}", value)
    for row_number in (7, 10, 11):
        balance[f"A{row_number}"].font = balance[f"B{row_number}"].font = Font(bold=True)

    net_income = gross_profit - selling_fees - ppc - storage - salaries
    kpis = [("ARR", net_revenue * 12, MONEY_FORMAT), ("Marketing % of Net Revenue", ppc / net_revenue if net_revenue else 0, '0.00%'), ("Blended CPA / CAC", ppc / imported["Units Sold Qty"] if imported["Units Sold Qty"] else 0, MONEY_FORMAT), ("Target Achievement %", net_revenue / gross_sales if gross_sales else 0, '0.00%'), ("Net Profit Margin", net_income / net_revenue if net_revenue else 0, '0.00%'), ("Balance Check", 0, MONEY_FORMAT)]
    for row_number, (label, value, number_format) in enumerate(kpis, 5):
        dashboard.cell(row_number, 1, label)
        set_value(dashboard, f"B{row_number}", value, number_format)
    chart = BarChart()
    chart.title = "Income Statement Overview"
    chart.add_data(Reference(income, min_col=2, min_row=7, max_row=14))
    chart.set_categories(Reference(income, min_col=1, min_row=7, max_row=14))
    dashboard.add_chart(chart, "D4")
    for sheet in workbook.worksheets:
        sheet.auto_filter.ref = f"A4:B{sheet.max_row}"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    workbook.save(output_path)


def main() -> None:
    summary_file = SUMMARY_FILE if SUMMARY_FILE.exists() else SUMMARY_PDF_FILE if SUMMARY_PDF_FILE.exists() else None
    if summary_file is None:
        raise FileNotFoundError("No Amazon summary report found. Add amazon_summary.csv or amazon_summary.csv.pdf beside app.py.")
    transaction_file = next((path for path in TRANSACTION_FILES if path.exists()), None)
    if transaction_file is None:
        raise FileNotFoundError("amazon_transactions.csv was not found beside app.py.")
    print("Enter manual non-Amazon inputs (press Enter to use 0.00).")
    manual = {"COGS per Unit": ask_number("COGS per unit"), "Units Sold Qty": ask_number("Units Sold Qty"), "Salaries": ask_number("Salaries"), "External Supplier Cash Outflows": ask_number("External Supplier Cash Outflows"), "Owner Drawings": ask_number("Owner Drawings"), "Asset Purchases": ask_number("Asset Purchases"), "Opening Bank Balance": ask_number("Opening Bank Balance")}
    summary = read_summary(summary_file) if summary_file.suffix.lower() == ".csv" else read_summary_pdf(summary_file)
    transactions = read_transactions(transaction_file)
    for label in ("Gross Sales", "Refunds", "Amazon Selling Fees"):
        if not summary.get(label):
            summary[label] = transactions[label]
    build_workbook({**summary, **transactions}, manual, OUTPUT_FILE)
    print(f"Created {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
