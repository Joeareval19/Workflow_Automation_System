import os
import re
import csv
from pathlib import Path
from datetime import datetime
from decimal import Decimal

# Base path configuration
BASE_PATH = Path(r"C:\Users\User\Desktop\JEAV")
RS_VS_QB_PATH = BASE_PATH / "Weekly Payments Reconcile (daily)" / "RS vs QB"

# Regex pattern for week folders
WEEK_FOLDER_PATTERN = re.compile(r"Week_\d+_\(\d{2}\.\d{2}\.\d{2}\)_\(\d{2}\.\d{2}\.\d{2}\)")

def find_latest_week_folder(base_dir: Path):
    valid_folders = [f for f in base_dir.iterdir() if f.is_dir() and WEEK_FOLDER_PATTERN.match(f.name)]
    if not valid_folders:
        raise FileNotFoundError("No valid week folder found in the specified directory.")
    return sorted(valid_folders, key=lambda f: f.stat().st_ctime, reverse=True)[0]

def extract_dates_from_folder_name(folder_name: str):
    date_pattern = re.compile(r"\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)")
    match = date_pattern.search(folder_name)
    if not match:
        raise ValueError("Invalid folder name format, cannot extract dates.")
    return match.groups()

def ensure_directory_exists(path: Path):
    path.mkdir(parents=True, exist_ok=True)

def append_to_log(logPath: Path, message: str):
    with logPath.open("a", encoding="utf-8") as lf:
        lf.write(f"[{datetime.now()}] {message}\n")

def load_csv_as_list_of_dicts(csv_path: Path):
    rows = []
    with csv_path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows

def convert_to_decimal(value_str, invoice_id, field_name, logPath: Path):
    cleaned_value = re.sub(r"[$,\s]", "", value_str or "")
    if not cleaned_value:
        append_to_log(logPath, f"Empty {field_name} for Invoice ID: {invoice_id}, defaulting to 0")
        return Decimal("0")
    try:
        return Decimal(cleaned_value)
    except Exception as e:
        append_to_log(logPath, f"Error converting {field_name} to decimal for Invoice ID: {invoice_id}. Value: '{cleaned_value}', Error: {str(e)}")
        return None

def main():
    latest_week_folder = find_latest_week_folder(RS_VS_QB_PATH)
    start_date, end_date = extract_dates_from_folder_name(latest_week_folder.name)
    
    rawDataPath = latest_week_folder / f"1.Gathering/RAW_({start_date})_({end_date})csv.csv"
    combinedInvoicePath = BASE_PATH / "Weekly Payments Reconcile (daily)" / "Weekly Payables" / "Combined_Invoice_Report.csv"
    outputPath = latest_week_folder / "2.Import TransactionPro"
    ensure_directory_exists(outputPath)
    
    ilsReportPath = outputPath / f"ILS_cc_fee_({start_date})_({end_date}).csv"
    shipReportPath = outputPath / f"SHIP_cc_fee_({start_date})_({end_date}).csv"
    logPath = outputPath / "ProcessingLog.txt"
    
    # Clear log file
    with logPath.open("w", encoding="utf-8"):
        pass
    append_to_log(logPath, f"Starting processing for folder: {latest_week_folder.name}")
    
    rawData = load_csv_as_list_of_dicts(rawDataPath)
    combinedInvoice = load_csv_as_list_of_dicts(combinedInvoicePath)
    append_to_log(logPath, f"Loaded rawData: {len(rawData)} rows, combinedInvoice: {len(combinedInvoice)} rows")
    
    ilsPayments = []
    shipPayments = []
    
    for row in rawData:
        customer_id = row.get("Customer Id", "").strip()
        invoice_id = row.get("Invoice Id(s)", "").strip()
        invoice_id_trimmed = invoice_id[-8:] if len(invoice_id) > 8 else invoice_id  # Flexible trimming
        customer_name = row.get("Customer", "").strip()
        payment_date = row.get("Payment Date", "").strip()
        amount = convert_to_decimal(row.get("Amount", ""), invoice_id, "Amount", logPath)
        invoice_total = convert_to_decimal(row.get("Invoice Total", ""), invoice_id, "Invoice Total", logPath)
        
        append_to_log(logPath, f"Processing Invoice ID: {invoice_id}, Amount: {amount}, Invoice Total: {invoice_total}")
        
        if not amount or not invoice_total:
            append_to_log(logPath, f"Skipping Invoice ID: {invoice_id} due to invalid Amount or Invoice Total")
            continue
        
        # Relaxed comparison with tolerance for rounding
        expected_amount = (invoice_total * Decimal("1.03")).quantize(Decimal("0.00"))
        if abs(amount - expected_amount) > Decimal("0.01"):  # Tolerance of 1 cent
            append_to_log(logPath, f"Skipping Invoice ID: {invoice_id}, Amount {amount} != {expected_amount} (1.03x Invoice Total)")
            continue
        
        matchingRows = [r for r in combinedInvoice if r.get("Invoice Number", "") == invoice_id or r.get("Invoice Number", "") == invoice_id_trimmed]
        if not matchingRows:
            append_to_log(logPath, f"No match in Combined report for Invoice ID: {invoice_id} or {invoice_id_trimmed}")
            continue
        
        for combinedRow in matchingRows:
            colB = combinedRow.get("Column B", "")
            isILS = bool(re.search(r"DHL|ILS", colB, re.IGNORECASE))
            isSHIP = bool(re.search(r"FEDEX|SHIP", colB, re.IGNORECASE))
            
            if not (isILS or isSHIP):
                append_to_log(logPath, f"Skipping Invoice ID: {invoice_id} - Unknown service type in Column B: {colB}")
                continue
            
            percentage = convert_to_decimal(combinedRow.get("Percentage of Total (%)", ""), invoice_id, "Percentage", logPath)
            if not percentage:
                append_to_log(logPath, f"Skipping Invoice ID: {invoice_id} - Invalid Percentage")
                continue
            
            adjustedAmount = ((amount - invoice_total) * percentage).quantize(Decimal("0.00"))
            
            payment_record = {
                "CUST NO": customer_id,
                "INV NO": "CC"+invoice_id_trimmed,
                "CUSTOMER": customer_name,
                "MEMO INV ITEM": "CC-Fee paid by customer",
                "DATE": payment_date,
                "TERMS": "NET 15",
                "ACCOUNT": "MERCHANT FEE",
                "AMOUNT": str(adjustedAmount),
                "REP": ""
            }
            
            if isILS:
                ilsPayments.append(payment_record)
                append_to_log(logPath, f"Added to ILS: Invoice ID {invoice_id}, Adjusted Amount {adjustedAmount}")
            else:
                shipPayments.append(payment_record)
                append_to_log(logPath, f"Added to SHIP: Invoice ID {invoice_id}, Adjusted Amount {adjustedAmount}")
    
    for path, data in [(ilsReportPath, ilsPayments), (shipReportPath, shipPayments)]:
        with path.open("w", encoding="utf-8", newline="") as f:
            if data:
                writer = csv.DictWriter(f, fieldnames=data[0].keys())
                writer.writeheader()
                writer.writerows(data)
    
    append_to_log(logPath, f"Processing complete. ILS Payments: {len(ilsPayments)}, SHIP Payments: {len(shipPayments)}")
    print("Processing complete. Files created:")
    print(ilsReportPath)
    print(shipReportPath)
    print(f"Check log for details: {logPath}")

if __name__ == "__main__":
    main()
