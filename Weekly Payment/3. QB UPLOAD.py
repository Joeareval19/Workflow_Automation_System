import os
import re
import csv
from pathlib import Path
from datetime import datetime
from decimal import Decimal

#
# ========== #Dynamic Path Configuration #=========
#
BASE_PATH = Path(r"C:\Users\User\Desktop\JEAV")
RS_VS_QB_PATH = BASE_PATH / "Weekly Payments Reconcile (daily)" / "RS vs QB"

# This pattern matches folder names like: Week_1_(01.08.23)_(01.14.23)
WEEK_FOLDER_PATTERN = re.compile(r"Week_\d+_\(\d{2}\.\d{2}\.\d{2}\)_\(\d{2}\.\d{2}\.\d{2}\)")

#
# ========== #Helper/Utility Functions #=========
#

def find_latest_week_folder(base_dir: Path):
    valid_folders = []
    for item in base_dir.iterdir():
        if item.is_dir() and WEEK_FOLDER_PATTERN.match(item.name):
            valid_folders.append(item)
    if not valid_folders:
        raise FileNotFoundError("No valid week folder found in the specified directory.")
    valid_folders.sort(key=lambda f: f.stat().st_ctime, reverse=True)
    return valid_folders[0]

def extract_dates_from_folder_name(folder_name: str):
    date_pattern = re.compile(r"\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)")
    match_obj = date_pattern.search(folder_name)
    if not match_obj:
        raise ValueError("Invalid folder name format, cannot extract dates.")
    return match_obj.groups()

def ensure_directory_exists(path: Path):
    path.mkdir(parents=True, exist_ok=True)

def clear_content_of_file(logPath: Path):
    with logPath.open("w", encoding="utf-8"):
        pass

def append_to_log(logPath: Path, message: str):
    with logPath.open("a", encoding="utf-8") as lf:
        lf.write(f"[{datetime.now()}] {message}\n")

def load_csv_as_list_of_dicts(csvPath: Path):
    rows = []
    with csvPath.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows

def get_customer_name(customer_list, customer_id, logPath: Path):
    for c in customer_list:
        if c.get("Customer Id") == customer_id:
            return c.get("Customer", "")
    append_to_log(logPath, f"Customer ID not found: {customer_id}")
    return "N/A"

def generate_ref_no(payment_method, invoice_id, logPath: Path):
    if not invoice_id or len(invoice_id) < 4:
        append_to_log(logPath, f"Warning: Invalid invoice ID length for {invoice_id}")
        return f"INVALID-{invoice_id}"

    if re.match(r"^\d+$", payment_method or ""):
        return "5-" + invoice_id[4:]
    elif re.search(r"VISA|MasterCard", payment_method or "", re.IGNORECASE):
        return "4-" + invoice_id[4:]
    elif re.search(r"AMEX", payment_method or "", re.IGNORECASE):
        return "3-" + invoice_id[4:]
    elif re.search(r"ACH", payment_method or "", re.IGNORECASE):
        return "6-" + invoice_id[4:]
    else:
        return "UNKNOWN-" + invoice_id[4:]

def convert_to_decimal(value_str, invoice_id, field_name, logPath: Path):
    cleaned_value = re.sub(r"[\$,\s]", "", value_str or "")
    if not cleaned_value:
        cleaned_value = "0"
    try:
        dec_val = Decimal(cleaned_value)
        if field_name == "percentage" and dec_val > 1:
            dec_val = dec_val / Decimal("100")
        return dec_val
    except Exception:
        append_to_log(logPath, f"Error converting {field_name} to decimal for Invoice ID: {invoice_id}. Value: '{cleaned_value}'")
        return None

#
# Main script logic
#
def main():
    try:
        latest_week_folder = find_latest_week_folder(RS_VS_QB_PATH)
    except FileNotFoundError as e:
        print(str(e))
        return

    folder_name = latest_week_folder.name

    try:
        start_date, end_date = extract_dates_from_folder_name(folder_name)
    except ValueError as e:
        print(str(e))
        return

    rawDataPath = latest_week_folder / f"1.Gathering/RAW_({start_date})_({end_date})csv.csv"
    customerListPath = BASE_PATH / "Customer List.csv"
    combinedInvoicePath = BASE_PATH / "Weekly Payments Reconcile (daily)" / "Weekly Payables" / "Combined_Invoice_Report.csv"
    importPath = latest_week_folder / "2.Import TransactionPro"
    ensure_directory_exists(importPath)

    ilsReportPath = importPath / f"ILS_Payment_Report_({start_date})_({end_date}).csv"
    shipReportPath = importPath / f"SHIP_Payment_Report_({start_date})_({end_date}).csv"
    logPath = importPath / "ProcessingLog.txt"

    REPORT_OUTPUT_PATH = RS_VS_QB_PATH / "1.Report"
    ensure_directory_exists(REPORT_OUTPUT_PATH)

    clear_content_of_file(logPath)
    append_to_log(logPath, f"Processing Log: {datetime.now()}")
    append_to_log(logPath, f"Processing folder: {folder_name}")

    try:
        rawData = load_csv_as_list_of_dicts(rawDataPath)
        customerList = load_csv_as_list_of_dicts(customerListPath)
        combinedInvoice = load_csv_as_list_of_dicts(combinedInvoicePath)
        append_to_log(logPath, "Successfully loaded all CSV files")
    except Exception as e:
        append_to_log(logPath, f"Error loading CSV files: {str(e)}")
        print("Failed to load required CSV files. See log for details.")
        return

    ilsPayments = []
    shipPayments = []
    rawFileCount = len(rawData)

    missingCustomerNames = []
    prepaidInvoices = []
    unmatchedInvoices = []
    totalProcessed = 0
    totalErrors = 0

    for row in rawData:
        customerId = row.get("Customer Id", "")
        invoiceId = row.get("Invoice Id(s)", "")
        notes = row.get("Notes", "")
        checkNo = row.get("Check #", "")
        bankName = row.get("Bank Name", "")
        paymentDate = row.get("Payment Date", "")

        if not checkNo.strip() and not bankName.strip():
            append_to_log(logPath, f"Skipped row with Invoice ID: {invoiceId} because both Check # and Bank Name are empty")
            continue

        amount = convert_to_decimal(row.get("Amount", ""), invoiceId, "amount", logPath)
        if amount is None:
            totalErrors += 1
            continue

        customerName = get_customer_name(customerList, customerId, logPath)
        if customerName == "N/A":
            missingCustomerNames.append(invoiceId)

        if re.search(r"prepaid|prepay", notes, re.IGNORECASE) or re.search(r"CC", checkNo, re.IGNORECASE):
            prepaidInvoices.append(invoiceId)
            append_to_log(logPath, f"Skipped payment for Invoice ID: {invoiceId} (Reason: Prepaid or CC payment)")
            continue

        matchingRows = [r for r in combinedInvoice if r.get("Invoice Number", "") == invoiceId]
        if not matchingRows:
            unmatchedInvoices.append(invoiceId)
            append_to_log(logPath, f"Unmatched invoice: {invoiceId}")
            continue

        for combinedRow in matchingRows:
            colB = combinedRow.get("Column B", "")
            isILS = bool(re.search(r"DHL|ILS", colB, re.IGNORECASE))
            isSHIP = bool(re.search(r"FEDEX|SHIP", colB, re.IGNORECASE))

            if not (isILS or isSHIP):
                append_to_log(logPath, f"Skipping row for Invoice ID: {invoiceId} - Unknown service type: {colB}")
                continue

            perc_str = combinedRow.get("Percentage of Total (%)", "")
            percentage = convert_to_decimal(perc_str, invoiceId, "percentage", logPath)
            if percentage is None:
                totalErrors += 1
                continue

            adjustedAmount = (amount * percentage).quantize(Decimal("0.00"))
            ref_no = generate_ref_no(checkNo, invoiceId, logPath)
            apply_to_invoice = invoiceId[4:] if len(invoiceId) > 4 else invoiceId
            deposit_to = "15000" if isILS else "12000"

            payment_record = {
                "CUSTOMER": customerName,
                "REF NO": ref_no,
                "DATE": paymentDate,
                "PAYMENT METHOD": checkNo,
                "APPLY_TO_INVOICE": apply_to_invoice,
                "AMOUNT": str(adjustedAmount),
                "DEPOSIT TO": deposit_to,
                "MEMO": notes if notes else invoiceId
            }

            if isILS:
                ilsPayments.append(payment_record)
            else:
                shipPayments.append(payment_record)

            totalProcessed += 1

    try:
        for path, data in [
            (ilsReportPath, ilsPayments),
            (shipReportPath, shipPayments),
        ]:
            with path.open("w", encoding="utf-8", newline="") as f:
                if data:
                    writer = csv.DictWriter(f, fieldnames=data[0].keys())
                    writer.writeheader()
                    writer.writerows(data)

        # Identify invoices in raw file but not in reports
        allReportedInvoices = {p["APPLY_TO_INVOICE"] for p in (ilsPayments + shipPayments)}
        rawNotInReport = [
            row["Invoice Id(s)"]
            for row in rawData
            if row.get("Invoice Id(s)") not in allReportedInvoices
        ]
  # Summaries in the log file
        append_to_log(logPath, "\n===========================================")
        append_to_log(logPath, "               RAW FILE SUMMARY            ")
        append_to_log(logPath, "===========================================\n")

        # 1. Raw file count
        rawFileCount = len(rawData)
        append_to_log(logPath, f"1. Raw file count: {rawFileCount}")

        # 2. Prepaid Invoices
        #   (We can re-filter the raw data for 'prepay' or 'prepaid')
        prepaidRawInvoices = [
            r for r in rawData if re.search(r"prepay|prepaid", r.get("Notes", ""), re.IGNORECASE)
        ]
        append_to_log(logPath, "2. Prepaid Invoices:")
        if prepaidRawInvoices:
            for inv in prepaidRawInvoices:
                append_to_log(logPath, f"Invoice ID: {inv.get('Invoice Id(s)', '')}")
        else:
            append_to_log(logPath, "No prepaid invoices found in raw file.")
        append_to_log(logPath, "")

        # 3. CC Payments
        ccPayments = [
            r for r in rawData if re.search(r"CC", r.get("Check #", ""), re.IGNORECASE)
        ]
        append_to_log(logPath, "3. CC Payments:")
        if ccPayments:
            for inv in ccPayments:
                append_to_log(logPath, f"Invoice ID: {inv.get('Invoice Id(s)', '')}")
        else:
            append_to_log(logPath, "No CC payments found in raw file.")
        append_to_log(logPath, "")

        # 4. Total Removed (Prepaid + CC)
        totalRemoved = len(prepaidRawInvoices) + len(ccPayments)
        append_to_log(logPath, f"4. Total Removed: {totalRemoved}")

        # 5. Total Invoices Processed
        totalInvoicesProcessed = rawFileCount - totalRemoved
        append_to_log(logPath, f"5. Total Invoices Processed: {totalInvoicesProcessed}")
        append_to_log(logPath, "")

        # PROCESSING SUMMARY
        append_to_log(logPath, "===========================================")
        append_to_log(logPath, "           PROCESSING SUMMARY             ")
        append_to_log(logPath, "===========================================\n")

        append_to_log(logPath, "PROCESSING COUNTS:")
        append_to_log(logPath, "------------------")
        append_to_log(logPath, f"Total Transactions Processed: {totalProcessed}")
        append_to_log(logPath, f"Total ILS Payments: {len(ilsPayments)}")
        append_to_log(logPath, f"Total SHIP Payments: {len(shipPayments)}")
        append_to_log(logPath, f"Total Errors Encountered: {totalErrors}")
        append_to_log(logPath, "")

        # Not in Combined file
        append_to_log(logPath, "NOT IN COMBINED FILE:")
        append_to_log(logPath, "--------------------")
        if unmatchedInvoices:
            for inv in sorted(unmatchedInvoices):
                append_to_log(logPath, f"Invoice ID: {inv}")
        else:
            append_to_log(logPath, "No unmatched invoices found.")
        append_to_log(logPath, "")

        # Missing customer names
        append_to_log(logPath, "MISSING IN CUSTOMER NAME FILE:")
        append_to_log(logPath, "----------------------------")
        if missingCustomerNames:
            for inv in sorted(missingCustomerNames):
                append_to_log(logPath, f"Invoice ID: {inv}")
        else:
            append_to_log(logPath, "No missing customer names found.")
        append_to_log(logPath, "")

        # New section: IN RAW NOT IN REPORT
        # We'll consider "in raw not in report" as raw lines that ended up not in either ILS or SHIP final data
        # The final data 'APPLY_TO_INVOICE' is the last  part of invoiceId if length>4
        processedInvoices = ilsPayments + shipPayments
        apply_to_invoices = [p["APPLY_TO_INVOICE"] for p in processedInvoices]

        rawNotInReportItems = []
        for rd in rawData:
            rawInv = rd.get("Invoice Id(s)", "")
            # We'll do a naive match: if rawInv ends with any "APPLY_TO_INVOICE".
            # This matches your approach in the code.
            found_match = False
            for applied_id in apply_to_invoices:
                if rawInv.endswith(applied_id):
                    found_match = True
                    break
            if not found_match:
                rawNotInReportItems.append(rd)

        missingInvoicesCount = len(rawNotInReportItems)
        append_to_log(logPath, f"IN RAW NOT IN REPORT: ({missingInvoicesCount} missing)")
        append_to_log(logPath, "--------------------")
        if missingInvoicesCount > 0:
            for inv in sorted(rawNotInReportItems, key=lambda r: r.get("Invoice Id(s)", "")):
                append_to_log(logPath, f"Invoice ID: {inv.get('Invoice Id(s)', '')}")
        else:
            append_to_log(logPath, "No invoices in raw file were excluded from the report.")
        append_to_log(logPath, "")

        # Add processing completion timestamp
        append_to_log(logPath, "===========================================")
        append_to_log(logPath, f" Processing completed on: {datetime.now()}")
        append_to_log(logPath, "===========================================")

        print(f"Processing complete. Log file generated at: {logPath}")

    except Exception as e:
        append_to_log(logPath, f"Error generating reports: {str(e)}")
        print("Failed to generate ILS/SHIP reports. See log for details.")
        return

    #
    # Step 9: Final Summary Output
    #
    print(f"Total Transactions Processed: {totalProcessed}")
    print(f"Total ILS Payments: {len(ilsPayments)}")
    print(f"Total SHIP Payments: {len(shipPayments)}")
    print(f"Total Errors Encountered: {totalErrors}")

    if unmatchedInvoices:
        print(f"Unmatched Invoices found: {len(unmatchedInvoices)}")
        print(f"Check log for details: {logPath}")
    else:
        print("All invoices matched successfully.")

    if missingCustomerNames:
        print(f"Missing customer names for some invoices: {len(missingCustomerNames)}")
        print(f"Check log for details: {logPath}")
    else:
        print("No missing customer names found.")

    if missingInvoicesCount > 0:
        print(f"Invoices in raw file but not in report: {missingInvoicesCount}")
        print(f"Check log for details: {logPath}")
    else:
        print("No invoices in raw file were excluded from the report.")

    print("Script execution completed.")



if __name__ == "__main__":
    main()
