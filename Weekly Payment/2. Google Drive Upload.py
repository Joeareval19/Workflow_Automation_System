import os
import re
import glob
import csv
import sys
from datetime import datetime
from collections import defaultdict

def parse_date(date_str):
    """
    Attempt to parse a date string using common formats.
    If parsing fails, return the original string.
    """
    for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d", "%m/%d/%y", "%m-%d-%y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return date_str  # Fallback if parsing fails

def main():
    # --- Step 1: Locate the Latest Week Folder ---
    baseDir = r"C:\Users\User\Desktop\JEAV\Weekly Payments Reconcile (daily)\RS vs QB"
    # Regex: Week_<week_number>_(<start_date>)_(<end_date>)
    folder_pattern = re.compile(r'^Week_(\d+)_\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)$')
    week_folders = []

    for entry in os.listdir(baseDir):
        full_path = os.path.join(baseDir, entry)
        if os.path.isdir(full_path):
            match = folder_pattern.match(entry)
            if match:
                try:
                    week_number = int(match.group(1))
                    # Dates in folder name use the format MM.dd.yy
                    start_date_str = match.group(2)
                    end_date_str = match.group(3)
                    end_date = datetime.strptime(end_date_str, "%m.%d.%y")
                    week_folders.append({
                        "folder_name": entry,
                        "full_path": full_path,
                        "week_number": week_number,
                        "end_date": end_date
                    })
                except Exception as e:
                    print(f"Warning: Failed to parse folder '{entry}': {e}", file=sys.stderr)
                    continue

    if not week_folders:
        print(f"Error: No folders matching the pattern were found in {baseDir}", file=sys.stderr)
        sys.exit(1)

    # Pick the folder with the latest end date
    latest_folder = max(week_folders, key=lambda x: x["end_date"])
    latest_folder_path = latest_folder["full_path"]

    # --- Step 2: Find the RAW CSV File in the '1.Gathering' Subfolder ---
    gathering_dir = os.path.join(latest_folder_path, "1.Gathering")
    raw_files = glob.glob(os.path.join(gathering_dir, "RAW_*.csv"))
    if not raw_files:
        print(f"Error: No RAW file found in {gathering_dir}", file=sys.stderr)
        sys.exit(1)
    input_file_path = raw_files[0]  # Pick the first RAW file

    # --- Step 3: Define Other File Paths ---
    combined_report_path = r"C:\Users\User\Desktop\JEAV\Weekly Payments Reconcile (daily)\Weekly Payables\Combined_Invoice_Report.csv"
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    output_file_path = os.path.join(desktop_path, "CleanedData.csv")

    # --- Step 4: Import CSV Data ---
    try:
        with open(input_file_path, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            raw_data = list(reader)
    except Exception as e:
        print(f"Error reading input CSV file: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        with open(combined_report_path, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            combined_data = list(reader)
    except Exception as e:
        print(f"Error reading combined invoice report CSV: {e}", file=sys.stderr)
        sys.exit(1)

    # --- Step 5: Build a Map of Invoice Info ---
    invoice_info_map = defaultdict(list)
    for invoice in combined_data:
        invoice_number = invoice.get("Invoice Number", "").strip()
        column_b_value = invoice.get("Column B", "").strip()
        percentage = invoice.get("Percentage of Total (%)", "").strip()
        if invoice_number:
            invoice_info_map[invoice_number].append({
                "ColumnB": column_b_value,
                "Percentage": percentage
            })

    # --- Step 6: Process and Clean the Data ---
    output_records = []
    payment_method_order = ['Check', 'ACH', 'Credit Card', 'Other']

    for row in raw_data:
        check_num = row.get("Check #", "").strip()
        notes = row.get("Notes", "")
        # Filter out rows:
        #   - Where Check # contains 'CC'
        #   - Where Notes contain 're-allocated' (case-insensitive)
        if "CC" in check_num:
            continue
        if re.search(r'(?i)re-allocated', notes):
            continue

        # If Check # is empty, use Bank Name
        if not check_num:
            check_num = row.get("Bank Name", "").strip()

        # Replace "American Express" with "Amex"
        if "American Express" in check_num:
            check_num = check_num.replace("American Express", "Amex")

        # --- Determine Payment Method ---
        check_num_upper = check_num.upper()
        if re.match(r'^(AMEX|MASTERCARD|MC|VISA|DISCOVER|CARD)', check_num_upper):
            payment_method = "Credit Card"
        elif "ACH" in check_num_upper:
            payment_method = "ACH"
        elif check_num.isdigit():
            payment_method = "Check"
        else:
            payment_method = "Other"

        # --- Process the Amount ---
        raw_amount_str = re.sub(r'[^\d\.-]', '', row.get("Amount", "0"))
        try:
            raw_amount = float(raw_amount_str)
        except ValueError:
            raw_amount = 0.0

        # --- Process Invoice IDs ---
        invoice_ids = [inv.strip() for inv in re.split(r'[,;]', row.get("Invoice Id(s)", "")) if inv.strip()]
        has_invoice_info = False

        for invoice_id in invoice_ids:
            if invoice_id in invoice_info_map:
                for info in invoice_info_map[invoice_id]:
                    column_b_value = info.get("ColumnB", "")
                    percentage_str = re.sub(r'[^\d\.-]', '', info.get("Percentage", ""))
                    if percentage_str:
                        try:
                            percentage_decimal = float(percentage_str)
                        except ValueError:
                            print(f"Warning: Invalid percentage for Invoice ID {invoice_id} in combined invoice report.", file=sys.stderr)
                            continue
                        invoice_amount = raw_amount * percentage_decimal
                        record = {
                            "Amount": invoice_amount,
                            "Column B": column_b_value,
                            "Check #": check_num,
                            "Payment Date": row.get("Payment Date", ""),
                            "Customer Id": row.get("Customer Id", ""),
                            "Customer": row.get("Customer", ""),
                            "Invoice Id(s)": invoice_id,
                            "Notes": notes,
                            "Payment Method": payment_method
                        }
                        output_records.append(record)
                        has_invoice_info = True
                    else:
                        print(f"Warning: Percentage for Invoice ID {invoice_id} not found or invalid in combined invoice report.", file=sys.stderr)
            else:
                print(f"Warning: Invoice ID {invoice_id} not found in combined invoice report.", file=sys.stderr)

        # If no invoice info was found, create a record with the raw amount
        if not has_invoice_info:
            record = {
                "Amount": raw_amount,
                "Column B": "",
                "Check #": check_num,
                "Payment Date": row.get("Payment Date", ""),
                "Customer Id": row.get("Customer Id", ""),
                "Customer": row.get("Customer", ""),
                "Invoice Id(s)": row.get("Invoice Id(s)", ""),
                "Notes": notes,
                "Payment Method": payment_method
            }
            output_records.append(record)

    # --- Step 7: Sort the Output Records ---
    def sort_key(rec):
        # Try to parse the Payment Date; fall back to the raw string if needed.
        pd = parse_date(rec.get("Payment Date", ""))
        if not isinstance(pd, datetime):
            pd = rec.get("Payment Date", "")
        try:
            method_index = payment_method_order.index(rec.get("Payment Method", "Other"))
        except ValueError:
            method_index = len(payment_method_order)
        return (pd, method_index, rec.get("Check #", ""))

    sorted_records = sorted(output_records, key=sort_key)

    # --- Step 8: Write the CSV with Grouped Data and Spacer Lines ---
    headers = ["Amount", "Column B", "Check #", "Payment Date", "Customer Id", "Customer", "Invoice Id(s)", "Notes", "Payment Method"]
    csv_lines = []
    # Write header line
    csv_lines.append(",".join(headers))

    # Group records by Payment Date
    # (Grouping is done based on exact string equality of the Payment Date.)
    grouped_by_date = {}
    for rec in sorted_records:
        key = rec.get("Payment Date", "")
        grouped_by_date.setdefault(key, []).append(rec)

    date_keys = list(grouped_by_date.keys())
    for i, date_key in enumerate(date_keys):
        records_for_date = grouped_by_date[date_key]
        # Group records by Payment Method within this date group (using desired order)
        method_groups = {method: [] for method in payment_method_order}
        for rec in records_for_date:
            method = rec.get("Payment Method", "Other")
            if method in method_groups:
                method_groups[method].append(rec)
            else:
                method_groups.setdefault(method, []).append(rec)
        # Process each Payment Method group in the desired order
        method_keys = [m for m in payment_method_order if m in method_groups and method_groups[m]]
        for j, method in enumerate(method_keys):
            for rec in method_groups[method]:
                # Use csv module to ensure proper quoting/escaping
                from io import StringIO
                output_io = StringIO()
                writer = csv.writer(output_io)
                writer.writerow([rec.get(col, "") for col in headers])
                csv_line = output_io.getvalue().strip()
                csv_lines.append(csv_line)
            # After each Payment Method group except the last, insert two empty lines.
            if j != len(method_keys) - 1:
                empty_line = "," * (len(headers) - 1)
                csv_lines.append(empty_line)
                csv_lines.append(empty_line)
        # After each Payment Date group (except the last), insert two empty lines.
        if i != len(date_keys) - 1:
            empty_line = "," * (len(headers) - 1)
            csv_lines.append(empty_line)
            csv_lines.append(empty_line)

    # Write all CSV lines to the output file.
    try:
        with open(output_file_path, "w", newline='', encoding='utf-8') as f:
            for line in csv_lines:
                f.write(line + "\n")
    except Exception as e:
        print(f"Error writing output CSV file: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Cleaning, filtering, and sorting completed. Cleaned data saved to: {output_file_path}")

if __name__ == "__main__":
    main()
