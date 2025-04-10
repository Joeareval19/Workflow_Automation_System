



from google.colab import drive
drive.mount('/content/drive')

import os
import re
import csv
import datetime
from pathlib import Path
import decimal


# -------------------------------
# Define Paths on Google Drive
# -------------------------------
customer_list_path = r"/content/drive/My Drive/Drive Code/Customer List.csv"
raw_file_base_path = Path(r"/content/drive/My Drive/Drive Code/EDI_Upload/2025/")

# -------------------------------
# Functions
# -------------------------------
def get_latest_week_folder(base_path):
    week_folders = [f for f in base_path.iterdir()
                    if f.is_dir() and re.match(r'^Week_\d+_\(\d{2}\.\d{2}\.\d{2}\)_\(\d{2}\.\d{2}\.\d{2}\)$', f.name)]
    if not week_folders:
        return None

    # Extract the week number from the folder name (e.g., "Week_5_(...)" -> 5)
    def extract_week_number(folder):
        parts = folder.name.split("_")
        try:
            return int(parts[1])
        except Exception:
            return 0

    # Sort by the week number in descending order
    latest_folder = sorted(week_folders, key=extract_week_number, reverse=True)[0]
    return latest_folder

def convert_to_date(inv_no):
    if inv_no == "N/A":
        return "N/A"
    date_code = inv_no[-4:]
    year = ord(date_code[0]) - ord('Y') + 2024
    month = ord(date_code[1]) - ord('A') + 1
    month -= 1 if date_code[1] > 'I' else 0
    day = int(date_code[2:])
    try:
        return datetime.date(year, month, day).strftime("%m/%d/%Y")
    except ValueError:
        return "Invalid Date"

def get_account(cust_no, carrier, sub_carrier):
    if carrier == "FedEx":
        if sub_carrier == "England":
            return "FEDEX COST (ENGLAND LOGISTICS)"
        elif sub_carrier == "RSIS":
            return "FEDEX COST (DESCARTES)"
        return "FEDEX COST"
    elif carrier == "UPS":
        return "UPS COST"
    return "UNKNOWN"

def get_class(sales_rep):
    return sales_rep[:5] if sales_rep else ""

# -------------------------------
# Locate the Latest Week Folder
# -------------------------------
latest_week_folder = get_latest_week_folder(raw_file_base_path)

if latest_week_folder:
    # Extract the date range from the folder name.
    match = re.search(r'\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)', latest_week_folder.name)
    if match:
        start_date_str = match.group(1)
        end_date_str = match.group(2)
        date_range = f"({start_date_str})_({end_date_str})"
    else:
        print("Invalid folder naming convention.")
        exit(1)

    # -------------------------------
    # Locate the Raw File Using the First Code's Approach
    # -------------------------------
    raw_folder_path = latest_week_folder / "1. Gathering"
    # Build the expected raw file name using the date range.
    input_file = raw_folder_path / f"RAW_EDI_RS_({start_date_str})_({end_date_str}).csv"

    if not input_file.exists():
        print(f"Error: Input file not found: {input_file}")
        exit(1)

    # -------------------------------
    # Read CSV Files
    # -------------------------------
    with open(input_file, newline='', encoding='utf-8-sig') as raw_file:
        raw_data = list(csv.DictReader(raw_file))

    with open(customer_list_path, newline='', encoding='utf-8-sig') as customer_file:
        customer_list = list(csv.DictReader(customer_file))

    report_data = []
    for row in raw_data:
        # Clean up keys and values.
        row = {k.strip(): v.strip() if isinstance(v, str) else v for k, v in row.items()}
        carrier = row.get('Carrier', '')
        cust_no = row.get('Customer #', '')
        carrier_inv_no = row.get('Carrier Inv. #', '')

        # Process only FedEx or UPS rows that pass the filter criteria.
        if carrier in ["FedEx", "UPS"] and not carrier_inv_no.startswith('D') and cust_no not in ["10003217", "10003324"]:
            # Lookup customer based on Customer Id.
            customer = next((c for c in customer_list if c['Customer Id'] == cust_no), None)
            if customer:
                inv_no = row.get('Invoice Number', 'N/A')[-8:]
                date_converted = convert_to_date(inv_no)
                sub_carrier = row.get('Sub Carrier', '').strip() or ("RSIS" if carrier == "FedEx" else "")

                vendor = ("DESCARTES" if sub_carrier == "RSIS"
                          else "ENGLAND LOGISTICS" if sub_carrier == "England"
                          else f"{carrier} ENGLAND")

                airbill_number_raw = row.get('Airbill Number') or row.get('Air Bill Number') or row.get('AirBill') or ""
                airbill_number_raw = airbill_number_raw.strip()
###########


                # --- START: CORRECTED Handle scientific notation ---
                airbill_number_formatted = airbill_number_raw # Default to raw value
                if airbill_number_raw: # Only process if not empty
                    try:
                        # Use Decimal for accurate conversion from potential scientific notation
                        dec_val = decimal.Decimal(str(airbill_number_raw)) # Ensure input is string

                        # Convert Decimal to integer (truncates any decimal part), then to string.
                        # This forces the full number representation without exponent.
                        airbill_number_formatted = str(int(dec_val))

                    except (ValueError, decimal.InvalidOperation, TypeError) as e:
                        # If conversion fails (e.g., non-numeric "ABC123XYZ", empty string after strip),
                        # keep the original raw value.
                        # Optional: Log which specific values failed if needed for debugging
                        # print(f"Debug: Row {row_index + 2}: Could not convert airbill '{airbill_number_raw}' to number: {e}")
                        airbill_number_formatted = airbill_number_raw # Keep original on error
                    except Exception as e: # Catch any other unexpected errors during conversion
                         print(f"Warning: Row {row_index + 2}: Unexpected error converting airbill '{airbill_number_raw}': {e}")
                         airbill_number_formatted = airbill_number_raw # Fallback to original
                # --- END: CORRECTED Handle scientific notation ---

                # Construct the memo field using the potentially formatted airbill number
                service_type = row.get('Service Type', '') # Get service type
                memo_bill_item = f"{carrier} | AIRBILL# {airbill_number_formatted} | {service_type} | {sub_carrier}"
                report_data.append({
                    'DATE': date_converted,
                    'CUST NO': cust_no,
                    'CUSTOMER': customer['Customer'],
                    'CLASS': get_class(row.get('Sales Rep')),
                    'BILL NO': carrier_inv_no,
                    'VENDOR': vendor,
                    'MEMO BILL ITEM': memo_bill_item,
                    'AMOUNT': row.get('Carrier Cost Total', '0'),
                    'ACCOUNT': get_account(cust_no, carrier, sub_carrier)
                })
            else:
                print(f"Customer not found: {cust_no}")

    sorted_report = sorted(report_data, key=lambda x: str(x['BILL NO']))

    # -------------------------------
    # Export the Report
    # -------------------------------
    output_folder_path = latest_week_folder / '2. Upload'
    output_folder_path.mkdir(parents=True, exist_ok=True)
    output_file_path = output_folder_path / f"SHIPIUM EDI ROCKSOLID UPS FedEx BILLS Cost_{date_range}.csv"

    with open(output_file_path, 'w', newline='', encoding='utf-8-sig') as output_file:
        writer = csv.DictWriter(output_file, fieldnames=['DATE', 'CUST NO', 'CUSTOMER', 'CLASS', 'BILL NO', 'VENDOR', 'MEMO BILL ITEM', 'AMOUNT', 'ACCOUNT'])
        writer.writeheader()
        writer.writerows(sorted_report)

    print(f"Report successfully exported to {output_file_path}")
else:
    print("No week folders found.")
