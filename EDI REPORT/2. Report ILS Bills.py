###(ILS_BILLS_REPORT)###


from google.colab import drive
# If needed, force a remount with: drive.mount("/content/drive", force_remount=True)
drive.mount('/content/drive')

import os
import csv
import re
import sys
import glob
from datetime import datetime

# -------------------------------
# Define Paths & Check Existence
# -------------------------------

# Adjust the customer list file path as needed.
customer_list_path = "/content/drive/My Drive/Drive Code/Customer List.csv"
base_path = "/content/drive/My Drive/Drive Code/EDI_Upload/2025/"

# Check if base path exists
if not os.path.exists(base_path):
    print(f"Error: Base path does not exist: {base_path}")
    sys.exit(1)

# -------------------------------
# Identify the Latest Week Folder
# -------------------------------

# List all folders starting with "Week_"
week_folders = [f for f in os.listdir(base_path) if re.match(r"Week_\d+", f)]
if not week_folders:
    print("Error: No week folders found in the specified directory.")
    sys.exit(1)

# Sort folders by week number in descending order and pick the first one
latest_week_folder = sorted(week_folders, key=lambda x: int(x.split("_")[1]), reverse=True)[0]

# Extract start and end dates from the folder name using regex
date_match = re.search(r"\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)", latest_week_folder)
if date_match:
    start_date_str = date_match.group(1)
    end_date_str = date_match.group(2)
    date_range = f"({start_date_str})_({end_date_str})"
else:
    print("Error: Could not extract date from the latest week folder name. Check the folder name format.")
    sys.exit(1)

# Define folder paths for raw (gathering) and processed (upload) files
raw_folder_path = os.path.join(base_path, latest_week_folder, "1. Gathering")
processed_folder_path = os.path.join(base_path, latest_week_folder, "2. Upload")

# Check if raw folder exists
if not os.path.exists(raw_folder_path):
    print(f"Error: Raw folder not found: {raw_folder_path}")
    sys.exit(1)

# -------------------------------
# Construct File Names
# -------------------------------

# Input file: RAW file inside the "1. Gathering" folder
input_file = os.path.join(raw_folder_path, f"RAW_EDI_RS_({start_date_str})_({end_date_str}).csv")
# Output file: Processed file inside the "2. Upload" folder
output_file = os.path.join(processed_folder_path, f"ILS EDI ROCKSOLID BILLS Cost_{date_range}.csv")

# Check if input file exists
if not os.path.exists(input_file):
    print(f"Error: Input file not found: {input_file}")
    sys.exit(1)

# -------------------------------
# Function Definitions
# -------------------------------

def convert_to_date(inv_no):
    """Converts the last 4 characters of INV NO to a formatted date.
       Returns "N/A" if inv_no is "N/A", or "Invalid Date" if conversion fails.
    """
    if inv_no == "N/A":
        return "N/A"
    date_code = inv_no[-4:]
    year_letter = date_code[0]
    month_letter = date_code[1]
    day_str = date_code[2:]
    # Calculate year based on the year letter (adjust base year as needed)
    year = ord(year_letter) - ord('Y') + 2024
    # Calculate month with adjustment for skipped month "I"
    month = ord(month_letter) - ord('A') + 1
    if month_letter >= 'J':
        month -= 1
    try:
        day = int(day_str)
        date_obj = datetime(year, month, day)
        return date_obj.strftime("%m/%d/%Y")
    except ValueError:
        return "Invalid Date"

def get_account(cust_no):
    """Determine account type based on Customer Number prefix."""
    if cust_no.startswith('1'):
        return "DHL COST"
    elif cust_no.startswith('5'):
        return "DHL COST FJ"
    elif cust_no.startswith('6'):
        return "DHL COST FS"
    else:
        return "UNKNOWN"

def get_class(sales_rep):
    """Return the first 5 characters of Sales Rep, if provided."""
    return sales_rep[:5] if sales_rep else ""

# -------------------------------
# Read Input CSV Files
# -------------------------------

# Read raw input file
try:
    with open(input_file, mode='r', newline='', encoding='utf-8-sig') as raw_file:
        raw_data = list(csv.DictReader(raw_file))
except Exception as e:
    print(f"Error reading input CSV file: {input_file}\nException: {e}")
    sys.exit(1)

# Check and read customer list file
if not os.path.exists(customer_list_path):
    print(f"Error: Customer list file not found: {customer_list_path}")
    sys.exit(1)

try:
    with open(customer_list_path, mode='r', newline='', encoding='utf-8-sig') as customer_file:
        customer_list = list(csv.DictReader(customer_file))
except Exception as e:
    print(f"Error reading customer list CSV file: {customer_list_path}\nException: {e}")
    sys.exit(1)

# Create a dictionary for fast customer lookup
customer_dict = {cust['Customer Id']: cust['Customer'] for cust in customer_list}

# -------------------------------
# Process Data and Build Report
# -------------------------------

report_data = []
for current_row in raw_data:
    carrier = current_row.get('Carrier', '')
    carrier_inv_no = current_row.get('Carrier Inv. #', '')
    customer_name = current_row.get('Customer', '')

    # Apply filters based on business logic
    if (carrier == "DHL" and
        not carrier_inv_no.startswith('D') and
        not re.match(r'^Curlmix', customer_name) and
        not re.match(r'^TRG', customer_name)):

        cust_no = current_row.get('Customer #', '')
        customer = customer_dict.get(cust_no)
        if customer:
            invoice_number = current_row.get('Invoice Number', '')
            inv_no = invoice_number[-8:] if invoice_number else "N/A"
            converted_date = convert_to_date(inv_no)
            report_row = {
                'DATE': converted_date,
                'CUST NO': cust_no,
                'CUSTOMER': customer,
                'CLASS': get_class(current_row.get('Sales Rep', '')),
                'BILL NO': carrier_inv_no,
                'VENDOR': "DHL EXPRESS",
                'MEMO BILL ITEM': f"{carrier} | AIRBILL# {current_row.get('Airbill Number', '')} | {current_row.get('Service Type', '')}",
                'AMOUNT': current_row.get('Carrier Cost Total', ''),
                'ACCOUNT': get_account(cust_no),
            }
            report_data.append(report_row)
        else:
            print(f"Warning: Customer not found for Customer #: {cust_no}")

# Sort the report data by BILL NO (A-Z)
sorted_report_data = sorted(report_data, key=lambda x: x['BILL NO'])

# -------------------------------
# Write the Processed Data to CSV
# -------------------------------

# Ensure the processed folder exists
os.makedirs(processed_folder_path, exist_ok=True)

try:
    with open(output_file, mode='w', newline='', encoding='utf-8-sig') as out_file:
        fieldnames = ['DATE', 'CUST NO', 'CUSTOMER', 'CLASS', 'BILL NO', 'VENDOR', 'MEMO BILL ITEM', 'AMOUNT', 'ACCOUNT']
        writer = csv.DictWriter(out_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(sorted_report_data)
except Exception as e:
    print(f"Error writing output CSV file: {output_file}\nException: {e}")
    sys.exit(1)

print(f"DHL Billing report exported to: {output_file} (sorted by BILL NO A-Z)")
print("Script completed successfully.")

# -------------------------------
# Display Summary of Problematic Data
# -------------------------------

invalid_dates = [row for row in sorted_report_data if row['DATE'] in ('Invalid Date', 'N/A')]
if invalid_dates:
    print("\nWarning: Some records have invalid or missing dates:")
    for row in invalid_dates:
        print(f"Customer: {row['CUST NO']}, Bill No: {row['BILL NO']}, Date: {row['DATE']}")

unknown_accounts = [row for row in sorted_report_data if row['ACCOUNT'] == 'UNKNOWN']
if unknown_accounts:
    print("\nWarning: Some records have unknown account types:")
    for row in unknown_accounts:
        print(f"Customer: {row['CUST NO']}, Bill No: {row['BILL NO']}, Account: {row['ACCOUNT']}")

zero_amounts = [row for row in sorted_report_data if float(row.get('AMOUNT', '0') or '0') == 0]
if zero_amounts:
    print("\nWarning: Some records have zero amounts:")
    for row in zero_amounts:
        print(f"Customer: {row['CUST NO']}, Bill No: {row['BILL NO']}, Amount: {row['AMOUNT']}")
