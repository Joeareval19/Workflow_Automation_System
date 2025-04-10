from google.colab import drive
drive.mount('/content/drive')

import os
import re
import sys
import csv
from datetime import datetime

# -------------------------------
# Define Paths & Check Existence
# -------------------------------

# Use the same base folder as in your first script
base_path = "/content/drive/My Drive/Drive Code/EDI_Upload/2025/"
customer_list_path = "/content/drive/My Drive/Drive Code/Customer List.csv"

if not os.path.exists(base_path):
    print(f"Error: Base path does not exist: {base_path}")
    sys.exit(1)
if not os.path.exists(customer_list_path):
    print(f"Error: Customer list file not found: {customer_list_path}")
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

# -------------------------------
# Construct File Paths for Input & Output
# -------------------------------

# Define the raw file folder similar to your first script
raw_folder_path = os.path.join(base_path, latest_week_folder, "1. Gathering")
# Construct the input file name using the date range
input_file = os.path.join(raw_folder_path, f"RAW_EDI_RS_({start_date_str})_({end_date_str}).csv")
if not os.path.exists(input_file):
    print(f"Error: Input file not found: {input_file}")
    sys.exit(1)

# The processed folder is also taken from the same base structure
processed_folder_path = os.path.join(base_path, latest_week_folder, "2. Upload")
os.makedirs(processed_folder_path, exist_ok=True)
# Define the output file for the Invoices & Income report
output_file_path = os.path.join(processed_folder_path, f"ILS EDI ROCKSOLID INVOICES & INCOME_{date_range}.csv")

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

# -------------------------------
# Read CSV Files
# -------------------------------

# Read raw input file
try:
    with open(input_file, mode='r', newline='', encoding='utf-8-sig') as raw_file:
        raw_data = list(csv.DictReader(raw_file))
except Exception as e:
    print(f"Error reading input CSV file: {input_file}\nException: {e}")
    sys.exit(1)

# Read customer list file
try:
    with open(customer_list_path, mode='r', newline='', encoding='utf-8-sig') as customer_file:
        customer_list = list(csv.DictReader(customer_file))
except Exception as e:
    print(f"Error reading customer list CSV file: {customer_list_path}\nException: {e}")
    sys.exit(1)

# Create a dictionary for fast customer lookup (keyed by Customer Id)
customer_dict = {cust['Customer Id']: cust for cust in customer_list}

# -------------------------------
# Process Data & Build Report
# -------------------------------

report_data = []
for current_row in raw_data:
    carrier = current_row.get('Carrier', '')
    customer_no = current_row.get('Customer #', '')
    carrier_inv_no = current_row.get('Carrier Inv. #', '')

    try:
        customer_no_value = float(customer_no)
    except ValueError:
        continue  # Skip if Customer # is not a number

    # Filter logic as per your second code:
    # Only process rows where:
    #  - Carrier is "DHL"
    #  - Customer # is less than 50,000,000
    #  - Carrier Inv. # does not start with 'D'
    if (carrier == "DHL" and
        customer_no_value < 50000000 and
        not carrier_inv_no.startswith('D')):

        customer = customer_dict.get(customer_no)
        if customer:
            # Extract last 8 characters of Invoice Number for INV NO
            invoice_number = current_row.get('Invoice Number', '')
            inv_no = invoice_number[-8:] if invoice_number else "N/A"
            converted_date = convert_to_date(inv_no)

            # Use the Ship Date for MEMO INV ITEM
            ship_date_string = current_row.get('Ship Date', 'N/A')
            memo_inv_item = f"{carrier} | AIRBILL# {current_row.get('Airbill Number', '')} | DATE {ship_date_string}"

            # TERMS: Use first two digits of 'Inv Terms' from the customer record
            inv_terms = customer.get('Inv Terms', '')
            terms = f"NET {inv_terms[:2]}" if inv_terms else "NET"

            report_row = {
                'CUST NO': customer_no,
                'INV NO': inv_no,
                'CUSTOMER': customer.get('Customer', ''),
                'MEMO INV ITEM': memo_inv_item,
                'DATE': converted_date,
                'TERMS': terms,
                'ACCOUNT': "DHL SALES",
                'AMOUNT': current_row.get('Customer Total', ''),
                'REP': customer.get('Customer Salesrep', ''),
            }
            report_data.append(report_row)
        else:
            print(f"Customer not found for Customer #: {customer_no}")

# -------------------------------
# Export the Report
# -------------------------------

try:
    fieldnames = ['CUST NO', 'INV NO', 'CUSTOMER', 'MEMO INV ITEM', 'DATE',
                  'TERMS', 'ACCOUNT', 'AMOUNT', 'REP']
    with open(output_file_path, mode='w', newline='', encoding='utf-8-sig') as output_file:
        writer = csv.DictWriter(output_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(report_data)
    print(f"Filtered report exported to: {output_file_path}")
    print("Script completed successfully.")
except Exception as e:
    print(f"Error writing output CSV file: {output_file_path}\nException: {e}")
    sys.exit(1)

# -------------------------------
# Display Summary of Problematic Data
# -------------------------------

invalid_dates = [row for row in report_data if row['DATE'] in ('Invalid Date', 'N/A')]
if invalid_dates:
    print("\nWarning: Some records have invalid or missing dates:")
    for row in invalid_dates:
        print(f"Customer: {row['CUST NO']}, Invoice: {row['INV NO']}, Date: {row['DATE']}")

missing_invoices = [row for row in report_data if row['INV NO'] == 'N/A']
if missing_invoices:
    print("\nWarning: Some records have missing invoice numbers:")
    for row in missing_invoices:
        print(f"Customer: {row['CUST NO']}, Date: {row['DATE']}")
