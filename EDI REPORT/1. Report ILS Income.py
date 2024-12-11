import os
import re
import sys
import csv
import glob
from datetime import datetime

# Define the constant path for the customer list file
customer_list_path = r"C:\Users\User\Desktop\JEAV\Customer List.csv"

# Define the base path for the raw file directory
raw_file_base_path = r"C:\Users\User\Desktop\JEAV\EDI Reconcile (monday)"

# Function to get the most recent week folder
def get_latest_week_folder():
    pattern = re.compile(r'^Week_\d+_\(\d{2}\.\d{2}\.\d{2}\)_\(\d{2}\.\d{2}\.\d{2}\)$')
    folders = []
    for entry in os.listdir(raw_file_base_path):
        full_path = os.path.join(raw_file_base_path, entry)
        if os.path.isdir(full_path) and pattern.match(entry):
            folders.append((full_path, os.path.getmtime(full_path)))
    if folders:
        # Sort folders by last modified time in descending order
        latest_folder = sorted(folders, key=lambda x: x[1], reverse=True)[0][0]
        return latest_folder
    else:
        return None

# Function to convert INV NO to date with adjusted month letters
def convert_to_date(inv_no):
    if inv_no == "N/A":
        return "N/A"
    date_code = inv_no[-4:]
    year_letter = date_code[0]
    month_letter = date_code[1]
    day_str = date_code[2:]

    # Calculate year based on the year letter
    year = ord(year_letter) - ord('Y') + 2024

    # Adjust for the skipped month 'I'
    month = ord(month_letter) - ord('A') + 1
    if month_letter >= 'J':
        month -= 1  # If September or later, decrease month by 1

    try:
        day = int(day_str)
        date_obj = datetime(year, month, day)
        return date_obj.strftime("%m/%d/%Y")
    except ValueError:
        return "Invalid Date"

# Get the latest week folder
latest_week_folder = get_latest_week_folder()

if latest_week_folder:
    # Extract date range from the folder name
    folder_name = os.path.basename(latest_week_folder)
    match = re.search(r'\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)', folder_name)
    if match:
        start_date_str = match.group(1)
        end_date_str = match.group(2)
        date_range = f"({start_date_str})_({end_date_str})"
    else:
        print(f"Unable to extract date range from folder name: {folder_name}")
        sys.exit(1)

    # Construct the full path to the raw file
    raw_file_path_pattern = os.path.join(
        latest_week_folder, "1.Gathering_Data", "RS", "RAW_EDI_RS_*.csv"
    )

    # Get the most recent raw file in the folder
    raw_files = glob.glob(raw_file_path_pattern)
    if raw_files:
        latest_raw_file = max(raw_files, key=os.path.getmtime)
        print(f"Using raw file: {latest_raw_file}")

        # Import CSV files
        with open(latest_raw_file, mode='r', newline='', encoding='utf-8-sig') as raw_file:
            filtered_data = list(csv.DictReader(raw_file))

        with open(customer_list_path, mode='r', newline='', encoding='utf-8-sig') as customer_file:
            customer_list = list(csv.DictReader(customer_file))

        # Create a dictionary for customer lookup
        customer_dict = {cust['Customer Id']: cust for cust in customer_list}

        report_data = []
        for current_row in filtered_data:
            carrier = current_row.get('Carrier', '')
            customer_no = current_row.get('Customer #', '')
            carrier_inv_no = current_row.get('Carrier Inv. #', '')
            try:
                customer_no_value = float(customer_no)
            except ValueError:
                continue  # Skip if Customer # is not a number

            if (
                carrier == "DHL" and
                customer_no_value < 50000000 and
                not carrier_inv_no.startswith('D')
            ):
                customer = customer_dict.get(customer_no)
                if customer:
                    # Extract last 8 characters of Invoice Number for INV NO
                    invoice_number = current_row.get('Invoice Number', '')
                    if invoice_number:
                        inv_no = invoice_number[-8:]
                    else:
                        inv_no = "N/A"

                    # Convert INV NO to date
                    converted_date = convert_to_date(inv_no)

                    # Use the Ship Date for MEMO INV ITEM
                    ship_date_string = current_row.get('Ship Date', 'N/A')

                    memo_inv_item = f"{carrier} | AIRBILL# {current_row.get('Airbill Number', '')} | DATE {ship_date_string}"

                    # TERMS: Use first two digits of 'Inv Terms'
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
        # Define output folder path
        output_folder_path = os.path.join(
            latest_week_folder, '4.QB_Report', 'ILS (DHL)'
        )

        # Ensure the output directory exists
        os.makedirs(output_folder_path, exist_ok=True)

        # Define output file name
        output_file_name = f"ILS EDI ROCKSOLID INVOICES & INCOME_{date_range}.csv"

        # Define output file path
        output_file_path = os.path.join(output_folder_path, output_file_name)

        # Export data to CSV
        fieldnames = [
            'CUST NO', 'INV NO', 'CUSTOMER', 'MEMO INV ITEM', 'DATE',
            'TERMS', 'ACCOUNT', 'AMOUNT', 'REP'
        ]
        with open(output_file_path, mode='w', newline='', encoding='utf-8-sig') as output_file:
            writer = csv.DictWriter(output_file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(report_data)

        print(f"Filtered report exported to: {output_file_path}")
        print("Script completed successfully.")

        # Display summary of problematic data
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

    else:
        print("No raw file found in the latest week folder.")
else:
    print("No week folders found in the specified directory.")
