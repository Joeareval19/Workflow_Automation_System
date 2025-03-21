
from google.colab import drive
drive.mount('/content/drive')

import os
import re
import csv
import datetime
from pathlib import Path

# -------------------------------
# Define constant paths on Google Drive
# -------------------------------
customer_list_path = r"/content/drive/My Drive/LAMINAR AUTOMATON/Customer List.csv"
raw_file_base_path = Path(r"/content/drive/My Drive/LAMINAR AUTOMATON/EDI_Upload/2025/")

# -------------------------------
# Function to get the latest week folder based on the week number in its name
# -------------------------------
def get_latest_week_folder(base_path):
    pattern = re.compile(r'^Week_(\d+)_\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)$')
    folders = [f for f in base_path.iterdir() if f.is_dir() and pattern.match(f.name)]
    if not folders:
        return None
    # Sort by the week number (the number after "Week_")
    def extract_week_number(folder):
        m = pattern.match(folder.name)
        return int(m.group(1)) if m else 0
    latest_folder = sorted(folders, key=extract_week_number, reverse=True)[0]
    return latest_folder

# -------------------------------
# Function to convert INV NO to date
# -------------------------------
def convert_to_date(inv_no):
    if inv_no == "N/A" or len(inv_no) < 4:
        return "Invalid Date"
    year_letter = inv_no[-4]
    month_letter = inv_no[-3]
    day_str = inv_no[-2:]
    year = ord(year_letter.upper()) - ord('Y') + 2024
    month = ord(month_letter.upper()) - ord('A') + 1
    if month_letter.upper() >= 'J':
        month -= 1
    try:
        day = int(day_str)
        date_obj = datetime.date(year, month, day)
        return date_obj.strftime("%m/%d/%Y")
    except ValueError:
        return "Invalid Date"

# -------------------------------
# Locate the latest week folder
# -------------------------------
latest_week_folder = get_latest_week_folder(raw_file_base_path)

if latest_week_folder:
    # Extract the date range from the folder name.
    date_match = re.search(r'\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)', latest_week_folder.name)
    if date_match:
        start_date_str = date_match.group(1)
        end_date_str = date_match.group(2)
        date_range = f"({start_date_str})_({end_date_str})"
    else:
        print("Invalid folder naming convention.")
        exit(1)
    
    # -------------------------------
    # Locate the raw file in "1. Gathering"
    # -------------------------------
    raw_folder_path = latest_week_folder / "1. Gathering"
    # Build the expected raw file name using the extracted dates.
    input_file = raw_folder_path / f"RAW_EDI_RS_({start_date_str})_({end_date_str}).csv"
    
    if not input_file.exists():
        print(f"Error: Input file not found: {input_file}")
        exit(1)
    
    print(f"Using raw file: {input_file}")
    
    # -------------------------------
    # Read CSV data from the raw file and customer list
    # -------------------------------
    with open(input_file, newline='', encoding='utf-8-sig') as raw_file:
        raw_data = list(csv.DictReader(raw_file))
    
    with open(customer_list_path, newline='', encoding='utf-8-sig') as customer_file:
        customer_list = list(csv.DictReader(customer_file))
    
    filtered_data = [
        row for row in raw_data
        if row['Carrier'] in ["FedEx", "UPS", "FREIGHT"] and
           not (row['Customer #'].startswith("10003217") or row['Customer #'].startswith("10003324"))
    ]
    
    report_data = []
    for row in filtered_data:
        customer = next((c for c in customer_list if c['Customer Id'] == row['Customer #']), None)
        inv_no = row['Invoice Number'][-8:] if row['Invoice Number'] else "N/A"
        converted_date = convert_to_date(inv_no)
        ship_date = row['Ship Date'] if row['Ship Date'] else "N/A"
        carrier = row['Carrier']
        sub_carrier = row['Sub Carrier'].upper() if row.get('Sub Carrier') else ""
        
        if carrier == "UPS":
            account = "UPS SALES"
        elif carrier == "FREIGHT":
            account = "FREIGHT & OTHER"
        elif carrier == "FedEx":
            account = "FEDEX SALES (ENGLAND LOGISTICS)" if sub_carrier == "ENGLAND" else "FEDEX SALES (DESCARTE)"
        else:
            account = "OTHER SALES"
        
        if carrier == "FREIGHT":
            memo_inv_item = f"FREIGHT (LTL) | AIRBILL# {row['Airbill Number']} | DATE {ship_date}"
        else:
            memo_inv_item = f"{carrier} | AIRBILL# {row['Airbill Number']} | DATE {ship_date}"
        
        amount = sum(float(row.get(col, 0) or 0) for col in [
            'Customer Base', 'Chg 1 Total', 'Chg 2 Total', 'Chg 3 Total',
            'Chg 4 Total', 'Chg 5 Total', 'Chg 6 Total', 'Chg 7 Total', 'Chg 8 Total'
        ])
        
        report_data.append({
            'CUST NO': row['Customer #'],
            'INV NO': inv_no,
            'CUSTOMER': customer['Customer'] if customer else "N/A",
            'MEMO INV ITEM': memo_inv_item,
            'DATE': converted_date,
            'TERMS': f"NET {customer['Inv Terms'][:2]}" if customer else "N/A",
            'ACCOUNT': account,
            'AMOUNT': amount,
            'REP': customer['Customer Salesrep'] if customer else "N/A",
        })
    
    # -------------------------------
    # Export the report to the proper folder structure
    # -------------------------------
    # Output folder: 4.QB_Report/SHIPIUM (NOT DHL)
    output_folder_path = latest_week_folder / '2. Upload'
    output_folder_path.mkdir(parents=True, exist_ok=True)
    output_file_name = f"SHIPIUM EDI ROCKSOLID UPS FedEx INVOICES & INCOME_{date_range}.csv"
    output_file_path = output_folder_path / output_file_name
    
    with open(output_file_path, 'w', newline='', encoding='utf-8-sig') as output_file:
        writer = csv.DictWriter(output_file, fieldnames=[
            'CUST NO', 'INV NO', 'CUSTOMER', 'MEMO INV ITEM', 'DATE', 'TERMS', 'ACCOUNT', 'AMOUNT', 'REP'
        ])
        writer.writeheader()
        writer.writerows(report_data)
    
    print(f"Filtered report exported to: {output_file_path}")
    print("Script completed successfully.")
else:
    print("No week folders found in the specified directory.")
