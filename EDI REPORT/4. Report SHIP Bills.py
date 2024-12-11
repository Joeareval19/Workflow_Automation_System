import os
import re
import csv
import datetime
from pathlib import Path

# Define the constant paths
customer_list_path = r"C:\Users\User\Desktop\JEAV\Customer List.csv"
raw_file_base_path = Path(r"C:\Users\User\Desktop\JEAV\EDI Reconcile (monday)")

# Function to get the most recent week folder
def get_latest_week_folder(base_path):
    week_folders = [f for f in base_path.iterdir() if f.is_dir() and re.match(r'^Week_\d+_\(\d{2}\.\d{2}\.\d{2}\)_\(\d{2}\.\d{2}\.\d{2}\)$', f.name)]
    if week_folders:
        return max(week_folders, key=lambda f: f.stat().st_mtime)
    return None

# Function to convert INV NO to date
def convert_to_date(inv_no):
    if inv_no == "N/A":
        return "N/A"
    
    date_code = inv_no[-4:]
    year_letter = date_code[0]
    month_letter = date_code[1]
    day = int(date_code[2:])
    
    year = ord(year_letter) - ord('Y') + 2024
    month = ord(month_letter) - ord('A') + 1
    if month_letter > 'I':
        month -= 1  # Decrement month if it's after 'I'
    
    try:
        return datetime.date(year, month, day).strftime("%m/%d/%Y")
    except ValueError:
        return "Invalid Date"

# Function to determine the account based on CUST NO, Carrier, and Sub Carrier
def get_account(cust_no, carrier, sub_carrier):
    if carrier == "FedEx":
        if sub_carrier == "England":
            return "FEDEX COST (ENGLAND LOGISTICS)"
        elif sub_carrier == "RSIS":
            return "FEDEX COST (DESCARTES)"
        else:
            return "FEDEX COST"
    elif carrier == "UPS":
        return "UPS COST"
    else:
        return "UNKNOWN"
    

# Function to process Sales Rep for CLASS
def get_class(sales_rep):
    return sales_rep[:5] if sales_rep else ""

# Get the latest week folder
latest_week_folder = get_latest_week_folder(raw_file_base_path)

if latest_week_folder:
    # Extract date range from the folder name
    match = re.search(r'\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)', latest_week_folder.name)
    if match:
        start_date_str, end_date_str = match.groups()
        date_range = f"({start_date_str})_({end_date_str})"
    else:
        print(f"Unable to extract date range from folder name: {latest_week_folder.name}")
        exit(1)

    # Construct the full path to the raw file
    raw_file_path_pattern = latest_week_folder / "1.Gathering_Data/RS/RAW_EDI_RS_*.csv"
    raw_files = list(raw_file_path_pattern.parent.glob(raw_file_path_pattern.name))
    latest_raw_file = max(raw_files, key=lambda f: f.stat().st_mtime) if raw_files else None

    if latest_raw_file:
        print(f"Using raw file: {latest_raw_file}")
        
        # Import CSV files
        with open(latest_raw_file, newline='') as raw_file:
            raw_data = list(csv.DictReader(raw_file))
        
        # Normalize the data to strip spaces and unify the format
        for row in raw_data:
            for key in row.keys():
                if key is not None:
                    row[key] = row[key].strip() if isinstance(row[key], str) else row[key]
        
        
        with open(customer_list_path, newline='') as customer_file:
            customer_list = list(csv.DictReader(customer_file))
        
        report_data = []
        for row in raw_data:
            if row['Carrier'] in ["FedEx", "UPS"] and not row['Carrier Inv. #'].startswith('D') and not (row['Customer #'] in ["10003217", "10003324"]):
                customer = next((c for c in customer_list if c['Customer Id'] == row['Customer #']), None)
                
                if customer:
                    inv_no = row['Invoice Number'][-8:] if row['Invoice Number'] else "N/A"
                    converted_date = convert_to_date(inv_no)
                    
                    if row['Carrier'] == "FedEx" and not row['Sub Carrier'].strip():
                        sub_carrier = "RSIS"
                        print("FedEx has empty subcarrier")
                    else:
                        sub_carrier = row['Sub Carrier']
                    vendor = "DESCARTES" if sub_carrier == "RSIS" else "ENGLAND LOGISTICS" if sub_carrier == "ENGLAND" else f"{row['Carrier']} EXPRESS"
                    
                    report_data.append({
                        'DATE': converted_date,
                        'CUST NO': row['Customer #'],
                        'CUSTOMER': customer['Customer'],
                        'CLASS': get_class(row['Sales Rep']),
                        'BILL NO': row['Carrier Inv. #'],
                        'VENDOR': vendor,
                        'MEMO BILL ITEM': f"{row['Carrier']} | AIRBILL# {row.get('Airbill Number', 'N/A')} | {row['Service Type']} | {sub_carrier}",
                        'AMOUNT': row['Carrier Cost Total'],
                        'ACCOUNT': get_account(row['Customer #'], row['Carrier'], sub_carrier)
                    })
                else:
                    print(f"Customer not found for Customer #: {row['Customer #']}")

        # Sort the report data by BILL NO from smallest to largest
        sorted_report_data = sorted(report_data, key=lambda x: int(x['BILL NO']))

        # Define output folder path
        output_folder_path = latest_week_folder / '4.QB_Report/SHIPIUM (NOT DHL)'
        try:
            output_folder_path.mkdir(parents=True, exist_ok=True)
        except PermissionError:
            print(f"Permission denied: Unable to create directory at {output_folder_path}")
            exit(1)

        # Define output file name and path
        output_file_name = f"SHIPIUM EDI ROCKSOLID UPS FedEx BILLS Cost_{date_range}.csv"
        output_file_path = output_folder_path / output_file_name

        # Export sorted data to CSV
        with open(output_file_path, 'w', newline='', encoding='utf-8') as output_file:
            fieldnames = ['DATE', 'CUST NO', 'CUSTOMER', 'CLASS', 'BILL NO', 'VENDOR', 'MEMO BILL ITEM', 'AMOUNT', 'ACCOUNT']
            writer = csv.DictWriter(output_file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(sorted_report_data)

        print(f"Carrier Billing report exported to: {output_file_path}")
        print("Script completed successfully.")

        # Display summary of problematic data
        invalid_dates = [row for row in sorted_report_data if row['DATE'] in ['Invalid Date', 'N/A']]
        if invalid_dates:
            print("\nWarning: Some records have invalid or missing dates:")
            for row in invalid_dates:
                print(f"Customer: {row['CUST NO']}, Bill No: {row['BILL NO']}, Date: {row['DATE']}")

        unknown_accounts = [row for row in sorted_report_data if row['ACCOUNT'] == 'UNKNOWN']
        if unknown_accounts:
            print("\nWarning: Some records have unknown account types:")
            for row in unknown_accounts:
                print(f"Customer: {row['CUST NO']}, Bill No: {row['BILL NO']}, Account: {row['ACCOUNT']}")

        zero_amounts = [row for row in sorted_report_data if float(row['AMOUNT']) == 0]
        if zero_amounts:
            print("\nWarning: Some records have zero amounts:")
            for row in zero_amounts:
                print(f"Customer: {row['CUST NO']}, Bill No: {row['BILL NO']}, Amount: {row['AMOUNT']}")
    else:
        print("No raw file found in the latest week folder.")
else:
    print("No week folders found in the specified directory.")
