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
    
    # Placeholder conversion logic, replace with actual logic as needed
    return "ConvertedDatePlaceholder"

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
        filtered_data = []
        with open(latest_raw_file, newline='') as raw_file:
            raw_data = list(csv.DictReader(raw_file))
            
            # Filter data
            filtered_data = [
                row for row in raw_data
                if row['Carrier'] in ["FedEx", "UPS", "FREIGHT"]
                and not (row['Customer #'] in ['10003217', '10003324'])
            ]

        with open(customer_list_path, newline='') as customer_file:
            customer_list = list(csv.DictReader(customer_file))
        
        report_data = []
        for row in filtered_data:
            customer = next((c for c in customer_list if c['Customer Id'] == row['Customer #']), None)

            inv_no = row['Invoice Number'][-8:] if row['Invoice Number'] else "N/A"
            converted_date = convert_to_date(inv_no)

            ship_date_string = row['Ship Date'] if row['Ship Date'] else "N/A"
            
            account_value = "OTHER SALES"
            if row['Carrier'] == "UPS":
                account_value = "UPS SALES"
            elif row['Carrier'] == "FREIGHT":
                account_value = "FREIGHT & OTHER"
            elif row['Carrier'] == "FedEx":
                sub_carrier = row['Sub Carrier']
                if sub_carrier == "ENGLAND":
                    account_value = "FEDEX SALES (ENGLAND LOGISTICS)"
                elif sub_carrier == "RSIS":
                    account_value = "FEDEX SALES (DESCARTE)"
                else:
                    account_value = "FEDEX SALES (DESCARTE)"
            
            memo_inv_item = (
                f"FREIGHT (LTL) | AIRBILL# {row.get('Airbill Number', 'N/A')} | DATE {ship_date_string}"
                if row['Carrier'] == "FREIGHT"
                else f"{row['Carrier']} | AIRBILL# {row.get('Airbill Number', 'N/A')} | DATE {ship_date_string}"
            )

            report_data.append({
                'CUST NO': row['Customer #'],
                'INV NO': inv_no,
                'CUSTOMER': customer['Customer'] if customer else "N/A",
                'MEMO INV ITEM': memo_inv_item,
                'DATE': converted_date,
                'TERMS': f"NET {customer['Inv Terms'][:2]}" if customer else "N/A",
                'ACCOUNT': account_value,
                'AMOUNT': sum(float(row.get(f'Chg {i} Total', 0)) for i in range(1, 9)),
                'REP': customer['Customer Salesrep'] if customer else "N/A",
            })

        # Define output folder path
        output_folder_path = latest_week_folder / '4.QB_Report/SHIPIUM (NOT DHL)'
        if not output_folder_path.exists():
            try:
                output_folder_path.mkdir(parents=True, exist_ok=True)
            except PermissionError:
                print(f"Permission denied: Unable to create directory at {output_folder_path}")
                exit(1)

        # Define output file name and path
        output_file_name = f"SHIPIUM EDI ROCKSOLID UPS FedEx INVOICES & INCOME_{date_range}.csv"
        output_file_path = output_folder_path / output_file_name

        # Export data to CSV
        with open(output_file_path, 'w', newline='', encoding='utf-8') as output_file:
            fieldnames = ['CUST NO', 'INV NO', 'CUSTOMER', 'MEMO INV ITEM', 'DATE', 'TERMS', 'ACCOUNT', 'AMOUNT', 'REP']
            writer = csv.DictWriter(output_file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(report_data)

        print(f"Filtered report exported to: {output_file_path}")
        print("Script completed successfully.")

    else:
        print("No raw file found in the latest week folder.")
else:
    print("No week folders found in the specified directory.")
