import os
import re
import csv
import datetime
from pathlib import Path

# Define constant paths
customer_list_path = r"C:\Users\User\Desktop\JEAV\Customer List.csv"
raw_file_base_path = Path(r"C:\Users\User\Desktop\JEAV\EDI Reconcile (monday)")

# Function to get the most recent week folder
def get_latest_week_folder(base_path):
    pattern = re.compile(r'^Week_\d+_\(\d{2}\.\d{2}\.\d{2}\)_\(\d{2}\.\d{2}\.\d{2}\)$')
    folders = [entry for entry in base_path.iterdir() if entry.is_dir() and pattern.match(entry.name)]
    return max(folders, key=lambda x: x.stat().st_mtime, default=None)

# Corrected conversion logic for INV NO to date
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

latest_week_folder = get_latest_week_folder(raw_file_base_path)

if latest_week_folder:
    date_match = re.search(r'\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)', latest_week_folder.name)
    date_range = f"({date_match.group(1)})_({date_match.group(2)})" if date_match else "Unknown_DateRange"

    raw_files = sorted(latest_week_folder.glob("1.Gathering_Data/RS/RAW_EDI_RS_*.csv"), key=lambda f: f.stat().st_mtime, reverse=True)
    latest_raw_file = raw_files[0] if raw_files else None

    if latest_raw_file:
        print(f"Using raw file: {latest_raw_file}")

        with open(latest_raw_file, newline='', encoding='utf-8-sig') as raw_file:
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
            ship_date = row['Ship Date'] or "N/A"

            carrier = row['Carrier']
            sub_carrier = row['Sub Carrier'].upper()
            if carrier == "UPS":
                account = "UPS SALES"
            elif carrier == "FREIGHT":
                account = "FREIGHT & OTHER"
            elif carrier == "FedEx":
                account = "FEDEX SALES (ENGLAND LOGISTICS)" if sub_carrier == "ENGLAND" else "FEDEX SALES (DESCARTE)"
            else:
                account = "OTHER SALES"

            memo_inv_item = f"FREIGHT (LTL) | AIRBILL# {row['Airbill Number']} | DATE {row['Ship Date']}" if carrier == "FREIGHT" else f"{carrier} | AIRBILL# {row['Airbill Number']} | DATE {row['Ship Date']}"

            amount = sum(float(row.get(col, 0) or 0) for col in [
                'Customer Base', 'Chg 1 Total', 'Chg 2 Total', 'Chg 3 Total',
                'Chg 4 Total', 'Chg 5 Total', 'Chg 6 Total', 'Chg 7 Total', 'Chg 8 Total'
            ])

            report_data.append({
                'CUST NO': row['Customer #'],
                'INV NO': inv_no,
                'CUSTOMER': customer['Customer'] if customer else "N/A",
                'MEMO INV ITEM': memo_inv_item,
                'DATE': convert_to_date(inv_no),
                'TERMS': f"NET {customer['Inv Terms'][:2]}" if customer else "N/A",
                'ACCOUNT': account,
                'AMOUNT': amount,
                'REP': customer['Customer Salesrep'] if customer else "N/A",
            })

        output_folder_path = latest_week_folder / '4.QB_Report' / 'SHIPIUM (NOT DHL)'
        output_folder_path.mkdir(parents=True, exist_ok=True)

        output_file_name = f"SHIPIUM EDI ROCKSOLID UPS FedEx INVOICES & INCOME_{date_range}.csv"
        output_file_path = output_folder_path / output_file_name

        with open(output_file_path, 'w', newline='', encoding='utf-8-sig') as output_file:
            writer = csv.DictWriter(output_file, fieldnames=['CUST NO', 'INV NO', 'CUSTOMER', 'MEMO INV ITEM', 'DATE', 'TERMS', 'ACCOUNT', 'AMOUNT', 'REP'])
            writer.writeheader()
            writer.writerows(report_data)

        print(f"Filtered report exported to: {output_file_path}")
        print("Script completed successfully.")
    else:
        print("No raw file found in the latest week folder.")
else:
    print("No week folders found in the specified directory.")
