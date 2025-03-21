import os
import re
import csv
import datetime
from pathlib import Path

customer_list_path = r"C:\Users\User\Desktop\JEAV\Customer List.csv"
raw_file_base_path = Path(r"C:\Users\User\Desktop\JEAV\EDI Reconcile (monday)")

def get_latest_week_folder(base_path):
    week_folders = [f for f in base_path.iterdir() if f.is_dir() and re.match(r'^Week_\d+_\(\d{2}\.\d{2}\.\d{2}\)_\(\d{2}\.\d{2}\.\d{2}\)$', f.name)]
    return max(week_folders, key=lambda f: f.stat().st_mtime) if week_folders else None

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

latest_week_folder = get_latest_week_folder(raw_file_base_path)

if latest_week_folder:
    match = re.search(r'\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)', latest_week_folder.name)
    if match:
        date_range = f"({match.group(1)})_({match.group(2)})"
    else:
        print("Invalid folder naming convention.")
        exit(1)

    raw_files = sorted((latest_week_folder / "1.Gathering_Data/RS").glob("RAW_EDI_RS_*.csv"), key=os.path.getmtime)
    latest_raw_file = raw_files[-1] if raw_files else None

    if latest_raw_file:
        with open(latest_raw_file, newline='', encoding='utf-8-sig') as raw_file:
            raw_data = list(csv.DictReader(raw_file))

        with open(customer_list_path, newline='', encoding='utf-8-sig') as customer_file:
            customer_list = list(csv.DictReader(customer_file))

        report_data = []
        for row in raw_data:
            row = {k.strip(): v.strip() if isinstance(v, str) else v for k, v in row.items()}
            carrier = row.get('Carrier', '')
            cust_no = row.get('Customer #', '')
            carrier_inv_no = row.get('Carrier Inv. #', '')

            if carrier in ["FedEx", "UPS"] and not carrier_inv_no.startswith('D') and cust_no not in ["10003217", "10003324"]:
                customer = next((c for c in customer_list if c['Customer Id'] == cust_no), None)

                if customer:
                    inv_no = row.get('Invoice Number', 'N/A')[-8:]
                    date_converted = convert_to_date(inv_no)
                    sub_carrier = row.get('Sub Carrier', '').strip() or ("RSIS" if carrier == "FedEx" else "")

                    vendor = "DESCARTES" if sub_carrier == "RSIS" else "ENGLAND LOGISTICS" if sub_carrier == "England" else f"{carrier} ENGLAND"
                    
                    airbill_number = row.get('Airbill Number') or row.get('Air Bill Number') or row.get('AirBill') or ""
                    airbill_number = airbill_number.strip()

                    memo_bill_item = f"{carrier} | AIRBILL# {airbill_number} | {row.get('Service Type', '')} | {sub_carrier}"

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

        sorted_report = sorted(report_data, key=lambda x: int(x['BILL NO']))

        output_folder_path = latest_week_folder / '4.QB_Report/SHIPIUM (NOT DHL)'
        output_folder_path.mkdir(parents=True, exist_ok=True)

        output_file_path = output_folder_path / f"SHIPIUM EDI ROCKSOLID UPS FedEx BILLS Cost_{date_range}.csv"

        with open(output_file_path, 'w', newline='', encoding='utf-8-sig') as output_file:
            writer = csv.DictWriter(output_file, fieldnames=['DATE', 'CUST NO', 'CUSTOMER', 'CLASS', 'BILL NO', 'VENDOR', 'MEMO BILL ITEM', 'AMOUNT', 'ACCOUNT'])
            writer.writeheader()
            writer.writerows(sorted_report)

        print(f"Report successfully exported to {output_file_path}")
    else:
        print("No raw files found.")
else:
    print("No week folders found.")
