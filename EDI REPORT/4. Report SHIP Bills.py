import os
import re
import csv
import datetime
from pathlib import Path
import decimal # Import the decimal module for precision

# --- Configuration ---
# !! Please double-check these paths are correct for your system !!
customer_list_path = r"C:\Users\User\Desktop\JEAV\Customer List.csv"
raw_file_base_path = Path(r"C:\Users\User\Desktop\JEAV\EDI Reconcile (monday)")
# --- End Configuration ---

def get_latest_week_folder(base_path):
    """Finds the most recently modified 'Week_...' folder in the base path."""
    try:
        week_folders = [f for f in base_path.iterdir() if f.is_dir() and re.match(r'^Week_\d+_\(\d{2}\.\d{2}\.\d{2}\)_\(\d{2}\.\d{2}\.\d{2}\)$', f.name)]
        # Sort by modification time, newest first
        week_folders.sort(key=lambda f: f.stat().st_mtime, reverse=True)
        return week_folders[0] if week_folders else None
    except FileNotFoundError:
        print(f"Error: Base path not found: {base_path}")
        return None
    except Exception as e:
        print(f"Error finding latest week folder: {e}")
        return None

def convert_to_date(inv_no_suffix):
    """Converts the last 4 characters of an invoice number (format YMDD) to MM/DD/YYYY."""
    if not isinstance(inv_no_suffix, str) or len(inv_no_suffix) != 4:
         # Handle cases where inv_no is shorter than expected or not 'N/A'
        if inv_no_suffix == "N/A":
            return "N/A"
        else:
            # Decide how to handle unexpected input. Return error or a default?
            # print(f"Warning: Invalid input format for date conversion: '{inv_no_suffix}'. Expected 4 chars or 'N/A'.")
            return "Invalid Suffix" # Or return "N/A" or raise an error

    date_code = inv_no_suffix
    try:
        # Determine year based on character code relative to 'Y' (2024)
        # Ensure the character is uppercase for consistent ord() values
        year_char = date_code[0].upper()
        if 'A' <= year_char <= 'Z': # Basic check if it's a letter
             year = ord(year_char) - ord('Y') + 2024
        else:
             return "Invalid Year Char"

        month_char = date_code[1].upper()
        if 'A' <= month_char <= 'Z':
            month = ord(month_char) - ord('A') + 1
            # Adjust for skipped 'I' if your convention does that
            if month_char > 'I':
                month -= 1
        else:
             return "Invalid Month Char"

        day_str = date_code[2:]
        if day_str.isdigit():
             day = int(day_str)
        else:
             return "Invalid Day Chars"


        # Validate month and day before creating date object
        if not (1 <= month <= 12):
            # print(f"Debug: Invalid month calculated: {month} from char '{month_char}'")
            return "Invalid Month"
        # Basic day validation (doesn't account for month length/leap years perfectly but catches obvious errors)
        if not (1 <= day <= 31):
             # print(f"Debug: Invalid day calculated: {day} from chars '{day_str}'")
             return "Invalid Day"

        # Final check using datetime to catch impossible dates (like Feb 30)
        return datetime.date(year, month, day).strftime("%m/%d/%Y")

    except ValueError:
        # Catches errors from int() if day part isn't numeric, or datetime.date() for invalid dates (e.g., Feb 30)
        # print(f"Error converting date code '{date_code}': Invalid date components (e.g., Feb 30).")
        return "Invalid Date"
    except TypeError as e:
         # Catches errors likely from ord() if non-char processed
        # print(f"Error converting date code '{date_code}': Type error ({e}).")
        return "Invalid Date Type"
    except IndexError:
         # print(f"Error: Date code '{date_code}' has unexpected length.")
         return "Invalid Date Format"


def get_account(cust_no, carrier, sub_carrier):
    """Determines the account string based on carrier and sub-carrier."""
    carrier = carrier.strip()
    sub_carrier = sub_carrier.strip()
    if carrier == "FedEx":
        if sub_carrier == "England":
            return "FEDEX COST (ENGLAND LOGISTICS)"
        elif sub_carrier == "RSIS":
            return "FEDEX COST (DESCARTES)"
        # Default FedEx account if sub-carrier doesn't match specific cases or is empty
        return "FEDEX COST"
    elif carrier == "UPS":
        # UPS doesn't seem to use sub-carrier for account in the provided logic
        return "UPS COST"
    # Default for unknown carriers
    # print(f"Warning: Unknown carrier '{carrier}' for account determination.")
    return "UNKNOWN"

def get_class(sales_rep):
    """Extracts the class (first 5 characters) from the sales rep string."""
    return sales_rep[:5] if sales_rep and isinstance(sales_rep, str) else ""

# --- Main Script Logic ---
# Script run time: Thursday, April 3, 2025 at 2:22:56 PM EDT
print(f"Starting script... ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")

latest_week_folder = get_latest_week_folder(raw_file_base_path)

if latest_week_folder:
    print(f"Processing folder: {latest_week_folder.name}")
    match = re.search(r'\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)', latest_week_folder.name)
    if match:
        # Format date range for the output filename
        date_range = f"({match.group(1)})_({match.group(2)})"
        print(f"Date range extracted: {date_range}")
    else:
        print(f"Error: Invalid folder naming convention for date extraction in '{latest_week_folder.name}'. Expected 'Week_..._(MM.DD.YY)_(MM.DD.YY)'.")
        exit(1) # Exit if the date range can't be determined

    # Define the path to the raw data files within the week folder
    raw_files_path = latest_week_folder / "1.Gathering_Data/RS"
    latest_raw_file = None # Initialize
    try:
        # Find all potential raw files and sort by modification time to get the latest
        raw_files = sorted(raw_files_path.glob("RAW_EDI_RS_*.csv"), key=os.path.getmtime, reverse=True) # Get newest first
        if raw_files:
            latest_raw_file = raw_files[0]
            print(f"Selected latest raw file: {latest_raw_file.name}")
        else:
             print(f"Error: No raw files matching 'RAW_EDI_RS_*.csv' found in {raw_files_path}")
             exit(1) # Exit if no raw file found
    except FileNotFoundError:
         print(f"Error: Raw files directory not found: {raw_files_path}")
         exit(1)
    except Exception as e:
        print(f"Error accessing or sorting raw files in {raw_files_path}: {e}")
        exit(1) # Exit on other errors finding the file

    # --- Customer List Loading ---
    customer_lookup = {}
    try:
        print(f"Loading customer list from: {customer_list_path}")
        with open(customer_list_path, newline='', encoding='utf-8-sig') as customer_file:
            customer_list_reader = csv.DictReader(customer_file)
            # Clean header names immediately after creating the reader
            customer_list_reader.fieldnames = [name.strip() for name in customer_list_reader.fieldnames]

            # Check if 'Customer Id' column exists
            if 'Customer Id' not in customer_list_reader.fieldnames:
                 print(f"Error: Critical column 'Customer Id' not found in header of {customer_list_path}. Please check the file.")
                 exit(1)
            if 'Customer' not in customer_list_reader.fieldnames: # Check for customer name column too
                 print(f"Warning: 'Customer' column not found in header of {customer_list_path}. Customer names will be 'N/A'.")


            # Load data into lookup dictionary
            for i, c_row in enumerate(customer_list_reader):
                # Clean values and handle potential missing 'Customer Id'
                cleaned_row = {k: v.strip() if isinstance(v, str) else v for k, v in c_row.items()}
                cust_id = cleaned_row.get('Customer Id')
                if cust_id: # Only add if Customer Id exists and is not empty
                     customer_lookup[cust_id] = cleaned_row
                else:
                     print(f"Warning: Skipping row {i+2} in customer list due to missing or empty 'Customer Id': {cleaned_row}")
        print(f"Loaded {len(customer_lookup)} customers.")
        if not customer_lookup:
             print("Warning: Customer lookup is empty. No customer data will be matched.")

    except FileNotFoundError:
        print(f"Error: Customer list file not found at {customer_list_path}")
        exit(1)
    except Exception as e:
        print(f"Error reading customer list {customer_list_path}: {e}")
        exit(1)

    # --- Raw Data Processing ---
    raw_data = []
    raw_data_reader = None # Initialize
    try:
        print(f"Reading raw data from: {latest_raw_file.name}")
        with open(latest_raw_file, newline='', encoding='utf-8-sig') as raw_file:
             # Read raw data using DictReader
            raw_data_reader = csv.DictReader(raw_file)
            # Clean header names immediately
            raw_data_reader.fieldnames = [name.strip() for name in raw_data_reader.fieldnames]
            raw_data = list(raw_data_reader)
        print(f"Read {len(raw_data)} rows from raw file.")
    except FileNotFoundError:
         print(f"Error: Raw file seems to have disappeared: {latest_raw_file}")
         exit(1)
    except Exception as e:
        print(f"Error reading raw data file {latest_raw_file}: {e}")
        exit(1)

    # Check for essential columns in the raw data *after* reading headers
    if raw_data_reader:
        expected_raw_columns = ['Carrier', 'Customer #', 'Carrier Inv. #', 'Invoice Number', 'Sub Carrier', 'Service Type', 'Sales Rep', 'Carrier Cost Total']
        # Check for at least one airbill column variant
        airbill_cols_present = any(col in raw_data_reader.fieldnames for col in ['Airbill Number', 'Air Bill Number', 'AirBill'])

        missing_cols = [col for col in expected_raw_columns if col not in raw_data_reader.fieldnames]
        if missing_cols:
            print(f"Warning: Missing expected columns in {latest_raw_file.name}: {', '.join(missing_cols)}")
        if not airbill_cols_present:
             print(f"Warning: Missing all possible Airbill columns ('Airbill Number', 'Air Bill Number', 'AirBill') in {latest_raw_file.name}. Airbills will be empty.")
    else:
         print("Error: Could not read headers from raw file.")
         exit(1)


    # --- Data Transformation Loop ---
    report_data = []
    processed_rows = 0
    skipped_filtered = 0 # Count rows skipped by initial filter
    customers_not_found = set() # Track unique customer IDs not found

    print("Processing rows...")
    for row_index, row_raw in enumerate(raw_data):
        # Clean keys and string values in the current row
        row = {k: v.strip() if isinstance(v, str) else v for k, v in row_raw.items()}

        carrier = row.get('Carrier', '')
        cust_no = row.get('Customer #', '')
        carrier_inv_no = row.get('Carrier Inv. #', '')

        # Apply filtering conditions
        if carrier in ["FedEx", "UPS"] and carrier_inv_no and not carrier_inv_no.startswith('D') and cust_no not in ["10003217", "10003324"]:

            # Look up customer using the pre-loaded dictionary
            customer = customer_lookup.get(cust_no)

            if customer:
                # --- Process fields if customer is found ---
                inv_no_raw = row.get('Invoice Number', 'N/A')
                # Extract last 8 chars, then last 4 safely
                inv_no_suffix = str(inv_no_raw)[-8:][-4:] if inv_no_raw != 'N/A' and len(str(inv_no_raw)) >= 4 else 'N/A'
                date_converted = convert_to_date(inv_no_suffix)

                sub_carrier = row.get('Sub Carrier', '').strip()
                # Assign default sub_carrier if empty based on carrier
                if not sub_carrier and carrier == "FedEx":
                    sub_carrier = "RSIS" # Assign default if needed, adjust if necessary

                vendor = "DESCARTES" if sub_carrier == "RSIS" else "ENGLAND LOGISTICS" if sub_carrier == "England" else f"{carrier} ENGLAND" # Determine vendor

                # Get the raw airbill number string from potential columns
                airbill_number_raw = row.get('Airbill Number') or row.get('Air Bill Number') or row.get('AirBill') or ""
                airbill_number_raw = airbill_number_raw.strip()

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

                # Append processed data for the report
                report_data.append({
                    'DATE': date_converted,
                    'CUST NO': cust_no,
                    'CUSTOMER': customer.get('Customer', 'N/A'), # Use .get() for safety
                    'CLASS': get_class(row.get('Sales Rep')),
                    'BILL NO': carrier_inv_no,
                    'VENDOR': vendor,
                    'MEMO BILL ITEM': memo_bill_item,
                    'AMOUNT': row.get('Carrier Cost Total', '0'), # Default amount to '0' if missing
                    'ACCOUNT': get_account(cust_no, carrier, sub_carrier)
                })
                processed_rows += 1
                # --- End processing for found customer ---
            else:
                # Customer not found in the lookup dictionary
                if cust_no not in customers_not_found:
                     # Print only once per missing customer ID for cleaner logs
                     # print(f"Info: Customer not found in Customer List for CUST NO: {cust_no} (First occurrence: Row {row_index + 2})")
                     customers_not_found.add(cust_no)
                # This row is skipped because customer wasn't found
        else:
             # Row skipped due to initial filtering criteria
             skipped_filtered += 1

    print(f"Finished processing rows. Processed for report: {processed_rows}")
    if skipped_filtered > 0:
         print(f"Rows skipped by initial filter (Carrier, Inv#, Cust#): {skipped_filtered}")
    if customers_not_found:
         print(f"Rows skipped because Customer # was not found in list: {len(customers_not_found)} unique IDs ({len([r for r in raw_data if r.get('Customer #', '') in customers_not_found])} total occurrences)")
         # Optional: list the IDs if needed
         # print(f"Missing Customer IDs: {', '.join(sorted(list(customers_not_found)))}")


    # --- Sorting Report Data ---
    print("Sorting report data by BILL NO...")
    sorted_report = []
    if report_data: # Only sort if there's data
        try:
            # Attempt to sort numerically, converting BILL NO to int
            # Handle potential non-digit characters gracefully during sort key generation
            sorted_report = sorted(report_data, key=lambda x: int(x['BILL NO']) if x.get('BILL NO', '').isdigit() else float('inf'))
            # float('inf') places non-numeric or missing BILL NOs at the end
            print("Sorting successful (numeric preferred).")
        except KeyError:
             print("Error: 'BILL NO' key missing in report data during sorting. Report will be unsorted.")
             sorted_report = report_data # Keep unsorted if key is missing
        except Exception as e:
             print(f"Warning: Unexpected error during sorting ({e}). Report may be unsorted or partially sorted.")
             sorted_report = report_data # Fallback to unsorted on other errors
    else:
        print("No data processed for the report, skipping sorting and writing.")


    # --- Writing Output File ---
    if sorted_report: # Only write file if there is data
        output_folder_path = latest_week_folder / '4.QB_Report/SHIPIUM (NOT DHL)'
        try:
            print(f"Ensuring output directory exists: {output_folder_path}")
            output_folder_path.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"Error: Could not create output directory {output_folder_path}: {e}")
            exit(1) # Exit if output dir can't be created

        # Define the output file path using the extracted date range
        output_file_name = f"SHIPIUM EDI ROCKSOLID UPS FedEx BILLS Cost_{date_range}.csv"
        output_file_path = output_folder_path / output_file_name

        print(f"Writing report to: {output_file_path}")
        try:
            with open(output_file_path, 'w', newline='', encoding='utf-8-sig') as output_file:
                # Define the header columns for the output CSV
                fieldnames = ['DATE', 'CUST NO', 'CUSTOMER', 'CLASS', 'BILL NO', 'VENDOR', 'MEMO BILL ITEM', 'AMOUNT', 'ACCOUNT']
                writer = csv.DictWriter(output_file, fieldnames=fieldnames, extrasaction='ignore') # Ignore extra fields if any crept in
                writer.writeheader()
                writer.writerows(sorted_report)
            print(f"Success! Report successfully exported ({len(sorted_report)} rows).")
        except Exception as e:
            print(f"Error writing output file {output_file_path}: {e}")

    else:
        print("No data to write to the output file.")

else:
    # No 'Week_...' folders were found in the base path
    print(f"Error: No valid 'Week_...' folders found in {raw_file_base_path}")

print(f"Script finished. ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
