import os
import csv
from pathlib import Path
from datetime import datetime
from decimal import Decimal
import shutil
import re

# ========== Dynamic Path Configuration =========
BASE_PATH = Path(r"C:\Users\User\Desktop\JEAV")
RS_VS_QB_PATH = BASE_PATH / "Weekly Payments Reconcile (daily)" / "RS vs QB"
REPORT_OUTPUT_PATH = RS_VS_QB_PATH / "1.Report"
HISTORY_FILE_PATH = RS_VS_QB_PATH / "1.Report" / "History Data.csv"

# Regex pattern for week folders (same as reference)
WEEK_FOLDER_PATTERN = re.compile(r"Week_\d+_\(\d{2}\.\d{2}\.\d{2}\)_\(\d{2}\.\d{2}\.\d{2}\)")


# ========== Helper/Utility Functions =========
def find_latest_week_folder(base_dir: Path):
    """Find the most recent week folder based on creation time."""
    valid_folders = []
    for item in base_dir.iterdir():
        if item.is_dir() and WEEK_FOLDER_PATTERN.match(item.name):
            valid_folders.append(item)
    if not valid_folders:
        raise FileNotFoundError("No valid week folder found in the specified directory.")
    valid_folders.sort(key=lambda f: f.stat().st_ctime, reverse=True)
    return valid_folders[0]

def extract_dates_from_folder_name(folder_name: str):
    """Extract start and end dates from the folder name."""
    date_pattern = re.compile(r"\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)")
    match_obj = date_pattern.search(folder_name)
    if not match_obj:
        raise ValueError("Invalid folder name format, cannot extract dates.")
    return match_obj.groups()

def ensure_directory_exists(path: Path):
    """Create directory if it doesn't exist."""
    path.mkdir(parents=True, exist_ok=True)

def load_csv_as_list_of_dicts(csv_path: Path):
    """Load a CSV file into a list of dictionaries."""
    rows = []
    with csv_path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows

def calculate_total_amount(data: list, amount_field: str = "AMOUNT"):
    """Calculate the sum of amounts in the data."""
    return sum(Decimal(row[amount_field].replace("$", "").replace(",", "")) for row in data)


# ========== Main Logic =========
def main():
    try:
        # Step 1: Find the latest week folder
        latest_week_folder = find_latest_week_folder(RS_VS_QB_PATH)
        folder_name = latest_week_folder.name
        start_date, end_date = extract_dates_from_folder_name(folder_name)
        import_path = latest_week_folder / "2.Import TransactionPro"

        # Step 2: Define source file paths
        ils_source_path = import_path / f"ILS_Payment_Report_({start_date})_({end_date}).csv"
        ship_source_path = import_path / f"SHIP_Payment_Report_({start_date})_({end_date}).csv"

        # Verify source files exist
        if not ils_source_path.exists():
            raise FileNotFoundError(f"ILS file not found: {ils_source_path}")
        if not ship_source_path.exists():
            raise FileNotFoundError(f"SHIP file not found: {ship_source_path}")

        # Step 3: Define destination file paths and ensure directory exists
        ensure_directory_exists(REPORT_OUTPUT_PATH)
        ils_destination_path = REPORT_OUTPUT_PATH / "ILS_Payment_Report.csv"
        ship_destination_path = REPORT_OUTPUT_PATH / "SHIP_Payment_Report.csv"

        # Step 4: Copy files to the report directory
        shutil.copy(ils_source_path, ils_destination_path)
        shutil.copy(ship_source_path, ship_destination_path)
        print(f"Copied ILS file to: {ils_destination_path}")
        print(f"Copied SHIP file to: {ship_destination_path}")

        # Step 5: Load data from source files for history calculation
        ils_data = load_csv_as_list_of_dicts(ils_source_path)
        ship_data = load_csv_as_list_of_dicts(ship_source_path)

        # Step 6: Calculate totals
        ils_total = calculate_total_amount(ils_data) if ils_data else Decimal("0.00")
        ship_total = calculate_total_amount(ship_data) if ship_data else Decimal("0.00")
        total_amount = ils_total + ship_total

        # Step 7: Prepare history data row with today's date (March 04, 2025)
        today_date = datetime.now().strftime("%Y-%m-%d")  # Format: YYYY-MM-DD
        history_row = {
            "Date": today_date,
            "Total Amount": f"{total_amount:.2f}",
            "ILS": f"{ils_total:.2f}",
            "SHIP": f"{ship_total:.2f}"
        }

        # Step 8: Append to History Data.csv (create if it doesn't exist)
        history_fieldnames = ["Date", "Total Amount", "ILS", "SHIP"]
        history_file_exists = HISTORY_FILE_PATH.exists()

        with HISTORY_FILE_PATH.open("a", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=history_fieldnames)
            if not history_file_exists:
                writer.writeheader()  # Write header if file is new
            writer.writerow(history_row)

        print(f"Appended to history file: {HISTORY_FILE_PATH}")
        print(f"Date: {today_date}")
        print(f"Total Amount: ${total_amount:.2f}")
        print(f"ILS Total: ${ils_total:.2f}")
        print(f"SHIP Total: ${ship_total:.2f}")

    except FileNotFoundError as e:
        print(f"Error: {str(e)}")
    except Exception as e:
        print(f"An unexpected error occurred: {str(e)}")


if __name__ == "__main__":
    main()
