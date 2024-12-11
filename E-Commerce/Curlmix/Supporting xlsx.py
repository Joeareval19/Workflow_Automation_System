import os
import pandas as pd
import re

# Function to find the latest week folder based on a specific pattern
def find_latest_week_folder(base_path):
    folders = [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))]
    week_folders = [f for f in folders if re.match(r'^Week_\d+_\(\d{2}\.\d{2}\.\d{2}\)_\(\d{2}\.\d{2}\.\d{2}\)$', f)]
    
    if not week_folders:
        raise FileNotFoundError("No week folders found in the specified directory.")
    
    latest_folder = max(week_folders, key=lambda f: os.path.getmtime(os.path.join(base_path, f)))
    return os.path.join(base_path, latest_folder)

# Function to find the latest raw file in the specified subdirectory
def find_latest_raw_file(folder_path):
    raw_folder_path = os.path.join(folder_path, "1.Gathering")
    if not os.path.exists(raw_folder_path):
        raise FileNotFoundError("The '1.Gathering' folder does not exist in the specified path.")

    raw_files = [f for f in os.listdir(raw_folder_path) if f.startswith("RAW_Curlmix_") and f.endswith(".xlsx")]
    
    if not raw_files:
        raise FileNotFoundError("No raw files found in the '1.Gathering' folder.")
    
    latest_raw_file = max(raw_files, key=lambda f: os.path.getmtime(os.path.join(raw_folder_path, f)))
    return os.path.join(raw_folder_path, latest_raw_file)

# Hardcoded GPM % table
gpm_lookup = {
    1: 0.11, 2: 0.11, 3: 0.11, 4: 0.11, 5: 0.11, 6: 0.11, 7: 0.11, 8: 0.16
}

# Hardcoded rate tables for ounces and pounds
rate_lookup_oz = {
    1: {1: 3.47, 2: 3.47, 3: 3.49, 4: 3.5, 5: 3.62, 6: 3.69, 7: 3.77, 8: 3.78},
    2: {1: 3.47, 2: 3.47, 3: 3.49, 4: 3.5, 5: 3.62, 6: 3.69, 7: 3.77, 8: 3.78},
    3: {1: 3.48, 2: 3.48, 3: 3.5, 4: 3.53, 5: 3.64, 6: 3.72, 7: 3.79, 8: 3.8},
    4: {1: 3.5, 2: 3.5, 3: 3.54, 4: 3.57, 5: 3.72, 6: 3.79, 7: 3.89, 8: 3.9},
    5: {1: 3.58, 2: 3.58, 3: 3.61, 4: 3.64, 5: 3.81, 6: 3.9, 7: 4.02, 8: 4.05},
    6: {1: 3.63, 2: 3.63, 3: 3.67, 4: 3.71, 5: 3.9, 6: 3.99, 7: 4.11, 8: 4.13},
    7: {1: 3.63, 2: 3.63, 3: 3.69, 4: 3.73, 5: 3.93, 6: 4.05, 7: 4.15, 8: 4.19},
    8: {1: 3.66, 2: 3.66, 3: 3.72, 4: 3.77, 5: 4.02, 6: 4.11, 7: 4.25, 8: 4.28},
    9: {1: 4.03, 2: 4.03, 3: 4.09, 4: 4.13, 5: 4.39, 6: 4.5, 7: 4.64, 8: 4.68},
    10: {1: 4.17, 2: 4.17, 3: 4.22, 4: 4.28, 5: 4.54, 6: 4.67, 7: 4.82, 8: 4.85},
    11: {1: 4.17, 2: 4.17, 3: 4.23, 4: 4.29, 5: 4.57, 6: 4.69, 7: 4.84, 8: 4.89},
    12: {1: 4.18, 2: 4.18, 3: 4.24, 4: 4.29, 5: 4.58, 6: 4.71, 7: 4.86, 8: 4.91},
    13: {1: 4.63, 2: 4.63, 3: 4.7, 4: 4.75, 5: 5.08, 6: 5.2, 7: 5.37, 8: 5.42},
    14: {1: 4.65, 2: 4.65, 3: 4.71, 4: 4.78, 5: 5.1, 6: 5.23, 7: 5.4, 8: 5.45},
    15: {1: 4.67, 2: 4.67, 3: 4.72, 4: 4.79, 5: 5.12, 6: 5.26, 7: 5.44, 8: 5.5},
    15.99: {1: 4.68, 2: 4.68, 3: 4.73, 4: 4.82, 5: 5.15, 6: 5.29, 7: 5.49, 8: 5.53},
    # ... (Add remaining data as per your original code)
}

rate_lookup_lb = {
   1: {1: 5.76, 2: 5.76, 3: 5.84, 4: 5.89, 5: 6.18, 6: 6.33, 7: 6.49, 8: 6.63},
    2: {1: 6.99, 2: 6.99, 3: 7.1, 4: 7.19, 5: 7.7, 6: 7.88, 7: 8.14, 8: 8.34},
    3: {1: 7.18, 2: 7.18, 3: 7.3, 4: 7.43, 5: 8.02, 6: 8.25, 7: 8.55, 8: 8.78},
    4: {1: 7.23, 2: 7.23, 3: 7.34, 4: 7.5, 5: 8.28, 6: 8.55, 7: 8.95, 8: 9.2},
    5: {1: 7.58, 2: 7.58, 3: 7.84, 4: 8.09, 5: 9.21, 6: 9.6, 7: 10.12, 8: 10.46},
    6: {1: 8.32, 2: 8.32, 3: 8.85, 4: 9.36, 5: 11.48, 6: 12.14, 7: 13.11, 8: 14.34},
    7: {1: 8.55, 2: 8.55, 3: 9.08, 4: 9.61, 5: 11.74, 6: 12.43, 7: 13.4, 8: 14.67},
    8: {1: 8.81, 2: 8.81, 3: 9.35, 4: 9.88, 5: 12.06, 6: 12.74, 7: 13.74, 8: 15.04},
    9: {1: 9.11, 2: 9.11, 3: 9.65, 4: 10.19, 5: 12.41, 6: 13.1, 7: 14.12, 8: 15.45},
    10: {1: 9.36, 2: 9.36, 3: 9.9, 4: 10.46, 5: 12.7, 6: 13.4, 7: 14.44, 8: 15.8},
    11: {1: 10.61, 2: 10.61, 3: 11.45, 4: 12.28, 5: 15.71, 6: 16.75, 7: 18.3, 8: 20.15},
    12: {1: 10.71, 2: 10.71, 3: 11.55, 4: 12.4, 5: 15.81, 6: 16.86, 7: 18.41, 8: 20.27},
    13: {1: 10.91, 2: 10.91, 3: 11.74, 4: 12.58, 5: 16.0, 6: 17.06, 7: 18.61, 8: 20.47},
    14: {1: 11.05, 2: 11.05, 3: 11.9, 4: 12.73, 5: 16.14, 6: 17.21, 7: 18.77, 8: 20.65},
    15: {1: 11.2, 2: 11.2, 3: 12.04, 4: 12.88, 5: 16.31, 6: 17.35, 7: 18.93, 8: 20.83},
    # ... (Add remaining data as per your original code)
}

# Configuration for file paths
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
base_raw_data_path = r"C:\Users\User\Desktop\JEAV\E-Commerce Invoice (tuesday)\Curl Mix"
report_paths = {
    "second_report": os.path.join(desktop_path, "Sup Curlmix.xlsx")
}

# Load the latest raw data
try:
    latest_week_folder = find_latest_week_folder(base_raw_data_path)
    print(f"Latest week folder found: {latest_week_folder}")
    
    raw_file_path = find_latest_raw_file(latest_week_folder)
    print(f"Using raw file: {raw_file_path}")
    
    raw_df = pd.read_excel(raw_file_path)
    print("Raw data loaded successfully.")
except Exception as e:
    print(f"Error: {e}")
    raise

# Calculate "Un-Rounded" column
raw_df['Un-Rounded'] = raw_df.apply(
    lambda row: row['Billed Weight'] if row['Actual Weight'] > 0.999999 else row['Actual Weight'] * 16,
    axis=1
)

# Calculate "LB/OZ" column based on Actual Weight
raw_df['LB/OZ'] = raw_df['Actual Weight'].apply(lambda x: 'LB' if x > 0.9999 else 'OZ')

# Calculate "Rounded Billed Weight" column based on specified conditions
raw_df['Rounded Billed Weight'] = raw_df.apply(
    lambda row: row['Billed Weight'] if row['Billed Service'] == "Parcel"
    else (15.99 if 15.001 <= row['Un-Rounded'] < 15.99 
          else (row['Billed Weight'] if row['Actual Weight'] > 0.9999 
                else row['Billed Weight'] * 16)),
    axis=1
)

# Calculate "LB/Oz" based on Billed Service and Actual Weight
raw_df['LB/Oz'] = raw_df.apply(
    lambda row: "" if pd.isna(row['OSM BOL'])
    else ("LB" if row['Billed Service'] == "Parcel" 
          else ("LB" if row['Actual Weight'] > 0.9999 else "OZ")),
    axis=1
)

# Function to calculate "Package Charge" based on specified logic
def calculate_total_charge(row):
    if row['LB/Oz'] == 'LB' and row['Rounded Billed Weight'] > 15:
        gpm_value = gpm_lookup.get(row['Zone'], None)
        if gpm_value is not None:
            return row['Billed Weight'] / (1 - gpm_value)
    if row['Billed Service'] == "Non Qualifying Under 1lb":
        return row['Total Charge']  
    if row['LB/Oz'] == 'OZ':
        weight_row = rate_lookup_oz.get(row['Rounded Billed Weight'], {})
        return weight_row.get(row['Zone'], None)
    else:
        weight_row = rate_lookup_lb.get(row['Rounded Billed Weight'], {})
        return weight_row.get(row['Zone'], None)

# Apply the function to calculate Package Charge
raw_df['Package Charge'] = raw_df.apply(calculate_total_charge, axis=1)

# Function to calculate Fuel Charge
def calculate_fuel_charge(row):
    if row['Billed Service'] in ["Non Qualifying Over 1lb", "Non Qualifying Under 1lb"]:
        return 0
    return round(row['Package Charge'] * 0.0975, 2)

# Apply Fuel Charge calculation
raw_df['Fuel Charge'] = raw_df.apply(calculate_fuel_charge, axis=1)

# Add a new column "DIM Rules Applied" with a constant value of 0 for all rows
raw_df['DIM Rules Applied'] = 0

# Calculate the "NET CHARGE" column as the sum of all specified charges
raw_df['Net Charge'] = raw_df[[ 
    'Package Charge', 'Fuel Charge', 'Delivery Confirmation Charge', 
    'OSM DIM Surcharge', 'Relabel Fee', 'Delivery Area Surcharge (DAS)', 
    'OCR Fee', 'Peak Season Surcharge', 'Nonstandard Length Fee 22 in', 
    'Nonstandard Length Fee 30 in', 'Nonstandard Length Fee 2 cu', 
    'Dimension Noncompliance', 'Irregular Shape Charge', 
    'Non Compliance/Unmanifested Charge', 'Signature Confirmation'
]].sum(axis=1)

# Calculate the "Plus/Minus" column as the difference between "Total Charge" and "NET CHARGE"
raw_df['Plus/Minus'] = raw_df['Net Charge'] - raw_df['Total Charge']

# Function to generate the first report with the specified output path and filename
def generate_first_report(dataframe, week_folder):
    output_path = os.path.join(week_folder, "2.Cleaning", "Curlmix_Review_data.xlsx")
    output_columns = [
        'OSM BOL', 'Package Id', 'Tracking ID', 'Processed Date/Time', 
        'Actual Weight', 'Billed Weight', 'Un-Rounded', 'LB/OZ', 
        'Rounded Billed Weight', 'LB/Oz', 'Billed Service', 'Zip', 'Zone', 
        'Package Charge', 'Fuel Charge', 'Delivery Confirmation Charge', 
        'OSM DIM Surcharge', 'Relabel Fee', 'Delivery Area Surcharge (DAS)',
        'OCR Fee', 'Peak Season Surcharge', 'Nonstandard Length Fee 22 in', 
        'Nonstandard Length Fee 30 in', 'Nonstandard Length Fee 2 cu', 
        'Dimension Noncompliance', 'Irregular Shape Charge', 
        'Non Compliance/Unmanifested Charge', 'Signature Confirmation',
        'DIM Rules Applied', 'Height', 'Length', 'Width', 'Customer Reference 1', 
        'Net Charge', 'Total Charge', 'Plus/Minus'
    ]
    report_df = dataframe[output_columns]

    # Calculate the sums for specified columns
    sum_row = {col: report_df[col].sum() for col in [
        'Package Charge', 'Fuel Charge', 'Delivery Confirmation Charge', 'OSM DIM Surcharge', 
        'Relabel Fee', 'Delivery Area Surcharge (DAS)', 'OCR Fee', 'Peak Season Surcharge', 
        'Nonstandard Length Fee 22 in', 'Nonstandard Length Fee 30 in', 'Nonstandard Length Fee 2 cu', 
        'Dimension Noncompliance', 'Irregular Shape Charge', 'Non Compliance/Unmanifested Charge', 
        'Signature Confirmation', 'Net Charge'
    ]}
    
    # Create a blank row for separation and then the sum row
    blank_row = pd.DataFrame({col: '' for col in report_df.columns}, index=[0])
    sum_row_df = pd.DataFrame(sum_row, index=[0])
    report_df = pd.concat([report_df, blank_row, sum_row_df], ignore_index=True)

    try:
        report_df.to_excel(output_path, index=False)
        print(f"First report successfully saved to {output_path}")
    except Exception as e:
        print(f"Error saving first report: {e}")

# Function to generate the second report with dynamic path and filename
def generate_second_report(dataframe, week_folder):
    # Increment the invoice number from 10165 to 10166
    invoice_number = 10165 + 1
    # Extract dates from the latest week folder
    week_dates = re.search(r'\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)', week_folder)
    
    if not week_dates:
        raise ValueError("Unable to extract date range from the latest week folder.")
    
    start_date = week_dates.group(1)  # 11.03.24
    end_date = week_dates.group(2)     # 11.09.24

    # Create the output filename
    output_filename = f"CM_Inv_{invoice_number}_Week_({start_date})_({end_date}).xlsx"
    output_path = os.path.join(week_folder, "3.Invoice", output_filename)

    dataframe['Fuel Charge'] = dataframe.apply(
        lambda row: row['Fuel Charge'] + abs(row['Plus/Minus']) if row['Plus/Minus'] < 0 else row['Fuel Charge'],
        axis=1
    )
    dataframe['Net Charge'] = dataframe[[ 
        'Package Charge', 'Fuel Charge', 'Delivery Confirmation Charge', 
        'OSM DIM Surcharge', 'Relabel Fee', 'Delivery Area Surcharge (DAS)', 
        'OCR Fee', 'Peak Season Surcharge', 'Nonstandard Length Fee 22 in', 
        'Nonstandard Length Fee 30 in', 'Nonstandard Length Fee 2 cu', 
        'Dimension Noncompliance', 'Irregular Shape Charge', 
        'Non Compliance/Unmanifested Charge', 'Signature Confirmation'
    ]].sum(axis=1)

    sup_report_columns = [
        'OSM BOL', 'Package Id', 'Tracking ID', 'Processed Date/Time', 
        'Actual Weight', 'Billed Weight', 'Billed Service', 'Zip', 'Zone', 
        'Package Charge', 'Fuel Charge', 'Delivery Confirmation Charge', 
        'OSM DIM Surcharge', 'Relabel Fee', 'Delivery Area Surcharge (DAS)', 
        'OCR Fee', 'Peak Season Surcharge', 'Nonstandard Length Fee 22 in', 
        'Nonstandard Length Fee 30 in', 'Nonstandard Length Fee 2 cu', 
        'Dimension Noncompliance', 'Irregular Shape Charge', 
        'Non Compliance/Unmanifested Charge', 'Signature Confirmation', 
        'DIM Rules Applied', 'Height', 'Length', 'Width', 
        'Customer Reference 1', 'Net Charge'
    ]
    sup_report_df = dataframe[sup_report_columns]

    # Calculate the sums for specified columns
    sum_row = {col: sup_report_df[col].sum() for col in [
        'Package Charge', 'Fuel Charge', 'Delivery Confirmation Charge', 'OSM DIM Surcharge', 
        'Relabel Fee', 'Delivery Area Surcharge (DAS)', 'OCR Fee', 'Peak Season Surcharge', 
        'Nonstandard Length Fee 22 in', 'Nonstandard Length Fee 30 in', 'Nonstandard Length Fee 2 cu', 
        'Dimension Noncompliance', 'Irregular Shape Charge', 'Non Compliance/Unmanifested Charge', 
        'Signature Confirmation', 'Net Charge'
    ]}

    # Create a blank row for separation and then the sum row
    blank_row = pd.DataFrame({col: '' for col in sup_report_df.columns}, index=[0])
    sum_row_df = pd.DataFrame(sum_row, index=[0])
    sup_report_df = pd.concat([sup_report_df, blank_row, sum_row_df], ignore_index=True)

    try:
        sup_report_df.to_excel(output_path, index=False)
        print(f"Second report successfully saved to {output_path}")
    except Exception as e:
        print(f"Error saving second report: {e}")

# Main function to generate both reports
def main():
    # Generate the first report dynamically
    generate_first_report(raw_df, latest_week_folder)
    
    # Generate the second report dynamically
    generate_second_report(raw_df, latest_week_folder)

if __name__ == "__main__":
    main()
