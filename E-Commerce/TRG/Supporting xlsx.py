import os
import pandas as pd
from datetime import datetime

# Set output path
output_path = r'C:\Users\User\Desktop'

def find_latest_folder(base_path):
    subdirs = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    try:
        return sorted(subdirs, key=lambda x: datetime.strptime(x, '%y.%m.%d'), reverse=True)[0]
    except (ValueError, IndexError):
        return None

def find_latest_excel_file(directory):
    excel_files = [f for f in os.listdir(directory) 
                  if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    return excel_files and os.path.join(directory, max(excel_files, 
        key=lambda f: os.path.getmtime(os.path.join(directory, f))))

def clean_data(df):
    text_cols = ['Product Name', 'Sender Company', 'Destination City', 
                 'Destination State', 'Destination Country']
    numeric_cols = ['Actual Weight', 'Length', 'Width', 'Height', 
                   'Dimensional Weight', 'Chargeable Weight', 'Billed Weight',
                   'Postage', 'Fuel Fee', 'Reship Fee', 'Remote Area Surcharge',
                   'Return Fee', 'Additional Charge for Special Products',
                   'Other Charges', 'Total']
    
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()
    
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[^\d.\-]', '', regex=True), 
                                  errors='coerce')
    
    date_cols = ['Invoice date', 'Arrival Time']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    df = df.fillna({
        'YT Reference': 'Unknown',
        'Order Number': 'Unknown',
        'Last Mile Tracking Number': 'Unknown',
        'Customer ID': 'Unknown'
    })
    
    return df.drop_duplicates()

def calculate_dimensional_weight(row):
    try:
        return (row['Length'] * row['Width'] * row['Height']) / 166
    except:
        return 0

def calculate_chargeable_weight(row):
    dim_weight = row['Dims Weight']
    actual_weight = row['Actual Weight']
    return max(dim_weight, actual_weight) if pd.notnull(dim_weight) and pd.notnull(actual_weight) else 0

def determine_weight_unit(weight):
    return 'OZ' if weight < 16 else 'LB'

def create_report(df_filtered, report_type):
    if df_filtered.empty:
        print(f"No data found for '{report_type}'")
        return
    
    df_report = pd.DataFrame({
        'Count': range(1, len(df_filtered) + 1),
        'Invoice Number': 'TRG10028',
        'Invoice Date': datetime.today().strftime('%Y-%m-%d'),
        'YT Reference': df_filtered['YT Reference'],
        'Order Number': df_filtered['Order Number'],
        'Last Mile Tracking Number': df_filtered['Last Mile Tracking Number'],
        'Shipper Name': df_filtered['Customer/Shipper Name'],
        'Ref. 1': df_filtered['Reference 1'],
        'Ref. 2': df_filtered['Reference 2'],
        'Customer ID': df_filtered['Customer ID'],
        'Arrival Time': df_filtered['Arrival Time'],
        'City': df_filtered['Destination City'],
        'State': df_filtered['Destination State'],
        'Zipcode': df_filtered['Destination Zipcode'],
        'Country': df_filtered['Destination Country'],
        'Product Name': df_filtered['Product Name'],
        'L': df_filtered['Length'],
        'W': df_filtered['Width'],
        'H': df_filtered['Height']
    })
    
    df_report['Dims Weight'] = df_filtered.apply(calculate_dimensional_weight, axis=1)
    df_report['Actual Weight'] = df_filtered['Actual Weight']
    df_report['Chargeable Weight'] = df_report.apply(calculate_chargeable_weight, axis=1)
    df_report['Weight'] = df_report['Chargeable Weight'].round(2)
    df_report['Round'] = df_report['Weight'].round()
    df_report['OZ/LB'] = df_report['Round'].apply(determine_weight_unit)
    df_report['Suggested Sell'] = df_filtered['Postage']
    df_report['ZONE'] = df_report.apply(
        lambda x: 'Remote' if x['Suggested Sell'] >= 9 or x['Suggested Sell'] <= 0 
        else str(int(x['Suggested Sell'])), axis=1)
    
    filename = f"TRG_Report_CIRRO_{report_type}_{datetime.today().strftime('%Y%m%d')}.csv"
    output_file = os.path.join(output_path, filename)
    df_report.to_csv(output_file, index=False)
    print(f"Report saved to: {output_file}")

def main():
    base_path = r'C:\Users\User\Desktop\JEAV\E-Commerce Invoice (tuesday)\TRG\TRG Report'
    
    latest_folder = find_latest_folder(base_path)
    if not latest_folder:
        raise ValueError("No valid folders found")
    
    gathering_path = os.path.join(base_path, latest_folder, '1.Gathering')
    raw_file_path = find_latest_excel_file(gathering_path)
    if not raw_file_path:
        raise ValueError("No Excel files found")
    
    try:
        df = pd.read_excel(raw_file_path)
        df_cleaned = clean_data(df)
        
        # Process WEST data
        df_west = df_cleaned[df_cleaned['Product Name'].str.contains('CIRRO ECONOMY WEST', 
                                                                    case=False, na=False)]
        create_report(df_west, 'WEST')
        
        # Process EAST data
        df_east = df_cleaned[df_cleaned['Product Name'].str.contains('CIRRO ECONOMY EAST', 
                                                                    case=False, na=False)]
        create_report(df_east, 'EAST')
        
    except Exception as e:
        print(f"Error processing file: {e}")
        raise

if __name__ == "__main__":
    main()
