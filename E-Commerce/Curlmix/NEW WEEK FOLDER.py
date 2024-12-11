import os
import re
import shutil
from datetime import datetime, timedelta

# Define the base path where the week folders are located
base_path = r"C:\Users\User\Desktop\JEAV\E-Commerce Invoice (tuesday)\Curl Mix"

# Define the format of the week folder name
folder_name_pattern = "Week_{0}_({1})_({2})"

# Function to add days to a date string
def add_days_to_date(date_str, days):
    date = datetime.strptime(date_str, "%m.%d.%y")
    new_date = date + timedelta(days=days)
    return new_date.strftime("%m.%d.%y")

# Function to update date within file names
def update_date_in_file_name(file_name, old_date_format_start, old_date_format_end, new_date_start, new_date_end):
    updated_file_name = file_name.replace(old_date_format_start, new_date_start)
    updated_file_name = updated_file_name.replace(old_date_format_end, new_date_end)
    return updated_file_name

# Get the list of week folders in the base path
week_folders = [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f)) and re.match(r"^Week_(\d+)_\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)$", f)]

# Initialize variables for the new week number and date range
new_week_number = 1
new_start_date = datetime.now()
new_end_date = new_start_date + timedelta(days=6)

# Initialize variable for the latest folder to copy from
latest_folder_path = ""

# Find the highest week number and get the latest folder to copy from
for folder in week_folders:
    match = re.match(r"^Week_(\d+)_\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)$", folder)
    if match:
        week_number = int(match.group(1))
        start_date = datetime.strptime(match.group(2), "%m.%d.%y")
        end_date = datetime.strptime(match.group(3), "%m.%d.%y")
        if week_number >= new_week_number:
            new_week_number = week_number + 1
            # Calculate new start date as the day after the previous end date
            new_start_date = end_date + timedelta(days=1)
            # Set new end date as 6 days after the new start date
            new_end_date = new_start_date + timedelta(days=6)
            latest_folder_path = os.path.join(base_path, folder)

# Format the new week number and date range
new_start_date_string = new_start_date.strftime("%m.%d.%y")
new_end_date_string = new_end_date.strftime("%m.%d.%y")

# Construct the new folder name
new_folder_name = folder_name_pattern.format(new_week_number, new_start_date_string, new_end_date_string)

# Define the full path for the new folder
new_folder_path = os.path.join(base_path, new_folder_name)

# Create the new folder
os.makedirs(new_folder_path, exist_ok=True)

# Copy and rename contents from the latest folder to the new folder
if latest_folder_path:
    for root, dirs, files in os.walk(latest_folder_path):
        # Create corresponding directory structure in the new folder
        for dir_name in dirs:
            dest_dir = os.path.join(new_folder_path, os.path.relpath(os.path.join(root, dir_name), latest_folder_path))
            os.makedirs(dest_dir, exist_ok=True)

        for file_name in files:
            source_file_path = os.path.join(root, file_name)
            dest_file_path = os.path.join(new_folder_path, os.path.relpath(root, latest_folder_path), file_name)

            # Extract old start and end dates from the file name
            old_start_end_match = re.search(r"\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)", file_name)
            if old_start_end_match:
                old_start_date = old_start_end_match.group(1)
                old_end_date = old_start_end_match.group(2)
                new_file_name = update_date_in_file_name(file_name, old_start_date, old_end_date, new_start_date_string, new_end_date_string)
            else:
                new_file_name = file_name
            
            new_file_path = os.path.join(os.path.dirname(dest_file_path), new_file_name)

            # Copy and rename file
            shutil.copy2(source_file_path, new_file_path)

# Output results
print(f"Created new folder and copied contents with updated names: {new_folder_path}")

# Log the creation
log_file = r"C:\Users\User\Desktop\JEAV\E-Commerce Invoice (tuesday)\Curl Mix\logfile.txt"
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
with open(log_file, "a") as log:
    log.write(f"{timestamp} - Created new folder and copied contents with updated names: {new_folder_path}\n")
