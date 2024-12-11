import os
import re
import shutil
from datetime import datetime, timedelta
import calendar

# Define the base path where the week folders are located
base_path = r"C:\Users\User\Desktop\JEAV\EDI Reconcile (Monday)"

# Define the format of the week folder name
folder_name_pattern = "Week_{0}_({1})_({2})"

# Function to update date within file names
def update_date_in_filename(filename, old_start_date_str, old_end_date_str, new_start_date_str, new_end_date_str):
    # Escape dots for regex
    old_start_date_pattern = re.escape(old_start_date_str)
    old_end_date_pattern = re.escape(old_end_date_str)
    # Replace old dates with new dates in the filename
    updated_filename = re.sub(old_start_date_pattern, new_start_date_str, filename)
    updated_filename = re.sub(old_end_date_pattern, new_end_date_str, updated_filename)
    return updated_filename

# Get the list of week folders in the base path
pattern = re.compile(r"^Week_(\d+)_\((\d{2}\.\d{2}\.\d{2})\)_\((\d{2}\.\d{2}\.\d{2})\)$")

week_folders = []

for entry in os.listdir(base_path):
    full_path = os.path.join(base_path, entry)
    if os.path.isdir(full_path):
        m = pattern.match(entry)
        if m:
            week_number = int(m.group(1))
            start_date_str = m.group(2)
            end_date_str = m.group(3)
            week_folders.append({
                'path': full_path,
                'name': entry,
                'week_number': week_number,
                'start_date_str': start_date_str,
                'end_date_str': end_date_str
            })

# Initialize variables for the new week number and date range
max_week_number = 0
latest_end_date = None
latest_folder = None

# Find the folder with the highest week number and latest end date
for folder in week_folders:
    week_number = folder['week_number']
    end_date_str = folder['end_date_str']
    end_date = datetime.strptime(end_date_str, "%m.%d.%y")
    if week_number > max_week_number or (week_number == max_week_number and end_date > latest_end_date):
        max_week_number = week_number
        latest_end_date = end_date
        latest_folder = folder

# Calculate new start and end dates
if latest_folder:
    new_week_number = max_week_number + 1
    new_start_date = latest_end_date + timedelta(days=1)
    new_end_date = new_start_date + timedelta(days=6)
    latest_folder_path = latest_folder['path']
    latest_folder_start_date_str = latest_folder['start_date_str']
    latest_folder_end_date_str = latest_folder['end_date_str']
else:
    # No existing week folders found; start from current date
    new_week_number = 1
    new_start_date = datetime.now()
    new_end_date = new_start_date + timedelta(days=6)
    latest_folder_path = None
    latest_folder_start_date_str = ''
    latest_folder_end_date_str = ''

# Format the new week number and date range
new_start_date_str = new_start_date.strftime("%m.%d.%y")
new_end_date_str = new_end_date.strftime("%m.%d.%y")

# Construct the new folder name
new_folder_name = folder_name_pattern.format(new_week_number, new_start_date_str, new_end_date_str)

# Define the full path for the new folder
new_folder_path = os.path.join(base_path, new_folder_name)

# Create the new folder
os.makedirs(new_folder_path, exist_ok=True)

# Copy and rename contents from the latest folder to the new folder
if latest_folder_path:
    for dirpath, dirnames, filenames in os.walk(latest_folder_path):
        # Compute relative path from latest_folder_path
        rel_path = os.path.relpath(dirpath, latest_folder_path)
        # Compute destination directory path
        dest_dir = os.path.join(new_folder_path, rel_path)
        # Create destination directory
        os.makedirs(dest_dir, exist_ok=True)
        # Copy files
        for filename in filenames:
            src_file = os.path.join(dirpath, filename)
            # Update dates in the filename
            new_filename = update_date_in_filename(
                filename,
                latest_folder_start_date_str,
                latest_folder_end_date_str,
                new_start_date_str,
                new_end_date_str
            )
            dest_file = os.path.join(dest_dir, new_filename)
            # Copy the file
            shutil.copy2(src_file, dest_file)

print(f"Created new folder and copied contents with updated names: {new_folder_path}")

# Log the creation
log_file = r"C:\Users\User\Desktop\JEAV\EDI Reconcile (Monday)\logfile.txt"  # Change this path if necessary
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
with open(log_file, 'a') as f:
    f.write(f"{timestamp} - Created new folder and copied contents with updated names: {new_folder_path}\n")
