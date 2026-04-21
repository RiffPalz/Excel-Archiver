import os
import shutil
from datetime import datetime

# Set directories
ROOT_FOLDER = r"C:\CSI Reporting Tool\prod\output"
ARCHIVE_BASE_FOLDER = r"C:\CSI Reporting Tool\prod\Empty_Excel_Archive"

# Set limit to 5KB (5 * 1024 bytes)
SIZE_THRESHOLD_BYTES = 5 * 1024 

# Get CURRENT date for folder filtering and archive naming
today = datetime.now()

# Set up archive folder path
folder_name = today.strftime("%B %d %Y")
DAILY_ARCHIVE_FOLDER = os.path.join(ARCHIVE_BASE_FOLDER, folder_name)

# Initialize tracking
target_files = []
files_scanned = 0

print("-" * 50)
print(f"mHub Archiver: Targeting folders from {folder_name} (<= 5KB)")
print("-" * 50)

# Scan through output folder
for root, dirs, files in os.walk(ROOT_FOLDER):
    
    # Skip the archive folder to avoid loops
    if ARCHIVE_BASE_FOLDER in root:
        continue

    try:
        # 1. Check the date of the folder itself
        folder_timestamp = os.path.getmtime(root)
        folder_date = datetime.fromtimestamp(folder_timestamp)

        # 2. Proceed ONLY if the folder matches the current month and current date
        if (folder_date.year == today.year and 
            folder_date.month == today.month and 
            folder_date.day == today.day):
            
            # 3. Check the files inside this valid folder
            for file in files:
                if file.lower().endswith((".xls", ".xlsx")):
                    files_scanned += 1
                    file_path = os.path.join(root, file)

                    # 4. Filter by file size
                    if os.path.getsize(file_path) <= SIZE_THRESHOLD_BYTES:
                        target_files.append(file_path)

    except Exception as e:
        print(f"Error checking folder {root}: {e}")

# Print summary
print("\n[ SCAN COMPLETE ]")
print(f"Excel files scanned in today's folders: {files_scanned}")
print(f"Matches to extract and archive:         {len(target_files)}")
print("-" * 50)

if not target_files:
    print("No matching files found.")
    input("\nPress Enter to exit...")
    exit()

# Confirm and move files
if input(f"Extract {len(target_files)} files to '{folder_name}'? (yes/no): ").lower() == "yes":
    os.makedirs(DAILY_ARCHIVE_FOLDER, exist_ok=True)
    
    successful_moves = 0
    for f in target_files:
        try:
            # This extracts ONLY the file. The original folder is ignored and left behind.
            file_name = os.path.basename(f)
            destination = os.path.join(DAILY_ARCHIVE_FOLDER, file_name)
            
            # Add timestamp if file already exists in archive
            if os.path.exists(destination):
                base, ext = os.path.splitext(file_name)
                destination = os.path.join(DAILY_ARCHIVE_FOLDER, f"{base}_{datetime.now().strftime('%H%M%S')}{ext}")

            shutil.move(f, destination)
            successful_moves += 1
            
        except Exception as e:
            print(f"Failed to move {file_name}: {e}")
            
    print(f"\n[ SUCCESS ] Moved {successful_moves} files.")
else:
    print("\nOperation cancelled.")

input("\nPress Enter to exit...")