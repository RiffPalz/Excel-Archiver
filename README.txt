======================================
mHub Excel Report Archiver
======================================

DESCRIPTION
This tool automatically scans the mHub production output folder to identify and archive empty Excel reports (files 5KB or smaller) before they are distributed to merchants. It targets files generated on the previous day and safely moves them into an organized daily archive folder rather than permanently deleting them.

HOW IT WORKS
1. Scans the specified output directory for `.xls` and `.xlsx` files.
2. Checks the "Date Modified" timestamp. It strictly targets files from exactly **yesterday**.
3. Checks the file size. If the file is 5KB or smaller, it is flagged for archiving.
4. Moves the flagged files into a specific, date-stamped folder: `Empty_Excel_Archive\[Month] [Day] [Year]` (e.g., March 24 2026).

HOW TO RUN
- If running the raw Python script: Open your terminal/command prompt and run `python excel_archiver.py`
- If compiled as a Windows Executable: Simply double-click `excel_archiver.exe`
- The tool will display a clean summary of how many files were scanned and flagged. It will pause and ask for a final "yes/no" confirmation before moving any files.
- Once finished, the window will stay open until you press "Enter" so you can read the final success or error messages.

SAFETY FEATURES
- Non-Destructive: Files are safely moved, never permanently deleted.
- Duplicate Protection: If a file with the exact same name already exists in the archive, the tool appends a unique timestamp to the incoming file so no historical data is accidentally overwritten.
- Archive Exclusion: The tool automatically ignores the archive folder during its scan so it doesn't re-process old files.

CURRENT CONFIGURATION PATHS
By default, the script looks at the following production directories:
- Target Folder: C:\CSI Reporting Tool\prod\output
- Archive Destination: C:\CSI Reporting Tool\prod\Empty_Excel_Archive

To change these target folders in the future, open `excel_archiver.py` in any standard text editor (like Notepad) and update the `ROOT_FOLDER` and `ARCHIVE_BASE_FOLDER` variables at the very top of the script.