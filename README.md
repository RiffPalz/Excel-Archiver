# 📊 Excel Report Archiver

> Automatically archives empty Excel reports from production folders into organized date-based directories.

---

## 📌 Description
The **Excel Report Archiver** is a utility that automatically scans a production output folder to identify and archive empty Excel reports.

It specifically targets Excel files that are:
- **5KB or smaller**
- Generated **exactly one day prior (yesterday)**

Instead of deleting these files, the tool safely moves them into a structured archive directory for future reference.

---

## ⚙️ How It Works
1. Scans the specified output directory for `.xls` and `.xlsx` files  
2. Filters files based on:
   - 📅 **Date Modified** → Must be **yesterday**
   - 📦 **File Size** → Must be **≤ 5KB**
3. Flags matching files for archiving  
4. Moves them into a date-based archive folder

---

## 🛡️ Safety Features

### 🔒 Non-Destructive
Files are **moved**, not deleted.

### 🧠 Duplicate Protection
If a file already exists in the archive:
- A unique timestamp is appended to prevent overwriting

### 🚫 Archive Exclusion
The archive directory is automatically ignored during scanning to prevent re-processing.

---

## 📎 Notes
- Designed for internal use
- Safe to run on a daily basis
- No files are permanently deleted