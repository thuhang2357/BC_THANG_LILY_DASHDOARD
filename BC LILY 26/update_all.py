#!/usr/bin/env python3
"""
update_all.py - Master script to update dashboard with new monthly data
Usage: python update_all.py

This script:
1. Reads all Excel files from monthly_reports/ folder
2. Processes data using process_data.py (extracts from BCKQHĐKD and Báo cáo cửa hàng)
3. Generates data.json
4. Generates expense_detail_summary.js and expense_detail.html
5. Copies dashboard.html to dashboard/ folder

For new months:
1. Add Excel file to monthly_reports/ with naming: "THÁNG XX.26 BÁO CÁO THÁNG XX.xlsx"
2. Run: python update_all.py
"""

import os
import sys
import json
import re
import shutil
sys.stdout.reconfigure(encoding='utf-8')

# Add scripts to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def find_excel_files():
    """Find all monthly report Excel files"""
    pattern = re.compile(r'THÁNG\s+(\d+)\.26\s+BÁO\s+CÁO\s+THÁNG\s+\d+\s*\.xlsx?', re.IGNORECASE)
    files = []
    search_path = os.path.join(os.path.dirname(__file__), 'monthly_reports')
    if os.path.exists(search_path):
        for f in os.listdir(search_path):
            if pattern.match(f):
                files.append(f)
    return sorted(files, key=lambda x: re.search(r'THÁNG\s+(\d+)', x, re.IGNORECASE).group(1))

def run_process_data():
    """Run process_data.py to generate data.json"""
    print("=" * 50)
    print("Step 1: Processing financial data...")
    print("=" * 50)
    os.system('python scripts/process_data.py')

def run_generate_expense_detail():
    """Run generate_expense_detail.py to generate expense detail files"""
    print("\n" + "=" * 50)
    print("Step 2: Generating expense detail...")
    print("=" * 50)
    os.system('python generate_expense_detail.py')

def copy_dashboard():
    """Copy dashboard.html to dashboard folder"""
    print("\n" + "=" * 50)
    print("Step 3: Copying dashboard files...")
    print("=" * 50)
    src = os.path.join(os.path.dirname(__file__), 'dashboard.html')
    dst = os.path.join(os.path.dirname(__file__), 'dashboard', 'dashboard.html')
    if os.path.exists(src):
        shutil.copy2(src, dst)
        print(f"Copied: {src} -> {dst}")

def main():
    print("\n" + "=" * 60)
    print("ULTD THỊNH PHÁT - DASHBOARD UPDATE SCRIPT")
    print("=" * 60)

    # Find Excel files
    files = find_excel_files()
    print(f"\nFound {len(files)} monthly report(s):")
    for f in files:
        print(f"  - {f}")

    # Run steps
    run_process_data()
    run_generate_expense_detail()
    copy_dashboard()

    print("\n" + "=" * 60)
    print("UPDATE COMPLETE!")
    print("=" * 60)
    print("\nNext steps:")
    print("1. Open dashboard/dashboard.html in browser")
    print("2. Verify data is correct")
    print("3. Deploy to server if needed")

if __name__ == '__main__':
    main()
