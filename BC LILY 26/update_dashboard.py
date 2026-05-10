#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Update dashboard - Run this after adding new monthly files
Usage: python update_dashboard.py
"""

import subprocess
import sys
import os
import io

# Fix encoding for Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def main():
    print("=" * 50)
    print("CAP NHAT DASHBOARD")
    print("=" * 50)

    # Step 1: Run ETL to process data
    print("\n[1/3] Dang xu ly du lieu tu Excel...")
    try:
        subprocess.run([sys.executable, "scripts/process_data.py"], check=True, encoding='utf-8', errors='replace')
        print("      [OK] Hoan thanh xu ly du lieu")
    except subprocess.CalledProcessError:
        print("      [ERROR] Loi khi xu ly du lieu!")
        return

    # Step 2: Export to Excel
    print("\n[2/3] Dang xuat Excel...")
    try:
        subprocess.run([sys.executable, "export_excel.py"], check=True, encoding='utf-8', errors='replace')
        print("      [OK] Hoan thanh xuat Excel")
    except subprocess.CalledProcessError:
        print("      [ERROR] Loi khi xuat Excel!")

    # Step 3: Regenerate detail pages
    print("\n[3/3] Dang tao trang chi tiet...")
    try:
        from generate_detail_pages import generate_all_detail_pages
        generate_all_detail_pages()
        print("      [OK] Hoan thanh tao trang chi tiet")
    except Exception as e:
        print(f"      [!] Canh bao: {e}")

    print("\n" + "=" * 50)
    print("HOAN THANH! Mo dashboard/dashboard.html de xem.")
    print("=" * 50)

if __name__ == "__main__":
    main()
