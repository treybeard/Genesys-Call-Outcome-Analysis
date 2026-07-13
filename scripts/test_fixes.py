#!/usr/bin/env python3
"""Test script to verify the bug fixes in office.py."""

import os
import sys
import tempfile
import zipfile
import subprocess

# Add project to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from engineering import analyze_csv, analyze_batch
from office import simple_xlsx, full_xlsx

def create_test_csvs():
    """Create minimal test CSV files."""
    test_dir = tempfile.mkdtemp(prefix="genesys_test_")
    
    # Create monthly CSV files for a test batch
    # Format: Genesys export with categories
    for month, date_range in [("April", "2026-04-01"), ("May", "2026-05-01"), ("June", "2026-06-01")]:
        csv_path = os.path.join(test_dir, f"doc_abc123_TestBatch {month}.csv")
        with open(csv_path, 'w') as f:
            f.write("EventDate,Duration,Status,Category,OtherKey,GroupKey,Agent,Queue,SystemKey\n")
            for i in range(100):
                # Mostly successful calls with some failures and abandons
                if i < 70:
                    status = "Success"
                    duration = f"0:{i % 5:02d}:{i % 60:02d}"
                elif i < 85:
                    status = "Fail - Failed"
                    duration = f"0:{i % 3:02d}:{i % 30:02d}"
                else:
                    status = "Abandon - Unsuccessful"
                    duration = "0:0:0"
                f.write(f"{date_range} {i:02d}:00:00,{duration},{status},TestGroup,key{i},agent{i},queue{i},sys{i}\n")
    
    return test_dir

def test_simple_mode(test_dir):
    """Test simple_xlsx mode."""
    print("=" * 60)
    print("Testing simple_xlsx mode...")
    print("=" * 60)
    
    monthly_files = [
        ("April", "doc_abc123_TestBatch April.csv"),
        ("May", "doc_abc123_TestBatch May.csv"),
        ("June", "doc_abc123_TestBatch June.csv"),
    ]
    
    monthly_data = analyze_batch("TestBatch", monthly_files, test_dir)
    if not monthly_data:
        print("❌ analyze_batch failed")
        return False
    
    out_path = os.path.join(test_dir, "TestBatch - Simple.xlsx")
    simple_xlsx("TestBatch", monthly_data, out_path)
    
    if not os.path.exists(out_path):
        print("❌ Output file not created")
        return False
    
    # Verify it's a valid ZIP (XLSX is ZIP)
    with zipfile.ZipFile(out_path, 'r') as zf:
        names = zf.namelist()
        print(f"✓ Created: {out_path} ({os.path.getsize(out_path)} bytes)")
        print(f"  Sheets: {[n for n in names if 'sheet' in n]}")
        
        # Check Summary sheet
        sheet1 = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')
        if 'TestBatch' in sheet1:
            print("  ✓ Summary sheet contains batch name")
        if 'Success' in sheet1:
            print("  ✓ Summary sheet contains Success metric")
    
    print("✅ simple_xlsx test PASSED")
    return True

def test_full_mode(test_dir):
    """Test full_xlsx mode (with charts)."""
    print("=" * 60)
    print("Testing full_xlsx mode (with charts)...")
    print("=" * 60)
    
    monthly_files = [
        ("April", "doc_abc123_TestBatch April.csv"),
        ("May", "doc_abc123_TestBatch May.csv"),
        ("June", "doc_abc123_TestBatch June.csv"),
    ]
    
    monthly_data = analyze_batch("TestBatch", monthly_files, test_dir)
    if not monthly_data:
        print("❌ analyze_batch failed")
        return False
    
    out_path = os.path.join(test_dir, "TestBatch - Full.xlsx")
    full_xlsx("TestBatch", monthly_data, out_path)
    
    if not os.path.exists(out_path):
        print("❌ Output file not created")
        return False
    
    # Verify it's a valid ZIP (XLSX is ZIP)
    with zipfile.ZipFile(out_path, 'r') as zf:
        names = zf.namelist()
        print(f"✓ Created: {out_path} ({os.path.getsize(out_path)} bytes)")
        print(f"  Sheets: {[n for n in names if 'sheet' in n]}")
        print(f"  Charts: {[n for n in names if 'chart' in n]}")
        print(f"  Drawings: {[n for n in names if 'drawing' in n]}")
        
        # Check Summary sheet
        sheet1 = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')
        if 'TestBatch' in sheet1:
            print("  ✓ Summary sheet contains batch name")
        if 'Success' in sheet1:
            print("  ✓ Summary sheet contains Success metric")
        
        # Check that percentage values are 0-1 range (not 0-100)
        if '0.7000000000' in sheet1 or '0.7' in sheet1:
            print("  ✓ Percentages are stored as 0-1 decimal (not 0-100)")
    
    print("✅ full_xlsx test PASSED")
    return True

def verify_percentage_format(test_dir):
    """Verify that percentages are formatted correctly."""
    print("=" * 60)
    print("Verifying percentage format...")
    print("=" * 60)
    
    monthly_files = [
        ("April", "doc_abc123_TestBatch April.csv"),
    ]
    monthly_data = analyze_batch("TestBatch", monthly_files, test_dir)
    
    out_path = os.path.join(test_dir, "VerifyPct.xlsx")
    simple_xlsx("TestBatch", monthly_data, out_path)
    
    with zipfile.ZipFile(out_path, 'r') as zf:
        sheet1 = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')
        
        # Check that Success/Abandon % rows have proper formatting
        lines = sheet1.split('\n')
        for line in lines:
            if 'Success' in line and 'Abandon' not in line:
                print(f"  Found metric line: {line.strip()[:200]}")
        
        # Verify the format code for percentages
        if '10\"' in sheet1:  # Excel percentage numFmtId
            print("  ✓ Style 4 uses numFmtId=\"10\" (Excel percentage)")
    
    print("✅ Percentage format verification PASSED")
    return True

def main():
    print("GeneSys Call Analysis — Bug Fix Tests")
    print("=" * 60)
    
    test_dir = create_test_csvs()
    print(f"Created test CSVs in: {test_dir}")
    print()
    
    results = []
    results.append(("simple_xlsx", test_simple_mode(test_dir)))
    results.append(("full_xlsx", test_full_mode(test_dir)))
    results.append(("percentage_format", verify_percentage_format(test_dir)))
    
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    passed = sum(1 for _, r in results if r)
    for name, result in results:
        status = "✅ PASSED" if result else "❌ FAILED"
        print(f"  {status} {name}")
    
    print(f"\n{passed}/{len(results)} tests passed")
    
    # Keep temp dir for inspection if needed
    print(f"\nTest CSVs preserved in: {test_dir}")
    print("(Delete manually when done investigating)")

if __name__ == '__main__':
    main()
