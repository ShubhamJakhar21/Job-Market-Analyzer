"""
test_with_sample.py  —  Tests the full analysis pipeline using sample_jobs.csv
Run this FIRST to make sure everything is installed and working correctly.
No internet connection or scraping needed.

  python test_with_sample.py
"""

import shutil
import os

print("=" * 55)
print("  TESTING WITH SAMPLE DATA (no scraping needed)")
print("=" * 55)

# Copy sample data as the main jobs.csv
shutil.copy("sample_jobs.csv", "jobs.csv")
print("✅ Copied sample_jobs.csv → jobs.csv")

# Run the full analysis
from analyzer import run_analysis
run_analysis(csv_path="jobs.csv", output_dir="output")

print("\n🎉 Test complete! Check the 'output' folder.")
print("   If you see charts and an Excel file, everything is working!\n")
