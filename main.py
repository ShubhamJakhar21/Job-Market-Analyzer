"""
main.py  —  Run this file to start the Job Market Analyzer

  python main.py

You can change the settings below before running.
"""

import os
import pandas as pd
from scraper import scrape_naukri
from analyzer import run_analysis


# ════════════════════════════════════════════════════
#   ⚙️  SETTINGS  —  Change these as you like
# ════════════════════════════════════════════════════

SEARCH_KEYWORD = "python developer"   # What job to search for
MAX_PAGES      = 3                    # How many pages to scrape (1 page ≈ 20 jobs)
                                      # Start with 3, increase to 10-25 for more data
CSV_FILE       = "jobs.csv"           # Where to save raw scraped data
OUTPUT_FOLDER  = "output"             # Folder for charts + Excel report

# Set this to True if you already have jobs.csv and just want to re-run the analysis
SKIP_SCRAPING  = False

# ════════════════════════════════════════════════════


def main():
    print("\n" + "=" * 55)
    print("   JOB MARKET ANALYZER")
    print(f"   Keyword : {SEARCH_KEYWORD}")
    print(f"   Pages   : {MAX_PAGES}")
    print("=" * 55)

    # ── STEP 1: Scrape ──────────────────────────────────
    if SKIP_SCRAPING and os.path.exists(CSV_FILE):
        print(f"\n⏭️  Skipping scraping. Using existing '{CSV_FILE}'")
    else:
        print("\n🚀 Starting scraper…")
        jobs = scrape_naukri(SEARCH_KEYWORD, max_pages=MAX_PAGES)

        if not jobs:
            print("\n❌ No jobs scraped. Possible reasons:")
            print("   • Naukri changed its HTML structure (update class names in scraper.py)")
            print("   • No results found for the keyword")
            print("   • Website blocked the request (try adding more sleep time)")
            print("\n💡 TIP: Run with SKIP_SCRAPING = True and a sample jobs.csv to test analysis.")
            return

        # Save to CSV
        df_raw = pd.DataFrame(jobs)
        df_raw.to_csv(CSV_FILE, index=False)
        print(f"\n💾 Saved {len(df_raw)} jobs to '{CSV_FILE}'")

    # ── STEP 2: Analyze ─────────────────────────────────
    run_analysis(
        csv_path    = CSV_FILE,
        output_dir  = OUTPUT_FOLDER,
        top_n_skills= 20,
    )

    print(f"\n📂 Open the '{OUTPUT_FOLDER}' folder to see:")
    print("   📊 chart_skills.png      — top 20 skills bar chart")
    print("   📊 chart_locations.png   — top hiring cities bar chart")
    print("   📊 chart_companies.png   — top companies bar chart")
    print("   📝 job_market_report.xlsx — full Excel report with 4 sheets")


if __name__ == "__main__":
    main()
