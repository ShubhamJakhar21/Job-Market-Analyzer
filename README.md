# 📊 Job Market Analyzer

Scrapes job listings from Naukri.com, analyzes in-demand skills,
and generates charts + a formatted Excel report.

---

## 🗂️ Project Structure

```
job_market_analyzer/
├── main.py               ← Run this to scrape + analyze
├── test_with_sample.py   ← Run this first to test (no scraping needed)
├── scraper.py            ← Playwright scraping logic
├── analyzer.py           ← Data cleaning, charts, Excel report
├── sample_jobs.csv       ← 50 sample jobs for testing
├── requirements.txt      ← Python libraries needed
└── output/               ← Created automatically
    ├── chart_skills.png
    ├── chart_locations.png
    ├── chart_companies.png
    └── job_market_report.xlsx
```

---

## 🚀 Setup (do this once)

### Step 1 — Open terminal in VS Code
Press  `Ctrl + `` ` (backtick) to open the terminal.

### Step 2 — Install libraries
```bash
pip install -r requirements.txt
playwright install chromium
```

### Step 3 — Test without scraping (recommended first step)
```bash
python test_with_sample.py
```
If you see charts and an Excel file in the `output/` folder → everything is working ✅

---

## ▶️ Running the Project

### Option A — Scrape real jobs + analyze
```bash
python main.py
```

### Option B — Just re-run analysis on existing data
Open `main.py` and set:
```python
SKIP_SCRAPING = True
```
Then run:
```bash
python main.py
```

---

## ⚙️ Customization

Open `main.py` and change these settings at the top:

| Setting | Default | Description |
|---------|---------|-------------|
| `SEARCH_KEYWORD` | `"python developer"` | Job title to search for |
| `MAX_PAGES` | `3` | Pages to scrape (1 page ≈ 20 jobs) |
| `SKIP_SCRAPING` | `False` | Skip scraping, use existing CSV |

---

## 📤 Output Files

| File | What it contains |
|------|-----------------|
| `jobs.csv` | Raw scraped data |
| `output/chart_skills.png` | Top 20 skills bar chart |
| `output/chart_locations.png` | Top hiring cities bar chart |
| `output/chart_companies.png` | Top companies bar chart |
| `output/job_market_report.xlsx` | Full Excel report (4 sheets) |

---

## ❗ Troubleshooting

**"No job cards found"**
→ Naukri may have updated their HTML. Open `scraper.py` and check
  the class names like `srp-jobtuple-wrapper`. Use browser DevTools
  (right-click → Inspect) to find the current class names.

**Charts look empty / errors in analyzer**
→ Run `python test_with_sample.py` first. If that works, the scraper
  is the issue, not the analysis code.

**Playwright not found**
→ Run: `pip install playwright` then `playwright install chromium`

---

## 📝 Resume Line

> "Built a job market analytics tool using Python, Playwright, and Pandas
> that scraped 200+ job listings from Naukri.com, extracted skill demand
> data, and produced automated Excel reports with trend charts."
