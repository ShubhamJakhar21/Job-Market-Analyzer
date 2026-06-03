"""
scraper.py  —  Scrapes job listings from Naukri.com
Uses Playwright to open the browser and BeautifulSoup to parse HTML.
"""

import time
import random
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup


def scrape_naukri(keyword: str = "python developer", max_pages: int = 5) -> list[dict]:
    """
    Scrapes Naukri.com for jobs matching the keyword.
    Returns a list of dicts: title, company, location, experience, skills, salary.
    """
    all_jobs = []

    print(f"\n🔍 Searching Naukri for: '{keyword}'")
    print(f"📄 Pages to scrape: {max_pages}")
    print("-" * 50)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,      # Set True to run in background (no browser window)
            slow_mo=500          # Slows down clicks so the site doesn't block us
        )
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        )
        page = context.new_page()

        for page_num in range(1, max_pages + 1):
            # Naukri URL format: keyword-jobs-1, keyword-jobs-2, etc.
            keyword_slug = keyword.strip().lower().replace(" ", "-")
            url = f"https://www.naukri.com/{keyword_slug}-jobs-{page_num}"

            print(f"📥 Scraping page {page_num}: {url}")

            try:
                page.goto(url, timeout=30000)
                # Wait for job cards to appear
                page.wait_for_selector(".srp-jobtuple-wrapper", timeout=15000)
                time.sleep(random.uniform(2, 4))  # Random wait to avoid being blocked

                html = page.content()
                soup = BeautifulSoup(html, "html.parser")

                # Each job listing is wrapped in this class
                job_cards = soup.find_all("div", class_="srp-jobtuple-wrapper")

                if not job_cards:
                    print(f"   ⚠️  No job cards found on page {page_num}. Stopping.")
                    break

                for card in job_cards:
                    job = parse_job_card(card)
                    if job:
                        all_jobs.append(job)

                print(f"   ✅ Found {len(job_cards)} jobs on page {page_num}")

            except Exception as e:
                print(f"   ❌ Error on page {page_num}: {e}")
                continue

        browser.close()

    print(f"\n✅ Total jobs scraped: {len(all_jobs)}")
    return all_jobs


def parse_job_card(card) -> dict | None:
    """Extracts info from a single Naukri job card."""
    try:
        # Job Title
        title_tag = card.find("a", class_="title")
        title = title_tag.get_text(strip=True) if title_tag else "N/A"

        # Company Name
        company_tag = card.find("a", class_="comp-name")
        if not company_tag:
            company_tag = card.find("a", class_="subTitle")
        company = company_tag.get_text(strip=True) if company_tag else "N/A"

        # Experience Required
        exp_tag = card.find("span", class_="expwdth")
        experience = exp_tag.get_text(strip=True) if exp_tag else "N/A"

        # Salary
        salary_tag = card.find("span", class_="sal")
        if not salary_tag:
            salary_tag = card.find("span", class_="salary")
        salary = salary_tag.get_text(strip=True) if salary_tag else "Not disclosed"

        # Location  (can have multiple cities)
        location_tags = card.find_all("span", class_="locWdth")
        if not location_tags:
            location_tags = card.find_all("li", class_="location")
        locations = [loc.get_text(strip=True) for loc in location_tags]
        location = ", ".join(locations) if locations else "N/A"

        # Skills  (shown as tag pills)
        skills_div = card.find("ul", class_="tags-gt")
        if not skills_div:
            skills_div = card.find("ul", class_="tags")
        skills_list = skills_div.find_all("li") if skills_div else []
        skills = ", ".join([s.get_text(strip=True) for s in skills_list]) if skills_list else "N/A"

        # Job description snippet
        desc_tag = card.find("span", class_="job-desc")
        description = desc_tag.get_text(strip=True) if desc_tag else "N/A"

        return {
            "title": title,
            "company": company,
            "location": location,
            "experience": experience,
            "salary": salary,
            "skills": skills,
            "description": description,
        }

    except Exception as e:
        print(f"   ⚠️  Could not parse a card: {e}")
        return None


if __name__ == "__main__":
    # Quick test — scrape 1 page
    jobs = scrape_naukri("python developer", max_pages=1)
    for j in jobs[:3]:
        print(j)
