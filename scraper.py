import os
import re
import json
import time
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

BASE_DIR = Path(__file__).parent.resolve()

def find_excel_file():
    env_file = os.getenv("EXCEL_FILE")
    if env_file and (BASE_DIR / env_file).exists():
        return BASE_DIR / env_file
    for name in ["Outlet name and Link.xlsx", "Outlet name and Link - Copy.xlsx"]:
        path = BASE_DIR / name
        if path.exists():
            return path
    raise FileNotFoundError("Could not find any target Excel file in the directory.")

def clean_number(text):
    if not text:
        return None
    cleaned = re.sub(r"[^\d]", "", str(text))
    return int(cleaned) if cleaned else None

def clean_rating(text):
    if not text:
        return None
    m = re.search(r"(\d\.\d)", str(text))
    return float(m.group(1)) if m else None

def extract_rating_and_reviews(page):
    """
    Extracts rating and review count using multiple hierarchical strategies:
    1. Google Search Knowledge Panel selectors
    2. Google Maps selectors
    3. Aria-labels
    4. Text regex matching
    """
    rating = None
    reviews = None

    # Wait briefly for dynamic elements
    try:
        page.wait_for_load_state("domcontentloaded", timeout=10000)
    except Exception:
        pass

    # Strategy 1: Search Knowledge Panel rating and review count
    try:
        # Search for knowledge panel rating (e.g. span with 4.4 next to stars)
        kp_rating_el = page.query_selector("span.aqN8e, span.FZ1T5, span.hGSR34, span.fontDisplayLarge")
        if kp_rating_el:
            r_val = clean_rating(kp_rating_el.inner_text())
            if r_val:
                rating = r_val

        # Search for knowledge panel reviews (e.g. "4,685 Google reviews")
        kp_reviews_el = page.query_selector("a:has-text('Google reviews'), span:has-text('Google reviews'), a:has-text('reviews')")
        if kp_reviews_el:
            rev_val = clean_number(kp_reviews_el.inner_text())
            if rev_val:
                reviews = rev_val
    except Exception:
        pass

    # Strategy 2: Google Maps panel selectors
    if rating is None or reviews is None:
        try:
            # Maps rating is often inside div.F7nice or fontHeadlineLarge
            maps_rating = page.query_selector("div.F7nice span[aria-hidden='true'], span.ceNzKf")
            if maps_rating and not rating:
                rating = clean_rating(maps_rating.inner_text())

            maps_reviews = page.query_selector("div.F7nice span[aria-label*='reviews'], div.F7nice button:has-text('reviews')")
            if maps_reviews and not reviews:
                reviews = clean_number(maps_reviews.inner_text() or maps_reviews.get_attribute("aria-label"))
        except Exception:
            pass

    # Strategy 3: Check aria-labels across the document
    if rating is None or reviews is None:
        try:
            star_elements = page.query_selector_all("[aria-label*='star'], [aria-label*='out of 5']")
            for el in star_elements:
                label = el.get_attribute("aria-label") or ""
                if not rating:
                    r_match = re.search(r"(\d\.\d)\s*(?:stars?|out of 5)", label, re.I)
                    if r_match:
                        rating = float(r_match.group(1))
                if not reviews:
                    rev_match = re.search(r"([\d,]+)\s*reviews?", label, re.I)
                    if rev_match:
                        reviews = clean_number(rev_match.group(1))
                if rating and reviews:
                    break
        except Exception:
            pass

    # Strategy 4: Fallback to Page Body Text Regex
    if rating is None or reviews is None:
        try:
            body_text = page.inner_text("body")
            
            # Review count pattern (e.g. '4,685 Google reviews' or '224 reviews')
            if not reviews:
                m_rev = re.search(r"([\d,]+)\s*(?:Google\s+)?reviews", body_text, re.I)
                if m_rev:
                    reviews = clean_number(m_rev.group(1))
                else:
                    m_paren = re.search(r"\(\s*([\d,]+)\s*\)\s*(?:reviews?)?", body_text, re.I)
                    if m_paren:
                        reviews = clean_number(m_paren.group(1))

            # Rating pattern (e.g. '4.4 ★' or 'Rated 4.4 out of 5')
            if not rating:
                m_rate = re.search(r"(\d\.\d)\s*(?:★|stars?|out of 5)", body_text, re.I)
                if m_rate:
                    rating = float(m_rate.group(1))
        except Exception:
            pass

    return rating, reviews

def scrape_store(page, store_code, brand, address, link):
    print(f"\n[Scraping] {store_code} - {brand} ({address})")
    print(f"  Target URL: {link}")

    rating = None
    reviews = None
    error_msg = None

    # Step 1: Attempt loading the direct link
    try:
        response = page.goto(link, timeout=35000, wait_until="domcontentloaded")
        page.wait_for_timeout(3000)
        rating, reviews = extract_rating_and_reviews(page)
    except Exception as e:
        error_msg = str(e)
        print(f"  [Direct link warning] {e}")

    # Step 2: Fallback to Google Search if rating or reviews were not captured
    if rating is None or reviews is None:
        search_query = f"{brand} {address}".replace(",", " ").strip()
        search_url = f"https://www.google.com/search?q={search_query.replace(' ', '+')}"
        print(f"  [Fallback Search] Querying: {search_url}")
        try:
            page.goto(search_url, timeout=35000, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)
            f_rating, f_reviews = extract_rating_and_reviews(page)
            if f_rating and not rating:
                rating = f_rating
            if f_reviews and not reviews:
                reviews = f_reviews
        except Exception as e:
            print(f"  [Fallback warning] {e}")

    print(f"  -> Result: Rating = {rating} | Reviews = {reviews}")
    return {
        "store_code": store_code,
        "brand": brand,
        "address": address,
        "link": link,
        "rating": rating,
        "reviews": reviews,
        "status": "success" if (rating is not None or reviews is not None) else "failed",
        "error": error_msg
    }

def update_excel_workbook(file_path, scraped_data):
    print(f"\n[Excel] Updating workbook: {file_path}")
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    today_dt = datetime.now()
    today_date_str = today_dt.strftime("%d/%m/%Y")
    
    # Check existing header dates in row 1
    # Store mapping of store_code to row number
    store_row_map = {}
    for r in range(3, ws.max_row + 1):
        code_val = ws.cell(row=r, column=2).value
        if code_val:
            store_row_map[str(code_val).strip()] = r

    # Determine columns for today
    # Look across row 1 to see if today's date already exists
    target_rating_col = None
    target_review_col = None

    for c in range(6, ws.max_column + 1):
        val = ws.cell(row=1, column=c).value
        if val:
            if isinstance(val, datetime):
                cell_date_str = val.strftime("%d/%m/%Y")
            else:
                cell_date_str = str(val).strip()
            if cell_date_str == today_date_str:
                target_rating_col = c
                target_review_col = c + 1
                break

    # If today's column does not exist yet, append at the end
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3")
    )

    if target_rating_col is None:
        # Find next empty column
        target_rating_col = ws.max_column + 1
        target_review_col = target_rating_col + 1

        print(f"  Adding new date columns: {target_rating_col} & {target_review_col} for {today_date_str}")
        
        # Merge row 1 for the date header
        ws.merge_cells(
            start_row=1, start_column=target_rating_col,
            end_row=1, end_column=target_review_col
        )
        date_cell = ws.cell(row=1, column=target_rating_col)
        date_cell.value = today_date_str
        date_cell.font = Font(name="Calibri", size=11, bold=True)
        date_cell.alignment = Alignment(horizontal="center", vertical="center")
        date_cell.fill = PatternFill(fill_type="solid", fgColor="FFD9D2E9")  # Lavender
        date_cell.border = thin_border
        ws.cell(row=1, column=target_review_col).border = thin_border

        # Subheaders in row 2
        rate_hdr = ws.cell(row=2, column=target_rating_col)
        rate_hdr.value = "GMB Live Rating"
        rate_hdr.font = Font(name="Calibri", size=10, bold=True)
        rate_hdr.alignment = Alignment(horizontal="center", vertical="center")
        rate_hdr.fill = PatternFill(fill_type="solid", fgColor="FFDD7E6B")  # Coral
        rate_hdr.border = thin_border

        rev_hdr = ws.cell(row=2, column=target_review_col)
        rev_hdr.value = "Review Count"
        rev_hdr.font = Font(name="Calibri", size=10, bold=True)
        rev_hdr.alignment = Alignment(horizontal="center", vertical="center")
        rev_hdr.fill = PatternFill(fill_type="solid", fgColor="FFFFF2CC")  # Yellow
        rev_hdr.border = thin_border
    else:
        print(f"  Updating existing date columns: {target_rating_col} & {target_review_col} for {today_date_str}")

    # Write data rows
    for item in scraped_data:
        code = item["store_code"]
        row_num = store_row_map.get(code)
        if not row_num:
            continue

        rating_cell = ws.cell(row=row_num, column=target_rating_col)
        if item["rating"] is not None:
            rating_cell.value = item["rating"]
        rating_cell.font = Font(name="Calibri", size=11, bold=True)
        rating_cell.alignment = Alignment(horizontal="center", vertical="center")
        rating_cell.fill = PatternFill(fill_type="solid", fgColor="FFD0E0E3")  # Light cyan
        rating_cell.border = thin_border

        rev_cell = ws.cell(row=row_num, column=target_review_col)
        if item["reviews"] is not None:
            rev_cell.value = item["reviews"]
        rev_cell.font = Font(name="Calibri", size=11, bold=False)
        rev_cell.alignment = Alignment(horizontal="center", vertical="center")
        rev_cell.fill = PatternFill(fill_type="solid", fgColor="FFF4CCCC")  # Light pink
        rev_cell.border = thin_border

    wb.save(file_path)
    print(f"  [Excel] Successfully saved {file_path}")

def run():
    excel_path = find_excel_file()
    print(f"Using Excel file: {excel_path}")

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    stores = []
    for row in range(3, ws.max_row + 1):
        s_no = ws.cell(row=row, column=1).value
        code = ws.cell(row=row, column=2).value
        brand = ws.cell(row=row, column=3).value or "Bikanervala"
        address = ws.cell(row=row, column=4).value or ""
        link = ws.cell(row=row, column=5).value

        if code and link:
            stores.append({
                "s_no": s_no,
                "code": str(code).strip(),
                "brand": str(brand).strip(),
                "address": str(address).strip(),
                "link": str(link).strip()
            })

    print(f"Loaded {len(stores)} stores to scrape.")

    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled"
            ]
        )
        context = browser.new_context(
            ignore_https_errors=True,
            locale="en-US",
            viewport={"width": 1280, "height": 800},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
        )
        page = context.new_page()

        for s in stores:
            res = scrape_store(page, s["code"], s["brand"], s["address"], s["link"])
            results.append(res)
            # Polite delay between stores to avoid rate limits
            time.sleep(2)

        browser.close()

    # Update Excel
    update_excel_workbook(excel_path, results)

    # Save results summary JSON
    summary_path = BASE_DIR / "scraped_results.json"
    summary_data = {
        "date": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "total_stores": len(stores),
        "successful": sum(1 for r in results if r["status"] == "success"),
        "failed": sum(1 for r in results if r["status"] == "failed"),
        "results": results
    }
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary_data, f, indent=2)
    print(f"\nSaved summary to {summary_path}")

if __name__ == "__main__":
    run()
