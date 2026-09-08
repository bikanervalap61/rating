import os
import re
import json
import time
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from playwright.sync_api import sync_playwright

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

def dismiss_consent_modal(page):
    """Dismisses Google cookie/consent dialog if it appears."""
    try:
        consent_buttons = [
            "button:has-text('Accept all')",
            "button:has-text('I agree')",
            "button:has-text('Accept')",
            "button[aria-label*='Accept all']"
        ]
        for selector in consent_buttons:
            btn = page.query_selector(selector)
            if btn and btn.is_visible():
                btn.click()
                page.wait_for_timeout(1500)
                break
    except Exception:
        pass

def extract_gmb_data(page):
    """
    Strictly extracts ONLY Google My Business / Google Maps rating & reviews.
    Excludes any 3rd party platforms (Zomato, Swiggy, TripAdvisor, Justdial).
    """
    rating = None
    reviews = None

    dismiss_consent_modal(page)

    # Strategy 1: Google Search Desktop Knowledge Graph (Right Hand Side: #rhs)
    # The Knowledge Panel strictly lives inside #rhs or div[data-attrid*='kc:/local:']
    try:
        rhs = page.query_selector("#rhs, div.kp-wholepage, div[data-attrid*='kc:/local:']")
        if rhs:
            # 1a. Reviews: Look strictly for "Google reviews" inside the knowledge card
            rev_el = rhs.query_selector("a:has-text('Google review'), span:has-text('Google review')")
            if rev_el:
                rev_text = rev_el.inner_text()
                m_rev = re.search(r"([\d,]+)\s*Google\s*reviews?", rev_text, re.I)
                if m_rev:
                    reviews = clean_number(m_rev.group(1))

            # 1b. Rating: Look for rating span inside the knowledge card
            rate_el = rhs.query_selector("span.aqN8e, span.FZ1T5, span[aria-hidden='true']:has-text('.')")
            if rate_el:
                r_text = rate_el.inner_text()
                m_rate = re.search(r"^(\d\.\d)$", r_text.strip())
                if m_rate:
                    rating = float(m_rate.group(1))

            # 1c. If rating wasn't in span.aqN8e, check aria-label within the ratings block
            if not rating:
                rate_block = rhs.query_selector("[data-attrid*='place_ratings'], .Ob27yc")
                if rate_block:
                    aria_el = rate_block.query_selector("[aria-label*='Rated'], [aria-label*='out of 5']")
                    if aria_el:
                        m_aria = re.search(r"(\d\.\d)\s*(?:stars?|out of 5)", aria_el.get_attribute("aria-label") or "", re.I)
                        if m_aria:
                            rating = float(m_aria.group(1))
                    if not rating:
                        # Direct regex on the rating block text
                        m_block = re.search(r"(\d\.\d)", rate_block.inner_text())
                        if m_block:
                            rating = float(m_block.group(1))
    except Exception as e:
        print(f"    [Knowledge Graph parser debug] {e}")

    # Strategy 2: Google Maps (e.g. if URL redirected to maps.google.com)
    if rating is None or reviews is None:
        try:
            maps_container = page.query_selector("div.F7nice, div.TIHn2, div.fontDisplayLarge")
            if maps_container:
                # Maps rating
                m_span = maps_container.query_selector("span[aria-hidden='true']")
                if m_span:
                    r_candidate = clean_rating(m_span.inner_text())
                    if r_candidate:
                        rating = r_candidate

                # Maps reviews
                m_rev_aria = maps_container.query_selector("[aria-label*='reviews']")
                if m_rev_aria:
                    label = m_rev_aria.get_attribute("aria-label") or ""
                    m = re.search(r"([\d,]+)\s*reviews?", label, re.I)
                    if m:
                        reviews = clean_number(m.group(1))
                if not reviews:
                    m_paren = re.search(r"\(\s*([\d,]+)\s*\)", maps_container.inner_text())
                    if m_paren:
                        reviews = clean_number(m_paren.group(1))
        except Exception as e:
            print(f"    [Maps parser debug] {e}")

    # Strategy 3: Target the specific phrase "Google reviews" anywhere on the page
    # This prevents matching ANY 3rd party site because only Google uses "Google reviews"
    if rating is None or reviews is None:
        try:
            # Find elements containing "Google review"
            google_rev_elements = page.query_selector_all("span:has-text('Google review'), a:has-text('Google review')")
            for g_el in google_rev_elements:
                text = g_el.inner_text()
                m_rev = re.search(r"([\d,]+)\s*Google\s*reviews?", text, re.I)
                if m_rev and not reviews:
                    reviews = clean_number(m_rev.group(1))

                # Look around this element for the rating
                if not rating:
                    # Get surrounding parent container
                    parent_text = g_el.evaluate("el => el.parentElement ? (el.parentElement.parentElement ? el.parentElement.parentElement.innerText : el.parentElement.innerText) : ''")
                    # Match pattern: 4.4 ★★★★★ 4,685 Google reviews
                    m_combined = re.search(r"(\d\.\d)\s*[\r\n\s]*★*[\r\n\s]*[\d,]+\s*Google\s*reviews?", parent_text, re.I)
                    if m_combined:
                        rating = float(m_combined.group(1))
                    else:
                        m_simple = re.search(r"(\d\.\d)", parent_text)
                        if m_simple:
                            candidate = float(m_simple.group(1))
                            if 1.0 <= candidate <= 5.0:
                                rating = candidate

                if rating and reviews:
                    break
        except Exception as e:
            print(f"    [Text parser debug] {e}")

    return rating, reviews

def scrape_store(page, store_code, brand, address, link):
    print(f"\n[Scraping] {store_code} - {brand} ({address})")
    print(f"  Target URL: {link}")

    rating = None
    reviews = None
    error_msg = None

    # Step 1: Open direct link
    try:
        page.goto(link, timeout=40000, wait_until="domcontentloaded")
        page.wait_for_timeout(3500)
        rating, reviews = extract_gmb_data(page)
    except Exception as e:
        error_msg = str(e)
        print(f"  [Direct link warning] {e}")

    # Step 2: Fallback to targeted search if either rating or review count wasn't captured
    if rating is None or reviews is None:
        clean_addr = address.replace(",", " ").strip()
        search_query = f"{brand} {clean_addr}"
        search_url = f"https://www.google.com/search?q={search_query.replace(' ', '+')}&hl=en"
        print(f"  [Fallback Search] Querying: {search_url}")
        try:
            page.goto(search_url, timeout=40000, wait_until="domcontentloaded")
            page.wait_for_timeout(3500)
            f_rating, f_reviews = extract_gmb_data(page)
            if f_rating and not rating:
                rating = f_rating
            if f_reviews and not reviews:
                reviews = f_reviews
        except Exception as e:
            print(f"  [Fallback warning] {e}")

    print(f"  -> Result: GMB Live Rating = {rating} | Review Count = {reviews}")
    return {
        "store_code": store_code,
        "brand": brand,
        "address": address,
        "link": link,
        "rating": rating,
        "reviews": reviews,
        "status": "success" if (rating is not None and reviews is not None) else "failed",
        "error": error_msg
    }

def update_excel_workbook(file_path, scraped_data):
    print(f"\n[Excel] Updating workbook: {file_path}")
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    today_dt = datetime.now()
    today_date_str = today_dt.strftime("%d/%m/%Y")

    store_row_map = {}
    for r in range(3, ws.max_row + 1):
        code_val = ws.cell(row=r, column=2).value
        if code_val:
            store_row_map[str(code_val).strip()] = r

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

    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3")
    )

    if target_rating_col is None:
        target_rating_col = ws.max_column + 1
        target_review_col = target_rating_col + 1

        print(f"  Adding new date columns: {target_rating_col} & {target_review_col} for {today_date_str}")
        
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
        rating_cell.fill = PatternFill(fill_type="solid", fgColor="FFD0E0E3")
        rating_cell.border = thin_border

        rev_cell = ws.cell(row=row_num, column=target_review_col)
        if item["reviews"] is not None:
            rev_cell.value = item["reviews"]
        rev_cell.font = Font(name="Calibri", size=11, bold=False)
        rev_cell.alignment = Alignment(horizontal="center", vertical="center")
        rev_cell.fill = PatternFill(fill_type="solid", fgColor="FFF4CCCC")
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
            viewport={"width": 1920, "height": 1080},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            extra_http_headers={
                "Accept-Language": "en-US,en;q=0.9"
            }
        )
        page = context.new_page()

        for s in stores:
            res = scrape_store(page, s["code"], s["brand"], s["address"], s["link"])
            results.append(res)
            time.sleep(2)

        browser.close()

    update_excel_workbook(excel_path, results)

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
