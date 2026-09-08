import os
import re
import json
import time
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from playwright.sync_api import sync_playwright

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

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

def parse_with_regex_fallback(html_content):
    """Pure regex fallback using standard library only."""
    try:
        # Maps check
        m_maps_rev = re.search(r'aria-label="([\d,]+)\s*reviews?"', html_content, re.I)
        m_maps_r = re.search(r'class="[^"]*F7nice[^"]*"[^>]*>.*?aria-hidden="true"[^>]*>(\d\.\d)<', html_content, re.S)
        if m_maps_r and m_maps_rev:
            return float(m_maps_r.group(1)), int(m_maps_rev.group(1).replace(",", ""))

        # Google reviews check
        m_rev = re.search(r'([\d,]+)\s*Google\s*reviews?', html_content, re.I)
        if m_rev:
            reviews = int(m_rev.group(1).replace(",", ""))
            start_pos = max(0, m_rev.start() - 600)
            end_pos = min(len(html_content), m_rev.end() + 200)
            snippet = html_content[start_pos:end_pos]
            
            m_rated = re.search(r'Rated\s*(\d\.\d)\s*out of 5', snippet, re.I)
            if m_rated:
                return float(m_rated.group(1)), reviews

            m_span = re.search(r'aria-hidden="true"[^>]*>(\d\.\d)<', snippet)
            if m_span:
                return float(m_span.group(1)), reviews

            m_simple = re.search(r'>(\d\.\d)<', snippet)
            if m_simple:
                candidate = float(m_simple.group(1))
                if 1.0 <= candidate <= 5.0:
                    return candidate, reviews

            return None, reviews
    except Exception as e:
        print(f"    [Regex fallback error] {e}")

    return None, None

def extract_gmb_data(page):
    """
    Extracts strictly the official Google Business Profile rating and review count.
    Primary: In-browser DOM traversal via JavaScript.
    Secondary: Pure regex fallback.
    """
    dismiss_consent_modal(page)

    js_code = r"""
    () => {
        // --- 1. Check Google Maps Panel ---
        const f7 = document.querySelector('div.F7nice');
        if (f7) {
            let r = null;
            let rev = null;

            const rSpan = f7.querySelector('span[aria-hidden="true"]');
            if (rSpan && /^\d\.\d$/.test(rSpan.innerText.trim())) {
                r = parseFloat(rSpan.innerText.trim());
            }

            const ariaRev = f7.querySelector('[aria-label*="reviews"]');
            if (ariaRev) {
                const label = ariaRev.getAttribute('aria-label') || '';
                const m = label.match(/([\d,]+)\s*reviews?/i);
                if (m) rev = parseInt(m[1].replace(/,/g, ''), 10);
            }
            if (!rev) {
                const mParen = f7.innerText.match(/\(\s*([\d,]+)\s*\)/);
                if (mParen) rev = parseInt(mParen[1].replace(/,/g, ''), 10);
            }

            if (r !== null && rev !== null) {
                return { rating: r, reviews: rev, method: 'maps' };
            }
        }

        // --- 2. Google Search Knowledge Graph ---
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
        let revNode = null;
        let revCount = null;

        while (walker.nextNode()) {
            const text = walker.currentNode.nodeValue;
            const m = text.match(/([\d,]+)\s*Google\s*reviews?/i);
            if (m) {
                revCount = parseInt(m[1].replace(/,/g, ''), 10);
                revNode = walker.currentNode.parentElement;
                break;
            }
        }

        if (revNode && revCount !== null) {
            let block = revNode.closest('[data-attrid*="place_ratings"], .Ob27yc');
            if (!block) {
                let p = revNode.parentElement;
                for (let i = 0; i < 3 && p; i++) {
                    if (p.querySelector('[aria-label*="Rated"], [aria-label*="out of 5"], span.aqN8e, span.yi40Hd')) {
                        block = p;
                        break;
                    }
                    p = p.parentElement;
                }
            }
            if (!block) {
                block = revNode.parentElement ? revNode.parentElement.parentElement : revNode;
            }

            let rating = null;

            const ariaEl = block.querySelector('[aria-label*="Rated"], [aria-label*="out of 5"]');
            if (ariaEl) {
                const label = ariaEl.getAttribute('aria-label') || '';
                const mRate = label.match(/Rated\s*(\d\.\d)\s*out of 5/i) || label.match(/(\d\.\d)\s*out of 5/i);
                if (mRate) {
                    rating = parseFloat(mRate[1]);
                }
            }

            if (rating === null) {
                const rateSpan = block.querySelector('span.aqN8e, span.yi40Hd, span.FZ1T5');
                if (rateSpan && /^\d\.\d$/.test(rateSpan.innerText.trim())) {
                    rating = parseFloat(rateSpan.innerText.trim());
                }
            }

            if (rating === null) {
                const spans = block.querySelectorAll('span, div');
                for (const s of spans) {
                    const txt = s.innerText ? s.innerText.trim() : '';
                    if (/^\d\.\d$/.test(txt)) {
                        const val = parseFloat(txt);
                        if (val >= 1.0 && val <= 5.0) {
                            rating = val;
                            break;
                        }
                    }
                }
            }

            if (rating === null) {
                const blockText = block.innerText || '';
                const mPattern = blockText.match(/(\d\.\d)\s*[\r\n\s]*★*[\r\n\s]*[\d,]+\s*Google\s*reviews?/i);
                if (mPattern) {
                    rating = parseFloat(mPattern[1]);
                }
            }

            return { rating: rating, reviews: revCount, method: 'knowledge_graph' };
        }

        return { rating: null, reviews: null, method: 'none' };
    }
    """

    try:
        data = page.evaluate(js_code)
        if data and data.get("rating") is not None and data.get("reviews") is not None:
            return data["rating"], data["reviews"]
    except Exception as e:
        print(f"    [Evaluate error] {e}")

    # Fallback to pure regex parser
    r_re, rev_re = parse_with_regex_fallback(page.content())
    if r_re is not None and rev_re is not None:
        return r_re, rev_re

    return None, None

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

    # Step 2: Fallback search if rating or reviews were not captured
    if rating is None or reviews is None:
        clean_addr = address.replace(",", " ").strip()
        search_query = f"{brand} {clean_addr}"
        search_url = f"https://www.google.com/search?q={search_query.replace(' ', '+')}&hl=en"
        print(f"  [Fallback Search] Querying: {search_url}")
        try:
            page.goto(search_url, timeout=40000, wait_until="domcontentloaded")
            page.wait_for_timeout(3500)
            f_rating, f_reviews = extract_gmb_data(page)
            if f_rating is not None and rating is None:
                rating = f_rating
            if f_reviews is not None and reviews is None:
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
        date_cell.fill = PatternFill(fill_type="solid", fgColor="FFD9D2E9")
        date_cell.border = thin_border
        ws.cell(row=1, column=target_review_col).border = thin_border

        rate_hdr = ws.cell(row=2, column=target_rating_col)
        rate_hdr.value = "GMB Live Rating"
        rate_hdr.font = Font(name="Calibri", size=10, bold=True)
        rate_hdr.alignment = Alignment(horizontal="center", vertical="center")
        rate_hdr.fill = PatternFill(fill_type="solid", fgColor="FFDD7E6B")
        rate_hdr.border = thin_border

        rev_hdr = ws.cell(row=2, column=target_review_col)
        rev_hdr.value = "Review Count"
        rev_hdr.font = Font(name="Calibri", size=10, bold=True)
        rev_hdr.alignment = Alignment(horizontal="center", vertical="center")
        rev_hdr.fill = PatternFill(fill_type="solid", fgColor="FFFFF2CC")
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
