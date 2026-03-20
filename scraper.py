"""
PDPC Decisions Scraper
Extracts all "Protection" obligation decisions from https://www.pdpc.gov.sg/all-commissions-decisions
and writes them to pdpc_decisions.xlsx

Uses Playwright for JS-rendered pagination, requests+BeautifulSoup for detail pages.
"""

import io
import re
import time
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from pypdf import PdfReader

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

BASE_URL     = "https://www.pdpc.gov.sg"
LISTING_URL  = f"{BASE_URL}/all-commissions-decisions"
OUTPUT_FILE  = "pdpc_decisions.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

DELAY_DETAIL = 1.5   # seconds between individual decision page requests

# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------

# Case citation: [YYYY] SGPDPC N or [YYYY] SGPDPCS N (newer summary format)
CITATION_RE = re.compile(
    r'\[\d{4}\]\s+SGPDPCS?\s+\d+(?:\s*\(NFA\))?',
    re.IGNORECASE,
)

# PDF link on detail pages
PDF_LINK_RE = re.compile(r'\.pdf', re.IGNORECASE)

# Context-aware financial penalty: "financial penalty of $X,XXX"
PENALTY_CONTEXT_RE = re.compile(
    r'(?:financial\s+penalty\s+of\s+|penalty\s+of\s+)'
    r'(?:S\$|\$)([\d,]+(?:\.\d{2})?)',
    re.IGNORECASE,
)

# Fallback: any standalone dollar amount
PENALTY_FALLBACK_RE = re.compile(
    r'(?:S\$|\$)([\d,]+(?:\.\d{2})?)',
    re.IGNORECASE,
)

# ---------------------------------------------------------------------------
# HTTP helpers (for detail pages)
# ---------------------------------------------------------------------------

def safe_get(session: requests.Session, url: str, retries: int = 3, backoff: float = 2.0):
    """GET url with retry + exponential backoff. Returns Response or None."""
    for attempt in range(retries):
        try:
            resp = session.get(url, timeout=20, headers=HEADERS)
            if resp.status_code == 429:
                wait = backoff * (2 ** attempt)
                print(f"  Rate limited. Waiting {wait:.0f}s...")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            return resp
        except requests.exceptions.RequestException as e:
            wait = backoff * (2 ** attempt)
            print(f"  Request failed (attempt {attempt + 1}/{retries}): {e}")
            if attempt < retries - 1:
                time.sleep(wait)
    return None

# ---------------------------------------------------------------------------
# Extraction helpers
# ---------------------------------------------------------------------------

def extract_citation_from_pdf(session: requests.Session, pdf_url: str) -> str:
    """Download the decision PDF and search the first few pages for the case citation."""
    try:
        resp = session.get(pdf_url, timeout=30)
        resp.raise_for_status()
        reader = PdfReader(io.BytesIO(resp.content))
        if not reader.pages:
            return ""
        # Search first 3 pages — citation can appear on page 0 or 1
        for page in reader.pages[:3]:
            text = page.extract_text() or ""
            match = CITATION_RE.search(text)
            if match:
                return match.group(0).strip()
        return ""
    except Exception:
        return ""


def extract_citation(text: str) -> str:
    match = CITATION_RE.search(text)
    return match.group(0).strip() if match else ""


def extract_penalty(text: str) -> str:
    # Try context-aware extraction first
    match = PENALTY_CONTEXT_RE.search(text)
    if match:
        return f"${match.group(1)}"

    # If no dollar amount at all, determine whether there simply is no penalty
    has_dollar = PENALTY_FALLBACK_RE.search(text)
    if not has_dollar:
        lower = text.lower()
        no_penalty_keywords = [
            "no further action", "warning", "not in breach",
            "advisory notice", "directions were issued", "directions issued",
            "no breach",
        ]
        if any(kw in lower for kw in no_penalty_keywords):
            return "None"
        # No dollar amount and no recognisable decision keyword — leave blank
        return ""

    # Fallback: first dollar amount anywhere
    match = PENALTY_FALLBACK_RE.search(text)
    return f"${match.group(1)}" if match else ""


def normalise_case_name(title: str) -> str:
    """
    Convert 'Breach of the Protection Obligation by Acme Pte Ltd'
    → 'Re Acme Pte Ltd'
    Handles 'No Breach', 'No breach', 'Breach of Data Protection...' etc.
    """
    match = re.search(r'\bby\s+(.+)$', title, re.IGNORECASE)
    if match:
        return "Re " + match.group(1).strip()
    return title


def extract_description(rte_soup) -> str:
    """Return the first substantive paragraph from the .rte block."""
    if rte_soup is None:
        return ""
    for p in rte_soup.find_all("p"):
        text = p.get_text(" ", strip=True)
        if len(text) > 80 and not CITATION_RE.match(text):
            return text
    return rte_soup.get_text(" ", strip=True)[:600]

# ---------------------------------------------------------------------------
# Phase 1: Use Playwright to get all listings with Protection filter
# ---------------------------------------------------------------------------

def fetch_all_listings_playwright() -> list[dict]:
    """
    Use a headless browser to:
    1. Navigate to the all-commissions-decisions page
    2. Tick the Protection checkbox
    3. Paginate through all pages, collecting decision metadata
    """
    print("Phase 1: Launching browser to scrape listing pages...")
    items = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page(user_agent=HEADERS["User-Agent"])

        # Intercept API responses to capture raw JSON data
        captured_items = []

        def handle_response(response):
            if "getenforcementcase" in response.url and response.status == 200:
                try:
                    data = response.json()
                    if "items" in data:
                        captured_items.extend(data["items"])
                except Exception:
                    pass

        page.on("response", handle_response)

        print(f"  Navigating to {LISTING_URL} ...")
        page.goto(LISTING_URL, wait_until="networkidle", timeout=30000)

        # Tick the Protection checkbox
        print("  Applying Protection filter...")
        try:
            # Try multiple selector approaches
            protection_cb = page.locator(
                "input[type='checkbox'][id='protection'], "
                "input[type='checkbox'][value='protection'], "
                "input[type='checkbox'][value='Protection']"
            ).first
            protection_cb.check()
            # Wait for the listing to update after filter is applied
            page.wait_for_load_state("networkidle", timeout=10000)
        except PWTimeout:
            print("  Warning: networkidle timeout after filter — continuing anyway")
        except Exception as e:
            print(f"  Warning: could not tick Protection checkbox: {e}")
            print("  Proceeding without filter — will filter client-side")

        # Find out how many pages there are
        try:
            total_pages_el = page.locator(".total-pages").first
            total_pages = int(total_pages_el.inner_text(timeout=5000).strip())
        except Exception:
            total_pages = 1
        print(f"  Detected {total_pages} page(s) of results")

        # Scrape current page items from the DOM as well (belt-and-suspenders)
        def scrape_page_items():
            cards = page.locator(".card").all()
            page_items = []
            for card in cards:
                try:
                    title = card.locator(".card__title").inner_text(timeout=2000).strip()
                    descs = card.locator(".card__desc").all()
                    nature   = descs[0].inner_text(timeout=1000).replace("Relevant DP obligation(s):", "").strip() if len(descs) > 0 else ""
                    decision = descs[1].inner_text(timeout=1000).replace("Decision:", "").strip() if len(descs) > 1 else ""
                    date_txt = descs[2].inner_text(timeout=1000).replace("Published Date:", "").strip() if len(descs) > 2 else ""
                    href = card.get_attribute("href") or ""
                    page_items.append({
                        "title": title,
                        "nature": nature,
                        "decision": decision,
                        "date": date_txt,
                        "url": href,
                    })
                except Exception:
                    pass
            return page_items

        # Collect items from page 1
        dom_items = scrape_page_items()
        if dom_items:
            items.extend(dom_items)
            print(f"  Page 1/{total_pages} — {len(items)} items so far (DOM)")

        # Navigate through remaining pages
        for page_num in range(2, total_pages + 1):
            try:
                # Click the "Next" pagination button
                next_btn = page.locator(".pagination-next a, a[data-page]").filter(
                    has_text=str(page_num)
                ).first
                if not next_btn.is_visible(timeout=2000):
                    # Fallback: click the generic "Next" button
                    next_btn = page.locator(".pagination-next a").first

                next_btn.click(timeout=5000)
                page.wait_for_load_state("networkidle", timeout=10000)
            except PWTimeout:
                print(f"  Timeout on page {page_num} navigation — continuing")
            except Exception as e:
                print(f"  Navigation error on page {page_num}: {e}")
                break

            dom_items = scrape_page_items()
            if dom_items:
                items.extend(dom_items)
            print(f"  Page {page_num}/{total_pages} — {len(items)} items so far")

        browser.close()

    # If API interception captured items, prefer those (they have cleaner data)
    if captured_items:
        print(f"  API interception captured {len(captured_items)} items (using these)")
        items = captured_items

    # Deduplicate by URL
    seen_urls = set()
    unique_items = []
    for item in items:
        url = item.get("url", "")
        if url and url not in seen_urls:
            seen_urls.add(url)
            unique_items.append(item)

    # Filter for Protection obligation (client-side safety net)
    protection_items = [
        i for i in unique_items
        if "protection" in i.get("nature", "").lower()
    ]

    print(f"  Total unique Protection decisions: {len(protection_items)}")
    return protection_items

# ---------------------------------------------------------------------------
# Phase 2: Fetch each detail page
# ---------------------------------------------------------------------------

def fetch_detail(session: requests.Session, item: dict) -> dict:
    """Fetch an individual decision page and return enriched row data."""
    relative_url = item.get("url", "")
    full_url = BASE_URL + relative_url if relative_url.startswith("/") else relative_url

    raw_title = item.get("title", "")
    row = {
        "case_name":     normalise_case_name(raw_title),
        "citation":      "",
        "date":          item.get("date", ""),
        "obligations":   item.get("nature", ""),
        "decision_type": item.get("decision", ""),
        "description":   "",
        "penalty":       "",
        "url":           full_url,
    }

    resp = safe_get(session, full_url)
    if resp is None:
        return row

    soup = BeautifulSoup(resp.text, "lxml")

    rte = soup.select_one(".rte")
    rte_text = rte.get_text(" ", strip=True) if rte else ""

    # Citation: extract from the decision PDF (most reliable source)
    pdf_link = soup.find("a", href=PDF_LINK_RE)
    if pdf_link:
        pdf_href = pdf_link.get("href", "")
        pdf_url = BASE_URL + pdf_href if pdf_href.startswith("/") else pdf_href
        row["citation"] = extract_citation_from_pdf(session, pdf_url)
    # Fallback: search page text (catches some older pages)
    if not row["citation"]:
        row["citation"] = extract_citation(soup.get_text(" ", strip=True))

    row["description"] = extract_description(rte)
    row["penalty"]     = extract_penalty(rte_text)

    return row

# ---------------------------------------------------------------------------
# Phase 3: Write Excel
# ---------------------------------------------------------------------------

HEADER = [
    "Case Name",
    "Case Citation",
    "Date",
    "Obligations",
    "Decision Type",
    "Case Description",
    "Financial Penalty",
    "URL",
]

COL_WIDTHS = [60, 22, 14, 35, 30, 80, 20, 80]


def write_excel(rows: list[dict], filename: str = OUTPUT_FILE) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "PDPC Decisions"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrap_align   = Alignment(vertical="top", wrap_text=True)

    for col_idx, header_text in enumerate(HEADER, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header_text)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center_align

    col_letters = ["A", "B", "C", "D", "E", "F", "G", "H"]
    for letter, width in zip(col_letters, COL_WIDTHS):
        ws.column_dimensions[letter].width = width

    ws.freeze_panes = "A2"

    for row_idx, row in enumerate(rows, start=2):
        values = [
            row["case_name"],
            row["citation"],
            row["date"],
            row["obligations"],
            row["decision_type"],
            row["description"],
            row["penalty"],
            row["url"],
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = wrap_align

    wb.save(filename)
    print(f"\nSaved {len(rows)} rows to {filename}")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    # Phase 1: collect all Protection listings via browser
    listings = fetch_all_listings_playwright()
    if not listings:
        print("No listings found. Exiting.")
        return

    # Phase 2: enrich each with detail page data
    session = requests.Session()
    print(f"\nPhase 2: Fetching detail pages for {len(listings)} decisions...")
    rows = []
    for i, item in enumerate(listings, start=1):
        title_preview = item.get("title", "")[:70]
        print(f"  [{i:3d}/{len(listings)}] {title_preview}")
        row = fetch_detail(session, item)
        rows.append(row)
        time.sleep(DELAY_DETAIL)

    # Phase 3: write Excel
    print("\nPhase 3: Writing Excel file...")
    write_excel(rows)
    print("Done.")


if __name__ == "__main__":
    main()
