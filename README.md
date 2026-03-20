# PDPC Decisions Scraper

A Python scraper that extracts all PDPC (Personal Data Protection Commission) enforcement decisions filtered by the **Protection obligation** and exports them to an Excel spreadsheet.

## Output

Produces `pdpc_decisions.xlsx` with the following columns:

| Column | Description |
|--------|-------------|
| **Case Name** | Normalised to `Re [Party Name]` format |
| **Case Citation** | Neutral citation (e.g. `[2024] SGPDPC 3`) extracted from the decision PDF |
| **Date** | Published date |
| **Obligations** | All DP obligations involved (e.g. `Protection`, `Accountability, Protection`) |
| **Decision Type** | Outcome (e.g. `Financial Penalty, Directions`, `Warning`) |
| **Case Description** | Summary paragraph from the decision page |
| **Financial Penalty** | Dollar amount imposed, or `None` for direction/warning-only decisions |
| **URL** | Link to the PDPC decision page |

## Setup

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

## Usage

```bash
python scraper.py
```

The script will:
1. Launch a headless browser, navigate to the [PDPC decisions page](https://www.pdpc.gov.sg/all-commissions-decisions), apply the Protection filter, and paginate through all results
2. Visit each decision page to extract the description and financial penalty
3. Download each decision PDF to extract the case citation
4. Write all data to `pdpc_decisions.xlsx`

Expected runtime: ~15 minutes for ~210 decisions (due to PDF downloads).

## Notes

- **Citation coverage**: ~75% of decisions have machine-readable PDFs. The remainder are scanned images or don't yet have a published PDF — these rows will have an empty citation field.
- **Financial penalty**: Decisions that received only Directions or a Warning are marked `None` in the penalty column.
- The scraper uses a 1.5-second delay between requests to avoid rate-limiting.

## Dependencies

| Package | Purpose |
|---------|---------|
| `playwright` | Headless browser for JS-rendered listing pages |
| `requests` | HTTP client for detail pages and PDFs |
| `beautifulsoup4` + `lxml` | HTML parsing |
| `pypdf` | PDF text extraction for case citations |
| `openpyxl` | Excel output |
