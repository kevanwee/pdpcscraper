# PDPC Decisions Scraper & Analyser

A Python toolkit that scrapes all PDPC (Personal Data Protection Commission) enforcement decisions filtered by the **Protection obligation**, exports them to Excel, and runs a trend analysis with charts and summary tables.

---

## Files

| File | Description |
|------|-------------|
| `scraper.py` | Scrapes PDPC decisions and writes `pdpc_decisions.xlsx` |
| `analyse.py` | Reads `pdpc_decisions.xlsx` and produces `pdpc_analysis.xlsx` + `pdpc_trends.png` |
| `requirements.txt` | Python dependencies |
| `pdpc_decisions.xlsx` | Raw decisions dataset (211 rows as of Mar 2026) |
| `pdpc_analysis.xlsx` | Summary tables and aggregates |
| `pdpc_trends.png` | 6-panel trend chart |

---

## Setup

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

---

## Scraper — `scraper.py`

### Usage

```bash
python scraper.py
```

### What it does

1. Launches a headless browser, navigates to the [PDPC decisions page](https://www.pdpc.gov.sg/all-commissions-decisions), applies the **Protection** filter, and paginates through all 34 pages
2. Visits each individual decision page to extract the summary description and financial penalty
3. Downloads each decision PDF to extract the neutral case citation
4. Writes everything to `pdpc_decisions.xlsx`

Expected runtime: ~15 minutes (~210 decisions, due to PDF downloads).

### Output — `pdpc_decisions.xlsx`

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

### Notes

- **Citation coverage**: ~75% of decisions have machine-readable PDFs. The remainder are scanned images or unpublished PDFs — those rows will have an empty citation field.
- **Financial penalty**: Decisions that received only Directions or a Warning are marked `None`.
- The scraper uses a 1.5-second delay between requests to avoid rate-limiting.

---

## Analyser — `analyse.py`

### Usage

```bash
python analyse.py
```

Reads `pdpc_decisions.xlsx` (must exist) and produces two output files.

### Output 1 — `pdpc_analysis.xlsx`

Five summary sheets:

| Sheet | Contents |
|-------|----------|
| **Yearly Summary** | Decisions per year, penalty count, total/average/max penalty, Protection-only count |
| **Penalty Distribution** | Breakdown of penalty amounts into bands ($1–$10K, $10K–$50K, etc.) |
| **Obligation Mix** | All obligation combinations observed, case counts, average penalty per group |
| **Decision Outcomes** | Breakdown by outcome type (Penalty + Directions, Directions Only, No Breach, etc.) |
| **Notable Cases** | Top 10 largest financial penalties with case name, date, and citation |

### Output 2 — `pdpc_trends.png`

Six-panel chart covering:

1. Decisions per year
2. Penalty vs direction-only decisions by year
3. Average financial penalty trend by year
4. Penalty amount distribution (bar chart)
5. Decision outcome breakdown (pie chart)
6. Obligations co-occurring with Protection

### Key findings (as of Mar 2026)

| Metric | Value |
|--------|-------|
| Total decisions | 211 |
| Date range | 2016 – 2026 |
| With financial penalty | 139 (66%) |
| Total penalties imposed | $3,498,050 |
| Average penalty | $25,165 |
| Median penalty | $11,000 |
| Largest single penalty | $315,000 |
| Protection-only cases | 175 (83%) |
| Most common co-obligation | Accountability (28 cases) |

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `playwright` | Headless browser for JS-rendered listing pages |
| `requests` | HTTP client for detail pages and PDFs |
| `beautifulsoup4` + `lxml` | HTML parsing |
| `pypdf` | PDF text extraction for case citations |
| `openpyxl` | Excel output |
| `matplotlib` + `seaborn` | Charts |
| `numpy` | Aggregation helpers |
