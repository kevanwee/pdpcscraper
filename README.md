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
```

No browser is needed.

---

## Scraper — `scraper.py`

### Usage

```bash
python scraper.py                 # every Protection Obligation decision (~20 minutes)
python scraper.py --limit 20      # the first 20 decision pages, for a quick check
python scraper.py --no-pdf        # skip the PDFs (faster; fewer citations)
python scraper.py --output out.xlsx --delay 2
```

### What it does

1. Reads the [enforcement decisions page](https://www.pdpc.gov.sg/organisations/regulations-decisions/enforcement-decisions?type=Commission%27s+Decisions). The site is server-rendered (Next.js) and embeds an index of every decision (about 390), so one request lists them all.
2. Visits each decision page for its published date, summary and financial penalty. Obligations come from the title; the decision type (financial penalty, directions, warning, not in breach) from the summary.
3. Keeps the Protection Obligation decisions: those whose title names the Protection Obligation, plus older "Data Protection Provisions" decisions and undertakings whose summary does.
4. Downloads each decision PDF to extract the neutral case citation.
5. Writes everything to `pdpc_decisions.xlsx`.

Expected runtime: ~20 minutes (about 390 pages at 1.5 seconds, plus PDFs).

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

### Key findings (as of Oct 2026)

| Metric | Value |
|--------|-------|
| Total decisions | 216 |
| Date range | 2016 – 2026 |
| With financial penalty | 140 (65%) |
| Total penalties imposed | $3,569,550 |
| Average penalty | $25,496 |
| Median penalty | $11,500 |
| Largest single penalty | $315,000 |
| Protection-only cases | 183 (85%) |
| Most common co-obligation | Accountability (22 cases) |

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `requests` | HTTP client for the listing, decision pages and PDFs |
| `beautifulsoup4` | HTML parsing |
| `pypdf` | PDF text extraction for case citations |
| `openpyxl` | Excel output |
| `matplotlib` | Charts |
| `numpy` | Aggregation helpers |
