"""
PDPC Decisions — Trend Analysis
Reads pdpc_decisions.xlsx and produces:
  1. pdpc_analysis.xlsx  — summary tables (yearly, obligation mix, penalty bands)
  2. pdpc_trends.png     — multi-panel chart
"""

import re
from collections import Counter, defaultdict
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── helpers ──────────────────────────────────────────────────────────────────

MONTH_MAP = {m: i+1 for i, m in enumerate(
    ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
)}

def parse_date(s):
    if not s:
        return None
    try:
        parts = s.strip().split()
        d, m, y = int(parts[0]), MONTH_MAP[parts[1]], int(parts[2])
        return datetime(y, m, d)
    except Exception:
        return None

def parse_penalty(s):
    if not s or s == "None":
        return None
    digits = re.sub(r"[^\d]", "", s)
    return int(digits) if digits else None

def normalise_obligations(s):
    """Return a frozenset of individual obligation strings."""
    if not s:
        return frozenset()
    parts = [p.strip() for p in re.split(r"[,;]", s) if p.strip()]
    return frozenset(parts)

# ── load data ─────────────────────────────────────────────────────────────────

wb_in = load_workbook("pdpc_decisions.xlsx")
ws_in = wb_in.active

records = []
for r in range(2, ws_in.max_row + 1):
    raw = {
        "name":        ws_in.cell(r, 1).value or "",
        "citation":    ws_in.cell(r, 2).value or "",
        "date_str":    ws_in.cell(r, 3).value or "",
        "obligations": ws_in.cell(r, 4).value or "",
        "decision":    ws_in.cell(r, 5).value or "",
        "description": ws_in.cell(r, 6).value or "",
        "penalty_str": ws_in.cell(r, 7).value or "",
        "url":         ws_in.cell(r, 8).value or "",
    }
    raw["date"]        = parse_date(raw["date_str"])
    raw["year"]        = raw["date"].year if raw["date"] else None
    raw["penalty_amt"] = parse_penalty(raw["penalty_str"])
    raw["has_penalty"] = raw["penalty_amt"] is not None
    raw["obligs"]      = normalise_obligations(raw["obligations"])
    raw["pure_protection"] = raw["obligs"] == frozenset(["Protection"])
    records.append(raw)

# ── computed aggregates ───────────────────────────────────────────────────────

years = sorted(y for r in records if (y := r["year"]) and y >= 2016)
year_range = list(range(min(years), max(years) + 1))

def by_year(records, key_fn, agg_fn=len):
    buckets = defaultdict(list)
    for r in records:
        if r["year"]:
            buckets[r["year"]].append(r)
    return {y: agg_fn(key_fn(buckets.get(y, []))) for y in year_range}

# Cases per year
cases_per_year = {y: len([r for r in records if r["year"] == y]) for y in year_range}

# Penalties per year (count and total $)
pen_count_year = {y: sum(1 for r in records if r["year"] == y and r["has_penalty"]) for y in year_range}
pen_total_year = {y: sum(r["penalty_amt"] for r in records if r["year"] == y and r["has_penalty"]) for y in year_range}
pen_avg_year   = {y: (pen_total_year[y] // pen_count_year[y]) if pen_count_year[y] else 0 for y in year_range}

# Obligation mix per year
pure_prot_year  = {y: sum(1 for r in records if r["year"] == y and r["pure_protection"]) for y in year_range}
multi_oblig_year = {y: sum(1 for r in records if r["year"] == y and not r["pure_protection"] and len(r["obligs"]) > 1) for y in year_range}

# Penalty bands
def band(amt):
    if amt is None: return None
    if amt == 0:            return "$0"
    if amt <= 10_000:       return "$1 – $10K"
    if amt <= 50_000:       return "$10K – $50K"
    if amt <= 100_000:      return "$50K – $100K"
    if amt <= 500_000:      return "$100K – $500K"
    return ">$500K"

BANDS = ["$1 – $10K", "$10K – $50K", "$50K – $100K", "$100K – $500K", ">$500K"]
band_counts = Counter(band(r["penalty_amt"]) for r in records if r["has_penalty"])

# Decision outcome breakdown
def outcome(r):
    d = r["decision"].lower()
    if "not in breach" in d or "no breach" in d or "no further action" in d:
        return "No Breach / NFA"
    if "financial penalty" in d and "directions" in d:
        return "Penalty + Directions"
    if "financial penalty" in d:
        return "Penalty Only"
    if "directions" in d:
        return "Directions Only"
    if "warning" in d:
        return "Warning"
    return "Other"

outcome_counts = Counter(outcome(r) for r in records)

# Co-obligation frequency (what pairs with Protection most)
co_oblig_counter = Counter()
for r in records:
    for o in r["obligs"]:
        if o != "Protection":
            co_oblig_counter[o] += 1

# ── matplotlib chart ──────────────────────────────────────────────────────────

BLUE  = "#1F4E79"
MID   = "#2E75B6"
LIGHT = "#9DC3E6"
RED   = "#C00000"
GREEN = "#375623"
GOLD  = "#BF8F00"

fig, axes = plt.subplots(2, 3, figsize=(18, 11))
fig.suptitle("PDPC Protection Obligation Decisions — Trend Analysis", fontsize=16, fontweight="bold", y=0.98)
fig.patch.set_facecolor("#F7F7F7")

def style_ax(ax, title):
    ax.set_title(title, fontsize=11, fontweight="bold", pad=8)
    ax.set_facecolor("#FAFAFA")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(labelsize=8)

# ── 1. Cases per year ──────────────────────────────────────────────────────────
ax = axes[0, 0]
style_ax(ax, "1. Decisions per Year")
bars = ax.bar(year_range, [cases_per_year[y] for y in year_range], color=BLUE, width=0.6)
ax.bar_label(bars, fontsize=7, padding=2)
ax.set_xlabel("Year", fontsize=9)
ax.set_ylabel("Number of decisions", fontsize=9)
ax.set_xticks(year_range)
ax.set_xticklabels(year_range, rotation=45, ha="right")

# ── 2. Penalty cases vs direction-only ────────────────────────────────────────
ax = axes[0, 1]
style_ax(ax, "2. Penalty vs Direction-Only by Year")
w = 0.4
xs = np.array(year_range)
pen_vals = [pen_count_year[y] for y in year_range]
dir_vals = [cases_per_year[y] - pen_count_year[y] for y in year_range]
b1 = ax.bar(xs - w/2, pen_vals, w, label="Financial Penalty", color=RED)
b2 = ax.bar(xs + w/2, dir_vals, w, label="No Penalty", color=LIGHT)
ax.legend(fontsize=8)
ax.set_xlabel("Year", fontsize=9)
ax.set_ylabel("Number of decisions", fontsize=9)
ax.set_xticks(year_range)
ax.set_xticklabels(year_range, rotation=45, ha="right")

# ── 3. Average penalty per year ────────────────────────────────────────────────
ax = axes[0, 2]
style_ax(ax, "3. Average Financial Penalty by Year (SGD)")
avg_vals = [pen_avg_year[y] for y in year_range]
ax.plot(year_range, avg_vals, marker="o", color=BLUE, linewidth=2, markersize=6)
ax.fill_between(year_range, avg_vals, alpha=0.15, color=BLUE)
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x/1000:.0f}K" if x >= 1000 else f"${x:.0f}"))
ax.set_xlabel("Year", fontsize=9)
ax.set_ylabel("Avg penalty (SGD)", fontsize=9)
ax.set_xticks(year_range)
ax.set_xticklabels(year_range, rotation=45, ha="right")
for x, y in zip(year_range, avg_vals):
    if y:
        ax.annotate(f"${y/1000:.0f}K" if y >= 1000 else f"${y:.0f}",
                    (x, y), textcoords="offset points", xytext=(0, 7),
                    ha="center", fontsize=7)

# ── 4. Penalty band distribution ───────────────────────────────────────────────
ax = axes[1, 0]
style_ax(ax, "4. Penalty Amount Distribution")
band_vals = [band_counts.get(b, 0) for b in BANDS]
colors = [LIGHT, MID, BLUE, "#0D2F50", RED]
bars = ax.barh(BANDS, band_vals, color=colors)
ax.bar_label(bars, fontsize=8, padding=3)
ax.set_xlabel("Number of decisions", fontsize=9)
ax.invert_yaxis()

# ── 5. Decision outcome breakdown ──────────────────────────────────────────────
ax = axes[1, 1]
style_ax(ax, "5. Decision Outcomes")
OUTCOME_ORDER = ["Penalty + Directions", "Penalty Only", "Directions Only",
                 "Warning", "No Breach / NFA", "Other"]
out_vals = [outcome_counts.get(o, 0) for o in OUTCOME_ORDER]
out_cols  = [RED, "#FF4040", LIGHT, GOLD, GREEN, "#888"]
wedges, texts, autotexts = ax.pie(
    out_vals, labels=None, autopct=lambda p: f"{p:.0f}%" if p > 3 else "",
    colors=out_cols, startangle=140, pctdistance=0.75,
    wedgeprops={"linewidth": 0.5, "edgecolor": "white"}
)
for at in autotexts:
    at.set_fontsize(8)
ax.legend(
    [f"{o} ({v})" for o, v in zip(OUTCOME_ORDER, out_vals)],
    fontsize=7.5, loc="lower center", bbox_to_anchor=(0.5, -0.22), ncol=2
)

# ── 6. Co-obligation frequency ─────────────────────────────────────────────────
ax = axes[1, 2]
style_ax(ax, "6. Obligations Co-occurring with Protection")
if co_oblig_counter:
    co_labels, co_vals = zip(*co_oblig_counter.most_common(8))
    bars = ax.barh(list(co_labels), list(co_vals), color=MID)
    ax.bar_label(bars, fontsize=8, padding=3)
    ax.set_xlabel("Number of decisions", fontsize=9)
    ax.invert_yaxis()
else:
    ax.text(0.5, 0.5, "All Protection-only", ha="center", va="center", transform=ax.transAxes)

plt.tight_layout(rect=[0, 0, 1, 0.96])
plt.savefig("pdpc_trends.png", dpi=150, bbox_inches="tight")
plt.close()
print("Saved pdpc_trends.png")

# ── Excel analysis workbook ───────────────────────────────────────────────────

wb_out = Workbook()

H_FONT  = Font(bold=True, color="FFFFFF", size=10)
H_FILL  = PatternFill("solid", fgColor="1F4E79")
S_FILL  = PatternFill("solid", fgColor="D6E4F0")
C_ALIGN = Alignment(horizontal="center", vertical="center")
W_ALIGN = Alignment(vertical="top", wrap_text=True)
BORDER  = Border(
    bottom=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
)

def hdr(ws, row, col, val, width=None):
    c = ws.cell(row, col, val)
    c.font, c.fill, c.alignment = H_FONT, H_FILL, C_ALIGN
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return c

def data(ws, row, col, val, fmt=None, bold=False, fill=None, align=C_ALIGN):
    c = ws.cell(row, col, val)
    c.alignment = align
    if fmt:    c.number_format = fmt
    if bold:   c.font = Font(bold=True)
    if fill:   c.fill = fill
    c.border = BORDER
    return c

# ── Sheet 1: Yearly Summary ───────────────────────────────────────────────────
ws1 = wb_out.active
ws1.title = "Yearly Summary"
ws1.freeze_panes = "A3"

hdr(ws1, 1, 1, "Yearly Summary — PDPC Protection Decisions", width=10)
ws1.merge_cells("A1:H1")
ws1.cell(1,1).font = Font(bold=True, size=13, color="1F4E79")
ws1.cell(1,1).fill = PatternFill("solid", fgColor="D6E4F0")
ws1.cell(1,1).alignment = C_ALIGN

cols = ["Year","Total Decisions","With Penalty","Without Penalty",
        "Total Penalties (SGD)","Avg Penalty (SGD)","Max Penalty (SGD)",
        "Protection-Only Cases"]
widths = [8,18,14,16,22,18,18,22]
for i,(h,w) in enumerate(zip(cols,widths),1):
    hdr(ws1, 2, i, h, width=w)

for i, y in enumerate(year_range, 3):
    yr_recs = [r for r in records if r["year"] == y]
    pen_recs = [r for r in yr_recs if r["has_penalty"]]
    amts = [r["penalty_amt"] for r in pen_recs]
    pure = sum(1 for r in yr_recs if r["pure_protection"])

    fill = S_FILL if i % 2 == 0 else None
    data(ws1, i, 1, y,                              bold=True, fill=fill)
    data(ws1, i, 2, len(yr_recs),                   fill=fill)
    data(ws1, i, 3, len(pen_recs),                  fill=fill)
    data(ws1, i, 4, len(yr_recs)-len(pen_recs),     fill=fill)
    data(ws1, i, 5, sum(amts) if amts else 0,       fmt='"$"#,##0', fill=fill)
    data(ws1, i, 6, int(np.mean(amts)) if amts else 0, fmt='"$"#,##0', fill=fill)
    data(ws1, i, 7, max(amts) if amts else 0,       fmt='"$"#,##0', fill=fill)
    data(ws1, i, 8, pure,                           fill=fill)

# Totals row
tot_row = len(year_range) + 3
all_pen = [r["penalty_amt"] for r in records if r["has_penalty"]]
data(ws1, tot_row, 1, "TOTAL", bold=True, fill=PatternFill("solid", fgColor="1F4E79"),
     align=C_ALIGN)
ws1.cell(tot_row, 1).font = Font(bold=True, color="FFFFFF")
data(ws1, tot_row, 2, len(records), bold=True)
data(ws1, tot_row, 3, len(all_pen), bold=True)
data(ws1, tot_row, 4, len(records)-len(all_pen), bold=True)
data(ws1, tot_row, 5, sum(all_pen), fmt='"$"#,##0', bold=True)
data(ws1, tot_row, 6, int(np.mean(all_pen)) if all_pen else 0, fmt='"$"#,##0', bold=True)
data(ws1, tot_row, 7, max(all_pen) if all_pen else 0, fmt='"$"#,##0', bold=True)
data(ws1, tot_row, 8, sum(1 for r in records if r["pure_protection"]), bold=True)

ws1.row_dimensions[1].height = 28
ws1.row_dimensions[2].height = 36

# ── Sheet 2: Penalty Distribution ────────────────────────────────────────────
ws2 = wb_out.create_sheet("Penalty Distribution")
ws2.freeze_panes = "A3"

hdr(ws2, 1, 1, "Penalty Amount Distribution")
ws2.merge_cells("A1:C1")
ws2.cell(1,1).font = Font(bold=True, size=13, color="1F4E79")
ws2.cell(1,1).fill = PatternFill("solid", fgColor="D6E4F0")
ws2.cell(1,1).alignment = C_ALIGN

for i,(h,w) in enumerate(zip(["Penalty Band","Count","% of Penalty Cases"],[22,10,20]),1):
    hdr(ws2, 2, i, h, width=w)

total_pen_cases = sum(band_counts.values())
for i, b in enumerate(BANDS, 3):
    cnt = band_counts.get(b, 0)
    fill = S_FILL if i % 2 == 0 else None
    data(ws2, i, 1, b,   fill=fill, align=Alignment(horizontal="left", vertical="center"))
    data(ws2, i, 2, cnt, fill=fill)
    data(ws2, i, 3, cnt/total_pen_cases if total_pen_cases else 0,
         fmt="0.0%", fill=fill)

# ── Sheet 3: Obligation Mix ───────────────────────────────────────────────────
ws3 = wb_out.create_sheet("Obligation Mix")
ws3.freeze_panes = "A3"

hdr(ws3, 1, 1, "Obligation Mix Analysis")
ws3.merge_cells("A1:D1")
ws3.cell(1,1).font = Font(bold=True, size=13, color="1F4E79")
ws3.cell(1,1).fill = PatternFill("solid", fgColor="D6E4F0")
ws3.cell(1,1).alignment = C_ALIGN

for i,(h,w) in enumerate(zip(["Obligations","Count","% of All Cases","Avg Penalty (SGD)"],[40,10,16,18]),1):
    hdr(ws3, 2, i, h, width=w)

oblig_groups = Counter(r["obligations"] for r in records)
total = len(records)
for i,(oblig,cnt) in enumerate(oblig_groups.most_common(), 3):
    pen_for_group = [r["penalty_amt"] for r in records if r["obligations"]==oblig and r["has_penalty"]]
    avg = int(np.mean(pen_for_group)) if pen_for_group else 0
    fill = S_FILL if i % 2 == 0 else None
    data(ws3, i, 1, oblig, fill=fill, align=Alignment(horizontal="left", vertical="center"))
    data(ws3, i, 2, cnt,   fill=fill)
    data(ws3, i, 3, cnt/total, fmt="0.0%", fill=fill)
    data(ws3, i, 4, avg,   fmt='"$"#,##0', fill=fill)

# ── Sheet 4: Decision Outcomes ────────────────────────────────────────────────
ws4 = wb_out.create_sheet("Decision Outcomes")
ws4.freeze_panes = "A3"

hdr(ws4, 1, 1, "Decision Outcome Summary")
ws4.merge_cells("A1:C1")
ws4.cell(1,1).font = Font(bold=True, size=13, color="1F4E79")
ws4.cell(1,1).fill = PatternFill("solid", fgColor="D6E4F0")
ws4.cell(1,1).alignment = C_ALIGN

for i,(h,w) in enumerate(zip(["Outcome","Count","% of All Cases"],[30,10,16]),1):
    hdr(ws4, 2, i, h, width=w)

for i, o in enumerate(OUTCOME_ORDER, 3):
    cnt = outcome_counts.get(o, 0)
    fill = S_FILL if i % 2 == 0 else None
    data(ws4, i, 1, o,   fill=fill, align=Alignment(horizontal="left", vertical="center"))
    data(ws4, i, 2, cnt, fill=fill)
    data(ws4, i, 3, cnt/total, fmt="0.0%", fill=fill)

# ── Sheet 5: Notable Cases ────────────────────────────────────────────────────
ws5 = wb_out.create_sheet("Notable Cases")

hdr(ws5, 1, 1, "Top 10 Largest Financial Penalties")
ws5.merge_cells("A1:D1")
ws5.cell(1,1).font = Font(bold=True, size=13, color="1F4E79")
ws5.cell(1,1).fill = PatternFill("solid", fgColor="D6E4F0")
ws5.cell(1,1).alignment = C_ALIGN

for i,(h,w) in enumerate(zip(["Case Name","Date","Penalty (SGD)","Citation"],[55,12,16,20]),1):
    hdr(ws5, 2, i, h, width=w)

top10 = sorted([r for r in records if r["has_penalty"]], key=lambda r: -r["penalty_amt"])[:10]
for i, r in enumerate(top10, 3):
    fill = S_FILL if i % 2 == 0 else None
    data(ws5, i, 1, r["name"],        fill=fill, align=Alignment(horizontal="left", vertical="center", wrap_text=True))
    data(ws5, i, 2, r["date_str"],    fill=fill)
    data(ws5, i, 3, r["penalty_amt"], fmt='"$"#,##0', bold=True, fill=fill)
    data(ws5, i, 4, r["citation"],    fill=fill)
    ws5.row_dimensions[i].height = 30

wb_out.save("pdpc_analysis.xlsx")
print("Saved pdpc_analysis.xlsx")

# ── print text summary ────────────────────────────────────────────────────────
sep = "=" * 60
print()
print(sep)
print("  PDPC Protection Decisions -- Key Stats")
print(sep)
print(f"  Total decisions analysed : {len(records)}")
print(f"  Date range               : {min(years)} - {max(years)}")
print(f"  With financial penalty   : {len(all_pen)} ({len(all_pen)/len(records)*100:.0f}%)")
print(f"  Without penalty          : {len(records)-len(all_pen)} ({(len(records)-len(all_pen))/len(records)*100:.0f}%)")
print(f"  Total penalties imposed  : ${sum(all_pen):,.0f}")
print(f"  Average penalty          : ${int(np.mean(all_pen)):,}")
print(f"  Median penalty           : ${int(np.median(all_pen)):,}")
print(f"  Largest single penalty   : ${max(all_pen):,}")
print(f"  Protection-only cases    : {sum(1 for r in records if r['pure_protection'])}")
print(f"  Multi-obligation cases   : {sum(1 for r in records if not r['pure_protection'])}")
print()
print("  Top co-occurring obligations:")
for o, c in co_oblig_counter.most_common(5):
    print(f"    {o:<30} {c} cases")
print(sep)
