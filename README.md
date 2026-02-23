# Idealz Google Reviews — Scraper & Analyzer
### Automated review collection, analysis, and weekly reporting for 3 Idealz stores

---

## Overview

This toolkit automatically collects Google Maps reviews for 3 Idealz stores in Colombo,
analyzes sentiment and categories, and produces professional Excel reports — either as a
**full historical deep-dive** or a **weekly tracking report** with week-on-week comparisons.

| Store | Location | Google Rating (Feb 2026) |
|---|---|---|
| Idealz Prime | 86 Galle Road, Colombo 4 | ⭐ 4.8 / 513 reviews |
| Idealz Lanka - Marino Mall | Galle Road, Colombo 3 | ⭐ 4.9 / 1,472 reviews |
| Idealz Lanka - Liberty Plaza | Liberty Plaza, Colombo 3 | ⭐ 4.5 / 1,881 reviews |

---

## Files in This Project

```
review/
│
├── idealz_weekly_tracker.py       ← ✅ Run every week
├── idealz_scraper_v4.py           ← Full scrape (run once a month)
├── idealz_analyzer_v2.py          ← Full analysis (run after full scrape)
├── README.md                      ← This file
│
├── snapshots/                     ← Auto-created. DO NOT DELETE.
│   ├── reviews_2026-02-23.json
│   ├── reviews_2026-03-02.json
│   └── ...
│
├── idealz_raw_reviews.json        ← Created by full scraper
├── weekly_report_2026-03-02.xlsx  ← Created by weekly tracker
└── idealz_reviews_report.xlsx     ← Created by full analyzer
```

---

## Setup (Run Once)

### 1. Install Python packages
```bash
pip install playwright textblob openpyxl pandas
```

### 2. Install the browser
```bash
playwright install chromium
```

### 3. Download TextBlob language data
```bash
python -m textblob.download_corpora
```

### 4. Add your Google Maps URLs

Open **each script** and paste the direct Google Maps `/maps/place/` URL for each store.

**How to get the URL:**
1. Go to [maps.google.com](https://maps.google.com)
2. Search for the store (e.g. "iDealz Prime Colombo")
3. Click the correct listing
4. Copy the full URL from your browser address bar
5. It should look like: `https://www.google.com/maps/place/iDealz+Prime/@6.88...`

Paste into the `STORES` list at the top of each script:
```python
STORES = [
    {"name": "Idealz Prime",                "url": "PASTE URL HERE", "expected": 513},
    {"name": "Idealz Lanka - Marino Mall",  "url": "PASTE URL HERE", "expected": 1472},
    {"name": "Idealz Lanka - Liberty Plaza","url": "PASTE URL HERE", "expected": 1881},
]
```

---

## Usage

### ✅ Option A — Weekly Tracking (Every Monday)

```bash
python idealz_weekly_tracker.py
```

Scrapes the newest 150 reviews per store (~5 minutes), compares with last week,
and saves a report named `weekly_report_YYYY-MM-DD.xlsx`.

**Weekly report sheets:**

| Sheet | Contents |
|---|---|
| 📊 Weekly Summary | KPI cards with ▲▼ change arrows vs last week |
| 🆕 New This Week | Every new review since last run, negatives sorted to top |
| 📈 Weekly Trends | Cumulative avg rating chart across all weeks |
| 🚨 Alerts | New 1–2 star reviews that need a response |

> **Week 1:** Saves a baseline only — no comparison yet.
> **Week 2+:** Full comparison report with change arrows.

---

### 📅 Option B — Full Historical Analysis (Once a Month)

**Step 1 — Scrape ALL reviews (~20–30 mins):**
```bash
python idealz_scraper_v4.py
```
Saves `idealz_raw_reviews.json`.

**Step 2 — Run full analysis (~2 mins):**
```bash
python idealz_analyzer_v2.py
```
Produces `idealz_reviews_report.xlsx`.

**Full report sheets:**

| Sheet | Contents |
|---|---|
| 📊 Dashboard | KPI cards + store comparison table |
| 📋 All Reviews | Every review with sentiment + category tags |
| ⭐ Ratings | Star distribution charts per store |
| 😊 Sentiment | Positive / Negative / Neutral breakdown + pie chart |
| 📂 Categories | Topics mentioned (service, product, pricing, etc.) |
| 📈 Trends | Yearly summary + monthly detail table + 2 line charts |
| 💬 Notable Reviews | Best, worst, and most-liked reviews |
| 🔤 Top Words | Most frequent words by sentiment and by store |

---

## Recommended Routine

| When | Command | Time |
|---|---|---|
| Every Monday | `python idealz_weekly_tracker.py` | ~5 mins |
| First of each month | `python idealz_scraper_v4.py` then `python idealz_analyzer_v2.py` | ~30 mins |

---

## How the Analysis Works

### Sentiment (TextBlob NLP)
Every review text is automatically scored:

| Label | Condition |
|---|---|
| Positive | Score > 0.1 |
| Neutral | Score between −0.1 and 0.1 |
| Negative | Score < −0.1 |
| No Text | Stars-only review — no text to analyze |

### Category Tagging (Keyword Matching)
Each review is tagged based on words it contains.
A single review can match multiple categories.

| Category | Example keywords |
|---|---|
| Product Quality | quality, defect, broken, genuine, damaged |
| Customer Service | staff, helpful, rude, friendly, attitude |
| Pricing & Value | price, expensive, affordable, value, worth |
| Store Experience | clean, organized, atmosphere, mall, crowded |
| Promotions / Deals | lucky draw, discount, voucher, cashback |
| After-Sales | refund, exchange, warranty, complaint |
| Variety / Stock | variety, selection, out of stock, range |
| Delivery | shipping, courier, delayed, tracking |

### Date Parsing
Google Maps shows relative timestamps like "3 months ago".
The script converts these to calendar months (e.g. `2025-11`) using today's date as anchor.
Handles: "just now", "yesterday", "a week ago", "2 months ago", "a year ago", etc.

---

## Understanding the Weekly Report

### ▲ ▼ Arrows in the Summary sheet

| Symbol | Meaning |
|---|---|
| ▲ +0.2 (green) | Metric went UP — generally good |
| ▼ −0.1 (red) | Metric went DOWN — investigate |
| — | No change vs last week |

> Note: For **% Negative**, a ▼ (down) is shown in green because fewer negatives is better.

### 🚨 Alerts Sheet
Shows only **new** 1–2 star reviews from this week.
Reviews without an owner reply are flagged with ❌.
This is the most important sheet to check first each week.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Only 10 reviews loading | Use the latest `v4` scraper and updated weekly tracker — old versions had a scroll bug |
| `PermissionError` on save | Close the `.xlsx` file in Excel before running the script |
| Browser opens but 0 reviews found | Paste a direct `/maps/place/` URL instead of a search URL |
| `playwright not installed` | Run `pip install playwright && playwright install chromium` |
| Review count lower than Google shows | Normal — Google hides spam/deleted reviews. Expect 85–95% capture rate |
| `textblob` corpus error | Run `python -m textblob.download_corpora` |
| Wrong store clicked | Paste direct Maps URL — search URLs sometimes pick the wrong listing |

---

## Important Rules

- ❌ **Never delete the `snapshots/` folder** — this is your entire weekly history
- ✅ The `snapshots/` folder must stay in the same directory as `idealz_weekly_tracker.py`
- ✅ Weekly report files (`weekly_report_*.xlsx`) are safe to move, copy, or archive anywhere
- ✅ Close Excel before running any script to avoid `PermissionError`
- ℹ️ Google's displayed review count includes deleted/hidden reviews — scraped count will always be slightly lower
- ℹ️ All tools are completely free — no API keys or paid services required

---

## Technical Requirements

| Package | Install command |
|---|---|
| Python 3.9+ | [python.org](https://python.org) |
| playwright | `pip install playwright` + `playwright install chromium` |
| textblob | `pip install textblob` |
| openpyxl | `pip install openpyxl` |
| pandas | `pip install pandas` |

---

## Project Info

Built for internal review monitoring of Idealz stores in Sri Lanka.
All data collected from publicly visible Google Maps reviews.
No login, authentication, or private data is accessed.
Last updated: February 2026.
