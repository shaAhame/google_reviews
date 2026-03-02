"""
╔══════════════════════════════════════════════════════════════════════╗
║       IDEALZ — WEEKLY REVIEW TRACKER                                 ║
║       Run every week to capture new reviews & compare vs last week   ║
╚══════════════════════════════════════════════════════════════════════╝

HOW IT WORKS:
  - Each run scrapes the latest reviews (sorted Newest first)
  - Saves a snapshot:  snapshots/reviews_YYYY-MM-DD.json
  - Compares with the previous week's snapshot
  - Outputs:           weekly_report_YYYY-MM-DD.xlsx

RUN EVERY MONDAY (or any fixed day):
    python idealz_weekly_tracker.py

FIRST RUN: Just saves a baseline — no comparison yet.
SECOND RUN (next week): Produces the first weekly comparison report.
"""

import json, re, sys
from pathlib import Path
from datetime import datetime, timedelta
from collections import Counter

import pandas as pd
from textblob import TextBlob
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

# ── CONFIG ────────────────────────────────────────────────────────────────────
# ── CONFIG ────────────────────────────────────────────────────────────────────
STORES = [
    {"name": "Idealz Prime","url": "https://www.google.com/maps/place/iDealz+Prime/@6.8912695,79.8560961,17z/data=!3m1!4b1!4m6!3m5!1s0x3ae259005a2260c1:0xd6febd8ffeac3a34!8m2!3d6.8912695!4d79.8560961!16s%2Fg%2F11w27bncwk?entry=ttu&g_ep=EgoyMDI2MDIxOC4wIKXMDSoASAFQAw%3D%3D", "expected": 513},
    {"name": "Idealz Lanka - Marino Mall","url": "https://www.google.com/maps/place/iDealz+Lanka+-+Marino+Mall/@6.9001796,79.8523305,17z/data=!3m1!4b1!4m6!3m5!1s0x3ae25957ebf8012b:0xe0e160f3a83edd3c!8m2!3d6.9001796!4d79.8523305!16s%2Fg%2F11gr41k7q8?entry=ttu&g_ep=EgoyMDI2MDIxOC4wIKXMDSoASAFQAw%3D%3D", "expected": 1472},
    {"name": "Idealz Lanka - Liberty Plaza","url": "https://www.google.com/maps/place/iDealz+Lanka+Pvt+Ltd/@6.9116839,79.8515051,17z/data=!3m1!4b1!4m6!3m5!1s0x3ae25911b0316acb:0xdd0d30f303baddf1!8m2!3d6.9116839!4d79.8515051!16s%2Fg%2F11b6dg62r6?entry=ttu&g_ep=EgoyMDI2MDIxOC4wIKXMDSoASAFQAw%3D%3D", "expected": 1881},
]
# Maximum reviews to scrape as a safety fallback (prevents infinite loops)
MAX_SAFETY_LIMIT = 500
# Target age: We stop when we see reviews older than 1 week
STOP_AGE_MARKERS = ["a week ago", "2 weeks ago", "3 weeks ago", "4 weeks ago", "month", "year"]
SNAPSHOT_DIR   = Path("snapshots")
TODAY          = datetime.today().strftime("%Y-%m-%d")
OUTPUT_XLSX    = f"weekly_report_{TODAY}.xlsx"

# ── COLOURS ───────────────────────────────────────────────────────────────────
C = {
    "navy": "1F3864", "blue": "2E75B6", "dkblue": "1A4F8A",
    "green": "1E5631", "lgreen": "C6EFCE", "red": "9C0006",
    "lred": "FFC7CE", "lyellow": "FFEB9C", "lgray": "F2F2F2",
    "altrow": "EBF3FB", "white": "FFFFFF", "border": "BFBFBF",
    "up": "C6EFCE", "down": "FFC7CE", "same": "FFEB9C",
}
SENT_CLR = {"Positive": C["lgreen"], "Negative": C["lred"],
            "Neutral": C["lyellow"], "No Text": C["lgray"]}
STAR_CLR = {5: "C6EFCE", 4: "DDEBF7", 3: "FFF2CC", 2: "FCE4D6", 1: "FFC7CE"}
STORE_CLR = ["1F4E79", "375623", "7B2C2C"]

# ── STYLE HELPERS ─────────────────────────────────────────────────────────────
def _b():
    s = Side(style="thin", color=C["border"])
    return Border(left=s, right=s, top=s, bottom=s)

def borders(ws, r1, r2, c1, c2):
    b = _b()
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            cell.border = b

def hdr(cell, text, bg=None, fg="FFFFFF", sz=11, bold=True):
    cell.value = text
    cell.font  = Font(name="Arial", bold=bold, size=sz, color=fg)
    cell.fill  = PatternFill("solid", start_color=bg or C["navy"])
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def dat(cell, value, center=False, sz=10, fg="000000", bold=False, indent=0):
    cell.value = value
    cell.font  = Font(name="Arial", size=sz, color=fg, bold=bold)
    cell.alignment = Alignment(horizontal="center" if center else "left",
                               vertical="center", indent=indent, wrap_text=True)

def alt(ws, row, ncols, c1=1):
    if row % 2 == 0:
        for c in range(c1, c1 + ncols):
            ws.cell(row, c).fill = PatternFill("solid", start_color=C["altrow"])

def banner(ws, text, cols="J", row=1, sz=15, bg=None):
    ws.merge_cells(f"A{row}:{cols}{row}")
    hdr(ws[f"A{row}"], text, bg=bg or C["navy"], sz=sz)
    ws.row_dimensions[row].height = 40

def arrow(val):
    """Return ▲ / ▼ / — with colour for change cells."""
    if val is None or val == 0:  return "—",  C["same"]
    if val > 0:                  return f"▲ +{val}", C["up"]
    return f"▼ {val}",           C["down"]

# ── SCRAPER (Time-Based — Newest only) ───────────────────────────────────────

def is_older_than_one_week(date_str):
    """Return True if the date string indicates a review older than ~7-10 days."""
    if not date_str: return False
    ds = date_str.lower()
    # If it says "2 weeks ago", "3 weeks ago", "a month ago", "a year ago" etc.
    return any(marker in ds for marker in STOP_AGE_MARKERS)

def scrape_newest(safety_limit=MAX_SAFETY_LIMIT):
    """Scrape only the most recent reviews for each store."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Run: pip install playwright && playwright install chromium")
        sys.exit(1)

    all_reviews = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False, slow_mo=30,
            args=["--disable-blink-features=AutomationControlled",
                  "--no-sandbox", "--start-maximized"]
        )
        ctx = browser.new_context(
            viewport={"width": 1400, "height": 900},
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 Chrome/122.0.0.0 Safari/537.36"),
            locale="si-LK", timezone_id="Asia/Colombo",
        )
        page = ctx.new_page()

        # Dismiss cookie popup
        page.goto("https://www.google.com/maps", wait_until="domcontentloaded")
        page.wait_for_timeout(2000)
        for sel in ['button[aria-label*="Accept"]', 'button:has-text("Accept all")']:
            try:
                btn = page.query_selector(sel)
                if btn: btn.click(); page.wait_for_timeout(800); break
            except: pass

        for store in STORES:
            print(f"\n  Scraping (searching for brand new reviews): {store['name']}")
            try:
                reviews = scrape_store_newest(page, store, safety_limit)
                print(f"  [OK] Collected {len(reviews)} reviews (stopped at older date boundary)")
                all_reviews.extend(reviews)
            except Exception as e:
                print(f"  X Failed: {e}")
            page.wait_for_timeout(3000)

        browser.close()

    return all_reviews


def scrape_store_newest(page, store, limit):
    url = store["url"] + ("&" if "?" in store["url"] else "?") + "hl=en"
    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(3000)

    # Click first result if search page
    if "/maps/place/" not in page.url:
        for sel in ['a[href*="/maps/place/"]', '.hfpxzc']:
            el = page.query_selector(sel)
            if el:
                el.click()
                page.wait_for_load_state("domcontentloaded")
                page.wait_for_timeout(4000)
                break

    if "/maps/place/" not in page.url:
        raise Exception("Not on a place page")

    # Click Reviews tab
    for sel in ['button[aria-label*="Reviews"]', 'button[role="tab"]:has-text("Reviews")']:
        el = page.query_selector(sel)
        if el:
            el.click()
            page.wait_for_timeout(3000)
            break

    # Sort by Newest
    for sel in ['button[aria-label^="Sort"]', 'button[data-value="Sort"]']:
        btn = page.query_selector(sel)
        if btn:
            btn.click()
            page.wait_for_timeout(1500)
            for opt in ['div[data-index="1"]', 'li:has-text("Newest")']:
                o = page.query_selector(opt)
                if o:
                    o.click()
                    page.wait_for_timeout(2500)
                    break
            break

    # Wait for review cards
    try:
        page.wait_for_selector('[data-review-id]', timeout=12000)
    except:
        return []

    # Scroll to load `limit` reviews
    seen_ids = set()
    stall = 0
    # Give the page an extra moment to settle after sort
    page.wait_for_timeout(2000)

    while stall < 20:
        # Check current dates in DOM to see if we can stop
        dates = page.evaluate("""() =>
            Array.from(document.querySelectorAll('.rsqaWe, .dehysf, .xRkPPb'))
                 .map(el => el.innerText.trim()).filter(Boolean)
        """)

        # If any date in the current view is older than 1 week, we hit our target
        found_old = any(is_older_than_one_week(d) for d in dates)

        if found_old:
            print(f"\n    [OK] Reached older review boundary (found reviews older than 1 week).")
            break

        # Scroll FIRST, then check
        page.evaluate("""() => {
            const selectors = [
                'div[role="feed"]',
                '.m6QErb.DxyBCb',
                '.m6QErb[aria-label]',
                '.m6QErb',
            ];
            for (const s of selectors) {
                const el = document.querySelector(s);
                if (el && el.scrollHeight > el.clientHeight) {
                    el.scrollTop += 5000;
                    return;
                }
            }
            // fallback
            document.querySelectorAll('div').forEach(el => {
                const r = el.getBoundingClientRect();
                if (r.left < 600 && el.scrollHeight - el.clientHeight > 200)
                    el.scrollTop += 5000;
            });
        }""")
        page.wait_for_timeout(2500)

        ids = set(page.evaluate("""() =>
            Array.from(document.querySelectorAll('[data-review-id]'))
                 .map(el => el.getAttribute('data-review-id')).filter(Boolean)
        """))
        new_ids = ids - seen_ids
        seen_ids = ids
        print(f"    {len(seen_ids)} / {limit} unique reviews loaded...", end="\r", flush=True)

        if len(seen_ids) >= limit:
            print(f"\n    [OK] Target reached ({limit}).")
            break

        if not new_ids:
            stall += 1
        else:
            stall = 0

    if stall >= 20:
        print(f"\n    [OK] No more reviews loading. Final: {len(seen_ids)}")
    print(f"    {len(seen_ids)} unique IDs found")

    # Expand truncated text
    for btn in page.query_selector_all('button.w8nwRe, button[aria-label="See more"]'):
        try: btn.click(); page.wait_for_timeout(60)
        except: pass

    # Extract
    cards = page.query_selector_all('[data-review-id]')
    reviews = []
    extracted_ids = set()
    for card in cards:
        try:
            rid = card.get_attribute("data-review-id") or ""
            if rid in extracted_ids: continue
            extracted_ids.add(rid)

            author = "Anonymous"
            for s in ['.d4r55', '.DUwDvf']:
                el = card.query_selector(s)
                if el and el.inner_text().strip():
                    author = el.inner_text().strip(); break

            rating = None
            for s in ['[aria-label*="star"]', '.kvMYJc']:
                el = card.query_selector(s)
                if el:
                    m = re.search(r'(\d+)', el.get_attribute("aria-label") or "")
                    if m: rating = int(m.group(1)); break

            text = ""
            for s in ['.wiI7pd', '.MyEned']:
                el = card.query_selector(s)
                if el and el.inner_text().strip():
                    text = el.inner_text().strip(); break

            date_str = ""
            for s in ['.rsqaWe', '.dehysf', '.xRkPPb']:
                el = card.query_selector(s)
                if el and el.inner_text().strip():
                    date_str = el.inner_text().strip(); break

            owner_reply = ""
            el = card.query_selector('.CDe7pd')
            if el: owner_reply = el.inner_text().strip()

            reviews.append({
                "review_id": rid, "store": store["name"],
                "author": author, "rating": rating,
                "date_raw": date_str, "text": text,
                "owner_reply": owner_reply,
                "scraped_on": TODAY,
            })
        except: continue

    # Filter to strictly include only reviews within the target age
    filtered_reviews = [r for r in reviews if not is_older_than_one_week(r.get("date_raw"))]
    
    return filtered_reviews[:limit]

# ── SNAPSHOT MANAGEMENT ───────────────────────────────────────────────────────

def save_snapshot(reviews):
    SNAPSHOT_DIR.mkdir(exist_ok=True)
    path = SNAPSHOT_DIR / f"reviews_{TODAY}.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(reviews, f, ensure_ascii=False, indent=2)
    print(f"  [OK] Snapshot saved: {path}")
    return path

def load_previous_snapshot():
    """Load the most recent snapshot BEFORE today."""
    if not SNAPSHOT_DIR.exists():
        return None, None
    files = sorted(SNAPSHOT_DIR.glob("reviews_*.json"))
    # Exclude today's file
    prev_files = [f for f in files if f.stem != f"reviews_{TODAY}"]
    if not prev_files:
        return None, None
    latest = prev_files[-1]
    date_str = latest.stem.replace("reviews_", "")
    with open(latest, "r", encoding="utf-8") as f:
        data = json.load(f)
    print(f"  [OK] Previous snapshot loaded: {latest} ({len(data)} reviews)")
    return data, date_str

def load_all_snapshots():
    """Load ALL snapshots for cumulative trend chart."""
    if not SNAPSHOT_DIR.exists():
        return {}
    result = {}
    for f in sorted(SNAPSHOT_DIR.glob("reviews_*.json")):
        date_str = f.stem.replace("reviews_", "")
        with open(f, "r", encoding="utf-8") as fh:
            result[date_str] = json.load(fh)
    return result

# ── ANALYSIS HELPERS ──────────────────────────────────────────────────────────

CATEGORIES = {
    "Product Quality":    ["quality","product","item","defect","broken","material",
                           "durable","genuine","fake","original","condition","damaged"],
    "Customer Service":   ["service","staff","employee","helpful","rude","assist",
                           "support","polite","attitude","friendly","ignored",
                           "customer care","respond","behavior","behaviour"],
    "Pricing & Value":    ["price","pricing","expensive","cheap","affordable","value",
                           "worth","overpriced","discount","offer","deal","cost","money"],
    "Store Experience":   ["store","shop","location","parking","clean","organized",
                           "display","ambiance","atmosphere","interior","mall","crowded",
                           "visit","branch","outlet"],
    "Promotions / Deals": ["promotion","sale","offer","discount","coupon","voucher",
                           "cashback","deal","lucky draw","win","prize","raffle"],
    "After-Sales":        ["return","refund","exchange","warranty","replace",
                           "complaint","issue","resolve","compensation","claim"],
    "Variety / Stock":    ["variety","selection","stock","available","range",
                           "choice","collection","assortment","limited","out of stock"],
}

def get_sentiment(text):
    if not text or not str(text).strip(): return "No Text", 0.0
    score = TextBlob(text).sentiment.polarity
    label = "Positive" if score > 0.1 else ("Negative" if score < -0.1 else "Neutral")
    return label, round(score, 3)

def categorize(text):
    if not text or not isinstance(text, str): return "General"
    tl = text.lower()
    matched = [cat for cat, kws in CATEGORIES.items() if any(kw in tl for kw in kws)]
    return ", ".join(matched) if matched else "General"

def enrich(reviews):
    """Add sentiment + category to a list of review dicts."""
    for r in reviews:
        sl, ss = get_sentiment(r.get("text", ""))
        r["sentiment"]   = sl
        r["sent_score"]  = ss
        r["categories"]  = categorize(r.get("text", ""))
    return reviews

def store_stats(reviews):
    """Return a dict of key stats for a list of reviews."""
    if not reviews:
        return {"count": 0, "avg_rating": None, "pos_pct": 0,
                "neg_pct": 0, "neutral_pct": 0, "reply_pct": 0,
                "star_dist": {}}
    df = pd.DataFrame(reviews)
    df["rating"] = pd.to_numeric(df.get("rating", pd.Series()), errors="coerce")
    cnt  = len(df)
    avg  = df["rating"].mean()
    pos  = 100*(df["sentiment"]=="Positive").sum()/max(cnt,1)
    neg  = 100*(df["sentiment"]=="Negative").sum()/max(cnt,1)
    neu  = 100*(df["sentiment"]=="Neutral").sum()/max(cnt,1)
    rep  = 100*(df["owner_reply"].str.strip().astype(bool)).sum()/max(cnt,1)
    sdist = {s: int((df["rating"]==s).sum()) for s in [5,4,3,2,1]}
    return {"count": cnt, "avg_rating": round(avg,2) if pd.notna(avg) else None,
            "pos_pct": round(pos,1), "neg_pct": round(neg,1),
            "neutral_pct": round(neu,1), "reply_pct": round(rep,1),
            "star_dist": sdist}

def find_new_reviews(current, previous):
    """Return reviews in current that are NOT in previous (by review_id)."""
    prev_ids = {r.get("review_id","") for r in (previous or [])}
    return [r for r in current if r.get("review_id","") not in prev_ids]

# ── EXCEL REPORT ──────────────────────────────────────────────────────────────

def write_weekly_report(current, previous, prev_date, all_snapshots):
    wb = Workbook()
    wb.remove(wb.active)

    enrich(current)
    if previous:
        enrich(previous)

    new_reviews = find_new_reviews(current, previous)
    enrich(new_reviews)

    print(f"  New reviews since {prev_date}: {len(new_reviews)}")

    write_weekly_summary(wb, current, previous, new_reviews, prev_date, all_snapshots)
    write_new_reviews_sheet(wb, new_reviews)
    write_weekly_trends(wb, all_snapshots)
    write_alerts(wb, new_reviews)

    wb.save(OUTPUT_XLSX)
    print(f"  [OK] Report saved: {OUTPUT_XLSX}")


def write_weekly_summary(wb, current, previous, new_reviews, prev_date, all_snapshots):
    ws = wb.create_sheet("📊 Weekly Summary")
    ws.sheet_view.showGridLines = False

    banner(ws, f"IDEALZ — WEEKLY REVIEW REPORT   |   Week ending: {TODAY}",
           cols="L", sz=15)
    # Subtitle
    ws.merge_cells("A2:L2")
    ws["A2"].value = (f"Compared with: {prev_date or 'N/A (first run)'}   |   "
                      f"New reviews this week: {len(new_reviews)}   |   "
                      f"Total in snapshot: {len(current)}")
    ws["A2"].font  = Font(name="Arial", size=10, italic=True, color="555555")
    ws["A2"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 18

    stores = [s["name"] for s in STORES]

    # ── PER-STORE KPI CHANGE CARDS ──
    ROW = 4
    for si, sname in enumerate(stores):
        curr_s = [r for r in current     if r["store"] == sname]
        prev_s = [r for r in (previous or []) if r["store"] == sname]
        new_s  = [r for r in new_reviews if r["store"] == sname]
        cs = store_stats(curr_s)
        ps = store_stats(prev_s)
        sc = STORE_CLR[si % len(STORE_CLR)]
        c1 = 1 + si * 4

        ws.merge_cells(start_row=ROW, start_column=c1, end_row=ROW, end_column=c1+3)
        hdr(ws.cell(ROW, c1), sname, bg=sc, sz=10)
        ws.row_dimensions[ROW].height = 24

        # New reviews badge
        ws.merge_cells(start_row=ROW+1, start_column=c1, end_row=ROW+1, end_column=c1+3)
        badge = ws.cell(ROW+1, c1)
        badge.value = f"[NEW]  {len(new_s)} new review{'s' if len(new_s)!=1 else ''} this week"
        badge.font  = Font(name="Arial", size=11, bold=True,
                           color="1E5631" if new_s else "888888")
        badge.fill  = PatternFill("solid", start_color="E8F5E9" if new_s else "F5F5F5")
        badge.alignment = Alignment(horizontal="center", vertical="center")
        apply_border_single(ws, ROW+1, c1, c1+3)
        ws.row_dimensions[ROW+1].height = 26

        kpis = [
            ("Avg Rating",  cs["avg_rating"], ps["avg_rating"], True),
            ("% Positive",  cs["pos_pct"],    ps["pos_pct"],    True),
            ("% Negative",  cs["neg_pct"],    ps["neg_pct"],    False),  # lower is better
            ("Owner Reply %", cs["reply_pct"], ps["reply_pct"], True),
        ]
        for ki, (label, curr_v, prev_v, higher_is_better) in enumerate(kpis):
            r = ROW + 2 + ki
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c1+1)
            lc = ws.cell(r, c1)
            lc.value = label
            lc.font  = Font(name="Arial", size=9, color="444444")
            lc.fill  = PatternFill("solid", start_color="F5F8FC")
            lc.alignment = Alignment(horizontal="left", indent=1, vertical="center")

            vc = ws.cell(r, c1+2)
            vc.value = f"{curr_v}" if curr_v is not None else "N/A"
            vc.font  = Font(name="Arial", size=11, bold=True, color="1A1A1A")
            vc.fill  = PatternFill("solid", start_color="F5F8FC")
            vc.alignment = Alignment(horizontal="center", vertical="center")

            # Change arrow
            ac = ws.cell(r, c1+3)
            if prev_v is not None and curr_v is not None:
                diff = round(curr_v - prev_v, 2)
                arrow_txt, arrow_bg = arrow(diff)
                # Flip colour if lower is better (e.g. negative %)
                if not higher_is_better and diff != 0:
                    arrow_bg = C["up"] if diff < 0 else C["down"]
                ac.value = arrow_txt
                ac.fill  = PatternFill("solid", start_color=arrow_bg)
            else:
                ac.value = "—"
                ac.fill  = PatternFill("solid", start_color=C["same"])
            ac.font  = Font(name="Arial", size=9, bold=True)
            ac.alignment = Alignment(horizontal="center", vertical="center")
            apply_border_single(ws, r, c1, c1+3)

    # ── SUMMARY TABLE ──
    ROW = 14
    ws.merge_cells(f"A{ROW}:L{ROW}")
    hdr(ws.cell(ROW, 1), "WEEK-ON-WEEK COMPARISON TABLE", bg=C["dkblue"], sz=11)
    ws.row_dimensions[ROW].height = 26
    ROW += 1

    th = ["Store", "New\nReviews", "Total\nSnapshot",
          "Avg ★\n(this wk)", "Avg ★\n(prev wk)", "Δ Rating",
          "Pos %\n(this wk)", "Pos %\n(prev wk)", "Δ Pos %",
          "Neg %\n(this wk)", "Neg %\n(prev wk)", "Δ Neg %"]
    for ci, h in enumerate(th, 1):
        hdr(ws.cell(ROW, ci), h, bg="365F91", sz=9)
    ws.row_dimensions[ROW].height = 32
    tbl_s = ROW; ROW += 1

    for sname in stores:
        curr_s = [r for r in current      if r["store"] == sname]
        prev_s = [r for r in (previous or []) if r["store"] == sname]
        new_s  = [r for r in new_reviews  if r["store"] == sname]
        cs = store_stats(curr_s)
        ps = store_stats(prev_s)

        d_rat = round(cs["avg_rating"] - ps["avg_rating"], 2) \
                if cs["avg_rating"] and ps["avg_rating"] else None
        d_pos = round(cs["pos_pct"] - ps["pos_pct"], 1) if ps["pos_pct"] else None
        d_neg = round(cs["neg_pct"] - ps["neg_pct"], 1) if ps["neg_pct"] else None

        vals = [sname, len(new_s), len(curr_s),
                cs["avg_rating"] or "N/A", ps["avg_rating"] or "N/A",
                d_rat, cs["pos_pct"], ps["pos_pct"], d_pos,
                cs["neg_pct"], ps["neg_pct"], d_neg]

        for ci, v in enumerate(vals, 1):
            cell = ws.cell(ROW, ci)
            dat(cell, v, center=(ci > 1))
            alt(ws, ROW, len(vals))

            # Colour change columns
            if ci in [6, 9] and v is not None and v != "N/A":
                _, bg = arrow(v)
                cell.fill = PatternFill("solid", start_color=bg)
            if ci == 12 and v is not None and v != "N/A":
                _, bg = arrow(-v)  # negative is good for neg%
                cell.fill = PatternFill("solid", start_color=bg)
        ROW += 1

    borders(ws, tbl_s, ROW-1, 1, 12)
    from openpyxl.utils import get_column_letter
    for ci, w in enumerate([26,10,10,13,13,10,12,12,10,12,12,10], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w


def apply_border_single(ws, row, c1, c2):
    b = _b()
    for c in range(c1, c2+1):
        ws.cell(row, c).border = b


def write_new_reviews_sheet(wb, new_reviews):
    ws = wb.create_sheet("🆕 New This Week")
    ws.sheet_view.showGridLines = False
    banner(ws, f"NEW REVIEWS THIS WEEK  ({len(new_reviews)} reviews)", cols="H")

    if not new_reviews:
        ws.merge_cells("A3:H3")
        ws["A3"].value = "No new reviews detected since last snapshot."
        ws["A3"].font  = Font(name="Arial", size=11, italic=True, color="888888")
        ws["A3"].alignment = Alignment(horizontal="center")
        return

    hdrs = ["Store", "★", "Sentiment", "Review Text", "Date", "Author", "Category", "Owner Reply?"]
    widths = [24, 6, 12, 65, 16, 18, 30, 13]
    for ci, (h, w) in enumerate(zip(hdrs, widths), 1):
        hdr(ws.cell(2, ci), h, bg=C["navy"], sz=10)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[2].height = 24

    # Sort: negatives first so they're easy to spot
    sorted_reviews = sorted(new_reviews,
                            key=lambda r: (r.get("sentiment","") != "Negative",
                                           -(r.get("rating") or 3)))

    for ri, r in enumerate(sorted_reviews, 3):
        vals = [r.get("store",""), r.get("rating",""),
                r.get("sentiment",""), r.get("text","")[:500],
                r.get("date_raw",""), r.get("author",""),
                r.get("categories",""),
                "Yes" if r.get("owner_reply","").strip() else "No"]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(ri, ci)
            cell.value = v
            cell.font  = Font(name="Arial", size=9)
            cell.alignment = Alignment(vertical="top", wrap_text=(ci==4))

        sent = r.get("sentiment","")
        if sent in SENT_CLR:
            ws.cell(ri, 3).fill = PatternFill("solid", start_color=SENT_CLR[sent])

        rat = r.get("rating")
        if rat and int(rat) in STAR_CLR:
            ws.cell(ri, 2).fill = PatternFill("solid", start_color=STAR_CLR[int(rat)])
            ws.cell(ri, 2).alignment = Alignment(horizontal="center", vertical="top")

        alt(ws, ri, len(hdrs))

    borders(ws, 2, len(sorted_reviews)+2, 1, len(hdrs))
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:H2"


def write_weekly_trends(wb, all_snapshots):
    """Cumulative weekly trend — avg rating and new review count per week."""
    ws = wb.create_sheet("📈 Weekly Trends")
    ws.sheet_view.showGridLines = False
    banner(ws, "CUMULATIVE WEEKLY TRENDS", cols="K")

    stores = [s["name"] for s in STORES]

    # Build week-by-week stats
    dates = sorted(all_snapshots.keys())
    rows  = []
    prev_ids_global = set()
    for di, date_str in enumerate(dates):
        snap   = all_snapshots[date_str]
        enrich(snap)
        curr_ids = {r.get("review_id","") for r in snap}
        new_cnt  = len(curr_ids - prev_ids_global)
        prev_ids_global = curr_ids

        for sname in stores:
            srev = [r for r in snap if r["store"] == sname]
            ss   = store_stats(srev)
            rows.append({
                "date": date_str,
                "store": sname,
                "total": ss["count"],
                "avg_rating": ss["avg_rating"],
                "pos_pct": ss["pos_pct"],
                "neg_pct": ss["neg_pct"],
            })

    if not rows:
        ws.cell(3, 1).value = "Not enough data yet — run the tracker again next week."
        return

    df = pd.DataFrame(rows)

    # ── Table ──
    ROW = 3
    all_dates = sorted(df["date"].unique())
    th = ["Date"] + [f"{s}\nAvg ★" for s in stores] + \
         [f"{s}\nPos %" for s in stores]
    for ci, h in enumerate(th, 1):
        hdr(ws.cell(ROW, ci), h, bg=C["navy"], sz=10)
        ws.column_dimensions[get_column_letter(ci)].width = 16
    ws.row_dimensions[ROW].height = 28
    tbl_s = ROW; ROW += 1

    for date_str in all_dates:
        ddf = df[df["date"] == date_str]
        vals = [date_str]
        for sname in stores:
            row_s = ddf[ddf["store"] == sname]
            vals.append(row_s["avg_rating"].values[0] if len(row_s) else "")
        for sname in stores:
            row_s = ddf[ddf["store"] == sname]
            vals.append(row_s["pos_pct"].values[0] if len(row_s) else "")
        for ci, v in enumerate(vals, 1):
            dat(ws.cell(ROW, ci), v, center=(ci>1))
            alt(ws, ROW, len(vals))
        ROW += 1

    borders(ws, tbl_s, ROW-1, 1, len(th))

    if len(all_dates) < 2:
        ws.cell(ROW+2, 1).value = "Chart will appear after 2+ weekly runs."
        return

    # ── Line chart: avg rating over weeks ──
    ROW += 2
    n = len(all_dates)
    chart = LineChart()
    chart.title = "Avg Star Rating by Store — Weekly"
    chart.y_axis.title = "Avg ★"
    chart.y_axis.scaling.min = 1
    chart.y_axis.scaling.max = 5
    chart.style = 10; chart.width = 30; chart.height = 16

    for si, sname in enumerate(stores):
        col_idx = 2 + si
        chart.add_data(
            Reference(ws, min_col=col_idx, max_col=col_idx,
                      min_row=tbl_s, max_row=tbl_s+n),
            titles_from_data=True
        )
    chart.set_categories(Reference(ws, min_col=1, min_row=tbl_s+1, max_row=tbl_s+n))
    ws.add_chart(chart, f"A{ROW}")


def write_alerts(wb, new_reviews):
    """Highlight new 1-2 star reviews that need attention."""
    ws = wb.create_sheet("🚨 Alerts")
    ws.sheet_view.showGridLines = False
    banner(ws, "⚠  ALERTS — NEW NEGATIVE REVIEWS REQUIRING ATTENTION",
           cols="G", bg="9C0006")

    critical = [r for r in new_reviews if (r.get("rating") or 5) <= 2]
    negative_sentiment = [r for r in new_reviews
                          if r.get("sentiment") == "Negative"
                          and (r.get("rating") or 5) >= 3]

    ROW = 3

    def section(title, revs, bg):
        nonlocal ROW
        if not revs:
            return
        ws.merge_cells(f"A{ROW}:G{ROW}")
        hdr(ws.cell(ROW, 1), f"{title}  ({len(revs)})", bg=bg, sz=11)
        ws.row_dimensions[ROW].height = 26
        ROW += 1
        for ci, h in enumerate(["Store","★","Review","Date","Author","Category","Reply?"],1):
            hdr(ws.cell(ROW, ci), h, bg="365F91", sz=10)
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 6
        ws.column_dimensions["C"].width = 70
        ws.column_dimensions["D"].width = 16
        ws.column_dimensions["E"].width = 18
        ws.column_dimensions["F"].width = 30
        ws.column_dimensions["G"].width = 8
        ROW += 1
        for r in revs:
            vals = [r.get("store",""), r.get("rating",""),
                    r.get("text","")[:600],
                    r.get("date_raw",""), r.get("author",""),
                    r.get("categories",""),
                    "Yes" if r.get("owner_reply","").strip() else "❌ No"]
            for ci, v in enumerate(vals, 1):
                cell = ws.cell(ROW, ci)
                cell.value = v
                cell.font  = Font(name="Arial", size=9)
                cell.alignment = Alignment(vertical="top", wrap_text=(ci==3))
            ws.cell(ROW, 7).fill = PatternFill("solid",
                start_color="C6EFCE" if vals[6]=="Yes" else "FFC7CE")
            ROW += 1
        borders(ws, ROW - len(revs) - 1, ROW-1, 1, 7)
        ROW += 2

    if not critical and not negative_sentiment:
        ws.merge_cells("A3:G3")
        ws["A3"].value = "[OK]  No new negative reviews this week. Great job!"
        ws["A3"].font  = Font(name="Arial", size=13, bold=True, color="1E5631")
        ws["A3"].alignment = Alignment(horizontal="center")
        ws["A3"].fill  = PatternFill("solid", start_color=C["lgreen"])
        return

    section("🔴  LOW STAR RATINGS  (1–2 Stars)",   critical,           "9C0006")
    section("🟡  NEGATIVE SENTIMENT (3–5 Stars)",  negative_sentiment, "7D6608")


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    print(f"\n{'='*60}")
    print(f"  IDEALZ WEEKLY TRACKER — {TODAY}")
    print(f"{'='*60}")

    # 1. Scrape newest reviews
    print("\n[1/4] Scraping newest reviews...")
    current = scrape_newest(MAX_SAFETY_LIMIT)
    if not current:
        print("  X No reviews scraped. Check your store URLs.")
        return
    print(f"  Total scraped: {len(current)}")

    # 2. Save snapshot
    print("\n[2/4] Saving snapshot...")
    save_snapshot(current)

    # 3. Load previous snapshot for comparison
    print("\n[3/4] Loading previous snapshot...")
    previous, prev_date = load_previous_snapshot()
    if not previous:
        print("  ! No previous snapshot found — this is your baseline.")
        print("  ! Run again next week to get a comparison report.")

    # 4. Load all snapshots for trend chart
    all_snapshots = load_all_snapshots()

    # 5. Generate report
    print("\n[4/4] Generating Excel report...")
    write_weekly_report(current, previous, prev_date, all_snapshots)

    print(f"\n{'='*60}")
    print(f"  [DONE]")
    print(f"  Report: {OUTPUT_XLSX}")
    print(f"  Snapshot: snapshots/reviews_{TODAY}.json")
    if previous:
        by_store = Counter(r["store"] for r in find_new_reviews(current, previous))
        print(f"  New reviews found:")
        for s, c in by_store.items():
            print(f"    {s}: {c}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()





