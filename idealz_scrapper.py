"""
╔══════════════════════════════════════════════════════════════════════╗
║       IDEALZ GOOGLE REVIEWS — SCRAPER v4                            ║
║       Fixed: deduplication using data-review-id                      ║
╚══════════════════════════════════════════════════════════════════════╝

SETUP (run once):
    pip install playwright textblob openpyxl pandas
    playwright install chromium

RUN:
    python idealz_scraper_v4.py
"""

import json, re, sys

# ── STORE CONFIGURATION ───────────────────────────────────────────────────────
# Paste each store's direct Google Maps /maps/place/ URL here.
# How: Open maps.google.com → search store → click it → copy URL from address bar

STORES = [
    {
        "name":         "Idealz Prime",
        "url":          "https://www.google.com/maps/place/iDealz+Prime/@6.8912695,79.8560961,17z/data=!3m1!4b1!4m6!3m5!1s0x3ae259005a2260c1:0xd6febd8ffeac3a34!8m2!3d6.8912695!4d79.8560961!16s%2Fg%2F11w27bncwk?entry=ttu&g_ep=EgoyMDI2MDIxOC4wIKXMDSoASAFQAw%3D%3D",
        "expected":     513,    # actual Google review count as of 22 Feb 2026
    },
    {
        "name":         "Idealz Lanka - Marino Mall",
        "url":          "https://www.google.com/maps/place/iDealz+Lanka+-+Marino+Mall/@6.9001796,79.8523305,17z/data=!3m1!4b1!4m6!3m5!1s0x3ae25957ebf8012b:0xe0e160f3a83edd3c!8m2!3d6.9001796!4d79.8523305!16s%2Fg%2F11gr41k7q8?entry=ttu&g_ep=EgoyMDI2MDIxOC4wIKXMDSoASAFQAw%3D%3D",
        "expected":     1472,
    },
    {
        "name":         "Idealz Lanka - Liberty Plaza",
        "url":          "https://www.google.com/maps/place/iDealz+Lanka+-+Liberty+Plaza/@6.9001796,79.8523305,17z/data=!3m1!4b1!4m6!3m5!1s0x3ae25957ebf8012b:0xe0e160f3a83edd3c!8m2!3d6.9001796!4d79.8523305!16s%2Fg%2F11gr41k7q8?entry=ttu&g_ep=EgoyMDI2MDIxOC4wIKXMDSoASAFQAw%3D%3D",
        "expected":     1881,
    },
]

OUTPUT_JSON = "idealz_raw_reviews.json"

# ─────────────────────────────────────────────────────────────────────────────

def wait(page, ms=2500):
    page.wait_for_timeout(ms)


def go_to_store(page, store):
    print(f"\n{'='*60}")
    print(f"  Store : {store['name']}  (expecting ~{store['expected']} reviews)")
    print(f"{'='*60}")

    # Append hl=en so UI is in English, but ALL reviews (any language) are shown
    url = store["url"]
    url += ("&" if "?" in url else "?") + "hl=en"
    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    wait(page, 3000)

    # If it's a search page, click first result
    if "/maps/place/" not in page.url:
        print("  -> Search results — clicking first listing...")
        try:
            page.wait_for_selector('a[href*="/maps/place/"], .hfpxzc', timeout=10000)
        except Exception:
            print("  X No results found.")
            return False
        for sel in ['a[href*="/maps/place/"]', '.hfpxzc', 'div.Nv2PK a']:
            el = page.query_selector(sel)
            if el:
                try:
                    el.click()
                    page.wait_for_load_state("domcontentloaded")
                    wait(page, 4000)
                    break
                except Exception:
                    continue

    if "/maps/place/" not in page.url:
        print(f"  X Not on a place page: {page.url[:100]}")
        return False

    print("  -> Place page loaded OK")
    return True


def click_reviews_tab(page):
    try:
        page.wait_for_selector('button[role="tab"], div[role="tablist"]', timeout=12000)
    except Exception:
        pass

    for sel in [
        'button[aria-label*="Reviews"]',
        'button[aria-label*="reviews"]',
        'button[role="tab"]:has-text("Reviews")',
    ]:
        el = page.query_selector(sel)
        if el:
            try:
                el.scroll_into_view_if_needed()
                el.click()
                wait(page, 3000)
                print(f"  -> Reviews tab clicked")
                return True
            except Exception:
                continue

    # Fallback: any tab with a number (review count)
    for tab in page.query_selector_all('button[role="tab"]'):
        try:
            txt = tab.inner_text().strip()
            if re.search(r'\d{2,}', txt) or "review" in txt.lower():
                tab.click()
                wait(page, 3000)
                print(f"  -> Reviews tab clicked (text: '{txt}')")
                return True
        except Exception:
            continue

    print("  ! Reviews tab not found")
    return False


def sort_newest(page):
    for sel in ['button[aria-label^="Sort"]', 'button[data-value="Sort"]']:
        btn = page.query_selector(sel)
        if btn:
            try:
                btn.click()
                wait(page, 1500)
                for opt in ['div[data-index="1"]', 'li[aria-label="Newest"]',
                            'div[role="menuitemradio"]:has-text("Newest")',
                            'li:has-text("Newest")']:
                    o = page.query_selector(opt)
                    if o:
                        o.click()
                        wait(page, 2500)
                        print("  -> Sorted by Newest")
                        return
            except Exception:
                continue
    print("  ! Could not sort (using default order)")


def scroll_to_load_all(page, target):
    """
    Scroll using data-review-id deduplication so we know exactly
    how many UNIQUE reviews are loaded — not DOM node count.
    Stops when unique count >= target OR no new IDs appear for 15 cycles.
    """
    seen_ids   = set()
    stall      = 0

    while True:
        # Collect all unique review IDs currently in DOM
        ids = page.evaluate("""() => {
            return Array.from(document.querySelectorAll('[data-review-id]'))
                        .map(el => el.getAttribute('data-review-id'))
                        .filter(id => id);
        }""")
        unique_now = set(ids)
        new_ids    = unique_now - seen_ids
        seen_ids   = unique_now
        n          = len(seen_ids)

        print(f"    {n} unique reviews loaded  (target: {target})...", end="\r", flush=True)

        if n >= target:
            print(f"\n    ✓ Reached target ({target}). Done scrolling.")
            break

        if not new_ids:
            stall += 1
            if stall >= 20:
                print(f"\n    ✓ No more reviews loading. Final unique count: {n}")
                break
        else:
            stall = 0

        # Scroll the reviews panel
        page.evaluate("""() => {
            const candidates = [
                document.querySelector('div[role="feed"]'),
                document.querySelector('.m6QErb.DxyBCb'),
                document.querySelector('.m6QErb[aria-label]'),
                document.querySelector('.m6QErb'),
            ];
            for (const el of candidates) {
                if (el && el.scrollHeight > el.clientHeight) {
                    el.scrollTop += 5000;
                    return;
                }
            }
            // Last resort
            document.querySelectorAll('div').forEach(el => {
                const r = el.getBoundingClientRect();
                if (r.left < 600 && el.scrollHeight - el.clientHeight > 200)
                    el.scrollTop += 5000;
            });
        }""")
        wait(page, 2500)

    return seen_ids  # return the set of unique IDs


def expand_text(page):
    count = 0
    for btn in page.query_selector_all(
        'button.w8nwRe, button[aria-label="See more"]'
    ):
        try:
            btn.scroll_into_view_if_needed()
            btn.click()
            page.wait_for_timeout(60)
            count += 1
        except Exception:
            pass
    if count:
        print(f"  -> Expanded {count} truncated reviews")


def extract(page, store_name):
    """
    Extract reviews, deduplicating by data-review-id so duplicates
    from Google's virtual scroll are removed.
    """
    cards = page.query_selector_all('[data-review-id]')
    seen_ids = set()
    reviews  = []

    for card in cards:
        try:
            # ── Deduplicate by review ID ──
            review_id = card.get_attribute("data-review-id") or ""
            if review_id in seen_ids:
                continue
            seen_ids.add(review_id)

            # ── Author ──
            author = "Anonymous"
            for s in ['.d4r55', '.DUwDvf']:
                el = card.query_selector(s)
                if el:
                    t = el.inner_text().strip()
                    if t:
                        author = t
                        break

            # ── Rating ──
            rating = None
            for s in ['[aria-label*="star"]', '[aria-label*="Star"]', '.kvMYJc']:
                el = card.query_selector(s)
                if el:
                    label = el.get_attribute("aria-label") or ""
                    m = re.search(r'(\d+)', label)
                    if m:
                        rating = int(m.group(1))
                        break

            # ── Review text ──
            text = ""
            for s in ['.wiI7pd', '.MyEned']:
                el = card.query_selector(s)
                if el:
                    t = el.inner_text().strip()
                    if t:
                        text = t
                        break

            # ── Date ──
            date_str = ""
            for s in ['.rsqaWe', '.dehysf', '.xRkPPb']:
                el = card.query_selector(s)
                if el:
                    t = el.inner_text().strip()
                    if t:
                        date_str = t
                        break

            # ── Likes ──
            likes = 0
            el = card.query_selector('.GBkF3d')
            if el:
                raw = el.inner_text().strip()
                if raw.isdigit():
                    likes = int(raw)

            # ── Owner reply ──
            owner_reply = ""
            el = card.query_selector('.CDe7pd')
            if el:
                owner_reply = el.inner_text().strip()

            reviews.append({
                "review_id":   review_id,
                "store":       store_name,
                "author":      author,
                "rating":      rating,
                "date_raw":    date_str,
                "text":        text,
                "likes":       likes,
                "owner_reply": owner_reply,
                "has_text":    bool(text),
            })

        except Exception:
            continue

    with_text  = sum(1 for r in reviews if r["has_text"])
    stars_only = len(reviews) - with_text
    print(f"  ✓ {len(reviews)} unique reviews  |  with text: {with_text}  |  stars-only: {stars_only}")
    return reviews


def scrape_store(page, store):
    if not go_to_store(page, store):
        return []
    click_reviews_tab(page)
    sort_newest(page)

    try:
        page.wait_for_selector('[data-review-id]', timeout=12000)
    except Exception:
        print("  X No review cards appeared.")
        return []

    # Scroll until we hit the expected count (add 10% buffer for new reviews)
    target = int(store["expected"] * 1.1) + 20
    scroll_to_load_all(page, target)
    expand_text(page)
    return extract(page, store["name"])


def main():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Run:  pip install playwright && playwright install chromium")
        sys.exit(1)

    all_reviews = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            slow_mo=30,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox",
                  "--start-maximized"]
        )
        ctx = browser.new_context(
            viewport={"width": 1400, "height": 900},
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/122.0.0.0 Safari/537.36"),
            locale="si-LK",  # Use Sri Lanka locale to get ALL reviews incl. Sinhala/Tamil
            timezone_id="Asia/Colombo",
        )
        page = ctx.new_page()

        # Dismiss cookie popup
        page.goto("https://www.google.com/maps", wait_until="domcontentloaded")
        wait(page, 2000)
        for sel in ['button[aria-label*="Accept"]', 'button:has-text("Accept all")',
                    'button:has-text("I agree")']:
            try:
                btn = page.query_selector(sel)
                if btn:
                    btn.click()
                    wait(page, 1000)
                    break
            except Exception:
                pass

        for store in STORES:
            try:
                reviews = scrape_store(page, store)
                all_reviews.extend(reviews)
                wait(page, 3000)
            except Exception as e:
                print(f"\n  X Failed on '{store['name']}': {e}")

        browser.close()

    # Save JSON
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(all_reviews, f, ensure_ascii=False, indent=2)

    print(f"\n{'='*60}")
    print(f"  DONE — {len(all_reviews)} total unique reviews → {OUTPUT_JSON}")
    print(f"  Breakdown:")
    by_store = {}
    for r in all_reviews:
        by_store[r["store"]] = by_store.get(r["store"], 0) + 1
    for s, c in by_store.items():
        print(f"    {s}: {c}")
    print(f"\n  Next step:  python idealz_analyzer.py")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()