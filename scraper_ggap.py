#!/usr/bin/env python3
"""
GLOBALG.A.P. Supply Chain Portal scraper.

Source: https://prod.osapiens.cloud/portal/webbundle/foodplus/field-service-os/
        supply-chain-portal (the GLOBALG.A.P certificate search, run by FoodPLUS
        on the osapiens platform).

There is no usable API — the portal is a GWT single-page app in a nested iframe
whose data comes over a proprietary *binary* RPC (POST /portal/rpc?B). So we
drive the UI with headless Chromium via Playwright, exactly like scraper_pefc.py.

The "Product" search method requires BOTH a product and a country (a product-only
search errors "Please input Country"), so the only way to enumerate is to walk
the grid of products x countries. We restrict countries to a configured set
(default: the Baltics) and iterate every product; results are producer rows
(GGN, Producer Name, City, Country, Producer Type) which we tag with the searched
product and de-duplicate.

Env:
    GGAP_COUNTRIES     comma list of country names (default "Latvia,Estonia,Lithuania")
    GGAP_MAX_PRODUCTS  cap products scanned (local testing); unset = all
    GGAP_PRODUCTS      comma list of products to scan instead of ggap_products.json
    GGAP_HEADFUL       set to 1 to watch the browser
    GGAP_ENUM          set to 1 to (re)enumerate the product list and rewrite
                       ggap_products.json, then exit (maintenance mode)
"""

import csv
import json
import os
import re
import sys
from datetime import datetime, timezone

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

URL = ("https://prod.osapiens.cloud/portal/webbundle/foodplus/field-service-os/"
       "supply-chain-portal?app-route-hash=%252Fcertificates")

COUNTRIES = [c.strip() for c in
             os.environ.get("GGAP_COUNTRIES", "Latvia,Estonia,Lithuania").split(",") if c.strip()]
MAX_PRODUCTS = int(os.environ["GGAP_MAX_PRODUCTS"]) if os.environ.get("GGAP_MAX_PRODUCTS") else None
HEADFUL = os.environ.get("GGAP_HEADFUL") == "1"
PRODUCTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ggap_products.json")

FIELDNAMES = ["GGN", "Producer Name", "City", "Country", "Producer Type", "Product"]
PTYPES = {"Producer", "Producer group", "Producer Group"}
GGN_RE = re.compile(r"^\d{8,14}$")
CHECKPOINT_EVERY = 50   # flush a partial CSV every N searches (crash safety)

# Crop-name vocabulary used only in GGAP_ENUM maintenance mode. The portal's
# product autocomplete needs >=3 chars to trigger a server search and caps each
# query at 25 alphabetical matches, so we probe it with these stems and union.
ENUM_STEMS = (
    "almond apple apricot artichoke asparagus aubergine eggplant avocado banana barley basil bean beet "
    "berry bilberry blackberry blackcurrant blueberry bok broccoli brussels buckwheat cabbage cane cantaloupe "
    "caper carrot cassava cauliflower celeriac celery chard cherry chervil chestnut chickpea chicory chili "
    "chilli chive cilantro citrus clementine coconut coffee collard coriander corn courgette cranberry cress "
    "cucumber cumin currant date dill durian eggplant elderberry endive fennel fenugreek fig flax garlic ginger "
    "gooseberry gourd grape grapefruit guava hazelnut herb hops horseradish jackfruit kale kiwi kohlrabi kumquat "
    "leek lemon lemongrass lentil lettuce lime lingonberry lychee macadamia maize mandarin mango mangosteen "
    "marjoram melon millet mint mulberry mushroom mustard nectarine nut oat okra olive onion orange oregano papaya "
    "parsley parsnip passion pea peach peanut pear pecan pepper persimmon physalis pineapple pistachio plantain "
    "plum pomegranate pomelo poppy potato pumpkin quince radicchio radish raisin rapeseed raspberry redcurrant "
    "rhubarb rice rocket rosemary rutabaga rye sage salad salsify savory scallion sesame shallot sorghum soy spelt "
    "spinach sprout squash strawberry sugar sunflower swede sweetcorn sweetpotato tangerine tarragon tea thyme "
    "tomato turmeric turnip vanilla walnut watercress watermelon wheat yam zucchini"
).split()


# ── Playwright helpers ────────────────────────────────────────────────────────
def content_frame(page):
    for fr in page.frames:
        if "index.html" in fr.url:
            return fr
    return page.main_frame


def options(fr):
    return fr.evaluate(
        "() => Array.from(document.querySelectorAll('li[role=option]'))"
        ".map(l => l.textContent.replace(/\\.prefix__[^}]*}/g,'').trim())")


def open_portal(p):
    browser = p.chromium.launch(headless=not HEADFUL)
    ctx = browser.new_context(viewport={"width": 1500, "height": 1100})
    page = ctx.new_page()
    page.goto(URL, wait_until="networkidle", timeout=90000)
    page.wait_for_timeout(8000)
    fr = content_frame(page)
    fr.get_by_text("Product", exact=True).first.click(timeout=15000)
    page.wait_for_timeout(1200)
    return browser, ctx, page, fr


def select_option(fr, page, combo, text, want_exact=True, retries=2):
    """Type into a MUI autocomplete and click the option matching `text`.
    Clears any prior selection first (filling over a chosen value does not
    reliably replace it). Returns the selected option string, or None."""
    for _ in range(retries + 1):
        try:
            combo.click(timeout=6000)
            combo.fill("")                 # clear any existing selection/text
            page.wait_for_timeout(150)
            combo.fill(text)
            page.wait_for_timeout(1000)
            opts = options(fr)
            target = None
            if opts:
                for o in opts:
                    if o.strip().lower() == text.strip().lower():
                        target = o
                        break
                if target is None and not want_exact:
                    target = opts[0]
            if target is not None:
                fr.locator("li[role=option]").nth(opts.index(target)).click(timeout=6000)
                page.wait_for_timeout(300)
                return target
        except Exception:
            pass
        page.wait_for_timeout(400)
    return None


# ── Product list ──────────────────────────────────────────────────────────────
def enumerate_products(fr, page):
    allp = set()
    prod = fr.get_by_role("combobox").nth(0)
    for i, q in enumerate(ENUM_STEMS):
        try:
            prod.click(); prod.fill(q); page.wait_for_timeout(800)
            for o in options(fr):
                if o:
                    allp.add(o)
        except Exception:
            pass
        if i % 25 == 0:
            print(f"  enum [{i}/{len(ENUM_STEMS)}] {q!r} -> {len(allp)}", flush=True)
    return sorted(allp)


def load_products():
    if os.environ.get("GGAP_PRODUCTS"):
        return [x.strip() for x in os.environ["GGAP_PRODUCTS"].split(",") if x.strip()]
    with open(PRODUCTS_FILE, encoding="utf-8") as f:
        return json.load(f)["products"]


# ── Search + extract ──────────────────────────────────────────────────────────
def result_state(fr):
    """Wait-free read of the current results area."""
    return fr.evaluate("""() => {
        const t = (document.body.innerText || '');
        const rows = Array.from(document.querySelectorAll('.MuiDataGrid-row'))
            .map(r => (r.innerText||'').split('\\n').map(s=>s.trim()).filter(Boolean));
        const m = t.match(/(\\d[\\d.,]*)\\s*[–-]\\s*(\\d[\\d.,]*)\\s*of\\s*(\\d[\\d.,]*)/i);
        return {noRes: /No results found/i.test(t),
                rows,
                total: m ? parseInt(m[3].replace(/[.,]/g,'')) : rows.length};
    }""")


def wait_results(fr, page, prev_sig, timeout_ms=13000):
    """Poll until the NEW search's results render or 'No results found' appears.
    The MUI DataGrid can briefly retain the previous search's rows, so we require
    the first GGN to differ from `prev_sig` before trusting rows (rather than
    reading on the first zero-delay poll)."""
    step, waited = 300, 0
    while waited < timeout_ms:
        page.wait_for_timeout(step)
        waited += step
        st = result_state(fr)
        if st["noRes"]:
            return st
        if st["rows"] and st["rows"][0][0] != prev_sig:
            page.wait_for_timeout(300)          # settle, then re-read
            return result_state(fr)
    return result_state(fr)


def parse_row(cells, country):
    """cells = list of non-empty strings for one DataGrid row →
    (GGN, Producer Name, City, Country, Producer Type). Column order is
    [GGN, Producer, City, Country, Type]; we take Producer/City positionally
    after removing the GGN and Producer-Type cells, and use the searched
    country (authoritative) rather than the on-screen Country cell."""
    ggn = next((c for c in cells if GGN_RE.match(c)), "")
    ptype = next((c for c in cells if c in PTYPES), "")
    rest = [c for c in cells if c != ggn and c != ptype]
    # rest = [Producer, City, (Country)] — Country cell (rest[2]) is ignored.
    producer = rest[0] if len(rest) >= 1 else ""
    city = rest[1] if len(rest) >= 2 else ""
    return {"GGN": ggn, "Producer Name": producer, "City": city,
            "Country": country, "Producer Type": ptype}


def click_next(fr):
    """Click the DataGrid 'next page' button; return True if it was clickable."""
    try:
        btn = fr.locator("button[aria-label='Go to next page']")
        if btn.count() and btn.first.is_enabled():
            btn.first.click(timeout=4000)
            return True
    except Exception:
        pass
    return False


def wait_page_change(fr, page, prev_first, timeout_ms=8000):
    """After clicking next, wait until the first GGN changes (guards against
    reading the old page before the new one renders)."""
    step, waited = 300, 0
    while waited < timeout_ms:
        page.wait_for_timeout(step)
        waited += step
        rows = result_state(fr)["rows"]
        if rows and rows[0][0] != prev_first:
            return True
        if not rows:
            return False
    return False


def set_country(fr, page, country):
    """Select the country once; it persists across product searches."""
    return select_option(fr, page, fr.get_by_role("combobox").nth(1), country) is not None


def search(fr, page, product, country):
    """Run one product search for the already-selected country; return records."""
    # signature of whatever is currently shown (the previous search's results)
    before = result_state(fr)
    prev_sig = before["rows"][0][0] if before["rows"] else None
    if select_option(fr, page, fr.get_by_role("combobox").nth(0), product) is None:
        return []            # product not offered by the portal
    # country is set once per country loop; re-assert only if it got cleared
    if fr.get_by_role("combobox").nth(1).input_value().strip().lower() != country.strip().lower():
        if not set_country(fr, page, country):
            return []
    fr.get_by_text("Start Search", exact=False).first.click(timeout=10000)
    st = wait_results(fr, page, prev_sig)
    if st["noRes"]:
        return []
    records, seen_ggn = [], set()
    for _ in range(200):                     # hard page cap
        page_rows = result_state(fr)["rows"]
        for cells in page_rows:
            rec = parse_row(cells, country)
            if rec["GGN"] and rec["GGN"] not in seen_ggn:
                seen_ggn.add(rec["GGN"])
                rec["Product"] = product
                records.append(rec)
        prev_first = page_rows[0][0] if page_rows else None
        if not click_next(fr):
            break
        if not wait_page_change(fr, page, prev_first):
            break
    return records


# ── CSV ───────────────────────────────────────────────────────────────────────
def dedupe(records):
    seen, out = set(), []
    for r in records:
        k = (r["GGN"], r["Country"], r["Product"])
        if k not in seen:
            seen.add(k)
            out.append(r)
    return out


def write_csv(records, announce=True):
    records = dedupe(records)
    date_str = datetime.now(timezone.utc).strftime("%Y.%m.%d")
    for path in (f"GGAP certificates {date_str}.csv", "GGAP certificates latest.csv"):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=FIELDNAMES)
            w.writeheader()
            w.writerows(records)
    if announce:
        print(f"Wrote {len(records)} rows → GGAP certificates latest.csv")
    return records


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    with sync_playwright() as p:
        browser, ctx, page, fr = open_portal(p)
        try:
            if os.environ.get("GGAP_ENUM") == "1":
                prods = enumerate_products(fr, page)
                json.dump({"products": prods,
                           "note": "GLOBALG.A.P product option strings (base + PPM)."},
                          open(PRODUCTS_FILE, "w"), ensure_ascii=False, indent=1)
                print(f"Enumerated {len(prods)} products → {PRODUCTS_FILE}")
                return

            products = load_products()
            if MAX_PRODUCTS:
                products = products[:MAX_PRODUCTS]
            print(f"Scanning {len(COUNTRIES)} countries × {len(products)} products "
                  f"= {len(COUNTRIES)*len(products)} searches", flush=True)

            records, done = [], 0
            for country in COUNTRIES:
                found = 0
                if not set_country(fr, page, country):
                    print(f"! could not select country {country}; skipping", flush=True)
                    continue
                for product in products:
                    try:
                        recs = search(fr, page, product, country)
                    except PWTimeout:
                        print(f"  ! timeout on {product} × {country}", flush=True)
                        recs = []
                    records.extend(recs)
                    found += len(recs)
                    done += 1
                    if recs:
                        print(f"  {country} · {product}: {len(recs)}", flush=True)
                    if done % CHECKPOINT_EVERY == 0:
                        write_csv(records, announce=False)
                        print(f"  …{done} searches, {len(records)} rows so far", flush=True)
                print(f"{country}: {found} producer-rows", flush=True)

            write_csv(records)
        finally:
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
