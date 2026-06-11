#!/usr/bin/env python3
"""
seed_product_rates.py — Populate ProductRate from a KSBCL indent PDF.

For each row: matches to a DB Product by volume_ml + brand name similarity.
If no match found, auto-creates a new Product record from the PDF item name.
Inserts into ProductRate only when rate or aroed changes vs the last known value.

This script does NOT create any Permit record — only rate/aroed data.

Usage:
    python3 seed_product_rates.py <path-to-pdf>
    python3 seed_product_rates.py <path-to-pdf> --dry-run
"""

import re
import sys
from datetime import datetime
from decimal import Decimal

import pdfplumber
import psycopg2

CONN = dict(host='195.201.119.186', port=5432, dbname='imfl_app',
            user='postgres', password='P@ssword4postgres')

MATCH_THRESHOLD = 0.50

# Generic packaging / grammar words that carry no brand meaning
_STOP = {
    'pack', 'brick', 'aseptic', 'btls', 'abp', 'tp', 'packs', 'ab',
    'the', 'and', 'a', 'an', 'of', 'with', 'by', 's', 'x',
    'pet', 'tetra', 'xxx', 'malt', 'indian',
}

# Category IDs in ProductCategory table
_CAT = {
    'BRANDY': 1, 'WHISKY': 2, 'RUM': 3, 'GIN': 4,
    'VODKA': 5, 'WINE': 6, 'OTHERS': 7,
    'BEER_BOTTLE': 8, 'BEER_TIN': 9,
}


def parse_date(s):
    return datetime.strptime(s.strip(), "%d-%m-%Y").date()


def extract_volume_ml(item_name):
    m = re.search(r'(\d+)\s*ML', item_name, re.IGNORECASE)
    return int(m.group(1)) if m else None


def extract_carton_size(item_name):
    m = re.search(r'[xX×](\d+)\s*(?:Btls|ABP|AB\.Pack|AB|P\.Btls|T\.Packs|Packs|TP|Cans|Tins|B\.Pack)',
                  item_name, re.IGNORECASE)
    return int(m.group(1)) if m else 12


def extract_brand_key(item_name):
    s = item_name
    s = re.sub(r'\s*[-–]\s*.*', '', s)
    s = re.sub(r'\d+\s*ML.*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\(.*?\)', '', s)
    return s.strip()


def infer_category_id(item_name):
    n = item_name.lower()
    if any(w in n for w in ['whisky', 'whiskey', 'scotch', 'grain whisky', 'malt whisky']):
        return _CAT['WHISKY']
    if 'brandy' in n:
        return _CAT['BRANDY']
    if 'rum' in n:
        return _CAT['RUM']
    if 'gin' in n:
        return _CAT['GIN']
    if 'vodka' in n:
        return _CAT['VODKA']
    if any(w in n for w in ['wine', 'cranberry', 'grape', 'merlot', 'cabernet', 'shiraz']):
        return _CAT['WINE']
    if any(w in n for w in ['beer', 'lager', 'stout', 'ale', 'pilsner']):
        if any(w in n for w in ['can', 'tin']):
            return _CAT['BEER_TIN']
        return _CAT['BEER_BOTTLE']
    return _CAT['OTHERS']


def _tokens(text):
    return set(re.findall(r'[a-z0-9]+', text.lower())) - _STOP


def token_overlap(brand_key, candidate_name):
    qa = _tokens(brand_key)
    cb = _tokens(candidate_name)
    if not qa:
        return 0.0
    return len(qa & cb) / len(qa)


def best_match(brand_key, volume_ml, products):
    candidates = [p for p in products if p[3] == volume_ml]
    if not candidates:
        return None, 0.0
    scored = []
    for p in candidates:
        names = [n for n in [p[1], p[2]] if n]
        score = max((token_overlap(brand_key, n) for n in names), default=0.0)
        scored.append((score, p))
    scored.sort(reverse=True)
    return scored[0][1], scored[0][0]


def parse_pdf(path):
    with pdfplumber.open(path) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    m = re.search(r'Permit\s+Dt\.\s+(\d{2}-\d{2}-\d{4})', text)
    if not m:
        raise ValueError("Could not find 'Permit Dt.' in PDF")
    permit_date = parse_date(m.group(1))

    rows = []
    for line in text.splitlines():
        m_sr = re.match(r'^\s*(\d+)\s+(.+)', line)
        if not m_sr:
            continue
        body = m_sr.group(2).strip()

        m_code = re.search(r'\b(\d{7})\b', body)
        if not m_code:
            continue

        item_name  = body[:m_code.start()].strip()
        after_code = body[m_code.end():]

        nums = re.findall(r'\d+(?:\.\d+)?', after_code)
        if len(nums) < 7:
            continue

        try:
            rate_per_cb      = Decimal(nums[0])
            aroed_per_bottle = Decimal(nums[6])
        except Exception:
            continue

        if not item_name:
            continue

        rows.append({
            "item_name":        item_name,
            "rate_per_cb":      rate_per_cb,
            "aroed_per_bottle": aroed_per_bottle,
        })

    return permit_date, rows


def get_latest_rate(cur, product_id):
    cur.execute("""
        SELECT rate_per_cb, aroed_per_bottle
        FROM "ProductRate"
        WHERE product_id = %s
        ORDER BY effective_date DESC, id DESC
        LIMIT 1
    """, (product_id,))
    row = cur.fetchone()
    return (Decimal(str(row[0])), Decimal(str(row[1]))) if row else None


def auto_create_product(cur, item_name, volume_ml, next_mc, dry_run):
    """
    Insert a new Product derived from the PDF item name.
    Returns (product_id, short_name). On dry_run uses a placeholder id (-1).
    """
    short_name   = extract_brand_key(item_name)[:60]
    carton_size  = extract_carton_size(item_name)
    category_id  = infer_category_id(item_name)

    if dry_run:
        return -1, short_name

    cur.execute("""
        INSERT INTO "Product"
            (master_code, short_name, ksbcl_item_name, volume_ml,
             carton_size, category_id, is_active, created_at, updated_at)
        VALUES (%s, %s, %s, %s, %s, %s, true, now(), now())
        RETURNING id
    """, (next_mc, short_name, item_name, volume_ml, carton_size, category_id))
    pid = cur.fetchone()[0]
    return pid, short_name


def seed(pdf_path, dry_run=False):
    effective_date, rows = parse_pdf(pdf_path)
    print(f"Permit Date  : {effective_date}")
    print(f"Rows parsed  : {len(rows)}")
    print()

    conn = psycopg2.connect(**CONN)
    cur  = conn.cursor()

    cur.execute('SELECT id, short_name, ksbcl_item_name, volume_ml FROM "Product"')
    all_products = list(cur.fetchall())

    cur.execute('SELECT COALESCE(MAX(master_code), 499) FROM "Product"')
    next_mc = max(cur.fetchone()[0] + 1, 500)

    hdr = f"{'#':<3}  {'PDF Item Name':<47} {'Vol':>4}  {'Matched Product':<32} {'Sc':>4}  {'Rate':>8}  {'Aroed':>6}  Status"
    print(hdr)
    print("-" * len(hdr))

    rate_inserted = rate_unchanged = prod_created = parse_skipped = 0

    for i, r in enumerate(rows, 1):
        item_name = r["item_name"]
        volume_ml = extract_volume_ml(item_name)
        brand_key = extract_brand_key(item_name)

        if volume_ml is None:
            print(f"{i:<3}  {item_name[:47]:<47}  ???  NO VOLUME")
            parse_skipped += 1
            continue

        prod, score = best_match(brand_key, volume_ml, all_products)

        if prod is None or score < MATCH_THRESHOLD:
            # Auto-create new Product
            pid, short_name = auto_create_product(cur, item_name, volume_ml, next_mc, dry_run)
            if not dry_run:
                # Add to in-memory list so later rows in this PDF can match it
                all_products.append((pid, short_name, item_name, volume_ml))
            next_mc += 1
            score  = 1.0
            status = "PRODUCT CREATED"
            prod_created += 1
        else:
            pid, short_name = prod[0], prod[1]
            status = None

        new_rate  = r["rate_per_cb"]
        new_aroed = r["aroed_per_bottle"]

        if status != "PRODUCT CREATED":
            latest = get_latest_rate(cur, pid)
            if latest and latest[0] == new_rate and latest[1] == new_aroed:
                status = "unchanged"
                rate_unchanged += 1
            else:
                status = "NEW" if latest is None else "CHANGED"
                rate_inserted += 1

        if status == "PRODUCT CREATED":
            rate_inserted += 1

        if status not in ("unchanged",) and not dry_run:
            cur.execute("""
                INSERT INTO "ProductRate"
                    (product_id, effective_date, rate_per_cb, aroed_per_bottle)
                VALUES (%s, %s, %s, %s)
            """, (pid, effective_date, str(new_rate), str(new_aroed)))

        sc_str = f"{score:.2f}" if status == "unchanged" or score < 1.0 else "    "
        print(f"{i:<3}  {item_name[:47]:<47} {volume_ml:>4}  {short_name:<32} {sc_str:>4}  {new_rate:>8}  {new_aroed:>6}  {status}")

    print()
    if dry_run:
        print(f"DRY RUN — {prod_created} products would be created, "
              f"{rate_inserted} rates would insert, {rate_unchanged} unchanged, "
              f"{parse_skipped} parse errors.")
        conn.rollback()
    else:
        conn.commit()
        print(f"Done — {prod_created} products created, "
              f"{rate_inserted} rates inserted, {rate_unchanged} unchanged, "
              f"{parse_skipped} parse errors.")

    cur.close()
    conn.close()


if __name__ == "__main__":
    args    = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry_run = "--dry-run" in sys.argv

    if not args:
        print("Usage: python3 seed_product_rates.py <path-to-pdf> [--dry-run]")
        sys.exit(1)

    seed(args[0], dry_run=dry_run)
