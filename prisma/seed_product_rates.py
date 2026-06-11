#!/usr/bin/env python3
"""
seed_product_rates.py — Populate ProductRate from a KSBCL indent PDF.

Extracts Permit Dt., then for each row: item name, Rate(Per CBs.),
AROED(Per Btls.). Matches each row to a DB Product by volume_ml + brand
name similarity, then upserts into ProductRate(product_id, permit_date).

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

MATCH_THRESHOLD = 0.50  # minimum token-overlap score to auto-match

# Generic packaging / grammar words that carry no brand meaning
_STOP = {
    'pack', 'brick', 'aseptic', 'btls', 'abp', 'tp', 'packs', 'ab',
    'the', 'and', 'a', 'an', 'of', 'with', 'by', 's', 'x',
    'pet', 'tetra', 'xxx', 'malt', 'indian',
}


def parse_date(s):
    return datetime.strptime(s.strip(), "%d-%m-%Y").date()


def extract_volume_ml(item_name):
    m = re.search(r'(\d+)\s*ML', item_name, re.IGNORECASE)
    return int(m.group(1)) if m else None


def extract_brand_key(item_name):
    """Return meaningful words before the dash or volume marker — the brand portion."""
    s = item_name
    s = re.sub(r'\s*[-–]\s*.*', '', s)
    s = re.sub(r'\d+\s*ML.*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\(.*?\)', '', s)
    return s.strip()


def _tokens(text):
    return set(re.findall(r'[a-z0-9]+', text.lower())) - _STOP


def token_overlap(brand_key, candidate_name):
    """Fraction of brand_key tokens found in candidate_name."""
    qa = _tokens(brand_key)
    cb = _tokens(candidate_name)
    if not qa:
        return 0.0
    return len(qa & cb) / len(qa)


def best_match(brand_key, volume_ml, products):
    """
    products: list of (id, short_name, ksbcl_item_name, volume_ml)
    Returns (product_row, score).
    """
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
    """
    Returns (permit_date, rows).
    rows = [{"item_name": str, "rate_per_cb": Decimal, "aroed_per_bottle": Decimal}]
    """
    with pdfplumber.open(path) as pdf:
        text = pdf.pages[0].extract_text()

    # Permit Dt.
    m = re.search(r'Permit\s+Dt\.\s+(\d{2}-\d{2}-\d{4})', text)
    if not m:
        raise ValueError("Could not find 'Permit Dt.' in PDF")
    permit_date = parse_date(m.group(1))

    rows = []
    for line in text.splitlines():
        # Lines start with a row number
        m_sr = re.match(r'^\s*(\d+)\s+(.+)', line)
        if not m_sr:
            continue

        body = m_sr.group(2).strip()

        # Item code is a 7-digit number — find it to split name from numbers
        m_code = re.search(r'\b(\d{7})\b', body)
        if not m_code:
            continue

        item_name = body[:m_code.start()].strip()
        after_code = body[m_code.end():]

        # Numbers after item code: rate, avail_cbs, avail_btls, indent_cbs,
        #   indent_btls, amount, aroed_per_btl, aroed_value
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
            "item_name":       item_name,
            "rate_per_cb":     rate_per_cb,
            "aroed_per_bottle": aroed_per_bottle,
        })

    return permit_date, rows


def get_latest_rate(cur, product_id):
    """Return (rate_per_cb, aroed_per_bottle) of the most recent entry, or None."""
    cur.execute("""
        SELECT rate_per_cb, aroed_per_bottle
        FROM "ProductRate"
        WHERE product_id = %s
        ORDER BY effective_date DESC, id DESC
        LIMIT 1
    """, (product_id,))
    row = cur.fetchone()
    if row:
        return Decimal(str(row[0])), Decimal(str(row[1]))
    return None


def seed(pdf_path, dry_run=False):
    effective_date, rows = parse_pdf(pdf_path)
    print(f"Permit Date  : {effective_date}")
    print(f"Rows parsed  : {len(rows)}")
    print()

    conn = psycopg2.connect(**CONN)
    cur = conn.cursor()

    cur.execute('SELECT id, short_name, ksbcl_item_name, volume_ml FROM "Product"')
    all_products = cur.fetchall()

    print(f"{'#':<3}  {'PDF Item Name':<45} {'Vol':>4}  {'Matched Product':<30} {'Score':>5}  {'Rate':>8}  {'Aroed':>6}  Status")
    print("-" * 130)

    inserted = unchanged = skipped = 0

    for i, r in enumerate(rows, 1):
        item_name = r["item_name"]
        volume_ml = extract_volume_ml(item_name)
        brand_key = extract_brand_key(item_name)

        if volume_ml is None:
            print(f"{i:<3}  {item_name[:45]:<45}  ???  {'NO VOLUME'}")
            skipped += 1
            continue

        prod, score = best_match(brand_key, volume_ml, all_products)

        if prod is None or score < MATCH_THRESHOLD:
            print(f"{i:<3}  {item_name[:45]:<45} {volume_ml:>4}  {'NO MATCH':<30} {score:>5.2f}")
            skipped += 1
            continue

        pid, short_name = prod[0], prod[1]
        new_rate  = r["rate_per_cb"]
        new_aroed = r["aroed_per_bottle"]
        flag = "  *** LOW CONF" if score < 0.5 else ""

        # Check if latest stored values are identical — skip if unchanged
        latest = get_latest_rate(cur, pid)
        if latest and latest[0] == new_rate and latest[1] == new_aroed:
            status = "unchanged"
            unchanged += 1
        else:
            status = "NEW" if latest is None else "CHANGED"
            if not dry_run:
                cur.execute("""
                    INSERT INTO "ProductRate"
                        (product_id, effective_date, rate_per_cb, aroed_per_bottle)
                    VALUES (%s, %s, %s, %s)
                """, (pid, effective_date, str(new_rate), str(new_aroed)))
            inserted += 1

        print(f"{i:<3}  {item_name[:45]:<45} {volume_ml:>4}  {short_name:<30} {score:>5.2f}  {new_rate:>8}  {new_aroed:>6}  {status}{flag}")

    print()

    if dry_run:
        print(f"DRY RUN — {inserted} would insert, {unchanged} unchanged, {skipped} skipped.")
        conn.rollback()
    else:
        conn.commit()
        print(f"Done — {inserted} inserted, {unchanged} unchanged, {skipped} skipped.")

    cur.close()
    conn.close()


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry_run = "--dry-run" in sys.argv

    if not args:
        print("Usage: python3 seed_product_rates.py <path-to-pdf> [--dry-run]")
        sys.exit(1)

    seed(args[0], dry_run=dry_run)
