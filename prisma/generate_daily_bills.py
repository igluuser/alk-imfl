#!/usr/bin/env python3
"""
Generate daily sales bills for IMFL and Beer — KSBCL billing rules:
  Liquor: max 4,320 ml per bill  →  series KLS-2026-L-1, L-2, …
  Beer  : max 3,900 ml per bill  →  series KLS-2026-B-1, B-2, …

Sequence is year-continuous (does NOT reset each day).
Running for the same date a second time replaces that day's bills.

Usage:
  python generate_daily_bills.py --date 2026-04-01
  python generate_daily_bills.py --date 2026-04-01 --shop KLS --no-save
"""

import argparse
import psycopg2
from decimal import Decimal
from datetime import date as date_type

# ── Config ────────────────────────────────────────────────────

CONN = dict(host="195.201.119.186", port=5432,
            dbname="imfl_app", user="postgres", password="P@ssword4postgres")

IMFL_LIMIT_ML = 4320
BEER_LIMIT_ML = 3900


# ── Billing algorithm ─────────────────────────────────────────

def generate_bills(products, limit_ml, type_char, shop_short, year, start_seq):
    """
    Splits products into bills respecting the ML limit.
    Returns (list_of_bills, next_sequence_no).
    Each bill: {sequence_no, bill_number, total_ml, total_amount, items[]}
    Each item: {product_id, short_name, bottles, ml_per_bottle, total_ml, mrp_per_bottle, total_amount}
    """
    bills = []
    cur_items = []
    cur_ml = 0
    cur_amount = Decimal("0")
    seq = start_seq

    def flush():
        nonlocal cur_items, cur_ml, cur_amount, seq
        if cur_items:
            bills.append({
                "sequence_no": seq,
                "bill_number": f"{shop_short}-{year}-{type_char}-{seq}",
                "total_ml": cur_ml,
                "total_amount": cur_amount,
                "items": cur_items,
            })
            seq += 1
        cur_items = []
        cur_ml = 0
        cur_amount = Decimal("0")

    for p in products:
        remaining = p["total_sold"]
        if remaining <= 0:
            continue
        vol = p["volume_ml"]
        mrp = Decimal(str(p["mrp"]))

        while remaining > 0:
            capacity = limit_ml - cur_ml
            if capacity < vol:
                flush()
                capacity = limit_ml

            in_bill = min(capacity // vol, remaining)
            ml_in  = in_bill * vol
            amt_in = Decimal(in_bill) * mrp

            cur_items.append({
                "product_id":    p["id"],
                "short_name":    p["short_name"],
                "bottles":       in_bill,
                "ml_per_bottle": vol,
                "total_ml":      ml_in,
                "mrp_per_bottle": mrp,
                "total_amount":  amt_in,
            })
            cur_ml     += ml_in
            cur_amount += amt_in
            remaining  -= in_bill

    flush()  # close last bill
    return bills, seq


# ── DB helpers ────────────────────────────────────────────────

def fetch_sales(cur, shop_id, sale_date):
    """Return shop-level totals (sum across all counters) ordered by master_code."""
    cur.execute("""
        SELECT
            p.id,
            p.short_name,
            p.volume_ml,
            p.master_code,
            pc.product_type,
            SUM(dcs.sold_btls)         AS total_sold,
            MAX(dcs.mrp_per_bottle)    AS mrp,
            SUM(dcs.sale_amount_mrp)   AS total_amount_mrp
        FROM "DailyCounterStock" dcs
        JOIN "Counter"          c  ON dcs.counter_id  = c.id
        JOIN "Product"          p  ON dcs.product_id  = p.id
        JOIN "ProductCategory"  pc ON p.category_id   = pc.id
        WHERE c.shop_id    = %s
          AND dcs.stock_date = %s
          AND dcs.sold_btls  > 0
        GROUP BY p.id, p.short_name, p.volume_ml, p.master_code, pc.product_type
        ORDER BY pc.product_type DESC, p.master_code
    """, (shop_id, sale_date))
    rows = cur.fetchall()
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, r)) for r in rows]


def next_seq(cur, shop_id, bill_year, bill_type):
    cur.execute("""
        SELECT COALESCE(MAX(sequence_no), 0) + 1
        FROM "DailySalesBill"
        WHERE shop_id  = %s
          AND bill_year  = %s
          AND bill_type  = %s
    """, (shop_id, bill_year, bill_type))
    return cur.fetchone()[0]


def delete_day_bills(cur, shop_id, bill_date):
    cur.execute("""
        DELETE FROM "DailySalesBillItem"
        WHERE bill_id IN (
            SELECT id FROM "DailySalesBill"
            WHERE shop_id = %s AND bill_date = %s
        )
    """, (shop_id, bill_date))
    cur.execute("""
        DELETE FROM "DailySalesBill"
        WHERE shop_id = %s AND bill_date = %s
    """, (shop_id, bill_date))


def save_bills(cur, bills, shop_id, bill_date, bill_year, bill_type):
    for b in bills:
        cur.execute("""
            INSERT INTO "DailySalesBill"
                (shop_id, bill_date, bill_year, bill_type, sequence_no, bill_number,
                 total_ml, total_amount)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id
        """, (shop_id, bill_date, bill_year, bill_type,
              b["sequence_no"], b["bill_number"],
              b["total_ml"], str(b["total_amount"])))
        bill_id = cur.fetchone()[0]

        rows = [
            (bill_id, i["product_id"], i["bottles"], i["ml_per_bottle"],
             i["total_ml"], str(i["mrp_per_bottle"]), str(i["total_amount"]))
            for i in b["items"]
        ]
        from psycopg2.extras import execute_values
        execute_values(cur, """
            INSERT INTO "DailySalesBillItem"
                (bill_id, product_id, bottles, ml_per_bottle,
                 total_ml, mrp_per_bottle, total_amount)
            VALUES %s
        """, rows)


# ── Report printer ────────────────────────────────────────────

def print_report(bills_imfl, bills_beer, shop_name, bill_date):
    W = 88
    line  = "─" * W
    dline = "═" * W

    print(dline)
    print(f"  DAILY SALES BILL REPORT  ·  {shop_name}  ·  {bill_date}")
    print(dline)

    def print_type_section(bills, label, limit):
        if not bills:
            print(f"\n  No {label} sales today.\n")
            return

        print(f"\n  ■ {label.upper()} BILLS   (limit: {limit:,} ml per bill)")
        print(f"  {line}")

        grand_ml  = 0
        grand_amt = Decimal("0")

        for b in bills:
            print(f"\n  Bill: {b['bill_number']:<28}  "
                  f"Total: {b['total_ml']:>6,} ml  |  "
                  f"₹{b['total_amount']:>10,.2f}")
            print(f"  {'Brand':<35} {'ml':>5}  {'Btls':>5}  {'Tot.ml':>7}  "
                  f"{'MRP/btl':>9}  {'Amount':>10}")
            print(f"  {'─'*35}  {'─'*5}  {'─'*5}  {'─'*7}  {'─'*9}  {'─'*10}")

            for it in b["items"]:
                print(f"  {it['short_name']:<35} "
                      f"{it['ml_per_bottle']:>5}  "
                      f"{it['bottles']:>5}  "
                      f"{it['total_ml']:>7,}  "
                      f"₹{it['mrp_per_bottle']:>8.2f}  "
                      f"₹{it['total_amount']:>9.2f}")

            print(f"  {'─'*35}  {'─'*5}  {'─'*5}  {'─'*7}  {'─'*9}  {'─'*10}")
            print(f"  {'BILL TOTAL':<35}  {'':5}  {'':5}  "
                  f"{b['total_ml']:>7,}  {'':9}  "
                  f"₹{b['total_amount']:>9.2f}")

            grand_ml  += b["total_ml"]
            grand_amt += b["total_amount"]

        print(f"\n  {label} subtotal: {len(bills)} bill(s)  |  "
              f"{grand_ml:,} ml total  |  ₹{grand_amt:,.2f}")
        print(f"  {line}")

    print_type_section(bills_imfl, "Liquor (IMFL)", IMFL_LIMIT_ML)
    print_type_section(bills_beer, "Beer",          BEER_LIMIT_ML)

    total_bills = len(bills_imfl) + len(bills_beer)
    total_amt   = (sum(b["total_amount"] for b in bills_imfl) +
                   sum(b["total_amount"] for b in bills_beer))
    print(f"\n  TOTAL — {total_bills} bill(s)  |  ₹{total_amt:,.2f}\n")
    print(dline)


# ── Main ──────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Generate KSBCL daily sales bills")
    ap.add_argument("--date",    required=True,
                    help="Sale date YYYY-MM-DD")
    ap.add_argument("--shop",    default="KLS",
                    help="Shop short_name (default: KLS)")
    ap.add_argument("--no-save", action="store_true",
                    help="Print report only — do not write to DB")
    args = ap.parse_args()

    sale_date = date_type.fromisoformat(args.date)
    year      = sale_date.year

    conn = psycopg2.connect(**CONN)
    conn.autocommit = False
    cur  = conn.cursor()

    # Resolve shop
    cur.execute('SELECT id, name, short_name FROM "Shop" WHERE short_name = %s',
                (args.shop,))
    row = cur.fetchone()
    if not row:
        print(f"ERROR: shop '{args.shop}' not found.")
        return
    shop_id, shop_name, shop_short = row

    # Fetch sales
    products = fetch_sales(cur, shop_id, sale_date)
    if not products:
        print(f"No sales found for {shop_short} on {sale_date}.")
        return

    imfl_products = [p for p in products if p["product_type"] == "IMFL"]
    beer_products = [p for p in products if p["product_type"] == "BEER"]

    if not args.no_save:
        # Delete existing bills for this day before regenerating
        delete_day_bills(cur, shop_id, sale_date)

    # Determine next sequence numbers (after deletion, sequence should resume
    # from where the year left off for other dates)
    imfl_start = next_seq(cur, shop_id, year, "IMFL") if not args.no_save else 1
    beer_start  = next_seq(cur, shop_id, year, "BEER") if not args.no_save else 1

    bills_imfl, _ = generate_bills(imfl_products, IMFL_LIMIT_ML, "L",
                                   shop_short, year, imfl_start)
    bills_beer, _ = generate_bills(beer_products, BEER_LIMIT_ML, "B",
                                   shop_short, year, beer_start)

    if not args.no_save and (bills_imfl or bills_beer):
        save_bills(cur, bills_imfl, shop_id, sale_date, year, "IMFL")
        save_bills(cur, bills_beer, shop_id, sale_date, year, "BEER")
        conn.commit()
        print(f"  Saved {len(bills_imfl)} liquor bill(s) and "
              f"{len(bills_beer)} beer bill(s) to DB.")

    print_report(bills_imfl, bills_beer, shop_name, sale_date)

    cur.close()
    conn.close()


if __name__ == "__main__":
    main()
