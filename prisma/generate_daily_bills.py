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
import io
import os
import sys
import psycopg2
from decimal import Decimal
from datetime import date as date_type

REPORTS_DIR = os.path.join(os.path.dirname(__file__), "..", "Reports")

# ── Config ────────────────────────────────────────────────────

CONN = dict(host="195.201.119.186", port=5432,
            dbname="imfl_app", user="postgres", password="P@ssword4postgres")

IMFL_LIMIT_ML = 4320
BEER_LIMIT_ML = 3900

EXPENSE_LABEL = {
    "TEA": "Tea", "PERMIT_RENT": "Permit Rent", "POOJA": "Pooja",
    "BREAKAGE": "Breakage", "OVER_CASH": "Over Cash",
    "ELECTRICITY_BILL": "Electricity Bill", "GOOGLE_PAY": "Google Pay",
    "BHATTA": "Bhatta", "GOOGLE_PAY_NEGATIVE": "Google Pay (Negative)",
    "PAK": "PAK", "OTHERS": "Others",
}


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


# ── Expenses / Cash / Summary fetchers ───────────────────────

def fetch_counter_summaries(cur, shop_id, sale_date):
    cur.execute("""
        SELECT c.id, c.name, c.display_order,
               dcs.staff_name,
               dcs.liquor_sale_amount, dcs.beer_sale_amount,
               dcs.total_tips, dcs.grand_total_by_sale,
               dcs.google_pay_amount, dcs.expenses_total,
               dcs.collection_total, dcs.drawer_cash_total,
               dcs.tips_cash_total, dcs.tally_difference
        FROM "Counter" c
        JOIN "DailyCounterSummary" dcs
             ON dcs.counter_id = c.id AND dcs.summary_date = %s
        WHERE c.shop_id = %s
        ORDER BY c.display_order
    """, (sale_date, shop_id))
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, r)) for r in cur.fetchall()]


def fetch_expenses(cur, shop_id, sale_date):
    """Returns {counter_id: [expense_dict, ...]}"""
    cur.execute("""
        SELECT c.id, de.category, de.amount, de.is_upi, de.upi_ref
        FROM "Counter" c
        JOIN "DailyExpense" de
             ON de.counter_id = c.id AND de.expense_date = %s
        WHERE c.shop_id = %s
        ORDER BY c.display_order, de.category
    """, (sale_date, shop_id))
    result = {}
    for cid, cat, amt, is_upi, ref in cur.fetchall():
        result.setdefault(cid, []).append(
            {"category": cat, "amount": amt, "is_upi": is_upi, "upi_ref": ref}
        )
    return result


def fetch_cash(cur, shop_id, sale_date):
    """Returns {counter_id: [cash_dict, ...]}"""
    cur.execute("""
        SELECT c.id, cr.cash_category, cr.denomination_type,
               cr.denomination_value, cr.count, cr.total_amount
        FROM "Counter" c
        JOIN "CashRecord" cr
             ON cr.counter_id = c.id AND cr.record_date = %s
        WHERE c.shop_id = %s
        ORDER BY c.display_order, cr.cash_category,
                 cr.denomination_type DESC, cr.denomination_value DESC
    """, (sale_date, shop_id))
    result = {}
    for cid, cat, dtype, dval, cnt, total in cur.fetchall():
        result.setdefault(cid, []).append(
            {"cash_category": cat, "denomination_type": dtype,
             "denomination_value": dval, "count": cnt, "total_amount": total}
        )
    return result


# ── Report printer ────────────────────────────────────────────

class _Tee:
    """Write to multiple streams at once."""
    def __init__(self, *streams):
        self._streams = streams
    def write(self, data):
        for s in self._streams:
            s.write(data)
    def flush(self):
        for s in self._streams:
            s.flush()


def print_report(bills_imfl, bills_beer, shop_name, bill_date, out=None):
    W = 88
    line  = "─" * W
    dline = "═" * W

    def p(*args, **kwargs):
        kwargs.setdefault("file", out)
        print(*args, **kwargs)

    p(dline)
    p(f"  DAILY SALES BILL REPORT  ·  {shop_name}  ·  {bill_date}")
    p(dline)

    def print_type_section(bills, label, limit):
        if not bills:
            p(f"\n  No {label} sales today.\n")
            return

        p(f"\n  ■ {label.upper()} BILLS   (limit: {limit:,} ml per bill)")
        p(f"  {line}")

        grand_ml  = 0
        grand_amt = Decimal("0")

        for b in bills:
            p(f"\n  Bill: {b['bill_number']:<28}  "
              f"Total: {b['total_ml']:>6,} ml  |  "
              f"₹{b['total_amount']:>10,.2f}")
            p(f"  {'Brand':<35} {'ml':>5}  {'Btls':>5}  {'Tot.ml':>7}  "
              f"{'MRP/btl':>9}  {'Amount':>10}")
            p(f"  {'─'*35}  {'─'*5}  {'─'*5}  {'─'*7}  {'─'*9}  {'─'*10}")

            for it in b["items"]:
                p(f"  {it['short_name']:<35} "
                  f"{it['ml_per_bottle']:>5}  "
                  f"{it['bottles']:>5}  "
                  f"{it['total_ml']:>7,}  "
                  f"₹{it['mrp_per_bottle']:>8.2f}  "
                  f"₹{it['total_amount']:>9.2f}")

            p(f"  {'─'*35}  {'─'*5}  {'─'*5}  {'─'*7}  {'─'*9}  {'─'*10}")
            p(f"  {'BILL TOTAL':<35}  {'':5}  {'':5}  "
              f"{b['total_ml']:>7,}  {'':9}  "
              f"₹{b['total_amount']:>9.2f}")

            grand_ml  += b["total_ml"]
            grand_amt += b["total_amount"]

        p(f"\n  {label} subtotal: {len(bills)} bill(s)  |  "
          f"{grand_ml:,} ml total  |  ₹{grand_amt:,.2f}")
        p(f"  {line}")

    print_type_section(bills_imfl, "Liquor (IMFL)", IMFL_LIMIT_ML)
    print_type_section(bills_beer, "Beer",          BEER_LIMIT_ML)

    total_bills = len(bills_imfl) + len(bills_beer)
    total_amt   = (sum(b["total_amount"] for b in bills_imfl) +
                   sum(b["total_amount"] for b in bills_beer))
    p(f"\n  TOTAL — {total_bills} bill(s)  |  ₹{total_amt:,.2f}\n")
    p(dline)


# ── Expenses / Cash / UPI section ────────────────────────────

def print_cash_section(p, cash_rows, category_key, label):
    rows = [r for r in cash_rows if r["cash_category"] == category_key]
    if not rows:
        return
    p(f"\n    Cash — {label}:")
    total = Decimal("0")
    for r in rows:
        dval  = r["denomination_value"]
        cnt   = r["count"]
        tamt  = Decimal(str(r["total_amount"]))
        dtype = r["denomination_type"]
        if dtype == "NOTE":
            p(f"      ₹{dval:>4} × {cnt:>5}  =  ₹{tamt:>9,.2f}")
        else:
            p(f"      Coins (₹{dval}) × {cnt:>5}  =  ₹{tamt:>9,.2f}")
        total += tamt
    p(f"      {'─'*38}")
    p(f"      {'Total':<30}  ₹{total:>9,.2f}")


def print_expenses_cash_report(summaries, expenses, cash, shop_name,
                               bill_date, out=None):
    W = 88
    line  = "─" * W
    dline = "═" * W

    def p(*args, **kwargs):
        kwargs.setdefault("file", out)
        print(*args, **kwargs)

    p(f"\n{dline}")
    p(f"  EXPENSES · CASH · UPI  ·  {shop_name}  ·  {bill_date}")
    p(dline)

    shop_totals = {
        "liquor": Decimal("0"), "beer": Decimal("0"), "tips": Decimal("0"),
        "grand": Decimal("0"), "upi": Decimal("0"), "expenses": Decimal("0"),
        "collection": Decimal("0"), "drawer": Decimal("0"), "tips_cash": Decimal("0"),
        "tally": Decimal("0"),
    }

    for s in summaries:
        cid   = s["id"]
        cname = s["name"]
        staff = s["staff_name"] or "—"
        exps  = expenses.get(cid, [])
        cash_rows = cash.get(cid, [])

        p(f"\n  ■ {cname}  —  {staff}")
        p(f"  {line}")

        # ── Expenses (non-UPI) ────────────────────────────────
        non_upi = [e for e in exps if not e["is_upi"]]
        if non_upi:
            p(f"\n    Expenses:")
            exp_total = Decimal("0")
            for e in non_upi:
                label = EXPENSE_LABEL.get(e["category"], e["category"])
                p(f"      {label:<30}  ₹{e['amount']:>9,.2f}")
                exp_total += Decimal(str(e["amount"]))
            p(f"      {'─'*38}")
            p(f"      {'Total Expenses':<30}  ₹{exp_total:>9,.2f}")

        # ── UPI / Google Pay ──────────────────────────────────
        upi_rows = [e for e in exps if e["is_upi"]]
        if upi_rows:
            p(f"\n    UPI / Google Pay:")
            upi_total = Decimal("0")
            for e in upi_rows:
                label = EXPENSE_LABEL.get(e["category"], e["category"])
                ref   = f"  [{e['upi_ref']}]" if e["upi_ref"] else ""
                p(f"      {label:<30}  ₹{e['amount']:>9,.2f}{ref}")
                upi_total += Decimal(str(e["amount"]))
            p(f"      {'─'*38}")
            p(f"      {'Total UPI':<30}  ₹{upi_total:>9,.2f}")

        # ── Cash denominations ────────────────────────────────
        for cat_key, cat_label in [
            ("COLLECTION",  "Collection"),
            ("DRAWER_CASH", "Drawer Cash"),
            ("TIPS_CASH",   "Tips Cash"),
        ]:
            print_cash_section(p, cash_rows, cat_key, cat_label)

        # ── Counter summary ───────────────────────────────────
        p(f"\n    Summary:")
        p(f"      {'Liquor Sale':<30}  ₹{s['liquor_sale_amount']:>9,.2f}")
        p(f"      {'Beer Sale':<30}  ₹{s['beer_sale_amount']:>9,.2f}")
        p(f"      {'Tips':<30}  ₹{s['total_tips']:>9,.2f}")
        p(f"      {'Grand Total (by sale)':<30}  ₹{s['grand_total_by_sale']:>9,.2f}")
        p(f"      {'─'*38}")
        p(f"      {'Google Pay (UPI)':<30}  ₹{s['google_pay_amount']:>9,.2f}")
        p(f"      {'Expenses Total':<30}  ₹{s['expenses_total']:>9,.2f}")
        p(f"      {'Collection':<30}  ₹{s['collection_total']:>9,.2f}")
        p(f"      {'Drawer Cash':<30}  ₹{s['drawer_cash_total']:>9,.2f}")
        p(f"      {'Tips Cash':<30}  ₹{s['tips_cash_total']:>9,.2f}")
        p(f"      {'─'*38}")
        tally = Decimal(str(s["tally_difference"]))
        sign  = "" if tally == 0 else ("+" if tally > 0 else "")
        p(f"      {'Tally Difference':<30}  {sign}₹{tally:>9,.2f}")

        for k, field in [
            ("liquor", "liquor_sale_amount"), ("beer", "beer_sale_amount"),
            ("tips", "total_tips"),           ("grand", "grand_total_by_sale"),
            ("upi", "google_pay_amount"),     ("expenses", "expenses_total"),
            ("collection", "collection_total"), ("drawer", "drawer_cash_total"),
            ("tips_cash", "tips_cash_total"), ("tally", "tally_difference"),
        ]:
            shop_totals[k] += Decimal(str(s[field]))

    # ── Shop total ────────────────────────────────────────────
    p(f"\n  {'─'*W}")
    p(f"  SHOP TOTAL — all counters")
    p(f"  {'─'*W}")
    p(f"    {'Liquor Sale':<30}  ₹{shop_totals['liquor']:>9,.2f}")
    p(f"    {'Beer Sale':<30}  ₹{shop_totals['beer']:>9,.2f}")
    p(f"    {'Tips':<30}  ₹{shop_totals['tips']:>9,.2f}")
    p(f"    {'Grand Total (by sale)':<30}  ₹{shop_totals['grand']:>9,.2f}")
    p(f"    {'─'*38}")
    p(f"    {'Google Pay (UPI)':<30}  ₹{shop_totals['upi']:>9,.2f}")
    p(f"    {'Expenses Total':<30}  ₹{shop_totals['expenses']:>9,.2f}")
    p(f"    {'Collection':<30}  ₹{shop_totals['collection']:>9,.2f}")
    p(f"    {'Drawer Cash':<30}  ₹{shop_totals['drawer']:>9,.2f}")
    p(f"    {'Tips Cash':<30}  ₹{shop_totals['tips_cash']:>9,.2f}")
    p(f"    {'─'*38}")
    t = shop_totals['tally']
    sign = "" if t == 0 else ("+" if t > 0 else "")
    p(f"    {'Tally Difference':<30}  {sign}₹{t:>9,.2f}")
    p(f"\n{dline}\n")


# ── Main ──────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Generate KSBCL daily sales bills")
    ap.add_argument("--date",    required=True,
                    help="Sale date YYYY-MM-DD")
    ap.add_argument("--shop",    default="KLS",
                    help="Shop short_name (default: KLS)")
    ap.add_argument("--no-save", action="store_true",
                    help="Print report only — do not write to DB")
    ap.add_argument("--output", metavar="FILE",
                    help="Save report to this path (default: Reports/{shop}-{date}-bills.txt)")
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

    # Fetch all data
    products   = fetch_sales(cur, shop_id, sale_date)
    summaries  = fetch_counter_summaries(cur, shop_id, sale_date)
    expenses   = fetch_expenses(cur, shop_id, sale_date)
    cash       = fetch_cash(cur, shop_id, sale_date)

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

    report_path = args.output or os.path.join(
        REPORTS_DIR, f"{shop_short}-{sale_date}-bills.txt"
    )
    os.makedirs(os.path.dirname(os.path.abspath(report_path)), exist_ok=True)
    with open(report_path, "w", encoding="utf-8") as f:
        out = _Tee(sys.stdout, f)
        print_report(bills_imfl, bills_beer, shop_name, sale_date, out=out)
        print_expenses_cash_report(summaries, expenses, cash,
                                   shop_name, sale_date, out=out)
    print(f"  Report saved to {report_path}")

    cur.close()
    conn.close()


if __name__ == "__main__":
    main()
