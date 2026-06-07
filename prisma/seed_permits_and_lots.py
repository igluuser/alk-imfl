#!/usr/bin/env python3
"""
seed_permits_and_lots.py — Populate Permit, PermitItem, and StockLot tables.

Creates:
  1. Permit records for each permit date
  2. PermitItem records with correct MRP per product per permit
  3. StockLot for each PermitItem (new stock at permit MRP)
  4. Opening-Balance StockLot per product (old stock, lot_date = first permit date)

MRP values come from the permit formula (stored here), NOT from the Excel Rate Matrix.
Formula: mrp_per_bottle = (rate_per_cb / carton_size + aroed_per_bottle) * 1.10

Run:
    python3 seed_permits_and_lots.py           # live
    python3 seed_permits_and_lots.py --dry-run # preview only
"""

import sys
import psycopg2
from decimal import Decimal
from datetime import date, timedelta

CONN = dict(host='195.201.119.186', port=5432, dbname='imfl_app',
            user='postgres', password='P@ssword4postgres')
SHOP_ID = 1

# ── Permit metadata ───────────────────────────────────────────────────────────
PERMIT_META = {
    date(2026, 5, 13): {"indent_no": "INDBG226001844"},
    date(2026, 5, 14): {"indent_no": "INDBG226002026"},
    date(2026, 5, 18): {"indent_no": "INDBG226002184"},
    date(2026, 5, 21): {"indent_no": "INDBG226002372"},
    date(2026, 5, 23): {"indent_no": "INDBG226002519"},
    date(2026, 5, 25): {"indent_no": "INDBG226002556"},
    date(2026, 5, 27): {"indent_no": "INDBG226002700"},
    date(2026, 5, 30): {"indent_no": "INDBG226002835"},
}

# ── Bottles received per permit date (master_code → bottles) ─────────────────
PERMIT_REC = {
    date(2026, 5, 13): {
         45:  6*48,  63:  3*48,  69:  1*48,  79:  3*96,
        108:  1*48, 109:  1*96, 113:  1*48, 145:  9*48,
        146:  3*96, 155:  3*48, 156: 49*96, 162:  1*48,
        169:  1*96, 174:  1*48, 175:  1*96, 206:  3*48,
        219:  1*48, 220:  1*96, 221:  1*48, 222:  1*96,
        255:  1*48,
    },
    date(2026, 5, 14): {
        185: 12*48, 201:  2*96, 366: 10*12, 369:  1*24,
        404: 20*12, 383:  2*12, 384:  1*24, 390:  2*24,
        388: 15*12, 376: 10*12, 406:  1*24, 405:  3*24,
        386:  3*12,
    },
    date(2026, 5, 18): {
         30:  1*48,  45:  3*48,  63:  2*48,  69:  1*48,
         83:  1*48, 108:  3*48, 113:  1*48, 114:  1*96,
        115:  1*12, 116:  1*24, 145: 15*48, 146:  3*96,
        155:  3*48, 156: 72*96, 159:  3*96, 162:  3*48,
        174:  3*48, 186:  1*48, 187:  1*96, 219:  1*48,
        234:  3*48, 255:  3*48, 332:  1*48, 363:  3*12,
        379:  1*12,
    },
    date(2026, 5, 21): {
         31:  1*96,  43:  1*12,  45:  1*48,  69:  1*48,
         73:  1*48, 109:  2*96, 145:  6*48, 146:  1*96,
        155:  6*48, 174:  1*48, 185:  9*48, 201:  3*96,
        206:  1*48, 235:  1*96, 363: 10*12, 366: 12*12,
        367:  2*24, 383:  1*12, 386: 12*12,
    },
    date(2026, 5, 23): {
         45:  2*48,  62:  1*24,  64:  1*96,  69:  1*48,
         72:  1*24, 115:  3*12, 117: 30*48, 145:  6*48,
        151:  1*12, 155:  6*48, 156:  6*96, 174:  1*48,
        187:  1*96, 206:  1*48, 219:  1*48, 220:  1*96,
        221:  1*48, 255:  2*48, 256:  1*96, 274:  1*48,
        378:  3*24, 383:  1*12,
    },
    date(2026, 5, 25): {404:  9*12},
    date(2026, 5, 27): {
         30:  1*48,  45:  2*48, 108:  1*48, 116:  2*24,
        145:  9*48, 146:  3*96, 155:  9*48, 174:  3*48,
        187:  1*96, 204:  1*96, 206:  1*48, 255:  3*48,
        328:  1*48, 366: 15*12, 383:  1*12, 388: 15*12,
        404: 15*12,
    },
    date(2026, 5, 30): {
         45:  3*48,  58:  1*48,  69:  1*48, 108:  1*48,
        114:  1*96, 155:  6*48, 156: 39*96, 174:  1*48,
        203:  1*48, 219:  1*48, 367:  1*24, 369:  3*24,
        376: 15*12, 379:  1*12, 383:  1*12, 386:  6*12,
        390:  2*24, 404:  9*12,
    },
}

# ── Permit MRP (computed from permit formula, stored here permanently) ────────
# Formula: (rate_per_cb / carton_size + aroed_per_bottle) * 1.10
PERMIT_MRP = {
    date(2026, 5, 13): {
         45: "155.08",  63: "175.02",  69: "110.10",  79: "50.28",
        108: "135.05", 109:  "70.27", 113: "175.13", 145: "105.32",
        146:  "55.41", 155:  "95.06", 156:  "50.28", 162: "165.10",
        169:  "85.30", 174:  "95.27", 175:  "55.16", 206: "135.05",
        219: "175.02", 220:  "90.25", 221: "110.33", 222:  "55.16",
        255: "105.29",
    },
    date(2026, 5, 14): {
        185: "105.48", 201:  "55.48", 366: "180.10", 369: "140.31",
        376: "170.41", 383: "250.00", 384: "200.38", 386: "120.00",
        388: "120.26", 390:  "95.36", 404: "120.00",
        # 405 (Bullet Strong CAN 500ML), 406 (Bullet 330ML) — MRP not confirmed, skipped
    },
    date(2026, 5, 18): {
         30: "145.20",  45: "155.08",  63: "175.02",  69: "110.10",
         83: "215.20", 108: "135.05", 113: "175.02", 114:  "90.25",
        115: "455.25", 116: "230.37", 145: "105.32", 146:  "55.41",
        155:  "95.06", 156:  "50.28", 159:  "70.27", 162: "165.10",
        174:  "95.27", 186: "135.05", 187:  "70.27", 219: "175.02",
        234: "165.10", 255: "105.29", 332:  "30.00", 363: "110.11",
        379: "205.04",
    },
    date(2026, 5, 21): {
         31:  "75.35",  43: "645.47",  45: "155.08",  69: "110.10",
         73: "215.20", 109:  "70.27", 145: "105.32", 146:  "55.41",
        155:  "95.06", 174:  "95.27", 185: "105.48", 201:  "55.48",
        206: "135.05", 235:  "85.30", 363: "110.11", 366: "180.10",
        367: "100.41", 383: "250.00", 386: "120.00",
    },
    date(2026, 5, 23): {
         45: "155.08",  62: "425.21",  64:  "90.25",  69: "110.10",
         72: "445.37", 115: "455.25", 117: "110.10", 145: "105.32",
        151: "385.21", 155:  "95.06", 156:  "50.28", 174:  "95.27",
        187:  "70.27", 206: "135.05", 219: "175.02", 220:  "90.25",
        221: "110.33", 255: "105.29", 256:  "55.39", 274: "215.20",
        378: "130.41", 383: "250.00",
    },
    date(2026, 5, 25): {404: "120.00"},
    date(2026, 5, 27): {
         30: "145.20",  45: "155.08", 108: "135.05", 116: "230.37",
        145: "105.32", 146:  "55.41", 155:  "95.06", 174:  "95.27",
        187:  "70.27", 204:  "55.16", 206: "135.05", 255: "105.29",
        328:  "40.30", 366: "180.10", 383: "250.00", 388: "120.26",
        404: "120.00",
    },
    date(2026, 5, 30): {},   # MRP to be entered when permit is seeded
}


def get_product_map(cur):
    """Returns {master_code: (product_id, short_name)}."""
    cur.execute('SELECT master_code, id, short_name FROM "Product"')
    return {r[0]: (r[1], r[2]) for r in cur.fetchall()}


def get_old_mrp(cur, product_id, before_date):
    """Returns DCS mrp_per_bottle from the most recent day before before_date."""
    for d in range(1, 30):
        check = before_date - timedelta(days=d)
        cur.execute("""
            SELECT mrp_per_bottle FROM "DailyCounterStock"
            WHERE product_id = %s AND stock_date = %s
            LIMIT 1
        """, (product_id, check))
        row = cur.fetchone()
        if row:
            return Decimal(str(row[0]))
    return None


def get_ob_qty(cur, product_id, permit_date):
    """
    Returns total old stock at start of permit_date:
      counter DCS opening_balance_btls + GodownStock opening_balance_btls
    """
    cur.execute("""
        SELECT COALESCE(SUM(opening_balance_btls), 0)
        FROM "DailyCounterStock"
        WHERE product_id = %s AND stock_date = %s
    """, (product_id, permit_date))
    counter_ob = cur.fetchone()[0] or 0

    cur.execute("""
        SELECT COALESCE(opening_balance_btls, 0)
        FROM "GodownStock"
        WHERE product_id = %s AND shop_id = %s AND stock_date = %s
    """, (product_id, SHOP_ID, permit_date))
    row = cur.fetchone()
    godown_ob = row[0] if row else 0

    return int(counter_ob) + int(godown_ob)


def seed(dry_run=False):
    con = psycopg2.connect(**CONN)
    cur = con.cursor()

    prod_map = get_product_map(cur)

    # ── Step 1: Clear existing data ─────────────────────────────────────────
    print("Step 1: Clearing existing Permit / PermitItem / StockLot data …")
    cur.execute('DELETE FROM "StockLot" WHERE shop_id = %s', (SHOP_ID,))
    cur.execute("""
        DELETE FROM "PermitItem"
        WHERE permit_id IN (SELECT id FROM "Permit" WHERE shop_id = %s)
    """, (SHOP_ID,))
    cur.execute('DELETE FROM "Permit" WHERE shop_id = %s', (SHOP_ID,))
    print("  Cleared.")

    # ── Step 2: Build first-permit-date map per master_code ──────────────────
    # first_permit_date[mc] = earliest date this product appears in any permit
    first_permit_date = {}
    for d in sorted(PERMIT_REC.keys()):
        for mc in PERMIT_REC[d]:
            if mc not in first_permit_date:
                first_permit_date[mc] = d

    # ── Step 3: Opening Balance StockLots ────────────────────────────────────
    print("\nStep 3: Creating Opening Balance StockLots …")
    ob_created = 0
    for mc, first_d in sorted(first_permit_date.items()):
        if mc not in prod_map:
            print(f"  WARNING: master_code {mc} not in Product table — skipped")
            continue
        pid, name = prod_map[mc]

        ob_qty  = get_ob_qty(cur, pid, first_d)
        old_mrp = get_old_mrp(cur, pid, first_d)
        if old_mrp is None:
            print(f"  WARNING: no prev DCS for mc={mc} {name} — skipped OB lot")
            continue

        lot_date = first_d  # OB lot belongs to first permit date so May 12 bills don't consume it

        if ob_qty > 0:
            print(f"  OB lot  mc={mc:>4} {name:<30} {lot_date}  mrp={old_mrp}  qty={ob_qty}")
            if not dry_run:
                cur.execute("""
                    INSERT INTO "StockLot"
                        (shop_id, product_id, mrp_per_bottle, initial_qty,
                         consumed_qty, lot_date, source, permit_item_id)
                    VALUES (%s, %s, %s, %s, 0, %s, 'OPENING_BALANCE', NULL)
                """, (SHOP_ID, pid, str(old_mrp), ob_qty, lot_date))
            ob_created += 1
        else:
            print(f"  OB=0    mc={mc:>4} {name:<30} {first_d} — no OB lot needed")

    print(f"  {ob_created} OB lots created.")

    # ── Step 4: Permit + PermitItem + Permit StockLots ────────────────────────
    print("\nStep 4: Creating Permit / PermitItem / StockLot records …")
    for pdate in sorted(PERMIT_REC.keys()):
        meta    = PERMIT_META.get(pdate, {})
        indent  = meta.get("indent_no", "UNKNOWN")
        mrp_map = PERMIT_MRP.get(pdate, {})
        rec_map = PERMIT_REC[pdate]

        print(f"\n  Permit {pdate}  {indent}")

        if not dry_run:
            cur.execute("""
                INSERT INTO "Permit"
                    (shop_id, indent_no, invoice_no, permit_date, status,
                     total_indent_cbs, total_indent_amount,
                     total_cnf_cbs, total_cnf_amount, has_rationed_items,
                     created_at, updated_at)
                VALUES (%s, %s, %s, %s, 'RECEIVED',
                        0, 0, 0, 0, FALSE,
                        NOW(), NOW())
                RETURNING id
            """, (SHOP_ID, indent, indent, pdate))
            permit_id = cur.fetchone()[0]
        else:
            permit_id = None

        items_created = 0
        for mc, recv_btls in sorted(rec_map.items()):
            if mc not in prod_map:
                print(f"    WARNING: mc={mc} not in Product — skipped")
                continue
            pid, name = prod_map[mc]

            mrp_str = mrp_map.get(mc)
            if mrp_str is None:
                print(f"    SKIP (no MRP): mc={mc:>4} {name:<30} recv={recv_btls}")
                continue

            mrp = Decimal(mrp_str)
            print(f"    mc={mc:>4} {name:<30} recv={recv_btls:>5}  mrp={mrp}")

            if not dry_run:
                cur.execute("""
                    INSERT INTO "PermitItem"
                        (permit_id, product_id, rate_per_cb, carton_size_snapshot,
                         indent_cbs, indent_btls, indent_amount,
                         cnf_cbs, cnf_btls, cnf_amount,
                         is_rationed, mrp_per_bottle, received_btls)
                    VALUES (%s, %s, 0, 0, 0, 0, 0, 0, %s, 0, FALSE, %s, %s)
                    RETURNING id
                """, (permit_id, pid, recv_btls, str(mrp), recv_btls))
                pi_id = cur.fetchone()[0]

                cur.execute("""
                    INSERT INTO "StockLot"
                        (shop_id, product_id, mrp_per_bottle, initial_qty,
                         consumed_qty, lot_date, source, permit_item_id)
                    VALUES (%s, %s, %s, %s, 0, %s, 'PERMIT', %s)
                """, (SHOP_ID, pid, str(mrp), recv_btls, pdate, pi_id))

            items_created += 1

        print(f"    → {items_created} items created")

    # ── Commit ────────────────────────────────────────────────────────────────
    if dry_run:
        print("\nDRY RUN — no changes written.")
        con.rollback()
    else:
        con.commit()
        print("\nAll changes committed.")

    cur.close()
    con.close()

    # ── Summary ───────────────────────────────────────────────────────────────
    if not dry_run:
        con2 = psycopg2.connect(**CONN)
        c2 = con2.cursor()
        c2.execute('SELECT COUNT(*) FROM "Permit" WHERE shop_id=%s', (SHOP_ID,))
        np = c2.fetchone()[0]
        c2.execute("""
            SELECT COUNT(*) FROM "PermitItem" pi
            JOIN "Permit" p ON p.id=pi.permit_id WHERE p.shop_id=%s
        """, (SHOP_ID,))
        ni = c2.fetchone()[0]
        c2.execute('SELECT COUNT(*) FROM "StockLot" WHERE shop_id=%s', (SHOP_ID,))
        nl = c2.fetchone()[0]
        print(f"\nFinal counts: Permits={np}  PermitItems={ni}  StockLots={nl}")
        c2.close()
        con2.close()


if __name__ == "__main__":
    dry_run = "--dry-run" in sys.argv
    seed(dry_run=dry_run)
