#!/usr/bin/env python3
"""
fix_may12_mrp_ref.py — Fix May 12 DCS mrp_per_bottle and selling_price_per_bottle
for products carried into May 13 as opening balance.

This ONLY updates the reference MRP/SP fields on May 12 DCS (not sale amounts or
totals) so that the May 13 FIFO bill generator uses the correct prev_mrp values.
May 12 bill PDFs and cash totals are NOT touched.
"""

import psycopg2
from decimal import Decimal

CONN = dict(host="195.201.119.186", port=5432,
            dbname="imfl_app", user="postgres", password="P@ssword4postgres")

DATE_12 = "2026-05-12"
DATE_13 = "2026-05-13"

# Products with OB-sold stock on May 13 whose prev_mrp was wrong.
# Correct MRP and SP (taken from corrected May 13 DCS values).
# mc → (correct_mrp, correct_sp)
CORRECT = {
     63: (Decimal("175.02"), Decimal("180.00")),  # DSP Black 180ML
    108: (Decimal("135.05"), Decimal("140.00")),  # Old Tavern 180ML
    113: (Decimal("175.13"), Decimal("180.00")),  # OC Special 180ML
    156: (Decimal("50.00"),  Decimal("50.00")),   # DK 90ML (capped at SP)
    174: (Decimal("95.27"),  Decimal("100.00")),  # Black Belt 180ML
    206: (Decimal("135.05"), Decimal("140.00")),  # Bagpiper XXX Rum 180ML
    219: (Decimal("175.02"), Decimal("180.00")),  # Old Monk 180ML
    221: (Decimal("110.00"), Decimal("110.00")),  # Raja Rum 180ML (capped at SP)
    222: (Decimal("55.16"),  Decimal("60.00")),   # Raja Rum 90ML
    255: (Decimal("105.29"), Decimal("110.00")),  # Carnival Gin 180ML
}


def fix_prev_mrp(dry_run=False):
    con = psycopg2.connect(**CONN)
    cur = con.cursor()

    cur.execute("SELECT master_code, id, short_name FROM \"Product\" WHERE master_code = ANY(%s)",
                (list(CORRECT.keys()),))
    mc_to_pid = {row[0]: (row[1], row[2]) for row in cur.fetchall()}

    print(f"{'mc':>5} {'name':<26} {'ctr':>4} {'old_mrp':>8} {'old_sp':>8} {'new_mrp':>8} {'new_sp':>8}")
    print("-" * 75)

    updates = []
    for mc, (new_mrp, new_sp) in sorted(CORRECT.items()):
        pid, nm = mc_to_pid[mc]
        cur.execute("""
            SELECT dcs.id, dcs.counter_id, dcs.mrp_per_bottle, dcs.selling_price_per_bottle
            FROM "DailyCounterStock" dcs
            WHERE dcs.stock_date = %s AND dcs.product_id = %s
            ORDER BY dcs.counter_id
        """, (DATE_12, pid))
        for dcs_id, ctr, old_mrp, old_sp in cur.fetchall():
            print(f"{mc:>5} {nm:<26} ctr{ctr} {float(old_mrp):>8.2f} {float(old_sp):>8.2f} "
                  f"{float(new_mrp):>8.2f} {float(new_sp):>8.2f}")
            updates.append((new_mrp, new_sp, dcs_id))

    print(f"\n{len(updates)} May-12 DCS rows to update (mrp+sp reference only).")

    if dry_run:
        print("DRY RUN — no changes written.")
        con.close()
        return

    for new_mrp, new_sp, dcs_id in updates:
        cur.execute("""
            UPDATE "DailyCounterStock"
            SET mrp_per_bottle = %s,
                selling_price_per_bottle = %s
            WHERE id = %s
        """, (str(new_mrp), str(new_sp), dcs_id))

    con.commit()
    cur.close()
    con.close()
    print("Done — May-12 DCS mrp/sp reference updated.")


if __name__ == "__main__":
    import sys
    dry_run = "--dry-run" in sys.argv
    fix_prev_mrp(dry_run=dry_run)
