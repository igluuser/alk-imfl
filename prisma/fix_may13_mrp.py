#!/usr/bin/env python3
"""
fix_may13_mrp.py — Fix wrong MRP values in DailyCounterStock for May 13, 2026.

The Excel had MRP set to price_base (rate/bottle + AROED) without the ×1.10 excise
markup for most permit products. Correct values come from permit INDBG226001844.

For two products (DK 90ML mc=156, Raja Rum 180ML mc=221), the calculated MRP slightly
exceeds the actual SP recorded (50.28>50 and 110.33>110), so we cap MRP at SP.
"""

import psycopg2
from decimal import Decimal

CONN = dict(host="195.201.119.186", port=5432,
            dbname="imfl_app", user="postgres", password="P@ssword4postgres")

SALE_DATE = "2026-05-13"

# master_code → correct MRP (from permit INDBG226001844 formula, capped at SP where needed)
CORRECT_MRP = {
     45: Decimal("155.08"),   # Bagpiper Deluxe Tetra 180ML
     63: Decimal("175.02"),   # DSP Black PET 180ML
     69: Decimal("110.10"),   # Haywards Cheers Tetra 180ML
     79: Decimal("50.28"),    # Bangalore Whisky 90ML
    108: Decimal("135.05"),   # Old Tavern Tetra 180ML
    109: Decimal("70.27"),    # Old Tavern Tetra 90ML
    113: Decimal("175.13"),   # OC Special ABP 180ML
    145: Decimal("105.32"),   # Wellington ABP 180ML
    146: Decimal("55.41"),    # Wellington ABP 90ML
    155: Decimal("95.06"),    # DK Double Kick ABP 180ML
    156: Decimal("50.00"),    # DK Double Kick ABP 90ML  (calc=50.28, capped at SP=50)
    162: Decimal("165.10"),   # OC Star Supreme ABP 180ML
    169: Decimal("85.30"),    # OC Star Supreme ABP 90ML
    174: Decimal("95.27"),    # Black Belt ABP 180ML
    175: Decimal("55.16"),    # Black Belt ABP 90ML
    206: Decimal("135.05"),   # Bagpiper XXX Rum Tetra 180ML
    219: Decimal("175.02"),   # Old Monk Rum ABP 180ML
    220: Decimal("90.25"),    # Old Monk Rum ABP 90ML
    221: Decimal("110.00"),   # Raja Super XXX Rum ABP 180ML  (calc=110.33, capped at SP=110)
    222: Decimal("55.16"),    # Raja Super XXX Rum ABP 90ML
    255: Decimal("105.29"),   # Carnival Dry Gin ABP 180ML
}


def fix_mrp(dry_run=False):
    con = psycopg2.connect(**CONN)
    cur = con.cursor()

    # Get product_id for each master_code
    cur.execute("SELECT master_code, id, short_name FROM \"Product\" WHERE master_code = ANY(%s)",
                (list(CORRECT_MRP.keys()),))
    mc_to_pid = {row[0]: (row[1], row[2]) for row in cur.fetchall()}

    print(f"{'mc':>5} {'name':<28} {'ctr':>4} {'old_mrp':>8} {'new_mrp':>8} "
          f"{'sold':>5} {'old_amt':>9} {'new_amt':>9} {'tips':>8}")
    print("-" * 95)

    updates = []
    for mc, new_mrp in sorted(CORRECT_MRP.items()):
        if mc not in mc_to_pid:
            print(f"  WARNING: master_code {mc} not found in Product table")
            continue
        pid, nm = mc_to_pid[mc]

        cur.execute("""
            SELECT dcs.id, dcs.counter_id, dcs.mrp_per_bottle, dcs.sold_btls,
                   dcs.sale_amount_mrp, dcs.total_sale_amount
            FROM "DailyCounterStock" dcs
            WHERE dcs.stock_date = %s AND dcs.product_id = %s
            ORDER BY dcs.counter_id
        """, (SALE_DATE, pid))
        rows = cur.fetchall()

        for dcs_id, ctr, old_mrp, sold, old_amt, total_sale in rows:
            new_amt   = Decimal(str(sold)) * new_mrp
            new_tips  = Decimal(str(total_sale)) - new_amt
            print(f"{mc:>5} {nm:<28} ctr{ctr} {float(old_mrp):>8.2f} {float(new_mrp):>8.2f} "
                  f"{sold:>5} {float(old_amt):>9.2f} {float(new_amt):>9.2f} {float(new_tips):>8.2f}")
            updates.append((new_mrp, new_amt, new_tips, dcs_id))

    print(f"\n{len(updates)} DailyCounterStock rows to update.")

    if dry_run:
        print("DRY RUN — no changes written.")
        con.close()
        return

    for new_mrp, new_amt, new_tips, dcs_id in updates:
        cur.execute("""
            UPDATE "DailyCounterStock"
            SET mrp_per_bottle   = %s,
                sale_amount_mrp  = %s,
                tips_amount      = %s
            WHERE id = %s
        """, (str(new_mrp), str(new_amt), str(new_tips), dcs_id))

    con.commit()
    cur.close()
    con.close()
    print("Done — DailyCounterStock updated.")


if __name__ == "__main__":
    import sys
    dry_run = "--dry-run" in sys.argv
    fix_mrp(dry_run=dry_run)
