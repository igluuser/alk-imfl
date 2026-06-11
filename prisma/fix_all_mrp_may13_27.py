#!/usr/bin/env python3
"""
fix_all_mrp_may13_27.py — Comprehensive MRP fix + bill regeneration for May 13-27.

Steps:
  1. Revert May-12 DCS to original MRP/SP (undo earlier incorrect fix)
  2. Delete all bills from May-13 onwards (incl. the partially fixed May-13 bills)
  3. For each permit delivery date, update DCS.mrp_per_bottle to the correct
     permit-calculated value (rate/carton + aroed) × 1.10
  4. Regenerate bills for every day May-13 → May-27

FIFO note: DCS for the PERMIT delivery date gets the NEW permit MRP.
  The previous day's DCS carries the OLD MRP. The bill generator uses
  prev_mrp for OB stock and current mrp for newly received stock — exactly
  what KSBCL FIFO requires.
"""

import os, subprocess, sys
from datetime import date, timedelta
from decimal import Decimal

import psycopg2

CONN = dict(host="195.201.119.186", port=5432,
            dbname="imfl_app", user="postgres", password="P@ssword4postgres")
SHOP_ID   = 1
BILL_YEAR = 2026
GENERATE  = os.path.join(os.path.dirname(__file__), "generate_daily_bills.py")

# ── Permit MRP data: date → {master_code: correct_mrp_str} ────────────────────
# Source: KSBCL permit PDFs; formula = (rate/carton + aroed) × 1.10
# ⚠ DK 90ML (mc=156) and Raja Rum 180ML (mc=221) have MRP slightly above the
#   Excel SP — accepted; tiny negative tips reflect the data discrepancy.
PERMIT_MRP = {
    # ── INDBG226001844  13-May ──────────────────────────────────────────────
    "2026-05-13": {
         45: "155.08",   # Bagpiper Deluxe Tetra 180ML
         63: "175.02",   # DSP Black PET 180ML
         69: "110.10",   # Haywards Cheers Tetra 180ML
         79: "50.28",    # Bangalore Whisky 90ML
        108: "135.05",   # Old Tavern Tetra 180ML
        109: "70.27",    # Old Tavern Tetra 90ML
        113: "175.13",   # OC Special ABP 180ML  (rate 7577.85 on this permit)
        145: "105.32",   # Wellington ABP 180ML
        146: "55.41",    # Wellington ABP 90ML
        155: "95.06",    # DK Double Kick ABP 180ML
        156: "50.28",    # DK Double Kick ABP 90ML
        162: "165.10",   # OC Star Supreme ABP 180ML
        169: "85.30",    # OC Star Supreme ABP 90ML
        174: "95.27",    # Black Belt ABP 180ML
        175: "55.16",    # Black Belt ABP 90ML
        206: "135.05",   # Bagpiper XXX Rum Tetra 180ML
        219: "175.02",   # Old Monk Rum ABP 180ML
        220: "90.25",    # Old Monk Rum ABP 90ML
        221: "110.33",   # Raja Super XXX Rum ABP 180ML
        222: "55.16",    # Raja Super XXX Rum ABP 90ML
        255: "105.29",   # Carnival Dry Gin ABP 180ML
    },

    # ── INDBG226002026  14-May (beer + Green Champion) ──────────────────────
    # Rate/AROED not in permit PDFs — using same-product rates from later permits
    "2026-05-14": {
        185: "105.48",   # Green Champion Superior 180ML  (May-21 permit)
        201: "55.48",    # Green Champion Superior 90ML   (May-21 permit)
        366: "180.10",   # KF Strong 650ML                (May-21 permit)
        369: "140.31",   # KF Strong CAN 500ML            (May-29 permit)
        376: "170.41",   # Tuborg Strong 650ML            (May-29 permit)
        383: "250.00",   # Budweiser Magnum 650ML         (May-21 permit)
        384: "200.38",   # Budweiser Magnum CAN 500ML     (Jun-03 permit)
        386: "120.00",   # Power Cool Strong 650ML        (May-21 permit)
        388: "120.26",   # RC Strong 650ML                (May-27 permit)
        390: "95.36",    # RC Strong CAN 500ML            (May-23 permit)
        404: "120.00",   # Bullet Super Strong 650ML      (May-27 permit)
        # 405 Bullet Strong CAN 500ML — rate unknown; skip (leave Excel value)
        # 406 Bullet Strong 330ML    — rate unknown; skip (leave Excel value)
    },

    # ── INDBG226002184  18-May ──────────────────────────────────────────────
    "2026-05-18": {
         30: "145.20",   # 8PM Special ABP 180ML
         45: "155.08",   # Bagpiper Deluxe Tetra 180ML
         63: "175.02",   # DSP Black PET 180ML
         69: "110.10",   # Haywards Cheers Tetra 180ML
         83: "215.20",   # McDowell's No.1 Reserve 180ML
        108: "135.05",   # Old Tavern Tetra 180ML
        113: "175.02",   # OC Special ABP 180ML  (rate 7629.11 on this permit)
        114: "90.25",    # OC Special ABP 90ML
        115: "455.25",   # ORC PET 750ML
        116: "230.37",   # ORC 375ML
        145: "105.32",   # Wellington ABP 180ML
        146: "55.41",    # Wellington ABP 90ML
        155: "95.06",    # DK Double Kick ABP 180ML
        156: "50.28",    # DK Double Kick ABP 90ML
        159: "70.27",    # MaQintosh ABP 90ML
        162: "165.10",   # OC Star Supreme ABP 180ML
        174: "95.27",    # Black Belt ABP 180ML
        186: "135.05",   # Original Traveller ABP 180ML
        187: "70.27",    # Original Traveller ABP 90ML
        219: "175.02",   # Old Monk Rum ABP 180ML
        234: "165.10",   # McDowell's Classic XXX Rum Tetra 180ML
        255: "105.29",   # Carnival Dry Gin ABP 180ML
        332: "30.00",    # Rico Red Wine PET 180ML
        363: "110.11",   # KF Premium Lager Beer 650ML
        379: "205.04",   # Budweiser Premium 650ML
    },

    # ── INDBG226002372  21-May ──────────────────────────────────────────────
    "2026-05-21": {
         31: "75.35",    # 8PM Special ABP 90ML
         43: "645.47",   # Bagpiper Deluxe PET 750ML
         45: "155.08",   # Bagpiper Deluxe Tetra 180ML
         69: "110.10",   # Haywards Cheers Tetra 180ML
         73: "215.20",   # Imperial Blue ABP 180ML
        109: "70.27",    # Old Tavern Tetra 90ML
        145: "105.32",   # Wellington ABP 180ML
        146: "55.41",    # Wellington ABP 90ML
        155: "95.06",    # DK Double Kick ABP 180ML
        174: "95.27",    # Black Belt ABP 180ML
        185: "105.48",   # Green Champion Superior 180ML
        201: "55.48",    # Green Champion Superior 90ML
        206: "135.05",   # Bagpiper XXX Rum Tetra 180ML
        235: "85.30",    # McDowell's Classic XXX Rum ABP 90ML
        363: "110.11",   # KF Premium Lager 650ML
        366: "180.10",   # KF Strong 650ML
        367: "100.41",   # KF Strong 330ML
        383: "250.00",   # Budweiser Magnum 650ML
        386: "120.00",   # Power Cool Strong 650ML
    },

    # ── INDBG226002519  23-May ──────────────────────────────────────────────
    "2026-05-23": {
         45: "155.08",   # Bagpiper Deluxe Tetra 180ML
         62: "425.21",   # DSP Black 375ML
         64: "90.25",    # DSP Black PET 90ML
         69: "110.10",   # Haywards Cheers Tetra 180ML
         72: "445.37",   # Imperial Blue 375ML
        115: "455.25",   # ORC PET 750ML
        117: "110.10",   # ORC ABP 180ML
        145: "105.32",   # Wellington ABP 180ML
        151: "385.21",   # Black Belt PET 750ML
        155: "95.06",    # DK Double Kick ABP 180ML
        156: "50.28",    # DK Double Kick ABP 90ML
        174: "95.27",    # Black Belt ABP 180ML
        187: "70.27",    # Original Traveller ABP 90ML
        206: "135.05",   # Bagpiper XXX Rum Tetra 180ML
        219: "175.02",   # Old Monk Rum ABP 180ML
        220: "90.25",    # Old Monk Rum ABP 90ML
        221: "110.33",   # Raja Super XXX Rum ABP 180ML
        255: "105.29",   # Carnival Dry Gin ABP 180ML
        256: "55.39",    # Carnival Dry Gin ABP 90ML
        274: "215.20",   # Romanov Orange Vodka ABP 180ML
        378: "130.41",   # Tuborg Strong CAN 500ML
        383: "250.00",   # Budweiser Magnum 650ML
    },

    # ── INDBG226002556  25-May  (rationed: only Bullet) ─────────────────────
    "2026-05-25": {
        404: "120.00",   # Bullet Super Strong 650ML
    },

    # ── INDBG226002700  27-May ──────────────────────────────────────────────
    "2026-05-27": {
         30: "145.20",   # 8PM Special ABP 180ML
         45: "155.08",   # Bagpiper Deluxe Tetra 180ML
        108: "135.05",   # Old Tavern Tetra 180ML
        116: "230.37",   # ORC 375ML
        145: "105.32",   # Wellington ABP 180ML
        146: "55.41",    # Wellington ABP 90ML
        155: "95.06",    # DK Double Kick ABP 180ML
        174: "95.27",    # Black Belt ABP 180ML
        187: "70.27",    # Original Traveller ABP 90ML
        204: "55.16",    # Carnival XXX Rum ABP 90ML
        206: "135.05",   # Bagpiper XXX Rum Tetra 180ML
        255: "105.29",   # Carnival Dry Gin ABP 180ML
        328: "40.30",    # Goan's Wine ABP 180ML
        366: "180.10",   # KF Strong 650ML
        383: "250.00",   # Budweiser Magnum 650ML
        388: "120.26",   # RC Strong 650ML
        404: "120.00",   # Bullet Super Strong 650ML
    },
}

# ── May-12 DCS original values to RESTORE (undo earlier incorrect fix) ────────
MAY12_RESTORE = {
     63: ("185.00", "185.00"),   # DSP Black 180ML
    108: ("115.00", "120.00"),   # Old Tavern 180ML
    113: ("150.00", "150.00"),   # OC Special 180ML
    156: ("40.00",  "45.00"),    # DK 90ML
    174: ("80.00",  "80.00"),    # Black Belt 180ML
    206: ("115.00", "115.00"),   # Bagpiper XXX Rum 180ML
    219: ("155.00", "160.00"),   # Old Monk 180ML
    221: ("95.00",  "110.00"),   # Raja Rum 180ML
    222: ("55.00",  "60.00"),    # Raja Rum 90ML
    255: ("95.00",  "110.00"),   # Carnival Gin 180ML
}

# ── Bill generation: dates with actual sales (May-28 had zero sales) ──────────
BILL_DATES = [
    date(2026, 5, 13), date(2026, 5, 14), date(2026, 5, 15), date(2026, 5, 16),
    date(2026, 5, 17), date(2026, 5, 18), date(2026, 5, 19), date(2026, 5, 20),
    date(2026, 5, 21), date(2026, 5, 22), date(2026, 5, 23), date(2026, 5, 24),
    date(2026, 5, 25), date(2026, 5, 26), date(2026, 5, 27),
]


def revert_may12(cur, mc_to_pid):
    """Restore May-12 DCS mrp/sp to original Excel values."""
    count = 0
    for mc, (mrp, sp) in MAY12_RESTORE.items():
        if mc not in mc_to_pid:
            continue
        pid = mc_to_pid[mc]
        cur.execute("""
            UPDATE "DailyCounterStock"
            SET mrp_per_bottle = %s, selling_price_per_bottle = %s
            WHERE stock_date = '2026-05-12' AND product_id = %s
        """, (mrp, sp, pid))
        count += cur.rowcount
    print(f"  ✓ Reverted May-12 DCS: {count} rows restored to original MRP/SP")


def delete_bills_from_may13(cur):
    """Delete all bills from May-13 onwards."""
    cur.execute("""
        DELETE FROM "DailySalesBillItem" dsbi
        USING "DailySalesBill" dsb
        WHERE dsbi.bill_id = dsb.id
          AND dsb.bill_date >= '2026-05-13'
          AND dsb.shop_id = %s
    """, (SHOP_ID,))
    items_deleted = cur.rowcount
    cur.execute("""
        DELETE FROM "DailySalesBill"
        WHERE bill_date >= '2026-05-13' AND shop_id = %s
    """, (SHOP_ID,))
    bills_deleted = cur.rowcount
    print(f"  ✓ Deleted {bills_deleted} bills ({items_deleted} items) from May-13+")


def fix_dcs_mrp(cur, mc_to_pid, dry_run=False):
    """Update DCS mrp_per_bottle + sale_amount_mrp for all permit dates."""
    total_rows = 0
    for date_str, mc_mrp in sorted(PERMIT_MRP.items()):
        rows_this_date = 0
        for mc, new_mrp_str in mc_mrp.items():
            if mc not in mc_to_pid:
                print(f"    WARNING: mc={mc} not in Product table (date={date_str})")
                continue
            pid = mc_to_pid[mc]
            new_mrp = Decimal(new_mrp_str)

            # Read current DCS rows for this product on this date
            cur.execute("""
                SELECT dcs.id, dcs.sold_btls, dcs.mrp_per_bottle,
                       dcs.selling_price_per_bottle, dcs.total_sale_amount
                FROM "DailyCounterStock" dcs
                WHERE dcs.stock_date = %s AND dcs.product_id = %s
                ORDER BY dcs.counter_id
            """, (date_str, pid))
            dcs_rows = cur.fetchall()

            for dcs_id, sold, old_mrp, cur_sp, total_sale in dcs_rows:
                new_amt  = Decimal(str(sold)) * new_mrp
                new_tips = Decimal(str(total_sale)) - new_amt  # may be negative if SP < MRP

                if not dry_run:
                    cur.execute("""
                        UPDATE "DailyCounterStock"
                        SET mrp_per_bottle  = %s,
                            sale_amount_mrp = %s,
                            tips_amount     = %s
                        WHERE id = %s
                    """, (new_mrp_str, str(new_amt), str(new_tips), dcs_id))
                rows_this_date += 1

        total_rows += rows_this_date
        print(f"  {date_str}: {rows_this_date} DCS rows updated")

    print(f"  ✓ Total DCS rows updated: {total_rows}")


def regenerate_bills(dry_run=False):
    """Run generate_daily_bills.py for each day in BILL_DATES."""
    if dry_run:
        print("  DRY RUN — skipping bill regeneration")
        return
    for d in BILL_DATES:
        d_str = d.isoformat()
        result = subprocess.run(
            [sys.executable, GENERATE, "--date", d_str],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            print(f"  ERROR on {d_str}:\n{result.stderr[:400]}")
        else:
            # Show just the summary line
            for line in result.stdout.strip().splitlines():
                if line.strip():
                    print(f"  {line}")


def main(dry_run=False):
    con = psycopg2.connect(**CONN)
    cur = con.cursor()

    # Build master_code → product_id map
    cur.execute('SELECT master_code, id FROM "Product" WHERE master_code IS NOT NULL')
    mc_to_pid = {row[0]: row[1] for row in cur.fetchall()}

    print("=== Step 1: Revert May-12 DCS to original MRP/SP ===")
    revert_may12(cur, mc_to_pid)

    print("\n=== Step 2: Delete all bills from May-13 onwards ===")
    delete_bills_from_may13(cur)

    print("\n=== Step 3: Fix DCS MRP for all permit dates May-13→27 ===")
    fix_dcs_mrp(cur, mc_to_pid, dry_run=dry_run)

    if not dry_run:
        con.commit()
        print("\n  DB changes committed.")
    else:
        con.rollback()
        print("\n  DRY RUN — all DB changes rolled back.")

    cur.close()
    con.close()

    print("\n=== Step 4: Regenerate bills May-13 → May-27 ===")
    regenerate_bills(dry_run=dry_run)

    print("\nDone.")


if __name__ == "__main__":
    dry_run = "--dry-run" in sys.argv
    if dry_run:
        print("=== DRY RUN MODE ===\n")
    main(dry_run=dry_run)
