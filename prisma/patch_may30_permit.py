#!/usr/bin/env python3
"""
patch_may30_permit.py — Add May 30 Permit + PermitItem + StockLot records.
Does NOT clear existing lots — appends only. Run after May 17-29 bills are done.
"""
import psycopg2
from decimal import Decimal
from datetime import date

CONN = dict(host="195.201.119.186", port=5432,
            dbname="imfl_app", user="postgres", password="P@ssword4postgres")
SHOP_ID = 1

PERMIT_DATE = date(2026, 5, 30)
INDENT_NO   = "INDBG226002835"

# master_code → (received_btls, mrp_per_bottle)
MAY30 = {
     45: (3*48,  "155.08"),  # Bag Piper 180ML
     58: (1*48,  "165.10"),  # DSP (SADA) 180ML
     69: (1*48,  "110.10"),  # HaywardsCheers 180ML
    108: (1*48,  "135.05"),  # Old Tavern 180ML
    114: (1*96,   "90.25"),  # Officers Choice 90ML
    155: (6*48,   "95.06"),  # DK 180ML
    156: (39*96,  "50.28"),  # DK 90ML
    174: (1*48,   "95.27"),  # Black Belt 180ML
    203: (1*48,   "95.27"),  # Carnival Rum 180ML
    219: (1*48,  "175.02"),  # Old Monk Rum 180ML
    367: (1*24,  "100.41"),  # KF Strong 330ML
    369: (3*24,  "140.31"),  # KF Strong 500ML
    376: (15*12, "170.41"),  # Tuborg Strong 650ML
    379: (1*12,  "205.04"),  # Budweiser Premium 650ML
    383: (1*12,  "250.00"),  # Budweiser Magnum 650ML
    386: (6*12,  "120.00"),  # Power Cool 650ML
    390: (2*24,   "95.36"),  # Royal Challenge 500ML
    404: (9*12,  "120.00"),  # Bullet Strong 650ML
}


def patch(dry_run=False):
    con = psycopg2.connect(**CONN)
    cur = con.cursor()

    # Check if May 30 permit already exists
    cur.execute('SELECT id FROM "Permit" WHERE shop_id=%s AND permit_date=%s',
                (SHOP_ID, PERMIT_DATE))
    row = cur.fetchone()
    if row:
        permit_id = row[0]
        print(f"May 30 Permit exists (id={permit_id}), will add items to it.")
    else:
        print(f"Creating Permit {PERMIT_DATE}  {INDENT_NO}")
        if not dry_run:
            cur.execute("""
                INSERT INTO "Permit"
                    (shop_id, indent_no, invoice_no, permit_date, status,
                     total_indent_cbs, total_indent_amount,
                     total_cnf_cbs, total_cnf_amount, has_rationed_items,
                     created_at, updated_at)
                VALUES (%s, %s, %s, %s, 'RECEIVED',
                        0, 0, 0, 0, FALSE, NOW(), NOW())
                RETURNING id
            """, (SHOP_ID, INDENT_NO, INDENT_NO, PERMIT_DATE))
            permit_id = cur.fetchone()[0]
        else:
            permit_id = None

    # Check no items already added
    cur.execute('SELECT COUNT(*) FROM "PermitItem" WHERE permit_id=%s', (permit_id,))
    existing = cur.fetchone()[0]
    if existing > 0:
        print(f"  {existing} PermitItems already exist — aborting to avoid duplicates.")
        con.close()
        return

    # Load product map
    cur.execute('SELECT master_code, id, short_name FROM "Product"')
    prod_map = {r[0]: (r[1], r[2]) for r in cur.fetchall()}

    items = 0
    for mc, (recv_btls, mrp) in sorted(MAY30.items()):
        if mc not in prod_map:
            print(f"  WARNING: mc={mc} not in Product — skipped")
            continue
        pid, name = prod_map[mc]
        print(f"  mc={mc:>4} {name:<30} recv={recv_btls:>5}  mrp={mrp}")

        if not dry_run:
            cur.execute("""
                INSERT INTO "PermitItem"
                    (permit_id, product_id, rate_per_cb, carton_size_snapshot,
                     indent_cbs, indent_btls, indent_amount,
                     cnf_cbs, cnf_btls, cnf_amount,
                     is_rationed, received_btls, mrp_per_bottle)
                VALUES (%s, %s, 0, 0, 0, 0, 0, 0, 0, 0, FALSE, %s, %s)
            """, (permit_id, pid, recv_btls, mrp))

            cur.execute("""
                INSERT INTO "StockLot"
                    (shop_id, product_id, mrp_per_bottle, initial_qty,
                     consumed_qty, lot_date, source, permit_item_id)
                VALUES (%s, %s, %s, %s, 0, %s, 'PERMIT', NULL)
            """, (SHOP_ID, pid, mrp, recv_btls, PERMIT_DATE))
        items += 1

    print(f"  → {items} items {'(dry run)' if dry_run else 'created'}")

    if not dry_run:
        con.commit()
        print("Committed.")
    con.close()


if __name__ == "__main__":
    import sys
    patch(dry_run="--dry-run" in sys.argv)
