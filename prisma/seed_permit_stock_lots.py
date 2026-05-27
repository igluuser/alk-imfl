"""
Create StockLot(source=PERMIT) rows for a permit receipt.

Two modes:

  Mode A — permit already in DB:
    python3 prisma/seed_permit_stock_lots.py --permit-id 1 [--dry-run]
    Uses PermitItem.mrp_per_bottle and PermitItem.received_btls.
    lot_date = Permit.received_date; permit_item_id is set on each lot.

  Mode B — permit not yet in DB (Excel-sourced data):
    python3 prisma/seed_permit_stock_lots.py \\
        --date 2026-04-15 --mrp-col M2 [--shop KLS] [--dry-run]
    Reads GodownStock.received_btls for that date and MRP from Rate Matrix.
    permit_item_id is left NULL.

Script is idempotent: existing PERMIT lots for the same
(shop, product, lot_date) are skipped.
"""

import sys
import os
import argparse
import openpyxl
import psycopg2
from decimal import Decimal
from urllib.parse import urlparse, unquote

# ── CLI ────────────────────────────────────────────────────────
parser = argparse.ArgumentParser()
parser.add_argument('--permit-id',  type=int,  help='DB Permit.id (Mode A)')
parser.add_argument('--date',       type=str,  help='Receipt date YYYY-MM-DD (Mode B)')
parser.add_argument('--mrp-col',    type=str,  help='Rate Matrix MRP column, e.g. M2 (Mode B)')
parser.add_argument('--shop',       type=str,  default='KLS')
parser.add_argument('--dry-run',    action='store_true')
args = parser.parse_args()

if not args.permit_id and not (args.date and args.mrp_col):
    parser.error('Provide either --permit-id (Mode A) or --date + --mrp-col (Mode B)')

DRY_RUN    = args.dry_run
SHOP_SHORT = args.shop
EXCEL_PATH = '/home/prasadkabadi/ALK/Permits/KLS-I-APR-2026.xlsx'

# Rate Matrix MRP column mapping: M1=col9, M2=col11, M3=col13, ...
# Pattern: Mn → col index = 9 + (n-1)*2
def mrp_col_index(col_name: str) -> int:
    col_name = col_name.upper()
    if not col_name.startswith('M'):
        raise ValueError(f'Unknown MRP column: {col_name}. Expected M1, M2, ..., M15')
    n = int(col_name[1:])
    return 9 + (n - 1) * 2

# ── DB ─────────────────────────────────────────────────────────
raw_url = os.environ.get('DATABASE_URL',
    'postgresql://postgres:P%40ssword4postgres@195.201.119.186:5432/imfl_app')
raw_url = raw_url.split('?')[0]
u = urlparse(raw_url)
conn = psycopg2.connect(
    host=u.hostname, port=u.port or 5432,
    dbname=u.path.lstrip('/'),
    user=u.username, password=unquote(u.password),
)
cur = conn.cursor()

# ── Shop ────────────────────────────────────────────────────────
cur.execute('SELECT id FROM "Shop" WHERE short_name = %s', (SHOP_SHORT,))
row = cur.fetchone()
if not row:
    sys.exit(f'Shop {SHOP_SHORT} not found')
shop_id = row[0]
print(f'Shop: {SHOP_SHORT} (id={shop_id})')

# ══════════════════════════════════════════════════════════════
# MODE A — permit in DB
# ══════════════════════════════════════════════════════════════
if args.permit_id:
    cur.execute("""
        SELECT p.id, p.received_date, p.status, s.id
        FROM "Permit" p
        JOIN "Shop" s ON s.id = p.shop_id
        WHERE p.id = %s
    """, (args.permit_id,))
    perm = cur.fetchone()
    if not perm:
        sys.exit(f'Permit id={args.permit_id} not found')
    _, received_date, status, perm_shop_id = perm
    if perm_shop_id != shop_id:
        sys.exit(f'Permit {args.permit_id} belongs to a different shop')
    if not received_date:
        sys.exit(f'Permit {args.permit_id} has no received_date — mark it received first')

    print(f'Permit {args.permit_id}: received_date={received_date}, status={status}')

    # Load permit items
    cur.execute("""
        SELECT pi.id, pi.product_id, pi.mrp_per_bottle, pi.received_btls
        FROM "PermitItem" pi
        WHERE pi.permit_id = %s AND pi.received_btls > 0
    """, (args.permit_id,))
    items = cur.fetchall()
    print(f'PermitItems with received_btls > 0: {len(items)}')

    # Idempotency check
    cur.execute("""
        SELECT product_id FROM "StockLot"
        WHERE shop_id = %s AND lot_date = %s AND source = 'PERMIT'
          AND permit_item_id = ANY(%s)
    """, (shop_id, received_date, [i[0] for i in items]))
    already_done = {r[0] for r in cur.fetchall()}

    rows = []
    for pi_id, product_id, mrp, rcv_btls in items:
        if product_id in already_done:
            print(f'  SKIP product_id={product_id} — lot already exists')
            continue
        rows.append((shop_id, product_id, mrp, int(rcv_btls), 0, received_date, pi_id))

    print(f'\nRows to insert: {len(rows)}')
    if rows:
        print(f'\n{"pid":<6} {"mrp":<10} {"qty":<8} {"pi_id":<8}')
        print('-' * 36)
        for _, pid, mrp, qty, _, _, pi_id in rows:
            print(f'{pid:<6} {mrp:<10} {qty:<8} {pi_id:<8}')

    if not DRY_RUN and rows:
        cur.executemany("""
            INSERT INTO "StockLot"
                (shop_id, product_id, mrp_per_bottle, initial_qty, consumed_qty,
                 lot_date, source, permit_item_id)
            VALUES (%s, %s, %s, %s, %s, %s, 'PERMIT', %s)
        """, rows)
        conn.commit()
        print(f'\n✓ Inserted {len(rows)} StockLot rows (PERMIT, permit_id={args.permit_id})')
    elif DRY_RUN:
        print('\n[DRY RUN] No changes made.')

# ══════════════════════════════════════════════════════════════
# MODE B — Excel GodownStock + Rate Matrix MRP
# ══════════════════════════════════════════════════════════════
else:
    lot_date  = args.date
    mrp_col_i = mrp_col_index(args.mrp_col)
    print(f'Mode B: date={lot_date}, mrp_col={args.mrp_col} (col index {mrp_col_i})')

    # Load Rate Matrix MRP for the specified column
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb['Rate Matrix']
    rate_mrp = {}   # master_code → Decimal mrp
    for row in ws.iter_rows(min_row=3, values_only=True):
        mc  = row[0]
        mrp = row[mrp_col_i]
        if mc is None:
            continue
        mc = int(float(mc))
        rate_mrp[mc] = Decimal(str(mrp)) if mrp is not None else None
    print(f'Rate Matrix: {len(rate_mrp)} products, column {args.mrp_col}')

    # Product master_code → id map
    cur.execute('SELECT id, master_code FROM "Product"')
    product_map = {mc: pid for pid, mc in cur.fetchall()}

    # GodownStock received_btls > 0 for that date
    cur.execute("""
        SELECT product_id, received_btls
        FROM "GodownStock"
        WHERE shop_id = %s AND stock_date = %s AND received_btls > 0
    """, (shop_id, lot_date))
    gs_rows = cur.fetchall()
    print(f'GodownStock received > 0 on {lot_date}: {len(gs_rows)} products')

    if not gs_rows:
        print('Nothing to do — no products with received_btls > 0 on this date.')
        cur.close(); conn.close(); sys.exit(0)

    # Idempotency
    cur.execute("""
        SELECT product_id FROM "StockLot"
        WHERE shop_id = %s AND lot_date = %s AND source = 'PERMIT'
    """, (shop_id, lot_date))
    already_done = {r[0] for r in cur.fetchall()}
    if already_done:
        print(f'WARNING: {len(already_done)} PERMIT lots already exist for {lot_date} — skipping those')

    # Reverse lookup for display
    mc_for_pid = {v: k for k, v in product_map.items()}

    rows    = []
    skipped = []
    for product_id, rcv_btls in gs_rows:
        if product_id in already_done:
            skipped.append(product_id)
            continue
        mc  = mc_for_pid.get(product_id)
        mrp = rate_mrp.get(mc) if mc else None
        if mrp is None:
            print(f'  WARNING: no MRP for product_id={product_id} mc={mc} — skipping')
            continue
        rows.append((shop_id, product_id, mrp, int(rcv_btls), 0, lot_date))

    print(f'\nRows to insert: {len(rows)}  (skipped already-done: {len(skipped)})')
    if rows:
        print(f'\n{"pid":<6} {"mc":<8} {"mrp":<10} {"qty":<8}')
        print('-' * 36)
        for _, pid, mrp, qty, _, _ in rows:
            mc_d = mc_for_pid.get(pid, '?')
            print(f'{pid:<6} {mc_d:<8} {mrp:<10} {qty:<8}')

    if not DRY_RUN and rows:
        cur.executemany("""
            INSERT INTO "StockLot"
                (shop_id, product_id, mrp_per_bottle, initial_qty, consumed_qty,
                 lot_date, source, permit_item_id)
            VALUES (%s, %s, %s, %s, %s, %s, 'PERMIT', NULL)
        """, rows)
        conn.commit()
        print(f'\n✓ Inserted {len(rows)} StockLot rows (PERMIT, {lot_date}, {args.mrp_col})')
    elif DRY_RUN:
        print('\n[DRY RUN] No changes made.')

cur.close()
conn.close()
