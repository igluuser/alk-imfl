"""
Step 2: Backfill StockLot rows for April 1, 2026 Opening Balance.

For each product with OB > 0 on April 1:
  - source = OPENING_BALANCE
  - mrp_per_bottle = Rate Matrix column J (M1 permit MRP)
  - initial_qty = GodownStock.opening_balance_btls for April 1
  - consumed_qty = total sold across all counters on April 1
    (bills were already generated without lot tracking, so we mark
     those bottles as consumed so the FIFO queue starts at the
     correct position for April 2 onward)
  - lot_date = 2026-04-01

Run with: python3 prisma/backfill_stock_lots_apr1.py [--dry-run]
"""

import sys
import os
import openpyxl
import psycopg2
from decimal import Decimal

DRY_RUN = '--dry-run' in sys.argv
STOCK_DATE = '2026-04-01'
EXCEL_PATH = '/home/prasadkabadi/ALK/Permits/KLS-I-APR-2026.xlsx'
SHOP_SHORT_NAME = 'KLS'

# Rate Matrix column indices (0-based)
COL_MASTER_CODE = 0
COL_M1_MRP      = 9   # column J — permit 1 MRP (OB MRP)

# ── DB connection ──────────────────────────────────────────────
from urllib.parse import urlparse, unquote

raw_url = os.environ.get('DATABASE_URL',
    'postgresql://postgres:P%40ssword4postgres@195.201.119.186:5432/imfl_app')
# Strip ?schema=public if present
raw_url = raw_url.split('?')[0]
u = urlparse(raw_url)
conn = psycopg2.connect(
    host=u.hostname,
    port=u.port or 5432,
    dbname=u.path.lstrip('/'),
    user=u.username,
    password=unquote(u.password),
)
cur  = conn.cursor()

# ── 1. Get shop id ─────────────────────────────────────────────
cur.execute('SELECT id FROM "Shop" WHERE short_name = %s', (SHOP_SHORT_NAME,))
row = cur.fetchone()
if not row:
    sys.exit(f'Shop {SHOP_SHORT_NAME} not found')
shop_id = row[0]
print(f'Shop: {SHOP_SHORT_NAME} (id={shop_id})')

# ── 2. Load Rate Matrix M1 MRP: master_code → mrp ─────────────
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
ws = wb['Rate Matrix']

rate_mrp = {}   # master_code (int) → M1 mrp (Decimal)
for row in ws.iter_rows(min_row=3, values_only=True):
    mc  = row[COL_MASTER_CODE]
    mrp = row[COL_M1_MRP]
    if mc is None:
        continue
    mc = int(float(mc))
    rate_mrp[mc] = Decimal(str(mrp)) if mrp is not None else None

print(f'Rate Matrix: {len(rate_mrp)} products loaded')

# ── 3. Build product master_code → id map ─────────────────────
cur.execute('SELECT id, master_code FROM "Product"')
product_map = {mc: pid for pid, mc in cur.fetchall()}
print(f'Products in DB: {len(product_map)}')

# ── 4. GodownStock OB for April 1 ─────────────────────────────
cur.execute("""
    SELECT product_id, opening_balance_btls
    FROM "GodownStock"
    WHERE shop_id = %s AND stock_date = %s AND opening_balance_btls > 0
""", (shop_id, STOCK_DATE))
ob_map = {pid: ob for pid, ob in cur.fetchall()}
print(f'Products with OB > 0 on {STOCK_DATE}: {len(ob_map)}')

# ── 5. Total sold April 1 across all counters ─────────────────
cur.execute("""
    SELECT dcs.product_id, SUM(dcs.sold_btls)
    FROM "DailyCounterStock" dcs
    JOIN "Counter" c ON c.id = dcs.counter_id
    WHERE c.shop_id = %s AND dcs.stock_date = %s
    GROUP BY dcs.product_id
""", (shop_id, STOCK_DATE))
sold_map = {pid: int(sold) for pid, sold in cur.fetchall()}
print(f'Products with sales on {STOCK_DATE}: {len(sold_map)}')

# ── 6. Check for existing StockLots (idempotency) ─────────────
cur.execute("""
    SELECT product_id FROM "StockLot"
    WHERE shop_id = %s AND lot_date = %s AND source = 'OPENING_BALANCE'
""", (shop_id, STOCK_DATE))
already_done = {r[0] for r in cur.fetchall()}
if already_done:
    print(f'WARNING: {len(already_done)} OB lots already exist for {STOCK_DATE} — skipping those')

# ── 7. Build insert rows ───────────────────────────────────────
rows = []
warnings = []

for mc, ob_qty in sorted(
    ((mc, ob_map[product_map[mc]]) for mc in rate_mrp if product_map.get(mc) in ob_map),
    key=lambda x: x[0]
):
    pid  = product_map[mc]
    if pid in already_done:
        continue

    mrp = rate_mrp[mc]
    if mrp is None:
        warnings.append(f'  master_code={mc}: M1 MRP is NULL — skipping')
        continue

    sold_qty     = sold_map.get(pid, 0)
    consumed_qty = min(sold_qty, ob_qty)   # can't consume more than initial

    rows.append((shop_id, pid, mrp, ob_qty, consumed_qty, STOCK_DATE))

print(f'\nRows to insert: {len(rows)}')
if warnings:
    print('Warnings:')
    for w in warnings: print(w)

# ── 8. Preview ─────────────────────────────────────────────────
print('\n{:<6} {:<8} {:<10} {:<10} {:<10}'.format(
    'pid', 'mc', 'mrp', 'initial', 'consumed'))
print('-' * 50)
for shop_id_, pid, mrp, initial, consumed, _ in rows[:10]:
    mc_display = next(mc for mc, p in product_map.items() if p == pid)
    print(f'{pid:<6} {mc_display:<8} {mrp:<10} {initial:<10} {consumed:<10}')
if len(rows) > 10:
    print(f'  ... and {len(rows)-10} more')

# ── 9. Insert ─────────────────────────────────────────────────
if DRY_RUN:
    print('\n[DRY RUN] No changes made.')
else:
    cur.executemany("""
        INSERT INTO "StockLot"
            (shop_id, product_id, mrp_per_bottle, initial_qty, consumed_qty,
             lot_date, source, permit_item_id)
        VALUES (%s, %s, %s, %s, %s, %s, 'OPENING_BALANCE', NULL)
    """, rows)
    conn.commit()
    print(f'\n✓ Inserted {len(rows)} StockLot rows (OPENING_BALANCE, {STOCK_DATE})')

cur.close()
conn.close()
