"""
Seed script — April 2, 2026 (Day 2).
Inserts: GodownStock, DailyCounterStock (C1/C2/C3),
         DailyCounterSummary, DailyExpense, CashRecord

Prerequisites: seed_april1.py must have run first (Shop, Products, Counters exist).
No new CounterSellingPrice rows — S2 values are identical to S1 for all products.
"""

import openpyxl
import psycopg2
from psycopg2.extras import execute_values
from datetime import date
from urllib.parse import urlparse, unquote
import os

raw_url = os.environ.get('DATABASE_URL',
    'postgresql://postgres:P%40ssword4postgres@195.201.119.186:5432/imfl_app')
raw_url = raw_url.split('?')[0]
u = urlparse(raw_url)
conn = psycopg2.connect(
    host=u.hostname, port=u.port or 5432,
    dbname=u.path.lstrip('/'), user=u.username, password=unquote(u.password),
)
conn.autocommit = False
cur = conn.cursor()

XLSX   = '/home/prasadkabadi/ALK/Permits/KLS-I-APR-2026.xlsx'
APR2   = date(2026, 4, 2)
SHOP   = 'KLS'

def safe_int(v, default=0):
    try: return int(float(v)) if v is not None else default
    except: return default

def safe_dec(v, default='0'):
    try: return str(round(float(v), 2)) if v is not None else default
    except: return default

# ── Load references ────────────────────────────────────────────
cur.execute('SELECT id FROM "Shop" WHERE short_name = %s', (SHOP,))
shop_id = cur.fetchone()[0]

cur.execute('SELECT id, master_code FROM "Product"')
prod_id = {mc: pid for pid, mc in cur.fetchall()}

cur.execute('SELECT id, display_order FROM "Counter" WHERE shop_id = %s ORDER BY display_order', (shop_id,))
counter_id = {order: cid for cid, order in cur.fetchall()}

print(f'Shop={shop_id}, Products={len(prod_id)}, Counters={counter_id}')

# ── Read sheet "2" ────────────────────────────────────────────
wb  = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
ws2 = wb['2']

# Column offsets — identical to sheet "1"
# C1: base=1  → ob=3, rec=4, total=5, sale=6, cb=7, amt=8, tips=9, totamt=10, mrp=11, sp=12
# C2: base=15 → ob=17, rec=18, total=19, sale=20, cb=21, amt=22, tips=23, totamt=24, mrp=25, sp=26
# C3: base=29 → ob=31, rec=32, total=33, sale=34, cb=35, amt=36, tips=37, totamt=38, mrp=39, sp=40
# Godown: ob=59, rec=60, sc1=62, sc2=63, sc3=64, cb=65

def counter_cols(row, base):
    return {
        'ob':     safe_int(row[base + 2]),
        'rec':    safe_int(row[base + 3]),
        'total':  safe_int(row[base + 4]),
        'sale':   safe_int(row[base + 5]),
        'cb':     safe_int(row[base + 6]),
        'amt':    safe_dec(row[base + 7]),
        'tips':   safe_dec(row[base + 8]),
        'totamt': safe_dec(row[base + 9]),
        'mrp':    safe_dec(row[base + 10]),
        'sp':     safe_dec(row[base + 11]),
    }

def godown_cols(row):
    return {
        'ob':  safe_int(row[59]),
        'rec': safe_int(row[60]),
        'sc1': safe_int(row[62]),
        'sc2': safe_int(row[63]),
        'sc3': safe_int(row[64]),
        'cb':  safe_int(row[65]),
    }

stock_c = {1: [], 2: [], 3: []}
stock_gd = []

for i, row in enumerate(ws2.iter_rows(values_only=True)):
    if i < 2 or i >= 420:
        continue
    mc = row[0]
    if mc is None or not isinstance(mc, (int, float)):
        continue
    mc  = int(mc)
    pid = prod_id.get(mc)
    if not pid:
        continue

    stock_c[1].append((pid, counter_cols(row, 1)))
    stock_c[2].append((pid, counter_cols(row, 15)))
    stock_c[3].append((pid, counter_cols(row, 29)))
    stock_gd.append((pid, godown_cols(row)))

print(f'Product rows in sheet: {len(stock_c[1])}')

# ═══════════════════════════════════════════════════════════════
# STEP 1 — DailyCounterStock
# ═══════════════════════════════════════════════════════════════
print('▶ Inserting DailyCounterStock...')

for cnum in [1, 2, 3]:
    rows = []
    for pid, d in stock_c[cnum]:
        rows.append((counter_id[cnum], pid, APR2,
                     d['ob'], d['rec'], d['total'], d['sale'], d['cb'],
                     d['mrp'], d['sp'], d['amt'], d['tips'], d['totamt']))
    execute_values(cur, """
        INSERT INTO "DailyCounterStock"
            (counter_id, product_id, stock_date,
             opening_balance_btls, received_from_godown_btls, total_btls,
             sold_btls, closing_balance_btls,
             mrp_per_bottle, selling_price_per_bottle,
             sale_amount_mrp, tips_amount, total_sale_amount)
        SELECT v.cid, v.pid, v.dt::date,
               v.ob::int, v.rec::int, v.total::int,
               v.sale::int, v.cb::int,
               v.mrp::numeric, v.sp::numeric,
               v.amt::numeric, v.tips::numeric, v.totamt::numeric
        FROM (VALUES %s) AS v(cid,pid,dt,ob,rec,total,sale,cb,mrp,sp,amt,tips,totamt)
        ON CONFLICT DO NOTHING
    """, rows)

print(f'   DailyCounterStock rows: {len(stock_c[1]) * 3}')

# ═══════════════════════════════════════════════════════════════
# STEP 2 — GodownStock
# ═══════════════════════════════════════════════════════════════
print('▶ Inserting GodownStock...')

gd_rows = [(shop_id, pid, APR2, gd['ob'], gd['rec'], gd['cb']) for pid, gd in stock_gd]
execute_values(cur, """
    INSERT INTO "GodownStock" (shop_id, product_id, stock_date,
                               opening_balance_btls, received_btls, closing_balance_btls)
    SELECT v.shop_id, v.pid, v.dt::date, v.ob::int, v.rec::int, v.cb::int
    FROM (VALUES %s) AS v(shop_id, pid, dt, ob, rec, cb)
    ON CONFLICT (shop_id, product_id, stock_date) DO NOTHING
""", gd_rows)

cur.execute('SELECT id, product_id FROM "GodownStock" WHERE shop_id = %s AND stock_date = %s',
            (shop_id, APR2))
gd_stock_id = {row[1]: row[0] for row in cur.fetchall()}

# ═══════════════════════════════════════════════════════════════
# STEP 3 — GodownDistribution
# ═══════════════════════════════════════════════════════════════
print('▶ Inserting GodownDistributions...')

dist_rows = []
for pid, gd in stock_gd:
    gsid = gd_stock_id.get(pid)
    if not gsid:
        continue
    for cnum, qty in [(1, gd['sc1']), (2, gd['sc2']), (3, gd['sc3'])]:
        if qty > 0:
            dist_rows.append((gsid, counter_id[cnum], qty, APR2))

if dist_rows:
    execute_values(cur, """
        INSERT INTO "GodownDistribution" (godown_stock_id, counter_id, distributed_btls, distribution_date)
        SELECT v.gsid, v.cid, v.qty::int, v.dt::date
        FROM (VALUES %s) AS v(gsid, cid, qty, dt)
    """, dist_rows)

print(f'   GodownDistribution rows: {len(dist_rows)}')

# ═══════════════════════════════════════════════════════════════
# STEP 4 — DailyCounterSummary
# ═══════════════════════════════════════════════════════════════
print('▶ Inserting DailyCounterSummary...')

# Values from sheet "2" summary rows 426-430
# (cnum, staff, liq, beer, tips, grand, gpay, exp, coll, drawer, tips_cash, tally)
summaries = [
    (1, 'Ramesh', '21370', '535',  '275', '21905', '2460', '2550', '12680', '6950', '275',  '-5'),
    (2, None,     '36361', '2120', '599', '38481', '9205', '9410', '27599', '2071', '599',   '0'),
    (3, 'Sudesh', '27355', '1580', '570', '28935', '5355', '5445', '21170', '2890', '570',   '0'),
]

for cnum, staff, liq, beer, tips, grand, gpay, exp, coll, drawer, tips_cash, tally in summaries:
    cur.execute("""
        INSERT INTO "DailyCounterSummary"
            (counter_id, summary_date, staff_name,
             liquor_sale_amount, beer_sale_amount, total_tips, grand_total_by_sale,
             google_pay_amount, expenses_total, collection_total,
             drawer_cash_total, tips_cash_total, tally_difference)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (counter_id, summary_date) DO NOTHING
    """, (counter_id[cnum], APR2, staff,
          liq, beer, tips, grand, gpay, exp, coll, drawer, tips_cash, tally))

# ═══════════════════════════════════════════════════════════════
# STEP 5 — DailyExpense
# ═══════════════════════════════════════════════════════════════
print('▶ Inserting DailyExpenses...')

expenses_data = [
    (1, 'TEA',        40,    False),
    (1, 'GOOGLE_PAY', 2460,  True),
    (1, 'BHATTA',     50,    False),
    (2, 'TEA',        40,    False),
    (2, 'BREAKAGE',   115,   False),
    (2, 'GOOGLE_PAY', 9205,  True),
    (2, 'BHATTA',     50,    False),
    (3, 'TEA',        40,    False),
    (3, 'GOOGLE_PAY', 5355,  True),
    (3, 'BHATTA',     50,    False),
]

for cnum, category, amount, is_upi in expenses_data:
    cur.execute("""
        INSERT INTO "DailyExpense" (counter_id, expense_date, category, amount, is_upi)
        VALUES (%s, %s, %s, %s, %s)
    """, (counter_id[cnum], APR2, category, str(amount), is_upi))

# ═══════════════════════════════════════════════════════════════
# STEP 6 — CashRecord
# ═══════════════════════════════════════════════════════════════
print('▶ Inserting CashRecords...')

# COLLECTION denominations (from counted amounts in summary section)
# DRAWER_CASH and TIPS_CASH stored as COIN denom=1 count=total (same pattern as Day 1)
cash_records = [
    # C1 collection
    (1, 'COLLECTION', 'NOTE', 500,   5),
    (1, 'COLLECTION', 'NOTE', 200,  12),
    (1, 'COLLECTION', 'NOTE', 100,  55),
    (1, 'COLLECTION', 'NOTE',  50,  30),
    (1, 'COLLECTION', 'NOTE',  20,  25),
    # C2 collection
    (2, 'COLLECTION', 'NOTE', 500,  32),
    (2, 'COLLECTION', 'NOTE', 200,  19),
    (2, 'COLLECTION', 'NOTE', 100,  58),
    (2, 'COLLECTION', 'NOTE',  50,  28),
    # C3 collection
    (3, 'COLLECTION', 'NOTE', 500,  23),
    (3, 'COLLECTION', 'NOTE', 200,  18),
    (3, 'COLLECTION', 'NOTE', 100,  35),
    (3, 'COLLECTION', 'NOTE',  50,  40),
    # Drawer cash (single COIN entry = total amount)
    (1, 'DRAWER_CASH', 'COIN', 1, 6950),
    (2, 'DRAWER_CASH', 'COIN', 1, 2071),
    (3, 'DRAWER_CASH', 'COIN', 1, 2890),
    # Tips cash
    (1, 'TIPS_CASH', 'COIN', 1, 275),
    (2, 'TIPS_CASH', 'COIN', 1, 599),
    (3, 'TIPS_CASH', 'COIN', 1, 570),
]

for cnum, cat, dtype, dval, count in cash_records:
    cur.execute("""
        INSERT INTO "CashRecord"
            (counter_id, record_date, cash_category, denomination_type,
             denomination_value, count, total_amount)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (counter_id, record_date, cash_category, denomination_type, denomination_value)
        DO NOTHING
    """, (counter_id[cnum], APR2, cat, dtype, dval, count, dval * count))

# ═══════════════════════════════════════════════════════════════
# COMMIT
# ═══════════════════════════════════════════════════════════════
conn.commit()
cur.close()
conn.close()

print()
print('✓ April 2, 2026 seed complete.')
print(f'  DailyCounterStock    : {len(stock_c[1]) * 3} rows')
print(f'  GodownStock          : {len(gd_rows)} rows')
print(f'  GodownDistributions  : {len(dist_rows)} rows')
print(f'  DailyExpenses        : {len(expenses_data)} rows')
print(f'  CashRecords          : {len(cash_records)} rows')
