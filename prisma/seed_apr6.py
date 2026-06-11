"""
Seed script — April 6, 2026 (Day 6).
Inserts: GodownStock, DailyCounterStock (C1/C2/C3),
         DailyCounterSummary, DailyExpense, CashRecord

No godown receipts. No permit received. No SP changes from Day 5.
C1 has PAK=350 (unlabeled row 443 in Excel). C1 Bhatta back to 50.
All counters tally=0.
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

XLSX = '/home/prasadkabadi/ALK/Permits/KLS-I-APR-2026.xlsx'
APR6 = date(2026, 4, 6)

def safe_int(v, default=0):
    try: return int(float(v)) if v is not None else default
    except: return default

def safe_dec(v, default='0'):
    try: return str(round(float(v), 2)) if v is not None else default
    except: return default

# ── References ─────────────────────────────────────────────────
cur.execute('SELECT id FROM "Shop" WHERE short_name = %s', ('KLS',))
shop_id = cur.fetchone()[0]
cur.execute('SELECT id, master_code FROM "Product"')
prod_id = {mc: pid for pid, mc in cur.fetchall()}
cur.execute('SELECT id, display_order FROM "Counter" WHERE shop_id = %s ORDER BY display_order', (shop_id,))
counter_id = {order: cid for cid, order in cur.fetchall()}
print(f'Shop={shop_id}, Products={len(prod_id)}, Counters={counter_id}')

# ── Read sheet "6" ─────────────────────────────────────────────
wb  = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
ws6 = wb['6']

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

stock_c  = {1: [], 2: [], 3: []}
stock_gd = []

for i, row in enumerate(ws6.iter_rows(values_only=True)):
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

# ── DailyCounterStock ──────────────────────────────────────────
print('▶ Inserting DailyCounterStock...')
for cnum in [1, 2, 3]:
    rows = []
    for pid, d in stock_c[cnum]:
        rows.append((counter_id[cnum], pid, APR6,
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

# ── GodownStock ────────────────────────────────────────────────
print('▶ Inserting GodownStock...')
gd_rows = [(shop_id, pid, APR6, gd['ob'], gd['rec'], gd['cb']) for pid, gd in stock_gd]
execute_values(cur, """
    INSERT INTO "GodownStock" (shop_id, product_id, stock_date,
                               opening_balance_btls, received_btls, closing_balance_btls)
    SELECT v.shop_id, v.pid, v.dt::date, v.ob::int, v.rec::int, v.cb::int
    FROM (VALUES %s) AS v(shop_id, pid, dt, ob, rec, cb)
    ON CONFLICT (shop_id, product_id, stock_date) DO NOTHING
""", gd_rows)

cur.execute('SELECT id, product_id FROM "GodownStock" WHERE shop_id = %s AND stock_date = %s',
            (shop_id, APR6))
gd_stock_id = {row[1]: row[0] for row in cur.fetchall()}

# ── GodownDistribution ─────────────────────────────────────────
print('▶ Inserting GodownDistributions...')
dist_rows = []
for pid, gd in stock_gd:
    gsid = gd_stock_id.get(pid)
    if not gsid:
        continue
    for cnum, qty in [(1, gd['sc1']), (2, gd['sc2']), (3, gd['sc3'])]:
        if qty > 0:
            dist_rows.append((gsid, counter_id[cnum], qty, APR6))
if dist_rows:
    execute_values(cur, """
        INSERT INTO "GodownDistribution" (godown_stock_id, counter_id, distributed_btls, distribution_date)
        SELECT v.gsid, v.cid, v.qty::int, v.dt::date
        FROM (VALUES %s) AS v(gsid, cid, qty, dt)
    """, dist_rows)
print(f'   GodownDistribution rows: {len(dist_rows)}')

# ── DailyCounterSummary ────────────────────────────────────────
print('▶ Inserting DailyCounterSummary...')
# (cnum, staff, liq, beer, tips, grand, gpay, exp, coll, drawer, tips_cash, tally)
summaries = [
    (1, 'Ramesh', '27125', '1805', '550', '28930', '5370', '5810', '21450', '2220', '550', '0'),
    (2, None,     '26929', '2065', '706', '28994', '7120', '7210', '21706',  '784', '706', '0'),
    (3, None,     '34070', '2095', '825', '36165', '4015', '4105', '30825', '2060', '825', '0'),
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
    """, (counter_id[cnum], APR6, staff,
          liq, beer, tips, grand, gpay, exp, coll, drawer, tips_cash, tally))

# ── DailyExpense ───────────────────────────────────────────────
print('▶ Inserting DailyExpenses...')
expenses_data = [
    (1, 'TEA',        40,    False),
    (1, 'PAK',        350,   False),
    (1, 'GOOGLE_PAY', 5370,  True),
    (1, 'BHATTA',     50,    False),
    (2, 'TEA',        40,    False),
    (2, 'GOOGLE_PAY', 7120,  True),
    (2, 'BHATTA',     50,    False),
    (3, 'TEA',        40,    False),
    (3, 'GOOGLE_PAY', 4015,  True),
    (3, 'BHATTA',     50,    False),
]
for cnum, category, amount, is_upi in expenses_data:
    cur.execute("""
        INSERT INTO "DailyExpense" (counter_id, expense_date, category, amount, is_upi)
        VALUES (%s, %s, %s, %s, %s)
    """, (counter_id[cnum], APR6, category, str(amount), is_upi))

# ── CashRecord ─────────────────────────────────────────────────
print('▶ Inserting CashRecords...')
cash_records = [
    # C1 collection
    (1, 'COLLECTION', 'NOTE', 500, 23),
    (1, 'COLLECTION', 'NOTE', 200, 17),
    (1, 'COLLECTION', 'NOTE', 100, 50),
    (1, 'COLLECTION', 'NOTE',  50, 20),
    # C2 collection
    (2, 'COLLECTION', 'NOTE', 500, 29),
    (2, 'COLLECTION', 'NOTE', 200, 15),
    (2, 'COLLECTION', 'NOTE', 100, 35),
    # C3 collection
    (3, 'COLLECTION', 'NOTE', 500, 42),
    (3, 'COLLECTION', 'NOTE', 200, 20),
    (3, 'COLLECTION', 'NOTE', 100, 25),
    (3, 'COLLECTION', 'NOTE',  50, 40),
    (3, 'COLLECTION', 'NOTE',  20, 25),
    # Drawer cash
    (1, 'DRAWER_CASH', 'COIN', 1, 2220),
    (2, 'DRAWER_CASH', 'COIN', 1,  784),
    (3, 'DRAWER_CASH', 'COIN', 1, 2060),
    # Tips cash
    (1, 'TIPS_CASH', 'COIN', 1, 550),
    (2, 'TIPS_CASH', 'COIN', 1, 706),
    (3, 'TIPS_CASH', 'COIN', 1, 825),
]
for cnum, cat, dtype, dval, count in cash_records:
    cur.execute("""
        INSERT INTO "CashRecord"
            (counter_id, record_date, cash_category, denomination_type,
             denomination_value, count, total_amount)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (counter_id, record_date, cash_category, denomination_type, denomination_value)
        DO NOTHING
    """, (counter_id[cnum], APR6, cat, dtype, dval, count, dval * count))

# ── Commit ─────────────────────────────────────────────────────
conn.commit()
cur.close()
conn.close()

print()
print('✓ April 6, 2026 seed complete.')
print(f'  DailyCounterStock   : {len(stock_c[1]) * 3} rows')
print(f'  GodownStock         : {len(gd_rows)} rows')
print(f'  GodownDistributions : {len(dist_rows)} rows')
print(f'  DailyExpenses       : {len(expenses_data)} rows')
print(f'  CashRecords         : {len(cash_records)} rows')
