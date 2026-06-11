#!/usr/bin/env python3
"""
seed_may13.py  — Seed May 13, 2026 from KLS-I-MAY-2026.xlsx sheet "13"
Godown received stock comes from permit INDBG226001844 CNF CBS (not Excel col AU/BI).
"""

import os, sys, subprocess
from datetime import date
import openpyxl
import psycopg2
from psycopg2.extras import execute_values

CONN = dict(host="195.201.119.186", port=5432,
            dbname="imfl_app", user="postgres", password="P@ssword4postgres")

XLSX      = "/home/prasadkabadi/ALK/Permits/KLS-I-MAY-2026.xlsx"
SALE_DATE = date(2026, 5, 13)
SHOP      = "KLS"

GENERATE_SCRIPT = os.path.join(os.path.dirname(__file__), "generate_daily_bills.py")

# ── Permit INDBG226001844 received stock (master_code → bottles) ──────────────
# CNF CBs × carton_size per product
PERMIT_REC = {
     45: 6 * 48,   # Bagpiper Deluxe 180ML  → 288
     63: 3 * 48,   # DSP Black PET 180ML    → 144
     69: 1 * 48,   # Haywards Cheers 180ML  →  48
     79: 3 * 96,   # Bangalore Whisky 90ML  → 288
    108: 1 * 48,   # Old Tavern 180ML       →  48
    109: 1 * 96,   # Old Tavern 90ML        →  96
    113: 1 * 48,   # OC Special 180ML       →  48
    145: 9 * 48,   # Wellington 180ML       → 432
    146: 3 * 96,   # Wellington 90ML        → 288
    155: 3 * 48,   # DK Double Kick 180ML   → 144
    156:49 * 96,   # DK Double Kick 90ML    →4704
    162: 1 * 48,   # OC Star Supreme 180ML  →  48
    169: 1 * 96,   # OC Star Supreme 90ML   →  96
    174: 1 * 48,   # Black Belt 180ML       →  48
    175: 1 * 96,   # Black Belt 90ML        →  96
    206: 3 * 48,   # Bagpiper XXX Rum 180ML → 144
    219: 1 * 48,   # Old Monk Rum 180ML     →  48
    220: 1 * 96,   # Old Monk Rum 90ML      →  96
    221: 1 * 48,   # Raja XXX Rum 180ML     →  48
    222: 1 * 96,   # Raja XXX Rum 90ML      →  96
    255: 1 * 48,   # Carnival Dry Gin 180ML →  48
}

CTR_BASES = {1: 1, 2: 15, 3: 29}

R_STAFF      = 423
R_LIQ        = 425
R_BEER       = 426
R_GTOTAL_ROW = 444
R_GPAY       = 439
R_EXP_TOTAL  = 445
R_TIPS_CASH  = 460

EXPENSE_ROWS = [
    (433, 'TEA',                 False),
    (434, 'PERMIT_RENT',         False),
    (435, 'POOJA',               False),
    (436, 'BREAKAGE',            False),
    (437, 'OVER_CASH',           False),
    (438, 'ELECTRICITY_BILL',    False),
    (439, 'GOOGLE_PAY',          True ),
    (440, 'BHATTA',              False),
    (441, 'GOOGLE_PAY_NEGATIVE', False),
    (442, 'PAK',                 False),
    (444, 'OTHERS',              False),
]
EXP_COLS       = {1: 2,  2: 16, 3: 30}
COLL_COLS      = {1: (5,6,7,8), 2: (19,20,21,22), 3: (33,34,35,36)}
DRAW_TOTAL_COL = {1: 13, 2: 27, 3: 41}
COLL_TOTAL_COL = {1:  8, 2: 22, 3: 36}
TIPS_TOTAL_COL = {1:  8, 2: 22, 3: 36}
LIQ_SALE_COL   = {1:  6, 2: 20, 3: 34}
LIQ_TIPS_COL   = {1:  7, 2: 21, 3: 35}
BEER_SALE_COL  = {1:  6, 2: 20, 3: 34}
BEER_TIPS_COL  = {1:  7, 2: 21, 3: 35}
TALLY_COL      = {1: 11, 2: 25, 3: 39}
GPAY_COL       = {1:  2, 2: 16, 3: 30}
STAFF_COL      = {1:  1, 2: 15, 3: 29}


def si(v, d=0):
    try: return int(float(v)) if v is not None else d
    except: return d

def sd(v, d='0'):
    try: return str(round(float(v), 2)) if v is not None else d
    except: return d

def sfloat(v, d=0.0):
    try: return float(v) if v is not None else d
    except: return d


def extract_stock(rows):
    stock_c = {1: [], 2: [], 3: []}
    stock_gd = []
    for i, row in enumerate(rows):
        if i < 2 or i >= 420:
            continue
        mc = row[0]
        if mc is None or not isinstance(mc, (int, float)):
            continue
        mc = int(mc)
        for cnum, base in CTR_BASES.items():
            stock_c[cnum].append((mc, {
                'ob':     si(row[base + 2]),
                'rec':    si(row[base + 3]),
                'total':  si(row[base + 4]),
                'sale':   si(row[base + 5]),
                'cb':     si(row[base + 6]),
                'amt':    sd(row[base + 7]),
                'tips':   sd(row[base + 8]),
                'totamt': sd(row[base + 9]),
                'mrp':    sd(row[base + 10]),
                'sp':     sd(row[base + 11]),
            }))
        # Godown: ob from Excel, rec from PERMIT override, sc1-3 from Excel, cb from Excel
        gd_rec = PERMIT_REC.get(mc, si(row[60]))
        stock_gd.append((mc, {
            'ob':  si(row[59]),
            'rec': gd_rec,
            'sc1': si(row[62]),
            'sc2': si(row[63]),
            'sc3': si(row[64]),
            'cb':  si(row[65]),
        }))
    return stock_c, stock_gd


def extract_summary(rows):
    result = {}
    for cnum in [1, 2, 3]:
        result[cnum] = {
            'staff':      rows[R_STAFF][STAFF_COL[cnum]],
            'liq_sale':   sfloat(rows[R_LIQ ][LIQ_SALE_COL [cnum]]),
            'beer_sale':  sfloat(rows[R_BEER][BEER_SALE_COL[cnum]]),
            'total_tips': sfloat(rows[R_LIQ ][LIQ_TIPS_COL [cnum]]) +
                          sfloat(rows[R_BEER][BEER_TIPS_COL[cnum]]),
            'gpay':       sfloat(rows[R_GPAY][GPAY_COL[cnum]]),
            'exp_total':  sfloat(rows[R_EXP_TOTAL][EXP_COLS[cnum]]),
            'collection': sfloat(rows[R_GTOTAL_ROW][COLL_TOTAL_COL[cnum]]),
            'drawer':     sfloat(rows[R_GTOTAL_ROW][DRAW_TOTAL_COL[cnum]]),
            'tips_cash':  sfloat(rows[R_TIPS_CASH ][TIPS_TOTAL_COL[cnum]]),
            'tally':      sfloat(rows[R_BEER][TALLY_COL[cnum]]),
        }
        s = result[cnum]
        s['grand'] = s['liq_sale'] + s['beer_sale']
    return result


def extract_expenses(rows):
    result = {1: [], 2: [], 3: []}
    for row_i, cat, is_upi in EXPENSE_ROWS:
        for cnum in [1, 2, 3]:
            amt = sfloat(rows[row_i][EXP_COLS[cnum]])
            if amt > 0:
                result[cnum].append((cat, amt, is_upi))
    return result


def extract_cash(rows):
    result = {1: [], 2: [], 3: []}
    for i in range(433, 444):
        row = rows[i]
        for cnum, (tc, dc, cc, totc) in COLL_COLS.items():
            dtype  = row[tc]
            dval   = sfloat(row[dc])
            count  = sfloat(row[cc])
            total  = sfloat(row[totc])
            dtype_db = 'NOTE' if dtype == 'Notes' else 'COIN' if dtype == 'Coins' else None
            if dtype_db and count > 0 and total > 0:
                result[cnum].append(('COLLECTION', dtype_db, int(dval), int(count), int(total)))
    for cnum in [1, 2, 3]:
        d = int(sfloat(rows[R_GTOTAL_ROW][DRAW_TOTAL_COL[cnum]]))
        if d > 0:
            result[cnum].append(('DRAWER_CASH', 'COIN', 1, d, d))
    for cnum in [1, 2, 3]:
        t = int(sfloat(rows[R_TIPS_CASH][TIPS_TOTAL_COL[cnum]]))
        if t > 0:
            result[cnum].append(('TIPS_CASH', 'COIN', 1, t, t))
    return result


def main():
    conn = psycopg2.connect(**CONN)
    conn.autocommit = False
    cur  = conn.cursor()

    cur.execute('SELECT id FROM "Shop" WHERE short_name=%s', (SHOP,))
    shop_id = cur.fetchone()[0]

    cur.execute('SELECT id, master_code FROM "Product"')
    prod_id = {mc: pid for pid, mc in cur.fetchall()}

    cur.execute('SELECT id, display_order FROM "Counter" WHERE shop_id=%s ORDER BY display_order', (shop_id,))
    counter_id = {order: cid for cid, order in cur.fetchall()}

    print(f"Shop={shop_id}  Products={len(prod_id)}  Counters={counter_id}")

    wb = openpyxl.load_workbook(XLSX, data_only=True)
    rows = list(wb['13'].iter_rows(values_only=True))

    stock_c, stock_gd = extract_stock(rows)
    summary            = extract_summary(rows)
    expenses           = extract_expenses(rows)
    cash               = extract_cash(rows)

    total_sold = sum(d['sale'] for cnum in [1,2,3] for _, d in stock_c[cnum])
    print(f"\n── {SALE_DATE}  (total bottles sold: {total_sold}) ──")
    for cnum in [1,2,3]:
        s = summary[cnum]
        print(f"  C{cnum}: liq={s['liq_sale']:.0f}  beer={s['beer_sale']:.0f}  "
              f"tips={s['total_tips']:.0f}  gpay={s['gpay']:.0f}  "
              f"coll={s['collection']:.0f}  drawer={s['drawer']:.0f}  tally={s['tally']:.0f}")

    # ── DailyCounterStock ─────────────────────────────────────────────────────
    for cnum in [1, 2, 3]:
        rows_db = []
        for mc, d in stock_c[cnum]:
            pid = prod_id.get(mc)
            if not pid:
                continue
            rows_db.append((counter_id[cnum], pid, SALE_DATE,
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
        """, rows_db)
    print(f"  DailyCounterStock inserted.")

    # ── GodownStock ───────────────────────────────────────────────────────────
    gd_rows = []
    for mc, gd in stock_gd:
        pid = prod_id.get(mc)
        if pid:
            gd_rows.append((shop_id, pid, SALE_DATE, gd['ob'], gd['rec'], gd['cb']))
    execute_values(cur, """
        INSERT INTO "GodownStock" (shop_id, product_id, stock_date,
                                   opening_balance_btls, received_btls, closing_balance_btls)
        SELECT v.sid, v.pid, v.dt::date, v.ob::int, v.rec::int, v.cb::int
        FROM (VALUES %s) AS v(sid, pid, dt, ob, rec, cb)
        ON CONFLICT (shop_id, product_id, stock_date) DO NOTHING
    """, gd_rows)
    print(f"  GodownStock inserted.")

    # ── GodownDistribution ────────────────────────────────────────────────────
    cur.execute('SELECT id, product_id FROM "GodownStock" WHERE shop_id=%s AND stock_date=%s',
                (shop_id, SALE_DATE))
    gd_stock_id = {r[1]: r[0] for r in cur.fetchall()}
    dist_rows = []
    for mc, gd in stock_gd:
        pid = prod_id.get(mc)
        gsid = gd_stock_id.get(pid)
        if not gsid:
            continue
        for cnum, qty in [(1, gd['sc1']), (2, gd['sc2']), (3, gd['sc3'])]:
            if qty > 0:
                dist_rows.append((gsid, counter_id[cnum], qty, SALE_DATE))
    if dist_rows:
        execute_values(cur, """
            INSERT INTO "GodownDistribution"
                (godown_stock_id, counter_id, distributed_btls, distribution_date)
            SELECT v.gsid, v.cid, v.qty::int, v.dt::date
            FROM (VALUES %s) AS v(gsid, cid, qty, dt)
        """, dist_rows)
    print(f"  GodownDistribution: {len(dist_rows)} rows.")

    # ── DailyCounterSummary ───────────────────────────────────────────────────
    for cnum in [1, 2, 3]:
        s = summary[cnum]
        cur.execute("""
            INSERT INTO "DailyCounterSummary"
                (counter_id, summary_date, staff_name,
                 liquor_sale_amount, beer_sale_amount, total_tips, grand_total_by_sale,
                 google_pay_amount, expenses_total, collection_total,
                 drawer_cash_total, tips_cash_total, tally_difference)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ON CONFLICT (counter_id, summary_date) DO NOTHING
        """, (counter_id[cnum], SALE_DATE, s['staff'],
              str(s['liq_sale']), str(s['beer_sale']), str(s['total_tips']),
              str(s['grand']), str(s['gpay']), str(s['exp_total']),
              str(s['collection']), str(s['drawer']), str(s['tips_cash']),
              str(s['tally'])))
    print(f"  DailyCounterSummary inserted.")

    # ── DailyExpense ──────────────────────────────────────────────────────────
    for cnum in [1, 2, 3]:
        for cat, amt, is_upi in expenses[cnum]:
            cur.execute("""
                INSERT INTO "DailyExpense" (counter_id, expense_date, category, amount, is_upi)
                VALUES (%s,%s,%s,%s,%s)
            """, (counter_id[cnum], SALE_DATE, cat, str(amt), is_upi))
    print(f"  DailyExpense inserted.")

    # ── CashRecord ────────────────────────────────────────────────────────────
    for cnum in [1, 2, 3]:
        for cat, dtype, dval, count, total in cash[cnum]:
            cur.execute("""
                INSERT INTO "CashRecord"
                    (counter_id, record_date, cash_category, denomination_type,
                     denomination_value, count, total_amount)
                VALUES (%s,%s,%s,%s,%s,%s,%s)
                ON CONFLICT (counter_id, record_date, cash_category, denomination_type, denomination_value)
                DO NOTHING
            """, (counter_id[cnum], SALE_DATE, cat, dtype, dval, count, total))
    print(f"  CashRecord inserted.")

    conn.commit()
    print(f"\n✓ May 13 seeded successfully.")
    conn.close()

    # ── Bill generation ───────────────────────────────────────────────────────
    print(f"\nGenerating bills for {SALE_DATE}...")
    r = subprocess.run(
        ['python3', GENERATE_SCRIPT, '--date', SALE_DATE.isoformat()],
        capture_output=True, text=True
    )
    print(r.stdout.rstrip())
    if r.returncode != 0:
        print(f"ERROR: {r.stderr.rstrip()}")


if __name__ == '__main__':
    main()
