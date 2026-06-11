"""
Import KSBCL Permit dated 17-05-2026 (Permit Dt: 18-05-2026)
Retailer: Prasad Ashok Kabadi (16884), Depot: KSBCL-BELAGAVI-2(64)
Total: 131 CBs, Net ₹5,99,871.16 + AROED ₹25,943.88 = Gross ₹6,37,813.04

MRP formula: ((rate_per_cb + aroed_per_btl * carton_size) / carton_size) * 1.1
"""

import psycopg2
from decimal import Decimal, ROUND_HALF_UP
import sys

CONN = dict(host="195.201.119.186", port=5432, dbname="imfl_app",
            user="postgres", password="P@ssword4postgres")
SHOP_ID = 1
PERMIT_DATE    = "2026-05-18"
INDENT_DATE    = "2026-05-17"
DRY_RUN = "--dry-run" in sys.argv

# (sr, product_id, pdf_code, rate_per_cb, aroed_per_btl, indent_cbs, indent_amount, carton_size)
# carton_size is taken from PDF item name (overrides DB where DB is wrong)
ITEMS = [
    ( 1, 13,  "0537010", Decimal("6240.38"), Decimal("1.99"),  1, Decimal("6240.38"),   48),
    ( 2, 28,  "0546010", Decimal("6731.49"), Decimal("0.74"),  3, Decimal("20194.47"),  48),
    ( 3, 46,  "0020015", Decimal("7629.11"), Decimal("0.17"),  2, Decimal("15258.22"),  48),
    ( 4, 52,  "0546010", Decimal("4760.00"), Decimal("0.92"),  1, Decimal("4760.00"),   48),
    ( 5, 66,  "0524010", Decimal("9295.78"), Decimal("1.97"),  1, Decimal("9295.78"),   48),
    ( 6, 91,  "0546010", Decimal("5869.60"), Decimal("0.49"),  3, Decimal("17608.80"),  48),
    ( 7, 96,  "0608010", Decimal("7629.11"), Decimal("0.17"),  1, Decimal("7629.11"),   48),
    ( 8, 97,  "0608010", Decimal("7629.11"), Decimal("2.58"),  1, Decimal("7629.11"),   96),
    ( 9, 98,  "0067010", Decimal("4936.56"), Decimal("2.48"),  1, Decimal("4936.56"),   12),
    (10, 99,  "0067010", Decimal("4936.56"), Decimal("3.74"),  1, Decimal("4936.56"),   24),
    (11, 128, "0014010", Decimal("4442.20"), Decimal("3.20"), 15, Decimal("66633.00"),  48),
    (12, 129, "0014010", Decimal("4442.20"), Decimal("4.10"),  3, Decimal("13326.60"),  96),
    (13, 138, "0022010", Decimal("4119.29"), Decimal("0.60"),  3, Decimal("12357.87"),  48),
    (14, 139, "0022010", Decimal("4119.29"), Decimal("2.80"), 72, Decimal("296588.88"), 96),
    (15, 142, "0602010", Decimal("5869.60"), Decimal("2.74"),  3, Decimal("17608.80"),  96),
    (16, 145, "0529019", Decimal("7156.97"), Decimal("0.99"),  3, Decimal("21470.91"),  48),
    (17, 157, "0024015", Decimal("4028.12"), Decimal("2.69"),  3, Decimal("12084.36"),  48),
    (18, 169, "0067010", Decimal("5869.60"), Decimal("0.49"),  1, Decimal("5869.60"),   48),
    (19, 170, "0067010", Decimal("5869.60"), Decimal("2.74"),  1, Decimal("5869.60"),   96),
    (20, 188, "0532030", Decimal("7629.11"), Decimal("0.17"),  1, Decimal("7629.11"),   48),
    (21, 203, "0524030", Decimal("7155.96"), Decimal("1.01"),  3, Decimal("21467.88"),  48),
    (22, 212, "0024045", Decimal("4457.03"), Decimal("2.86"),  3, Decimal("13371.09"),  48),
    (23, 268, "0010075", Decimal("1309.11"), Decimal("0.00"),  1, Decimal("1309.11"),   48),
    (24, 280, "0210090", Decimal("1187.75"), Decimal("1.12"),  3, Decimal("3563.25"),   12),
    (25, 295, "0217090", Decimal("2232.11"), Decimal("0.39"),  1, Decimal("2232.11"),   12),
]


def calc_mrp(rate, aroed_per_btl, carton):
    price_per_btl = rate / carton + aroed_per_btl
    return (price_per_btl * Decimal("1.1")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def run():
    conn = psycopg2.connect(**CONN)
    cur  = conn.cursor()

    # ── Fix wrong carton_size in DB for Original Traveller ──
    # id=169 180ml was stored as 24 (wrong), should be 48
    # id=170  90ml was stored as 48 (wrong), should be 96
    fixes = [(48, 169), (96, 170)]
    for correct_cs, pid in fixes:
        cur.execute('SELECT carton_size, ksbcl_item_name FROM "Product" WHERE id=%s', (pid,))
        row = cur.fetchone()
        if row and row[0] != correct_cs:
            print(f"  Fix Product #{pid} carton_size: {row[0]} → {correct_cs}  [{row[1]}]")
            if not DRY_RUN:
                cur.execute('UPDATE "Product" SET carton_size=%s WHERE id=%s', (correct_cs, pid))

    # ── Print calculated MRPs ──
    print(f"\n{'Sr':>3}  {'ProductID':>9}  {'Rate':>9}  {'AROED/Btl':>9}  {'Carton':>6}  {'MRP':>10}")
    print("-" * 65)
    for sr, pid, code, rate, aroed, cbs, amount, carton in ITEMS:
        mrp = calc_mrp(rate, aroed, carton)
        print(f"{sr:>3}  {pid:>9}  {rate:>9}  {aroed:>9}  {carton:>6}  {mrp:>10}")

    if DRY_RUN:
        print("\nDRY RUN — no DB writes.")
        conn.close()
        return

    # ── Insert Permit ──
    cur.execute("""
        INSERT INTO "Permit"
          (shop_id, indent_no, invoice_no, permit_date, total_indent_cbs, total_indent_amount, status, received_date, created_at, updated_at)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NOW(), NOW())
        RETURNING id
    """, (SHOP_ID, f"KSBCL-IND-{INDENT_DATE}", f"KSBCL-INV-{PERMIT_DATE}", PERMIT_DATE,
          131, Decimal("599871.16"), "RECEIVED", PERMIT_DATE))
    permit_id = cur.fetchone()[0]
    print(f"\nCreated Permit id={permit_id}")

    # ── Insert PermitItems + StockLots ──
    for sr, pid, code, rate, aroed, cbs, amount, carton in ITEMS:
        mrp = calc_mrp(rate, aroed, carton)
        total_btls = cbs * carton

        cur.execute("""
            INSERT INTO "PermitItem"
              (permit_id, product_id, rate_per_cb, carton_size_snapshot,
               indent_cbs, indent_btls, indent_amount,
               cnf_cbs, cnf_btls, cnf_amount,
               is_rationed, mrp_per_bottle, received_btls)
            VALUES (%s,%s,%s,%s,%s,0,%s,%s,0,%s,false,%s,%s)
            RETURNING id
        """, (permit_id, pid, rate, carton,
              cbs, amount,
              cbs, amount,
              mrp, total_btls))
        pi_id = cur.fetchone()[0]

        cur.execute("""
            INSERT INTO "StockLot"
              (shop_id, product_id, mrp_per_bottle, initial_qty, consumed_qty, lot_date, source, permit_item_id)
            VALUES (%s,%s,%s,%s,0,%s,'PERMIT',%s)
            RETURNING id
        """, (SHOP_ID, pid, mrp, total_btls, PERMIT_DATE, pi_id))
        sl_id = cur.fetchone()[0]

        print(f"  Sr{sr:>2}: PermitItem={pi_id}  StockLot={sl_id}  mrp=₹{mrp}  qty={total_btls}btls")

    conn.commit()
    conn.close()
    print("\nDone. All records committed.")


if __name__ == "__main__":
    run()
