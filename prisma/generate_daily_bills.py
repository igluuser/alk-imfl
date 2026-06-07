#!/usr/bin/env python3
"""
generate_daily_bills.py — KSBCL daily bills, A4 PDF (3-column layout)

Unified series: KLS-2026-1, KLS-2026-2, ...  (year-continuous, no L/B split)
UPI bills first (cumulative ≤ google_pay total), then Cash bills.
IMFL: max 4,320 ml/bill.  Beer: max 3,900 ml/bill.

Usage:
  python generate_daily_bills.py --date 2026-04-01
  python generate_daily_bills.py --date 2026-04-01 --no-save
  python generate_daily_bills.py --date 2026-04-01 --delete-all-first
"""

import argparse, os, random, sys
import psycopg2
from psycopg2.extras import execute_values
from decimal import Decimal
from datetime import date as date_type

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

REPORTS_DIR = os.path.join(os.path.dirname(__file__), "..", "Reports")

CONN = dict(host="195.201.119.186", port=5432,
            dbname="imfl_app", user="postgres", password="P@ssword4postgres")

IMFL_LIMIT_ML = 4320
BEER_LIMIT_ML  = 3900

EXPENSE_LABEL = {
    "TEA": "Tea", "PERMIT_RENT": "Permit Rent", "POOJA": "Pooja",
    "BREAKAGE": "Breakage", "OVER_CASH": "Over Cash",
    "ELECTRICITY_BILL": "Elec. Bill", "GOOGLE_PAY": "Google Pay",
    "BHATTA": "Bhatta", "GOOGLE_PAY_NEGATIVE": "Google Pay (Neg)",
    "PAK": "PAK", "OTHERS": "Others",
}


# ── Font registration ─────────────────────────────────────────────────────────

FONT  = "Helvetica"
FONTB = "Helvetica-Bold"
RS    = "Rs."   # rupee symbol fallback

def _setup_fonts():
    global FONT, FONTB, RS
    normal = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    bold   = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    if os.path.exists(normal):
        pdfmetrics.registerFont(TTFont("DV", normal))
        FONT = "DV"
        RS   = "₹"   # ₹
    if os.path.exists(bold):
        pdfmetrics.registerFont(TTFont("DVB", bold))
        FONTB = "DVB"


# ── PDF layout constants ──────────────────────────────────────────────────────

PAGE_W, PAGE_H = A4          # 595.27, 841.89 pt
MARGIN   = 12 * mm           # ~34 pt
GUTTER   = 7                 # pt between columns
N_COLS   = 3
COL_W    = (PAGE_W - 2*MARGIN - (N_COLS-1)*GUTTER) / N_COLS   # ~170 pt
COL_X    = [MARGIN + i*(COL_W + GUTTER) for i in range(N_COLS)]
TOP_Y    = PAGE_H - MARGIN
BOT_Y    = MARGIN

# Bill drawing sizes
FS_BILL  = 7.5    # bill header font size
FS_HDR   = 6.5    # column-header font size
FS_ITEM  = 6.5    # item row font size
FS_TOT   = 7.0    # total row font size
LH_BILL  = 11     # line height for bill header
LH_ITEM  = 9.5    # line height for item rows
BILL_GAP = 14     # gap between bills
_RG      = 7      # rule-to-text gap (avoids rule cutting into letter ascenders)


def _bill_height(bill):
    # count normal items + parcel-bag items (bottles may be 0 for bags)
    n_items = sum(1 for it in bill["items"]
                  if it["bottles"] > 0 or it.get("is_bag"))
    return (LH_BILL                 # bill number + time line
            + _RG                   # gap: header rule → column-header text
            + LH_ITEM               # column header row
            + n_items * LH_ITEM     # item rows
            + _RG                   # gap: items rule → Mode text
            + LH_ITEM               # Mode + TOTAL row
            + BILL_GAP)             # whitespace below bill


def _draw_rule(c, x, y, w, thick=0.3):
    c.setLineWidth(thick)
    c.line(x, y, x + w, y)


def _draw_bill(c, bill, x, top_y):
    """Draw one bill starting at (x, top_y), returns height used."""
    y  = top_y
    w  = COL_W
    cx = x + w

    # ── Bill number + time ───────────────────────────────────────
    c.setFont(FONTB, FS_BILL)
    c.drawString(x, y, bill["bill_number"])
    tw = c.stringWidth(bill["time_str"], FONTB, FS_BILL)
    c.drawString(cx - tw, y, bill["time_str"])
    y -= LH_BILL

    # Rule — drawn after header, gap before column-header text
    _draw_rule(c, x, y, w)
    y -= _RG   # enough clearance so rule doesn't cut into ascenders below

    # ── Column headers ───────────────────────────────────────────
    # Columns: Brand(0–49%) | ml(~55%) | Btls(~64%) | MRP/btl(~82%) | Amount(100%)
    c.setFont(FONTB, FS_HDR)
    c.drawString(x,           y, "Brand")
    c.drawRightString(x + w*0.55, y, "ml")
    c.drawRightString(x + w*0.64, y, "Btls")
    c.drawRightString(x + w*0.82, y, "MRP/btl")
    c.drawRightString(cx,         y, "Amount")
    y -= LH_ITEM

    # ── Items ────────────────────────────────────────────────────
    c.setFont(FONT, FS_ITEM)
    max_brand_w = w * 0.49    # 49% of column for brand name
    for it in bill["items"]:
        is_bag = it.get("is_bag", False)
        if it["bottles"] <= 0 and not is_bag:
            continue

        brand = it["short_name"]
        while c.stringWidth(brand, FONT, FS_ITEM) > max_brand_w and len(brand) > 4:
            brand = brand[:-2] + "."

        if is_bag:
            # Parcel bag: show name + amount only, dashes for other columns
            c.drawString(x,   y, brand)
            c.drawRightString(x + w*0.55, y, "—")
            c.drawRightString(x + w*0.64, y, "—")
            c.drawRightString(x + w*0.82, y, "—")
            c.drawRightString(cx,         y, f"{RS}{it['total_amount']:,.0f}")
        else:
            c.drawString(x,           y, brand)
            c.drawRightString(x + w*0.55, y, str(it["ml_per_bottle"]))
            c.drawRightString(x + w*0.64, y, str(it["bottles"]))
            c.drawRightString(x + w*0.82, y, f"{RS}{it['mrp_per_bottle']:,.2f}")
            c.drawRightString(cx,         y, f"{RS}{it['total_amount']:,.0f}")
        y -= LH_ITEM

    # Rule — drawn after items, gap before Mode text
    _draw_rule(c, x, y, w)
    y -= _RG

    # ── Mode + Total ─────────────────────────────────────────────
    c.setFont(FONTB, FS_TOT)
    c.drawString(x, y, f"Mode: {bill['mode']}")
    tot_str = f"TOTAL  {RS}{bill['total_amount']:,.0f}"
    c.drawRightString(cx, y, tot_str)
    y -= LH_ITEM + BILL_GAP

    return top_y - y


# ── PDF generation ────────────────────────────────────────────────────────────

def _page_header(c, shop_name, sale_date, right_label=""):
    c.setFont(FONTB, 8)
    c.drawString(MARGIN, PAGE_H - 9*mm,
                 f"{shop_name}  ·  Daily Bills  ·  {sale_date}")
    if right_label:
        c.drawRightString(PAGE_W - MARGIN, PAGE_H - 9*mm, right_label)
    c.setLineWidth(0.4)
    c.line(MARGIN, PAGE_H - 10*mm, PAGE_W - MARGIN, PAGE_H - 10*mm)


def generate_bills_pdf(bills, shop_name, sale_date, pdf_path):
    """Generate A4 PDF with 3-column bill layout (bills only)."""
    _setup_fonts()
    c   = rl_canvas.Canvas(pdf_path, pagesize=A4)
    top = TOP_Y - 8   # start below page header rule
    col = 0
    cy  = top

    _page_header(c, shop_name, sale_date)

    for bill in bills:
        bh = _bill_height(bill)
        if cy - bh < BOT_Y:
            col += 1
            if col >= N_COLS:
                c.showPage()
                _page_header(c, shop_name, sale_date)
                col = 0
            cy = top
        _draw_bill(c, bill, COL_X[col], cy)
        cy -= bh

    c.save()


def generate_summary_pdf(summaries, expenses_map, cash_map,
                         shop_name, sale_date, pdf_path):
    """Generate A4 PDF — 2-column summary layout (counters side-by-side)."""
    _setup_fonts()
    c = rl_canvas.Canvas(pdf_path, pagesize=A4)
    _page_header(c, shop_name, sale_date, "Expenses / Cash / Summary")
    _draw_summary_two_col(c, summaries, expenses_map, cash_map, shop_name, sale_date)
    c.save()


# ── 2-column summary helpers ──────────────────────────────────────────────────

_S_FS  = 7.5    # summary font size
_S_LH  = 10.5   # summary line height
_S_GAP = 8      # vertical gap between rows of blocks
_S_DIV = 7      # horizontal gutter between the two summary columns


def _s_col_w():
    return (PAGE_W - 2*MARGIN - _S_DIV) / 2   # ~260 pt


def _render_counter_block(c, s, exps, crows, cx, y, cw):
    """
    Draw one counter's full detail block at (cx, y) with column-width cw.
    Returns final y (lower than start — blocks draw downward in PDF coords).
    No automatic page breaks — caller must ensure enough space.
    """
    FS, LH = _S_FS, _S_LH
    rx = cx + cw   # right edge

    def wl(left, right=None, bold=False, size=FS, indent=0, step=LH*0.85):
        f = FONTB if bold else FONT
        c.setFont(f, size)
        if left:
            c.drawString(cx + indent, y[0], left)
        if right:
            c.drawRightString(rx, y[0], right)
        y[0] -= step

    def rule(thick=0.25, step=LH*0.55):
        c.setLineWidth(thick)
        c.line(cx, y[0], rx, y[0])
        y[0] -= step

    y = [y]   # mutable ref

    # Counter header
    wl(f"■ {s['name']}  —  {s['staff_name'] or '—'}",
       bold=True, size=FS+0.5, step=LH*0.4)
    rule(0.35)

    # Expenses (non-UPI)
    non_upi = [e for e in exps if not e["is_upi"]]
    if non_upi:
        wl("Expenses", bold=True, indent=4, step=LH*0.75)
        for e in non_upi:
            lbl = EXPENSE_LABEL.get(e["category"], e["category"])
            wl(lbl, f"{RS}{Decimal(str(e['amount'])):,.0f}", indent=10)

    # UPI
    upi_rows = [e for e in exps if e["is_upi"]]
    if upi_rows:
        wl("UPI / Google Pay", bold=True, indent=4, step=LH*0.75)
        for e in upi_rows:
            lbl = EXPENSE_LABEL.get(e["category"], e["category"])
            wl(lbl, f"{RS}{Decimal(str(e['amount'])):,.0f}", indent=10)

    # Cash denominations
    for cat_key, cat_lbl in [("COLLECTION", "Collection"),
                               ("DRAWER_CASH","Drawer Cash"),
                               ("TIPS_CASH",  "Tips Cash")]:
        rows = [r for r in crows if r["cash_category"] == cat_key]
        if not rows:
            continue
        wl(cat_lbl, bold=True, indent=4, step=LH*0.75)
        tot = Decimal(0)
        for r in rows:
            dv  = r["denomination_value"]
            cnt = r["count"]
            ta  = Decimal(str(r["total_amount"]))
            lbl = (f"{RS}{dv:>4} × {cnt}"
                   if r["denomination_type"] == "NOTE"
                   else f"Coins × {cnt}")
            wl(lbl, f"{RS}{ta:,.0f}", indent=10)
            tot += ta
        wl("Total", f"{RS}{tot:,.0f}", bold=True, indent=10)

    # Counter summary block
    wl("Summary", bold=True, indent=4, step=LH*0.75)
    for lbl, field in [
        ("Liquor Sale", "liquor_sale_amount"),
        ("Beer Sale",   "beer_sale_amount"),
        ("Tips",        "total_tips"),
        ("Grand Total", "grand_total_by_sale"),
        ("Google Pay",  "google_pay_amount"),
        ("Expenses",    "expenses_total"),
        ("Collection",  "collection_total"),
        ("Drawer Cash", "drawer_cash_total"),
        ("Tips Cash",   "tips_cash_total"),
    ]:
        wl(lbl, f"{RS}{Decimal(str(s[field])):,.0f}", indent=10)

    tally = Decimal(str(s["tally_difference"]))
    sign  = "+" if tally > 0 else ""
    wl("Tally", f"{sign}{RS}{tally:,.0f}", bold=True, indent=10)

    y[0] -= LH * 0.3
    rule(0.2, step=0)   # closing rule, no extra step

    return y[0]


def _render_shop_total_block(c, shop_totals, cx, y, cw):
    """Draw shop total block. Returns final y."""
    FS, LH = _S_FS, _S_LH
    rx = cx + cw
    y  = [y]

    def wl(left, right=None, bold=False, size=FS, indent=0, step=LH*0.85):
        f = FONTB if bold else FONT
        c.setFont(f, size)
        if left:
            c.drawString(cx + indent, y[0], left)
        if right:
            c.drawRightString(rx, y[0], right)
        y[0] -= step

    def rule(thick=0.25, step=LH*0.55):
        c.setLineWidth(thick)
        c.line(cx, y[0], rx, y[0])
        y[0] -= step

    wl("■ SHOP TOTAL — all counters", bold=True, size=FS+0.5, step=LH*0.4)
    rule(0.5)
    for lbl, k in [
        ("Liquor Sale",     "liquor"),
        ("Beer Sale",       "beer"),
        ("Tips",            "tips"),
        ("Grand Total",     "grand"),
        ("Google Pay (UPI)","upi"),
        ("Expenses",        "expenses"),
        ("Collection",      "collection"),
        ("Drawer Cash",     "drawer"),
        ("Tips Cash",       "tips_cash"),
    ]:
        wl(lbl, f"{RS}{shop_totals[k]:,.0f}", indent=4)

    t    = shop_totals["tally"]
    sign = "+" if t > 0 else ""
    wl("Tally Difference", f"{sign}{RS}{t:,.0f}", bold=True, indent=4)
    y[0] -= LH * 0.3
    rule(0.2, step=0)
    return y[0]


def _draw_summary_two_col(c, summaries, expenses_map, cash_map, shop_name, sale_date):
    """
    Render summary in a 2-column grid:
      Row 1:  Counter 1  |  Counter 2
      Row 2:  Counter 3  |  Shop Total
    A vertical divider separates the columns; a horizontal rule separates rows.
    """
    cw   = _s_col_w()
    cx   = [MARGIN, MARGIN + cw + _S_DIV]
    top  = TOP_Y - 8

    # Compute shop totals (needed for shop total block)
    totals = {k: Decimal(0) for k in
              ["liquor","beer","tips","grand","upi","expenses",
               "collection","drawer","tips_cash","tally"]}
    for s in summaries:
        for k, field in [
            ("liquor","liquor_sale_amount"), ("beer","beer_sale_amount"),
            ("tips","total_tips"),           ("grand","grand_total_by_sale"),
            ("upi","google_pay_amount"),     ("expenses","expenses_total"),
            ("collection","collection_total"), ("drawer","drawer_cash_total"),
            ("tips_cash","tips_cash_total"), ("tally","tally_difference"),
        ]:
            totals[k] += Decimal(str(s[field]))

    def _draw_vertical_div(y_top, y_bot):
        mid = cx[1] - _S_DIV / 2
        c.setLineWidth(0.3)
        c.line(mid, y_top, mid, y_bot)

    def _draw_horiz_div(y_pos):
        c.setLineWidth(0.4)
        c.line(MARGIN, y_pos, PAGE_W - MARGIN, y_pos)

    def _new_page():
        c.showPage()
        _page_header(c, shop_name, sale_date, "Summary (cont.)")
        return TOP_Y - 8

    # ── Row 1: Counter 1 (left) | Counter 2 (right) ──────────────
    row_top = top
    s0 = summaries[0]
    s1 = summaries[1]
    y_left  = _render_counter_block(c, s0,
                                    expenses_map.get(s0["id"], []),
                                    cash_map.get(s0["id"], []),
                                    cx[0], row_top, cw)
    y_right = _render_counter_block(c, s1,
                                    expenses_map.get(s1["id"], []),
                                    cash_map.get(s1["id"], []),
                                    cx[1], row_top, cw)
    row_bot = min(y_left, y_right)
    _draw_vertical_div(row_top, row_bot)

    # ── Row separator ─────────────────────────────────────────────
    sep_y = row_bot - _S_GAP / 2
    _draw_horiz_div(sep_y)
    row_top = sep_y - _S_GAP / 2

    # New page if Row 2 won't fit (conservative estimate: 280 pt)
    if row_top - 280 < BOT_Y + 20:
        row_top = _new_page()

    # ── Row 2: Counter 3 (left) | Shop Total (right) ─────────────
    s2 = summaries[2]
    y_left  = _render_counter_block(c, s2,
                                    expenses_map.get(s2["id"], []),
                                    cash_map.get(s2["id"], []),
                                    cx[0], row_top, cw)
    y_right = _render_shop_total_block(c, totals, cx[1], row_top, cw)
    row_bot = min(y_left, y_right)
    _draw_vertical_div(row_top, row_bot)


# ── Billing algorithm ─────────────────────────────────────────────────────────

def fetch_stock_lots(cur, shop_id, sale_date):
    """
    Returns {product_id: [(lot_id, mrp, remaining_qty), ...]} sorted oldest-first.

    remaining_qty = initial_qty - consumed_qty (maintained by save_bills /
    delete_day_bills).  Only lots with remaining_qty > 0 are included.

    Products with no StockLot entry fall back to DCS mrp (legacy path).
    """
    cur.execute("""
        SELECT sl.id, sl.product_id, sl.mrp_per_bottle,
               sl.initial_qty - sl.consumed_qty AS remaining
        FROM "StockLot" sl
        WHERE sl.shop_id = %s
          AND sl.lot_date <= %s
          AND (sl.initial_qty - sl.consumed_qty) > 0
        ORDER BY sl.product_id, sl.lot_date, sl.id
    """, (shop_id, sale_date))
    result = {}
    for lot_id, pid, mrp, remaining in cur.fetchall():
        if pid not in result:
            result[pid] = []
        result[pid].append((lot_id, Decimal(str(mrp)), int(remaining)))
    return result


def fetch_sales_fifo(cur, shop_id, sale_date):
    """
    Returns (bill_products, counter_data).

    MRP for each product comes from StockLot (permit-sourced, DB-stored).
    Products with no StockLot fall back to DCS mrp_per_bottle.

    FIFO: oldest StockLot consumed first.  When old stock exhausted the
    bill splits automatically at the lot boundary.
    """
    cur.execute("""
        SELECT dcs.counter_id, p.id, p.short_name, p.volume_ml, p.master_code,
               pc.product_type,
               dcs.sold_btls,
               dcs.mrp_per_bottle        AS dcs_mrp,
               dcs.selling_price_per_bottle AS dcs_sp
        FROM "DailyCounterStock" dcs
        JOIN "Counter"          c  ON c.id = dcs.counter_id
        JOIN "Product"          p  ON p.id = dcs.product_id
        JOIN "ProductCategory"  pc ON p.category_id = pc.id
        WHERE c.shop_id = %s AND dcs.stock_date = %s AND dcs.sold_btls > 0
        ORDER BY pc.product_type DESC, p.master_code
    """, (shop_id, sale_date))
    cols = [d[0] for d in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]

    stock_lots = fetch_stock_lots(cur, shop_id, sale_date)

    # Phase 1: aggregate sold qty per product across counters
    prod_agg = {}   # pid → {total_sold, meta, counter rows}
    ctr_rows = {}   # pid → [{ctr, sold, dcs_sp}]

    for row in rows:
        ctr   = row["counter_id"]
        pid   = row["id"]
        sold  = row["sold_btls"]
        dcs_sp = Decimal(str(row["dcs_sp"]))
        ptype  = row["product_type"]

        if pid not in prod_agg:
            prod_agg[pid] = {
                "total_sold":  0,
                "short_name":  row["short_name"],
                "volume_ml":   row["volume_ml"],
                "master_code": row["master_code"],
                "ptype":       ptype,
                "dcs_mrp":     Decimal(str(row["dcs_mrp"])),
                "dcs_sp":      dcs_sp,
            }
            ctr_rows[pid] = []
        prod_agg[pid]["total_sold"] += sold
        ctr_rows[pid].append({"ctr": ctr, "sold": sold, "dcs_sp": dcs_sp})

    # Phase 2: FIFO from StockLot, build bill_map and counter_data
    bill_map     = {}   # (pid, lot_id) → dict
    counter_data = {}   # counter_id → {liq_mrp, beer_mrp, tips}

    for pid, agg in prod_agg.items():
        total_sold = agg["total_sold"]
        ptype      = agg["ptype"]
        dcs_sp     = agg["dcs_sp"]
        lots       = stock_lots.get(pid)

        if lots:
            # StockLot path: FIFO from permit-sourced lots
            fifo_lots = []  # [(lot_id, mrp, qty)]
            remaining  = total_sold

            for lot_id, mrp, avail in lots:
                if remaining <= 0:
                    break
                take = min(remaining, avail)
                fifo_lots.append((lot_id, mrp, take))
                remaining -= take

            if remaining > 0:
                # Sold more than all known lots — extend last lot
                if fifo_lots:
                    lid, lmrp, lqty = fifo_lots[-1]
                    fifo_lots[-1] = (lid, lmrp, lqty + remaining)
                else:
                    fifo_lots.append((None, agg["dcs_mrp"], remaining))
        else:
            # Legacy path: no StockLot — use DCS mrp directly
            fifo_lots = [(None, agg["dcs_mrp"], total_sold)]

        # Accumulate into bill_map
        for lot_id, mrp, qty in fifo_lots:
            key = (pid, lot_id)
            if key not in bill_map:
                bill_map[key] = {
                    "id":           pid,
                    "short_name":   agg["short_name"],
                    "volume_ml":    agg["volume_ml"],
                    "master_code":  agg["master_code"],
                    "product_type": ptype,
                    "total_sold":   0,
                    "mrp":          mrp,
                    "stock_lot_id": lot_id,
                }
            bill_map[key]["total_sold"] += qty

        # Distribute FIFO quantities to counter_data (proportional by counter sales)
        tot_dec = Decimal(str(total_sold)) if total_sold else Decimal("1")
        for cr in ctr_rows[pid]:
            ctr    = cr["ctr"]
            c_sold = Decimal(str(cr["sold"]))
            c_sp   = cr["dcs_sp"]

            if ctr not in counter_data:
                counter_data[ctr] = {
                    "liq_mrp":  Decimal(0),
                    "beer_mrp": Decimal(0),
                    "tips":     Decimal(0),
                }
            for lot_id, mrp, lot_qty in fifo_lots:
                ratio    = Decimal(str(lot_qty)) / tot_dec
                qty_c    = c_sold * ratio
                mrp_amt  = qty_c * mrp
                tips_amt = qty_c * (c_sp - mrp)
                if ptype == "IMFL":
                    counter_data[ctr]["liq_mrp"] += mrp_amt
                else:
                    counter_data[ctr]["beer_mrp"] += mrp_amt
                counter_data[ctr]["tips"] += tips_amt

    # Sort: IMFL first, then by master_code, higher MRP first (old lot before new)
    bill_products = sorted(
        bill_map.values(),
        key=lambda x: (0 if x["product_type"] == "IMFL" else 1,
                       x["master_code"],
                       -float(x["mrp"])),
    )
    return bill_products, counter_data


def recalculate_summary(cur, shop_id, sale_date, counter_data):
    """
    Rewrite liquor_sale_amount, beer_sale_amount, total_tips, grand_total_by_sale
    in DailyCounterSummary from FIFO-correct bill figures.
    collection_total, drawer_cash_total, google_pay_amount, expenses_total,
    tips_cash_total, tally_difference are left unchanged.
    """
    for counter_id, d in counter_data.items():
        grand = d["liq_mrp"] + d["beer_mrp"] + d["tips"]
        cur.execute("""
            UPDATE "DailyCounterSummary"
               SET liquor_sale_amount  = %s,
                   beer_sale_amount    = %s,
                   total_tips          = %s,
                   grand_total_by_sale = %s
             WHERE counter_id   = %s
               AND summary_date = %s
        """, (str(d["liq_mrp"]), str(d["beer_mrp"]),
              str(d["tips"]),    str(grand),
              counter_id, sale_date))


def _pack_bills(products, limit_ml, shop_short, year, start_seq):
    """Pack product list into bills respecting limit_ml. Returns (bills, next_seq)."""
    bills   = []
    cur_its = []
    cur_ml  = 0
    cur_amt = Decimal(0)
    seq     = start_seq

    def flush():
        nonlocal cur_its, cur_ml, cur_amt, seq
        if cur_its:
            bills.append({
                "sequence_no":  seq,
                "bill_number":  f"{shop_short}-{year}-{seq}",
                "total_ml":     cur_ml,
                "total_amount": cur_amt,
                "items":        cur_its,
            })
            seq += 1
        cur_its, cur_ml, cur_amt = [], 0, Decimal(0)

    for p in products:
        remaining = p["total_sold"]
        if remaining <= 0:
            continue
        vol = p["volume_ml"]
        mrp = Decimal(str(p["mrp"]))
        while remaining > 0:
            cap = limit_ml - cur_ml
            if cap < vol:
                flush()
                cap = limit_ml
            take = min(cap // vol, remaining)
            if take == 0:
                flush()
                continue
            cur_its.append({
                "product_id":     p["id"],
                "short_name":     p["short_name"],
                "bottles":        take,
                "ml_per_bottle":  vol,
                "total_ml":       take * vol,
                "mrp_per_bottle": mrp,
                "total_amount":   Decimal(take) * mrp,
                "stock_lot_id":   p.get("stock_lot_id"),
            })
            cur_ml  += take * vol
            cur_amt += Decimal(take) * mrp
            remaining -= take

    flush()
    return bills, seq


def _split_bill_at(bill, budget):
    """
    Split a bill's items: greedily allocate items to UPI (≤ budget),
    rest goes to Cash.  Returns (upi_bill_dict | None, cash_bill_dict | None).
    """
    upi_items  = []
    cash_items = []
    rem = budget

    for item in bill["items"]:
        if rem <= 0:
            cash_items.append(item)
            continue
        mrp      = item["mrp_per_bottle"]
        can_take = min(int(rem / mrp), item["bottles"])
        if can_take > 0:
            upi_items.append({**item,
                "bottles":      can_take,
                "total_ml":     can_take * item["ml_per_bottle"],
                "total_amount": Decimal(can_take) * mrp,
            })
            rem -= Decimal(can_take) * mrp
        leftover = item["bottles"] - can_take
        if leftover > 0:
            cash_items.append({**item,
                "bottles":      leftover,
                "total_ml":     leftover * item["ml_per_bottle"],
                "total_amount": Decimal(leftover) * item["mrp_per_bottle"],
            })

    def _make(items):
        if not items:
            return None
        return {
            "sequence_no":  bill["sequence_no"],
            "bill_number":  bill["bill_number"],
            "total_ml":     sum(it["total_ml"]     for it in items),
            "total_amount": sum(it["total_amount"] for it in items),
            "items":        items,
        }
    return _make(upi_items), _make(cash_items)


def generate_unified_bills(imfl_products, beer_products,
                           upi_target, shop_short, year, start_seq):
    """
    Generate all bills (IMFL then Beer) with a unified sequence.
    UPI bills come first; the boundary bill is split so UPI total ≤ upi_target
    (greedy item-level split — may be off by < one bottle's MRP).
    Returns list of bills, each with mode='UPI'|'CASH' and time_str.
    """
    imfl_bills, nxt = _pack_bills(imfl_products, IMFL_LIMIT_ML, shop_short, year, start_seq)
    beer_bills, nxt = _pack_bills(beer_products, BEER_LIMIT_ML, shop_short, year, nxt)
    all_bills = imfl_bills + beer_bills

    upi_bills  = []
    cash_bills = []
    remaining  = upi_target

    for b in all_bills:
        if remaining <= 0:
            cash_bills.append({**b, "mode": "CASH"})
        elif b["total_amount"] <= remaining:
            upi_bills.append({**b, "mode": "UPI"})
            remaining -= b["total_amount"]
        else:
            # Boundary bill — split it
            upi_part, cash_part = _split_bill_at(b, remaining)
            if upi_part:
                upi_bills.append({**upi_part, "mode": "UPI"})
                remaining -= upi_part["total_amount"]
            if cash_part:
                cash_bills.append({**cash_part, "mode": "CASH"})

    # ── Parcel-bag top-up: spread UPI deficit evenly across all UPI bills ──
    if upi_bills:
        upi_total = sum(b["total_amount"] for b in upi_bills)
        deficit   = upi_target - upi_total
        if deficit > 0:
            n    = len(upi_bills)
            base = deficit // n       # Decimal floor division
            rem  = int(deficit % n)   # how many bills get an extra ₹1
            for i, bill in enumerate(upi_bills):
                bag_amount = base + (Decimal(1) if i < rem else Decimal(0))
                if bag_amount <= 0:
                    continue
                bill["items"].append({
                    "product_id":     None,
                    "short_name":     "Parcel Bag",
                    "bottles":        0,
                    "ml_per_bottle":  0,
                    "total_ml":       0,
                    "mrp_per_bottle": bag_amount,
                    "total_amount":   bag_amount,
                    "stock_lot_id":   None,
                    "is_bag":         True,
                })
                bill["total_amount"] += bag_amount

    # Renumber all bills in combined order
    combined = upi_bills + cash_bills
    for k, b in enumerate(combined):
        b["sequence_no"] = start_seq + k
        b["bill_number"] = f"{shop_short}-{year}-{start_seq + k}"

    # Assign random ascending billing times (11:00 AM – 10:30 PM)
    _assign_times(combined)

    return combined, start_seq + len(combined)


def _assign_times(bills):
    """Assign random ascending times to bills (11:00–22:30)."""
    if not bills:
        return
    start_min = 11 * 60        # 11:00 AM
    end_min   = 22 * 60 + 30   # 10:30 PM
    span      = end_min - start_min
    n         = len(bills)
    times     = sorted(random.randint(start_min, end_min) for _ in range(n))
    for b, t in zip(bills, times):
        h    = t // 60
        m    = t % 60
        s    = random.randint(0, 59)
        ampm = "AM" if h < 12 else "PM"
        h12  = h % 12 or 12
        b["time_str"] = f"{h12}:{m:02d}:{s:02d} {ampm}"


# ── DB helpers ────────────────────────────────────────────────────────────────

def ensure_payment_mode_column(cur):
    cur.execute("""
        ALTER TABLE "DailySalesBill"
        ADD COLUMN IF NOT EXISTS payment_mode TEXT NOT NULL DEFAULT 'CASH'
    """)


def fetch_sales(cur, shop_id, sale_date):
    cur.execute("""
        SELECT p.id, p.short_name, p.volume_ml, p.master_code,
               pc.product_type,
               SUM(dcs.sold_btls)         AS total_sold,
               MAX(dcs.mrp_per_bottle)    AS mrp,
               SUM(dcs.sale_amount_mrp)   AS total_amount_mrp
        FROM "DailyCounterStock" dcs
        JOIN "Counter"          c  ON dcs.counter_id  = c.id
        JOIN "Product"          p  ON dcs.product_id  = p.id
        JOIN "ProductCategory"  pc ON p.category_id   = pc.id
        WHERE c.shop_id    = %s
          AND dcs.stock_date = %s
          AND dcs.sold_btls  > 0
        GROUP BY p.id, p.short_name, p.volume_ml, p.master_code, pc.product_type
        ORDER BY pc.product_type DESC, p.master_code
    """, (shop_id, sale_date))
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, r)) for r in cur.fetchall()]
    # NOTE: kept for reference; main() uses fetch_sales_fifo() instead


def fetch_upi_target(cur, shop_id, sale_date):
    cur.execute("""
        SELECT COALESCE(SUM(dcs.google_pay_amount), 0)
        FROM "DailyCounterSummary" dcs
        JOIN "Counter" c ON c.id = dcs.counter_id
        WHERE c.shop_id = %s AND dcs.summary_date = %s
    """, (shop_id, sale_date))
    return Decimal(str(cur.fetchone()[0]))


def next_seq(cur, shop_id, bill_year):
    cur.execute("""
        SELECT COALESCE(MAX(sequence_no), 0) + 1
        FROM "DailySalesBill"
        WHERE shop_id  = %s
          AND bill_year = %s
          AND bill_type = 'IMFL'
    """, (shop_id, bill_year))
    return cur.fetchone()[0]


def delete_day_bills(cur, shop_id, bill_date):
    cur.execute("""
        SELECT dsbi.stock_lot_id, SUM(dsbi.bottles)
        FROM "DailySalesBillItem" dsbi
        JOIN "DailySalesBill"     dsb ON dsb.id = dsbi.bill_id
        WHERE dsb.shop_id  = %s
          AND dsb.bill_date = %s
          AND dsbi.stock_lot_id IS NOT NULL
        GROUP BY dsbi.stock_lot_id
    """, (shop_id, bill_date))
    for lot_id, bottles in cur.fetchall():
        cur.execute("""
            UPDATE "StockLot"
            SET consumed_qty = GREATEST(0, consumed_qty - %s)
            WHERE id = %s
        """, (int(bottles), lot_id))
    cur.execute("""
        DELETE FROM "DailySalesBillItem"
        WHERE bill_id IN (
            SELECT id FROM "DailySalesBill"
            WHERE shop_id = %s AND bill_date = %s
        )
    """, (shop_id, bill_date))
    cur.execute("""
        DELETE FROM "DailySalesBill"
        WHERE shop_id = %s AND bill_date = %s
    """, (shop_id, bill_date))


def delete_all_bills(cur, shop_id, year):
    """Delete all bills for the shop+year (clean slate for year sequence)."""
    cur.execute("""
        SELECT dsb.shop_id, dsbi.stock_lot_id, SUM(dsbi.bottles)
        FROM "DailySalesBillItem" dsbi
        JOIN "DailySalesBill"     dsb ON dsb.id = dsbi.bill_id
        WHERE dsb.shop_id  = %s
          AND dsb.bill_year = %s
          AND dsbi.stock_lot_id IS NOT NULL
        GROUP BY dsb.shop_id, dsbi.stock_lot_id
    """, (shop_id, year))
    for _, lot_id, bottles in cur.fetchall():
        cur.execute("""
            UPDATE "StockLot"
            SET consumed_qty = GREATEST(0, consumed_qty - %s)
            WHERE id = %s
        """, (int(bottles), lot_id))
    cur.execute("""
        DELETE FROM "DailySalesBillItem"
        WHERE bill_id IN (
            SELECT id FROM "DailySalesBill"
            WHERE shop_id = %s AND bill_year = %s
        )
    """, (shop_id, year))
    cur.execute("""
        DELETE FROM "DailySalesBill"
        WHERE shop_id = %s AND bill_year = %s
    """, (shop_id, year))


def save_bills(cur, bills, shop_id, bill_date, bill_year):
    for b in bills:
        cur.execute("""
            INSERT INTO "DailySalesBill"
                (shop_id, bill_date, bill_year, bill_type, sequence_no, bill_number,
                 total_ml, total_amount, payment_mode)
            VALUES (%s, %s, %s, 'IMFL', %s, %s, %s, %s, %s)
            RETURNING id
        """, (shop_id, bill_date, bill_year,
              b["sequence_no"], b["bill_number"],
              b["total_ml"], str(b["total_amount"]),
              b["mode"]))
        bill_id = cur.fetchone()[0]
        rows = [
            (bill_id, it["product_id"], it.get("stock_lot_id"),
             it["bottles"], it["ml_per_bottle"],
             it["total_ml"], str(it["mrp_per_bottle"]), str(it["total_amount"]))
            for it in b["items"]
            if it.get("product_id") is not None   # skip parcel-bag surcharges
        ]
        execute_values(cur, """
            INSERT INTO "DailySalesBillItem"
                (bill_id, product_id, stock_lot_id, bottles, ml_per_bottle,
                 total_ml, mrp_per_bottle, total_amount)
            VALUES %s
        """, rows)
        # Update StockLot consumed_qty for FIFO tracking
        from collections import defaultdict
        lot_consumption = defaultdict(int)
        for it in b["items"]:
            if it.get("stock_lot_id") and it.get("product_id") is not None:
                lot_consumption[it["stock_lot_id"]] += it["bottles"]
        for lot_id, consumed in lot_consumption.items():
            cur.execute("""
                UPDATE "StockLot" SET consumed_qty = consumed_qty + %s WHERE id = %s
            """, (consumed, lot_id))


def fetch_counter_summaries(cur, shop_id, sale_date):
    cur.execute("""
        SELECT c.id, c.name, c.display_order, dcs.staff_name,
               dcs.liquor_sale_amount, dcs.beer_sale_amount,
               dcs.total_tips, dcs.grand_total_by_sale,
               dcs.google_pay_amount, dcs.expenses_total,
               dcs.collection_total, dcs.drawer_cash_total,
               dcs.tips_cash_total, dcs.tally_difference
        FROM "Counter" c
        JOIN "DailyCounterSummary" dcs
             ON dcs.counter_id = c.id AND dcs.summary_date = %s
        WHERE c.shop_id = %s
        ORDER BY c.display_order
    """, (sale_date, shop_id))
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, r)) for r in cur.fetchall()]


def fetch_expenses(cur, shop_id, sale_date):
    cur.execute("""
        SELECT c.id, de.category, de.amount, de.is_upi, de.upi_ref
        FROM "Counter" c
        JOIN "DailyExpense" de
             ON de.counter_id = c.id AND de.expense_date = %s
        WHERE c.shop_id = %s
        ORDER BY c.display_order, de.category
    """, (sale_date, shop_id))
    result = {}
    for cid, cat, amt, is_upi, ref in cur.fetchall():
        result.setdefault(cid, []).append(
            {"category": cat, "amount": amt, "is_upi": is_upi, "upi_ref": ref})
    return result


def fetch_cash(cur, shop_id, sale_date):
    cur.execute("""
        SELECT c.id, cr.cash_category, cr.denomination_type,
               cr.denomination_value, cr.count, cr.total_amount
        FROM "Counter" c
        JOIN "CashRecord" cr
             ON cr.counter_id = c.id AND cr.record_date = %s
        WHERE c.shop_id = %s
        ORDER BY c.display_order, cr.cash_category,
                 cr.denomination_type DESC, cr.denomination_value DESC
    """, (sale_date, shop_id))
    result = {}
    for cid, cat, dtype, dval, cnt, total in cur.fetchall():
        result.setdefault(cid, []).append(
            {"cash_category": cat, "denomination_type": dtype,
             "denomination_value": dval, "count": cnt, "total_amount": total})
    return result


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Generate KSBCL daily bills — A4 PDF")
    ap.add_argument("--date",  required=True, help="Sale date YYYY-MM-DD")
    ap.add_argument("--shop",  default="KLS")
    ap.add_argument("--no-save",  action="store_true",
                    help="Print summary only — do not write to DB or PDF")
    ap.add_argument("--delete-all-first", action="store_true",
                    help="Delete all bills for the year before generating (resets sequence)")
    ap.add_argument("--output", metavar="FILE",
                    help="PDF output path (default: Reports/{shop}-{date}-bills.pdf)")
    args = ap.parse_args()

    sale_date = date_type.fromisoformat(args.date)
    year      = sale_date.year

    conn = psycopg2.connect(**CONN)
    conn.autocommit = False
    cur  = conn.cursor()

    cur.execute('SELECT id, name, short_name FROM "Shop" WHERE short_name = %s',
                (args.shop,))
    row = cur.fetchone()
    if not row:
        print(f"ERROR: shop '{args.shop}' not found.")
        return
    shop_id, shop_name, shop_short = row

    if not args.no_save:
        ensure_payment_mode_column(cur)

    if args.delete_all_first and not args.no_save:
        print(f"▶ Deleting all {year} bills for {shop_short}...")
        delete_all_bills(cur, shop_id, year)
        conn.commit()
        print("  Done.")

    products, counter_data = fetch_sales_fifo(cur, shop_id, sale_date)
    summaries = fetch_counter_summaries(cur, shop_id, sale_date)
    expenses  = fetch_expenses(cur, shop_id, sale_date)
    cash      = fetch_cash(cur, shop_id, sale_date)

    if not products:
        print(f"No sales found for {shop_short} on {sale_date}.")
        return

    upi_target = fetch_upi_target(cur, shop_id, sale_date)

    if not args.no_save:
        delete_day_bills(cur, shop_id, sale_date)

    imfl_products = [p for p in products if p["product_type"] == "IMFL"]
    beer_products = [p for p in products if p["product_type"] == "BEER"]

    start = next_seq(cur, shop_id, year) if not args.no_save else 1
    all_bills, _ = generate_unified_bills(
        imfl_products, beer_products, upi_target, shop_short, year, start)

    # Summary
    total_bills = len(all_bills)
    upi_bills   = [b for b in all_bills if b["mode"] == "UPI"]
    cash_bills  = [b for b in all_bills if b["mode"] == "CASH"]
    total_amt   = sum(b["total_amount"] for b in all_bills)
    upi_amt     = sum(b["total_amount"] for b in upi_bills)
    cash_amt    = sum(b["total_amount"] for b in cash_bills)

    print(f"\n  {shop_short}  {sale_date}  —  {total_bills} bill(s)")
    print(f"  Sequence : {start} → {start + total_bills - 1}")
    print(f"  UPI bills: {len(upi_bills)}   ₹{upi_amt:,.2f}  (target ₹{upi_target:,.2f})")
    print(f"  Cash bills:{len(cash_bills)}   ₹{cash_amt:,.2f}")
    print(f"  Total    : ₹{total_amt:,.2f}")

    if not args.no_save and all_bills:
        save_bills(cur, all_bills, shop_id, sale_date, year)
        recalculate_summary(cur, shop_id, sale_date, counter_data)
        conn.commit()
        print(f"  Saved {total_bills} bill(s) to DB.")

    os.makedirs(os.path.abspath(REPORTS_DIR), exist_ok=True)

    if not args.no_save:
        bills_pdf   = args.output or os.path.join(
            REPORTS_DIR, f"{shop_short}-{sale_date}-bills.pdf")
        summary_pdf = bills_pdf.replace("-bills.pdf", "-summary.pdf")

        generate_bills_pdf(all_bills, shop_name, sale_date, bills_pdf)
        print(f"  Bills PDF   : {bills_pdf}")

        generate_summary_pdf(summaries, expenses, cash,
                             shop_name, sale_date, summary_pdf)
        print(f"  Summary PDF : {summary_pdf}")

    cur.close()
    conn.close()


if __name__ == "__main__":
    main()
