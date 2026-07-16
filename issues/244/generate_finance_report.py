#!/usr/bin/env python3
"""
Personal Finance ROI Report Generator
======================================
Generates a detailed Word document report for remote worker in Sri Lanka
investing USD salary over a 1-year period across:
  Path A  – Sri Lanka local bank deposits (LKR, simple interest)
  Path B  – Crypto-exchange stablecoin earn products (USD, compound interest)
  Path C  – Mixed 50/50 strategy

Active investment carried forward:
  OKX USDT, 500 USDT @ 10 % APR (180-day PROMO), started 26 Feb 2026,
  matures 26 Aug 2026, expected maturity value 525.27 USDT.
"""

from datetime import date, timedelta
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ─── CONSTANTS ──────────────────────────────────────────────────────────────
USD_LKR_RATE = 300.0          # 1 USD = 300 LKR (Mar 2026 approximate mid-rate)
REPORT_START  = date(2026, 2, 23)   # First salary received
REPORT_END    = date(2027, 2, 23)   # 12 months later
TOTAL_DAYS    = (REPORT_END - REPORT_START).days   # 365

# ─── SALARY SCHEDULE ────────────────────────────────────────────────────────
# Bi-monthly for first 6 months (months 2,4,6 → Feb,Apr,Jun 23)
# Monthly from month 7 onward (Jul–Dec 23)
SALARY_DATES = [
    date(2026, 2, 23),   # Salary 1 – already invested in OKX (Path B only)
    date(2026, 4, 23),   # Salary 2
    date(2026, 6, 23),   # Salary 3
    date(2026, 7, 23),   # Salary 4
    date(2026, 8, 23),   # Salary 5
    date(2026, 9, 23),   # Salary 6
    date(2026, 10, 23),  # Salary 7
    date(2026, 11, 23),  # Salary 8
    date(2026, 12, 23),  # Salary 9
]
SALARY_USD = 500.0
TOTAL_INCOME_USD = len(SALARY_DATES) * SALARY_USD   # $4,500

# ─── OKX ACTIVE INVESTMENT ──────────────────────────────────────────────────
OKX_START       = date(2026, 2, 26)
OKX_END         = date(2026, 8, 26)
OKX_PRINCIPAL   = 500.0
OKX_MATURITY    = 525.27    # as per tracker
OKX_INTEREST    = OKX_MATURITY - OKX_PRINCIPAL   # 25.27

# ─── PATH A – BEST LKR BANK PRODUCTS ────────────────────────────────────────
# Rates in decimal per annum (simple interest), chosen for highest rate per tenor
BANK_PRODUCTS_LKR = {
    "1M":  {"bank": "DFCC",     "type": "FD",  "rate": 0.0750, "days": 30},
    "3M":  {"bank": "PAN ASIA", "type": "FD",  "rate": 0.0825, "days": 91},
    "6M":  {"bank": "PAN ASIA", "type": "FD",  "rate": 0.0850, "days": 183},
    "9M":  {"bank": "DFCC",     "type": "FD",  "rate": 0.0850, "days": 274},   # ~200D product
    "1Y":  {"bank": "PAN ASIA", "type": "FD",  "rate": 0.0900, "days": 365},
    "400D":{"bank": "SAMPATH",  "type": "FD",  "rate": 0.0900, "days": 400},
    "2Y":  {"bank": "SAMPATH",  "type": "FD",  "rate": 0.1000, "days": 730},
}

# ─── PATH B – BEST EXCHANGE EARN PRODUCTS (COMPOUNDING) ─────────────────────
# For each $500 new salary we use:
#   300 USDT → MEXC Flexible 15 % APR (tier-1 limit 300 USDT)  [daily compound]
#   200 USDT → WEEX Flexible 13 % APR (tier-1 limit 200 USDT)  [daily compound]
MEXC_TIER1_LIMIT = 300.0
MEXC_TIER1_RATE  = 0.15
MEXC_TIER2_RATE  = 0.06

WEEX_TIER1_LIMIT = 200.0
WEEX_TIER1_RATE  = 0.13

# After OKX matures (525.27 USDT) reinvest same allocation:
#   300 USDT → MEXC 15 %
#   200 USDT → WEEX 13 %
#    25.27 → MEXC tier-2 6 %

def compound_interest(principal: float, apr: float, days: int) -> float:
    """Daily compound interest (flexible earn products). Returns final value."""
    return principal * ((1 + apr / 365) ** days)

def compound_interest_gain(principal: float, apr: float, days: int) -> float:
    return compound_interest(principal, apr, days) - principal

def simple_interest_gain(principal_lkr: float, rate_pa: float, days: int) -> float:
    """Simple interest for bank FDs. Principal in LKR."""
    return principal_lkr * rate_pa * (days / 365)

# ─── HELPER: choose best LKR product by days remaining ──────────────────────
def best_lkr_strategy(days_available: int):
    """
    Returns list of (bank, tenor_label, rate, days_used, lkr_interest) tuples
    for a 150,000 LKR deposit given 'days_available' until report end.
    Greedy: fill with best long-tenor products first, remainder in shorter.
    """
    principal = SALARY_USD * USD_LKR_RATE   # 150,000 LKR
    strategy = []
    remaining_days = days_available

    # Sorted descending by days (use longest first)
    tenors = sorted(BANK_PRODUCTS_LKR.items(), key=lambda x: x[1]["days"], reverse=True)

    for label, prod in tenors:
        if remaining_days <= 0:
            break
        d = prod["days"]
        if d <= remaining_days:
            interest = simple_interest_gain(principal, prod["rate"], d)
            strategy.append({
                "tenor": label,
                "bank":  prod["bank"],
                "type":  prod["type"],
                "rate":  prod["rate"],
                "days":  d,
                "principal_lkr": principal,
                "interest_lkr":  round(interest, 2),
            })
            remaining_days -= d
            # For simplicity we reinvest same principal (not adding interest to principal
            # since local banks pay at maturity for simple-interest FDs)
            break  # one main block per salary tranche

    # If remaining days left, add a short-term follow-on
    if remaining_days > 0 and strategy:
        # Use 1M or 3M depending on remaining
        if remaining_days >= 91:
            follow_label = "3M"
        elif remaining_days >= 30:
            follow_label = "1M"
        else:
            follow_label = None

        if follow_label:
            prod = BANK_PRODUCTS_LKR[follow_label]
            fd = min(remaining_days, prod["days"])
            interest = simple_interest_gain(principal, prod["rate"], fd)
            strategy.append({
                "tenor": follow_label,
                "bank":  prod["bank"],
                "type":  prod["type"],
                "rate":  prod["rate"],
                "days":  fd,
                "principal_lkr": principal,
                "interest_lkr":  round(interest, 2),
            })

    return strategy


# ═══════════════════════════════════════════════════════════════════════════
#   CALCULATE ALL THREE PATHS
# ═══════════════════════════════════════════════════════════════════════════

def calculate_path_a():
    """
    Path A: Convert every salary to LKR and deposit in best local bank FD.
    Returns list of per-salary records plus aggregate totals.
    """
    results = []
    for i, sal_date in enumerate(SALARY_DATES):
        days_left = (REPORT_END - sal_date).days
        lkr_principal = SALARY_USD * USD_LKR_RATE
        strategy = best_lkr_strategy(days_left)

        total_interest_lkr = sum(s["interest_lkr"] for s in strategy)
        total_value_lkr    = lkr_principal + total_interest_lkr

        results.append({
            "salary_no":       i + 1,
            "salary_date":     sal_date,
            "usd_amount":      SALARY_USD,
            "lkr_amount":      lkr_principal,
            "days_available":  days_left,
            "strategy":        strategy,
            "total_int_lkr":   round(total_interest_lkr, 2),
            "total_val_lkr":   round(total_value_lkr, 2),
        })

    total_principal_lkr   = TOTAL_INCOME_USD * USD_LKR_RATE
    total_interest_lkr    = sum(r["total_int_lkr"] for r in results)
    total_value_lkr       = total_principal_lkr + total_interest_lkr
    roi_pct               = (total_interest_lkr / total_principal_lkr) * 100

    return results, {
        "total_principal_lkr":  total_principal_lkr,
        "total_interest_lkr":   round(total_interest_lkr, 2),
        "total_value_lkr":      round(total_value_lkr, 2),
        "roi_pct":              round(roi_pct, 2),
        "roi_usd":              round(total_interest_lkr / USD_LKR_RATE, 2),
    }


def calculate_path_b():
    """
    Path B: Keep USD, compound in crypto exchange earn products.
    Salary 1 (Feb 23) → already in OKX, matures Aug 26 → reinvest matured sum.
    Salaries 2-9 → MEXC 15 % + WEEX 13 % flex (daily compound).
    """
    results = []

    # ── Salary 1: OKX active (already invested) ──────────────────────────
    sal1_end   = REPORT_END
    okx_days   = (OKX_END - OKX_START).days                    # 181
    okx_mature = OKX_MATURITY                                   # 525.27

    # After OKX matures on Aug 26, reinvest until Feb 23, 2027
    post_okx_days = (sal1_end - OKX_END).days                  # 181 days
    # Split: 300 MEXC 15%, 200 WEEX 13%, 25.27 MEXC tier2 6%
    okx_mexc_gain  = compound_interest_gain(300.00, MEXC_TIER1_RATE, post_okx_days)
    okx_weex_gain  = compound_interest_gain(200.00, WEEX_TIER1_RATE, post_okx_days)
    okx_extra_gain = compound_interest_gain(25.27,  MEXC_TIER2_RATE, post_okx_days)
    post_okx_total = 300 + 200 + 25.27 + okx_mexc_gain + okx_weex_gain + okx_extra_gain

    total_sal1_value = post_okx_total
    total_sal1_gain  = total_sal1_value - OKX_PRINCIPAL

    results.append({
        "salary_no":    1,
        "salary_date":  date(2026, 2, 23),
        "usd_principal":OKX_PRINCIPAL,
        "phase1": {
            "product":   "OKX USDT Flexible (PROMO 180D)",
            "rate":      0.10,
            "days":      okx_days,
            "principal": OKX_PRINCIPAL,
            "interest":  round(OKX_INTEREST, 2),
            "maturity":  round(okx_mature, 2),
        },
        "phase2": {
            "product":  "MEXC 15% + WEEX 13% Flex",
            "days":     post_okx_days,
            "start":    OKX_END,
            "end":      sal1_end,
            "principal": round(okx_mature, 2),
            "interest":  round(okx_mexc_gain + okx_weex_gain + okx_extra_gain, 4),
            "maturity":  round(total_sal1_value, 2),
        },
        "total_gain": round(total_sal1_gain, 2),
        "final_value": round(total_sal1_value, 2),
    })

    # ── Salaries 2-9: MEXC 15% (300 USDT) + WEEX 13% (200 USDT) ─────────
    for i, sal_date in enumerate(SALARY_DATES[1:], start=2):
        days_avail = (REPORT_END - sal_date).days
        mexc_gain  = compound_interest_gain(MEXC_TIER1_LIMIT, MEXC_TIER1_RATE, days_avail)
        weex_gain  = compound_interest_gain(WEEX_TIER1_LIMIT, WEEX_TIER1_RATE, days_avail)
        total_gain = mexc_gain + weex_gain
        final_val  = SALARY_USD + total_gain

        results.append({
            "salary_no":    i,
            "salary_date":  sal_date,
            "usd_principal":SALARY_USD,
            "allocation": [
                {"exchange": "MEXC",  "asset": "USDT", "product": "Flexible Earn",
                 "rate": MEXC_TIER1_RATE, "principal": MEXC_TIER1_LIMIT,
                 "days": days_avail, "interest": round(mexc_gain, 4)},
                {"exchange": "WEEX",  "asset": "USDT", "product": "Flexible Earn",
                 "rate": WEEX_TIER1_RATE, "principal": WEEX_TIER1_LIMIT,
                 "days": days_avail, "interest": round(weex_gain, 4)},
            ],
            "total_gain":  round(total_gain, 4),
            "final_value": round(final_val, 4),
        })

    total_principal_usd = TOTAL_INCOME_USD
    total_interest_usd  = sum(r["total_gain"] for r in results)
    total_value_usd     = total_principal_usd + total_interest_usd
    roi_pct             = (total_interest_usd / total_principal_usd) * 100

    return results, {
        "total_principal_usd": total_principal_usd,
        "total_interest_usd":  round(total_interest_usd, 2),
        "total_value_usd":     round(total_value_usd, 2),
        "roi_pct":             round(roi_pct, 2),
    }


def calculate_path_c():
    """
    Path C: Mixed 50/50 split every salary.
    $250 → LKR FD in best bank
    $250 → Crypto exchange (MEXC 15 % 150 USDT + WEEX 13 % 100 USDT)
    Salary 1 special: existing OKX 500 USDT (carry as-is, treated as 100% exchange)
    """
    results = []

    # Salary 1: already fully in OKX – keep as exchange only
    sal1_okx_days  = (OKX_END - OKX_START).days
    post_okx_days  = (REPORT_END - OKX_END).days
    okx_mexc_gain  = compound_interest_gain(300.00, MEXC_TIER1_RATE, post_okx_days)
    okx_weex_gain  = compound_interest_gain(200.00, WEEX_TIER1_RATE, post_okx_days)
    okx_extra_gain = compound_interest_gain(25.27,  MEXC_TIER2_RATE, post_okx_days)
    sal1_exchange_final = 300 + 200 + 25.27 + okx_mexc_gain + okx_weex_gain + okx_extra_gain
    sal1_exchange_gain  = sal1_exchange_final - OKX_PRINCIPAL

    results.append({
        "salary_no":        1,
        "salary_date":      date(2026, 2, 23),
        "usd_principal":    OKX_PRINCIPAL,
        "note":             "Fully in OKX (existing). Reinvested after maturity in MEXC+WEEX.",
        "bank_gain_lkr":    0,
        "bank_gain_usd":    0,
        "exchange_gain_usd":round(sal1_exchange_gain, 2),
        "exchange_final_usd":round(sal1_exchange_final, 2),
        "total_gain_usd":   round(sal1_exchange_gain, 2),
        "total_final_usd":  round(sal1_exchange_final, 2),
    })

    for i, sal_date in enumerate(SALARY_DATES[1:], start=2):
        days_avail    = (REPORT_END - sal_date).days
        half_usd      = SALARY_USD / 2.0          # $250
        half_lkr      = half_usd * USD_LKR_RATE   # 75,000 LKR

        # Bank side: best product for 75,000 LKR
        bank_strat = best_lkr_strategy_half(days_avail)
        bank_int_lkr = sum(s["interest_lkr"] for s in bank_strat)
        bank_int_usd = bank_int_lkr / USD_LKR_RATE

        # Exchange side: $250 → MEXC 150 USDT 15% + WEEX 100 USDT 13%
        mexc_gain  = compound_interest_gain(150.0, MEXC_TIER1_RATE, days_avail)
        weex_gain  = compound_interest_gain(100.0, WEEX_TIER1_RATE, days_avail)
        ex_gain    = mexc_gain + weex_gain

        total_gain_usd  = bank_int_usd + ex_gain
        total_final_usd = SALARY_USD + total_gain_usd

        results.append({
            "salary_no":         i,
            "salary_date":       sal_date,
            "usd_principal":     SALARY_USD,
            "bank_principal_lkr":half_lkr,
            "bank_strategy":     bank_strat,
            "bank_gain_lkr":     round(bank_int_lkr, 2),
            "bank_gain_usd":     round(bank_int_usd, 4),
            "exchange_split": [
                {"exchange":"MEXC","asset":"USDT","rate":MEXC_TIER1_RATE,
                 "principal":150.0,"days":days_avail,"interest":round(mexc_gain,4)},
                {"exchange":"WEEX","asset":"USDT","rate":WEEX_TIER1_RATE,
                 "principal":100.0,"days":days_avail,"interest":round(weex_gain,4)},
            ],
            "exchange_gain_usd": round(ex_gain, 4),
            "total_gain_usd":    round(total_gain_usd, 4),
            "total_final_usd":   round(total_final_usd, 4),
        })

    total_principal_usd = TOTAL_INCOME_USD
    total_interest_usd  = sum(r["total_gain_usd"] for r in results)
    total_value_usd     = total_principal_usd + total_interest_usd
    roi_pct             = (total_interest_usd / total_principal_usd) * 100

    return results, {
        "total_principal_usd": total_principal_usd,
        "total_interest_usd":  round(total_interest_usd, 2),
        "total_value_usd":     round(total_value_usd, 2),
        "roi_pct":             round(roi_pct, 2),
    }


def best_lkr_strategy_half(days_available: int):
    """Same as best_lkr_strategy but for $250 → 75,000 LKR."""
    principal = (SALARY_USD / 2.0) * USD_LKR_RATE   # 75,000 LKR
    strategy = []
    remaining_days = days_available

    tenors = sorted(BANK_PRODUCTS_LKR.items(), key=lambda x: x[1]["days"], reverse=True)

    for label, prod in tenors:
        if remaining_days <= 0:
            break
        d = prod["days"]
        if d <= remaining_days:
            interest = simple_interest_gain(principal, prod["rate"], d)
            strategy.append({
                "tenor": label,
                "bank":  prod["bank"],
                "type":  prod["type"],
                "rate":  prod["rate"],
                "days":  d,
                "principal_lkr": principal,
                "interest_lkr":  round(interest, 2),
            })
            remaining_days -= d
            break

    if remaining_days > 0 and strategy:
        if remaining_days >= 91:
            follow_label = "3M"
        elif remaining_days >= 30:
            follow_label = "1M"
        else:
            follow_label = None

        if follow_label:
            prod = BANK_PRODUCTS_LKR[follow_label]
            fd = min(remaining_days, prod["days"])
            interest = simple_interest_gain(principal, prod["rate"], fd)
            strategy.append({
                "tenor": follow_label,
                "bank":  prod["bank"],
                "type":  prod["type"],
                "rate":  prod["rate"],
                "days":  fd,
                "principal_lkr": principal,
                "interest_lkr":  round(interest, 2),
            })

    return strategy


# ═══════════════════════════════════════════════════════════════════════════
#   WORD DOCUMENT BUILDER
# ═══════════════════════════════════════════════════════════════════════════

DARK_BLUE  = RGBColor(0x1F, 0x49, 0x7D)
MID_BLUE   = RGBColor(0x2E, 0x75, 0xB6)
ACCENT_GRN = RGBColor(0x38, 0x96, 0x38)
ACCENT_RED = RGBColor(0xC0, 0x00, 0x00)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
GOLD       = RGBColor(0xFF, 0xC0, 0x00)

def set_cell_bg(cell, rgb_hex: str):
    """Set background color of a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), rgb_hex)
    tcPr.append(shd)

def add_header_row(table, headers: list, bg_hex="1F497D"):
    """Add a styled header row."""
    row = table.rows[0]
    for i, hdr in enumerate(headers):
        cell = row.cells[i]
        cell.text = ""
        set_cell_bg(cell, bg_hex)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(hdr)
        run.bold = True
        run.font.color.rgb = WHITE
        run.font.size = Pt(9)

def add_data_row(table, row_idx: int, values: list, shade_alt=False):
    """Add a data row (optionally alternate-shaded)."""
    row = table.rows[row_idx]
    for i, val in enumerate(values):
        cell = row.cells[i]
        if shade_alt and row_idx % 2 == 0:
            set_cell_bg(cell, "EBF3FB")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(val))
        run.font.size = Pt(9)

def set_col_widths(table, widths_cm: list):
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            row.cells[i].width = Cm(w)

def heading(doc, text, level=1, color=None):
    p = doc.add_heading(text, level=level)
    if color:
        for run in p.runs:
            run.font.color.rgb = color
    return p

def body(doc, text, bold=False, italic=False, color=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color
    run.font.size = Pt(11)
    return p

def bold_kv(doc, key, value, key_color=None, val_color=None):
    p = doc.add_paragraph()
    r1 = p.add_run(f"{key}: ")
    r1.bold = True
    r1.font.size = Pt(11)
    if key_color:
        r1.font.color.rgb = key_color
    r2 = p.add_run(str(value))
    r2.font.size = Pt(11)
    if val_color:
        r2.font.color.rgb = val_color
    return p

def add_page_break(doc):
    doc.add_page_break()


# ─── Main document build function ───────────────────────────────────────────

def build_document(path_a_rows, summary_a, path_b_rows, summary_b,
                   path_c_rows, summary_c):

    doc = Document()

    # ── Page margins ──────────────────────────────────────────────────────
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ══════════════════════════════════════════════════════════════════════
    # TITLE PAGE
    # ══════════════════════════════════════════════════════════════════════
    doc.add_paragraph()
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title_p.add_run("PERSONAL FINANCE ROI REPORT")
    tr.bold = True
    tr.font.size = Pt(24)
    tr.font.color.rgb = DARK_BLUE

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub_p.add_run("1-Year Investment Strategy Analysis\n"
                       "Remote Worker (USA Company) · Sri Lanka Residency\n"
                       "Report Period: 23 Feb 2026 – 23 Feb 2027")
    sr.font.size = Pt(13)
    sr.font.color.rgb = MID_BLUE

    doc.add_paragraph()
    meta_p = doc.add_paragraph()
    meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = meta_p.add_run("Compiled: March 2026  |  Exchange Rate: 1 USD = 300 LKR")
    mr.font.size = Pt(10)
    mr.italic = True

    add_page_break(doc)

    # ══════════════════════════════════════════════════════════════════════
    # TABLE OF CONTENTS (manual)
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "Table of Contents", level=1, color=DARK_BLUE)
    toc_items = [
        "1.  Executive Summary",
        "2.  Scenario Overview & Salary Schedule",
        "3.  Active Investment Status (OKX)",
        "4.  Path A – Local Bank Investment (LKR, Simple Interest)",
        "5.  Path B – Crypto Exchange Investment (USD, Compound Interest)",
        "6.  Path C – Mixed Strategy (50 % Bank + 50 % Exchange)",
        "7.  ROI Comparison & Verdict",
        "8.  Recommended Investment Timeline (Apr 2026 – Dec 2026)",
        "9.  Key Risks & Assumptions",
        "10. Appendix – Interest Rate Reference",
    ]
    for item in toc_items:
        p = doc.add_paragraph(item, style="List Bullet")
        p.runs[0].font.size = Pt(11)

    add_page_break(doc)

    # ══════════════════════════════════════════════════════════════════════
    # 1. EXECUTIVE SUMMARY
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "1. Executive Summary", level=1, color=DARK_BLUE)

    winner = max(
        ("Path A (Local Bank)", summary_a["roi_pct"]),
        ("Path B (Crypto Exchange)", summary_b["roi_pct"]),
        ("Path C (Mixed)", summary_c["roi_pct"]),
        key=lambda x: x[1],
    )

    exec_text = (
        f"This report analyses three investment strategies for a remote worker employed by a "
        f"USA company and residing in Sri Lanka. The employee receives USD 500 salary bi-monthly "
        f"for the first six months (3 salaries), then USD 500 monthly thereafter — totalling "
        f"USD {TOTAL_INCOME_USD:,.0f} over the 12-month period ending 23 February 2027. "
        f"100 % of each salary is committed to investment.\n\n"
        f"Three investment paths are evaluated:\n"
        f"  • Path A: Convert USD → LKR and invest in Sri Lanka bank fixed deposits (simple interest).\n"
        f"  • Path B: Keep USD and earn via crypto-exchange stablecoin products (daily compounding).\n"
        f"  • Path C: Split each salary 50/50 between Path A and Path B.\n\n"
        f"An active OKX investment of USD 500 (PROMO 10 % APR, 180 days, maturing 26 Aug 2026) "
        f"is carried forward into Path B and Path C calculations.\n"
    )
    body(doc, exec_text)

    # Summary table
    doc.add_paragraph()
    tbl = doc.add_table(rows=4, cols=5)
    tbl.style = "Table Grid"
    add_header_row(tbl,
        ["Investment Path", "Total Capital (USD)", "Total Interest (USD)",
         "Portfolio Value (USD)", "ROI (%)"])

    rows_data = [
        ("Path A – Local Bank (LKR)",
         f"$ {summary_a['total_principal_lkr']/USD_LKR_RATE:,.2f}",
         f"$ {summary_a['roi_usd']:,.2f}",
         f"$ {(summary_a['total_value_lkr']/USD_LKR_RATE):,.2f}",
         f"{summary_a['roi_pct']:.2f} %"),
        ("Path B – Crypto Exchange",
         f"$ {summary_b['total_principal_usd']:,.2f}",
         f"$ {summary_b['total_interest_usd']:,.2f}",
         f"$ {summary_b['total_value_usd']:,.2f}",
         f"{summary_b['roi_pct']:.2f} %"),
        ("Path C – Mixed (50/50)",
         f"$ {summary_c['total_principal_usd']:,.2f}",
         f"$ {summary_c['total_interest_usd']:,.2f}",
         f"$ {summary_c['total_value_usd']:,.2f}",
         f"{summary_c['roi_pct']:.2f} %"),
    ]

    for row_i, rd in enumerate(rows_data, start=1):
        add_data_row(tbl, row_i, rd, shade_alt=True)
        # Bold the winner row in green
        if rd[0].startswith(winner[0].split("(")[0].strip()):
            for cell in tbl.rows[row_i].cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                        run.font.color.rgb = ACCENT_GRN

    set_col_widths(tbl, [4.5, 3.5, 3.5, 3.8, 2.0])

    doc.add_paragraph()
    winner_p = doc.add_paragraph()
    wr = winner_p.add_run(
        f"🏆  Highest ROI: {winner[0]}  ({winner[1]:.2f} %)  "
        f"— recommended for maximum returns."
    )
    wr.bold = True
    wr.font.size = Pt(12)
    wr.font.color.rgb = ACCENT_GRN

    add_page_break(doc)

    # ══════════════════════════════════════════════════════════════════════
    # 2. SCENARIO OVERVIEW & SALARY SCHEDULE
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "2. Scenario Overview & Salary Schedule", level=1, color=DARK_BLUE)

    body(doc,
        "You are employed remotely by a US company while living in Sri Lanka. "
        "The following details define the income and investment scenario:")

    details = [
        ("Residence",               "Sri Lanka"),
        ("Employer",                "USA Company (remote work)"),
        ("Salary Amount",           "USD 500 per salary"),
        ("Pay Frequency (Months 1–6)", "Bi-monthly (every 2 months) → 3 salaries"),
        ("Pay Frequency (Months 7–12)","Monthly → 6 salaries"),
        ("Total Salaries",          "9 over 12 months"),
        ("Total Income",            f"USD {TOTAL_INCOME_USD:,.0f}"),
        ("Salary Day",              "23rd of each applicable month"),
        ("Investment Rate",         "100 % of each salary invested immediately"),
        ("Report Period",           "23 Feb 2026 – 23 Feb 2027"),
        ("USD / LKR Rate",          f"1 USD = {USD_LKR_RATE:.0f} LKR (Mar 2026)"),
    ]
    for k, v in details:
        bold_kv(doc, k, v, key_color=MID_BLUE)

    heading(doc, "Salary Schedule", level=2, color=MID_BLUE)
    sal_tbl = doc.add_table(rows=len(SALARY_DATES) + 1, cols=5)
    sal_tbl.style = "Table Grid"
    add_header_row(sal_tbl,
        ["#", "Salary Date", "USD Amount", "LKR Equivalent", "Phase"])

    phases = ["Bi-Monthly"] * 3 + ["Monthly"] * 6
    for i, (sd, ph) in enumerate(zip(SALARY_DATES, phases)):
        row_vals = [
            str(i + 1),
            sd.strftime("%d %b %Y"),
            f"$ {SALARY_USD:,.0f}",
            f"LKR {SALARY_USD * USD_LKR_RATE:,.0f}",
            ph,
        ]
        add_data_row(sal_tbl, i + 1, row_vals, shade_alt=True)

    set_col_widths(sal_tbl, [1.2, 3.0, 3.0, 3.8, 3.0])
    add_page_break(doc)

    # ══════════════════════════════════════════════════════════════════════
    # 3. ACTIVE INVESTMENT STATUS
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "3. Active Investment Status (OKX)", level=1, color=DARK_BLUE)

    body(doc,
        "The first salary (USD 500) was received on 23 February 2026 and invested "
        "in OKX's USDT Flexible Earn PROMO product on 26 February 2026 at 10 % APR "
        "for 180 days. This product is locked until maturity.")

    act_tbl = doc.add_table(rows=2, cols=7)
    act_tbl.style = "Table Grid"
    add_header_row(act_tbl,
        ["Exchange", "Asset", "Principal", "Rate (APR)", "Start Date",
         "Maturity Date", "Maturity Value"])
    add_data_row(act_tbl, 1, [
        "OKX", "USDT",
        f"$ {OKX_PRINCIPAL:,.2f}", "10.00 %",
        OKX_START.strftime("%d %b %Y"),
        OKX_END.strftime("%d %b %Y"),
        f"$ {OKX_MATURITY:,.2f}",
    ], shade_alt=True)
    set_col_widths(act_tbl, [2.2, 1.5, 2.5, 2.5, 2.8, 2.8, 3.0])

    doc.add_paragraph()
    body(doc,
        f"Note on NEXO exclusion: NEXO offers 11.00 % APR on USDT and 9.00 % on USDC; "
        f"however, as noted, insufficient capital is currently available to meet NEXO's "
        f"minimum investment threshold. NEXO is therefore excluded from all calculations.")

    add_page_break(doc)

    # ══════════════════════════════════════════════════════════════════════
    # 4. PATH A – LOCAL BANK INVESTMENT
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "4. Path A – Local Bank Investment (LKR)", level=1, color=DARK_BLUE)

    body(doc,
        "In this path each salary is converted from USD to LKR at 300 LKR/USD and "
        "deposited in the highest-yielding Sri Lanka bank fixed deposit (FD) product "
        "available for the remaining term within the 12-month window. Local banks apply "
        "simple (non-compounding) interest.")

    heading(doc, "4.1  Product Selection Logic", level=2, color=MID_BLUE)
    body(doc,
        "For each salary received, the greedy algorithm selects the longest tenor FD "
        "that fits within the days remaining until 23 Feb 2027, followed by a shorter "
        "follow-on deposit for any residual days. Priority is always given to the highest "
        "rate for the available period.")

    # Best LKR products table
    best_tbl = doc.add_table(rows=len(BANK_PRODUCTS_LKR) + 1, cols=4)
    best_tbl.style = "Table Grid"
    add_header_row(best_tbl,
        ["Tenor", "Bank", "Type", "Rate (p.a.)"])
    for ri, (lbl, prod) in enumerate(
            sorted(BANK_PRODUCTS_LKR.items(), key=lambda x: x[1]["days"]),
            start=1):
        add_data_row(best_tbl, ri,
            [lbl, prod["bank"], prod["type"], f"{prod['rate']*100:.2f} %"],
            shade_alt=True)
    set_col_widths(best_tbl, [2.5, 3.5, 2.5, 3.0])

    heading(doc, "4.2  Per-Salary Investment Detail", level=2, color=MID_BLUE)

    for rec in path_a_rows:
        subhead = doc.add_paragraph()
        sr = subhead.add_run(
            f"Salary {rec['salary_no']}  ·  {rec['salary_date'].strftime('%d %b %Y')}  "
            f"·  USD {rec['usd_amount']:,.0f}  (LKR {rec['lkr_amount']:,.0f})  "
            f"·  {rec['days_available']} days available"
        )
        sr.bold = True
        sr.font.size = Pt(10)
        sr.font.color.rgb = MID_BLUE

        n_rows = len(rec["strategy"]) + 1
        strat_tbl = doc.add_table(rows=n_rows, cols=6)
        strat_tbl.style = "Table Grid"
        add_header_row(strat_tbl,
            ["Bank", "Type", "Tenor", "Rate (p.a.)", "Days", "Interest (LKR)"])
        for ri, s in enumerate(rec["strategy"], start=1):
            add_data_row(strat_tbl, ri, [
                s["bank"], s["type"], s["tenor"],
                f"{s['rate']*100:.2f} %",
                str(s["days"]),
                f"LKR {s['interest_lkr']:,.2f}",
            ], shade_alt=True)
        set_col_widths(strat_tbl, [2.8, 2.2, 2.0, 2.5, 2.0, 3.5])

        total_p = doc.add_paragraph()
        tr2 = total_p.add_run(
            f"  → Total Interest: LKR {rec['total_int_lkr']:,.2f}  "
            f"  → Portfolio Value: LKR {rec['total_val_lkr']:,.2f}"
        )
        tr2.bold = True
        tr2.font.size = Pt(10)
        tr2.font.color.rgb = ACCENT_GRN
        doc.add_paragraph()

    heading(doc, "4.3  Path A Summary", level=2, color=MID_BLUE)

    sum_a_tbl = doc.add_table(rows=2, cols=5)
    sum_a_tbl.style = "Table Grid"
    add_header_row(sum_a_tbl,
        ["Total Capital (LKR)", "Total Interest (LKR)", "Total Value (LKR)",
         "Total Interest (USD)", "ROI (%)"])
    add_data_row(sum_a_tbl, 1, [
        f"LKR {summary_a['total_principal_lkr']:,.0f}",
        f"LKR {summary_a['total_interest_lkr']:,.2f}",
        f"LKR {summary_a['total_value_lkr']:,.2f}",
        f"$ {summary_a['roi_usd']:,.2f}",
        f"{summary_a['roi_pct']:.2f} %",
    ], shade_alt=True)
    set_col_widths(sum_a_tbl, [3.8, 3.8, 3.8, 3.5, 2.5])

    add_page_break(doc)

    # ══════════════════════════════════════════════════════════════════════
    # 5. PATH B – CRYPTO EXCHANGE INVESTMENT
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "5. Path B – Crypto Exchange Investment (USD)", level=1, color=DARK_BLUE)

    body(doc,
        "In this path USD salaries are kept and invested in stablecoin earn products "
        "on regulated crypto exchanges. Earn products pay interest on a daily basis "
        "(compounding). The OKX active investment is continued and reinvested after "
        "maturity. New salaries are split across MEXC (15 % APR, ≤ 300 USDT) and "
        "WEEX (13 % APR, ≤ 200 USDT) for optimised blended yield.")

    heading(doc, "5.1  Exchange Rate Reference", level=2, color=MID_BLUE)

    ex_tbl = doc.add_table(rows=4, cols=5)
    ex_tbl.style = "Table Grid"
    add_header_row(ex_tbl,
        ["Exchange", "Asset", "Product", "Rate (APR)", "Tier Limit"])
    ex_rates = [
        ("OKX",  "USDT", "Flexible Earn (PROMO 180D)", "10.00 %", "500 USDT (active)"),
        ("MEXC", "USDT", "Flexible Earn",               "15.00 %", "300 USDT (Tier 1)"),
        ("WEEX", "USDT", "Flexible Earn",               "13.00 %", "200 USDT (Tier 1)"),
    ]
    for ri, r in enumerate(ex_rates, start=1):
        add_data_row(ex_tbl, ri, list(r), shade_alt=True)
    set_col_widths(ex_tbl, [2.5, 2.0, 4.5, 2.8, 4.0])

    heading(doc, "5.2  Per-Salary Investment Detail", level=2, color=MID_BLUE)

    for rec in path_b_rows:
        subhead = doc.add_paragraph()
        sr2 = subhead.add_run(
            f"Salary {rec['salary_no']}  ·  {rec['salary_date'].strftime('%d %b %Y')}  "
            f"·  USD {rec['usd_principal']:,.2f}"
        )
        sr2.bold = True
        sr2.font.size = Pt(10)
        sr2.font.color.rgb = MID_BLUE

        if rec["salary_no"] == 1:
            # OKX detail
            p1 = rec["phase1"]
            p2 = rec["phase2"]
            n = 3
            detail_tbl = doc.add_table(rows=n, cols=6)
            detail_tbl.style = "Table Grid"
            add_header_row(detail_tbl,
                ["Phase", "Product", "Principal (USDT)", "Rate (APR)", "Days", "Interest (USDT)"])
            add_data_row(detail_tbl, 1, [
                "1 – OKX Lock",
                p1["product"],
                f"$ {p1['principal']:,.2f}",
                f"{p1['rate']*100:.2f} %",
                str(p1["days"]),
                f"$ {p1['interest']:,.2f}",
            ], shade_alt=True)
            add_data_row(detail_tbl, 2, [
                f"2 – Reinvest ({p2['start'].strftime('%d %b')}–{p2['end'].strftime('%d %b %Y')})",
                p2["product"],
                f"$ {p2['principal']:,.2f}",
                "Blended",
                str(p2["days"]),
                f"$ {p2['interest']:,.4f}",
            ], shade_alt=True)
            set_col_widths(detail_tbl, [3.5, 4.5, 3.0, 2.5, 2.0, 3.0])
        else:
            alloc = rec["allocation"]
            n = len(alloc) + 1
            detail_tbl = doc.add_table(rows=n, cols=6)
            detail_tbl.style = "Table Grid"
            add_header_row(detail_tbl,
                ["Exchange", "Asset", "Principal (USDT)", "Rate (APR)", "Days", "Interest (USDT)"])
            for ri, a in enumerate(alloc, start=1):
                add_data_row(detail_tbl, ri, [
                    a["exchange"], a["asset"],
                    f"$ {a['principal']:,.2f}",
                    f"{a['rate']*100:.2f} %",
                    str(a["days"]),
                    f"$ {a['interest']:,.4f}",
                ], shade_alt=True)
            set_col_widths(detail_tbl, [2.5, 2.0, 3.0, 2.8, 2.0, 3.5])

        tp = doc.add_paragraph()
        tr3 = tp.add_run(
            f"  → Total Gain: $ {rec['total_gain']:,.4f}  "
            f"  → Final Value: $ {rec['final_value']:,.4f}"
        )
        tr3.bold = True
        tr3.font.size = Pt(10)
        tr3.font.color.rgb = ACCENT_GRN
        doc.add_paragraph()

    heading(doc, "5.3  Path B Summary", level=2, color=MID_BLUE)

    sum_b_tbl = doc.add_table(rows=2, cols=4)
    sum_b_tbl.style = "Table Grid"
    add_header_row(sum_b_tbl,
        ["Total Capital (USD)", "Total Interest (USD)", "Total Value (USD)", "ROI (%)"])
    add_data_row(sum_b_tbl, 1, [
        f"$ {summary_b['total_principal_usd']:,.2f}",
        f"$ {summary_b['total_interest_usd']:,.2f}",
        f"$ {summary_b['total_value_usd']:,.2f}",
        f"{summary_b['roi_pct']:.2f} %",
    ], shade_alt=True)
    set_col_widths(sum_b_tbl, [4.0, 4.0, 4.0, 2.5])

    add_page_break(doc)

    # ══════════════════════════════════════════════════════════════════════
    # 6. PATH C – MIXED STRATEGY
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "6. Path C – Mixed Strategy (50 % Bank + 50 % Exchange)", level=1, color=DARK_BLUE)

    body(doc,
        "Salary 1 (Feb 23) remains fully in OKX (active investment, unchanged). "
        "For each subsequent salary, USD 250 is converted to LKR 75,000 and placed "
        "in the best-tenor local bank FD, while USD 250 remains as USDT and is "
        "invested across MEXC (150 USDT @ 15 % APR) and WEEX (100 USDT @ 13 % APR).")

    for rec in path_c_rows:
        subhead = doc.add_paragraph()
        sr3 = subhead.add_run(
            f"Salary {rec['salary_no']}  ·  {rec['salary_date'].strftime('%d %b %Y')}  "
            f"·  USD {rec['usd_principal']:,.2f}"
        )
        sr3.bold = True
        sr3.font.size = Pt(10)
        sr3.font.color.rgb = MID_BLUE

        if rec["salary_no"] == 1:
            body(doc, f"  Fully in OKX (active). Gain: $ {rec['total_gain_usd']:,.2f}  "
                      f"Final: $ {rec['total_final_usd']:,.2f}")
        else:
            # bank side
            bp = doc.add_paragraph()
            bpr = bp.add_run(f"  Bank Side (LKR {rec['bank_principal_lkr']:,.0f}):")
            bpr.bold = True
            bpr.font.size = Pt(9)

            n = len(rec["bank_strategy"]) + 1
            bt = doc.add_table(rows=n, cols=5)
            bt.style = "Table Grid"
            add_header_row(bt,
                ["Bank", "Type", "Tenor", "Rate", "Interest (LKR)"])
            for ri, s in enumerate(rec["bank_strategy"], start=1):
                add_data_row(bt, ri, [
                    s["bank"], s["type"], s["tenor"],
                    f"{s['rate']*100:.2f} %",
                    f"LKR {s['interest_lkr']:,.2f}",
                ], shade_alt=True)
            set_col_widths(bt, [2.8, 2.2, 2.0, 2.5, 4.0])

            # exchange side
            ep = doc.add_paragraph()
            epr = ep.add_run("  Exchange Side (USD 250):")
            epr.bold = True
            epr.font.size = Pt(9)

            ne = len(rec["exchange_split"]) + 1
            et = doc.add_table(rows=ne, cols=5)
            et.style = "Table Grid"
            add_header_row(et,
                ["Exchange", "Asset", "Principal", "Rate", "Interest (USDT)"])
            for ri, a in enumerate(rec["exchange_split"], start=1):
                add_data_row(et, ri, [
                    a["exchange"], a["asset"],
                    f"$ {a['principal']:,.2f}",
                    f"{a['rate']*100:.2f} %",
                    f"$ {a['interest']:,.4f}",
                ], shade_alt=True)
            set_col_widths(et, [2.8, 2.0, 3.0, 2.5, 3.5])

            tp2 = doc.add_paragraph()
            tr4 = tp2.add_run(
                f"  → Bank Interest: LKR {rec['bank_gain_lkr']:,.2f} ($ {rec['bank_gain_usd']:,.4f})  "
                f"  → Exchange Gain: $ {rec['exchange_gain_usd']:,.4f}  "
                f"  → Total Gain: $ {rec['total_gain_usd']:,.4f}  "
                f"  → Final Value: $ {rec['total_final_usd']:,.4f}"
            )
            tr4.bold = True
            tr4.font.size = Pt(9)
            tr4.font.color.rgb = ACCENT_GRN
        doc.add_paragraph()

    heading(doc, "6.1  Path C Summary", level=2, color=MID_BLUE)

    sum_c_tbl = doc.add_table(rows=2, cols=4)
    sum_c_tbl.style = "Table Grid"
    add_header_row(sum_c_tbl,
        ["Total Capital (USD)", "Total Interest (USD)", "Total Value (USD)", "ROI (%)"])
    add_data_row(sum_c_tbl, 1, [
        f"$ {summary_c['total_principal_usd']:,.2f}",
        f"$ {summary_c['total_interest_usd']:,.2f}",
        f"$ {summary_c['total_value_usd']:,.2f}",
        f"{summary_c['roi_pct']:.2f} %",
    ], shade_alt=True)
    set_col_widths(sum_c_tbl, [4.0, 4.0, 4.0, 2.5])

    add_page_break(doc)

    # ══════════════════════════════════════════════════════════════════════
    # 7. ROI COMPARISON & VERDICT
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "7. ROI Comparison & Verdict", level=1, color=DARK_BLUE)

    comp_tbl = doc.add_table(rows=4, cols=6)
    comp_tbl.style = "Table Grid"
    add_header_row(comp_tbl,
        ["Path", "Total Invested (USD)", "Total Interest (USD)",
         "Final Value (USD)", "ROI (%)", "Compounding?"])

    comp_data = [
        ("A – Local Bank (LKR)",
         f"$ {summary_a['total_principal_lkr']/USD_LKR_RATE:,.2f}",
         f"$ {summary_a['roi_usd']:,.2f}",
         f"$ {summary_a['total_value_lkr']/USD_LKR_RATE:,.2f}",
         f"{summary_a['roi_pct']:.2f} %",
         "No (Simple)"),
        ("B – Crypto Exchange",
         f"$ {summary_b['total_principal_usd']:,.2f}",
         f"$ {summary_b['total_interest_usd']:,.2f}",
         f"$ {summary_b['total_value_usd']:,.2f}",
         f"{summary_b['roi_pct']:.2f} %",
         "Yes (Daily)"),
        ("C – Mixed (50/50)",
         f"$ {summary_c['total_principal_usd']:,.2f}",
         f"$ {summary_c['total_interest_usd']:,.2f}",
         f"$ {summary_c['total_value_usd']:,.2f}",
         f"{summary_c['roi_pct']:.2f} %",
         "Partial"),
    ]

    for row_i, rd in enumerate(comp_data, start=1):
        add_data_row(comp_tbl, row_i, rd, shade_alt=True)
        # Highlight winner
        if row_i == ["A", "B", "C"].index(winner[0][5]) + 1:
            for cell in comp_tbl.rows[row_i].cells:
                set_cell_bg(cell, "C6EFCE")
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                        run.font.color.rgb = ACCENT_GRN

    set_col_widths(comp_tbl, [3.5, 3.5, 3.5, 3.8, 2.5, 3.0])

    doc.add_paragraph()
    heading(doc, "Verdict", level=2, color=MID_BLUE)

    rois = {
        "A": summary_a["roi_pct"],
        "B": summary_b["roi_pct"],
        "C": summary_c["roi_pct"],
    }
    sorted_paths = sorted(rois.items(), key=lambda x: x[1], reverse=True)
    rank1, rank2, rank3 = sorted_paths[0], sorted_paths[1], sorted_paths[2]

    path_names = {"A": "Path A (Local Bank)", "B": "Path B (Crypto Exchange)", "C": "Path C (Mixed)"}

    verdict_text = (
        f"Based on the calculations above:\n\n"
        f"  1st Place: {path_names[rank1[0]]} — ROI {rank1[1]:.2f} %\n"
        f"  2nd Place: {path_names[rank2[0]]} — ROI {rank2[1]:.2f} %\n"
        f"  3rd Place: {path_names[rank3[0]]} — ROI {rank3[1]:.2f} %\n\n"
        f"{path_names[rank1[0]]} generates the highest return over the 1-year period. "
        f"The key advantage is daily compounding on flexible earn products vs. the simple "
        f"interest model used by local banks, plus generally higher gross rates available "
        f"on crypto exchanges for stablecoin deposits. "
        f"However, crypto exchange platforms carry counterparty and regulatory risks that "
        f"local bank FDs do not. Investors with low risk tolerance should consider "
        f"Path A or a conservative blend (Path C)."
    )
    body(doc, verdict_text)

    add_page_break(doc)

    # ══════════════════════════════════════════════════════════════════════
    # 8. RECOMMENDED INVESTMENT TIMELINE
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "8. Recommended Investment Timeline (Apr 2026 – Dec 2026)", level=1, color=DARK_BLUE)

    body(doc,
        "The following timeline shows the recommended action on each salary receipt date "
        "for the highest-ROI path (Path B – Crypto Exchange), while noting which products "
        "to use for the remaining months. Salary 1 is already deployed and no action is needed.")

    tl_tbl = doc.add_table(rows=len(SALARY_DATES) + 1, cols=5)
    tl_tbl.style = "Table Grid"
    add_header_row(tl_tbl,
        ["Salary #", "Date", "Action", "Product / Allocation", "Notes"])

    timeline_actions = [
        (1, date(2026, 2, 23),
         "✅ Already Invested",
         "OKX USDT 10 % APR (180D PROMO) → matures 26 Aug 2026",
         "No action needed. Monitor maturity."),
        (2, date(2026, 4, 23),
         "Invest $500 USDT",
         "300 USDT → MEXC Flexible 15 % APR\n200 USDT → WEEX Flexible 13 % APR",
         "Transfer within 24h of salary receipt."),
        (3, date(2026, 6, 23),
         "Invest $500 USDT",
         "300 USDT → MEXC Flexible 15 % APR\n200 USDT → WEEX Flexible 13 % APR",
         "Maintain existing positions. Add new tranche."),
        (4, date(2026, 7, 23),
         "Invest $500 USDT",
         "300 USDT → MEXC Flexible 15 % APR\n200 USDT → WEEX Flexible 13 % APR",
         "Monthly phase begins."),
        (5, date(2026, 8, 23),
         "Invest $500 USDT + Reinvest OKX (26 Aug)",
         "300 USDT → MEXC Flexible 15 % APR\n200 USDT → WEEX Flexible 13 % APR\n"
         "+ OKX maturity 525.27: 300→MEXC, 200→WEEX, 25.27→MEXC Tier2 6 %",
         "OKX matures 26 Aug. Withdraw and reinvest immediately."),
        (6, date(2026, 9, 23),
         "Invest $500 USDT",
         "300 USDT → MEXC Flexible 15 % APR\n200 USDT → WEEX Flexible 13 % APR",
         "Compound interest accumulating."),
        (7, date(2026, 10, 23),
         "Invest $500 USDT",
         "300 USDT → MEXC Flexible 15 % APR\n200 USDT → WEEX Flexible 13 % APR",
         "Q4 accumulation phase."),
        (8, date(2026, 11, 23),
         "Invest $500 USDT",
         "300 USDT → MEXC Flexible 15 % APR\n200 USDT → WEEX Flexible 13 % APR",
         "Consider reviewing rates; switch if better promo available."),
        (9, date(2026, 12, 23),
         "Invest $500 USDT",
         "300 USDT → MEXC Flexible 15 % APR\n200 USDT → WEEX Flexible 13 % APR",
         "Last salary of period. Mature all on / after 23 Feb 2027."),
    ]

    for ri, (sn, sd, action, alloc, notes) in enumerate(timeline_actions, start=1):
        row = tl_tbl.rows[ri]
        row.cells[0].text = str(sn)
        row.cells[1].text = sd.strftime("%d %b %Y")
        row.cells[2].text = action
        row.cells[3].text = alloc
        row.cells[4].text = notes
        for ci in range(5):
            for para in row.cells[ci].paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in para.runs:
                    run.font.size = Pt(8)
        if ri % 2 == 0:
            for ci in range(5):
                set_cell_bg(row.cells[ci], "EBF3FB")

    set_col_widths(tl_tbl, [1.5, 2.5, 3.0, 5.5, 4.5])
    add_page_break(doc)

    # ══════════════════════════════════════════════════════════════════════
    # 9. RISKS & ASSUMPTIONS
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "9. Key Risks & Assumptions", level=1, color=DARK_BLUE)

    risks = [
        ("USD/LKR Exchange Rate Risk",
         "Calculations use a fixed rate of 300 LKR/USD. Actual rates fluctuate. "
         "LKR depreciation benefits Path A (more LKR per USD); appreciation reduces it."),
        ("Crypto Exchange Counterparty Risk",
         "Exchanges such as MEXC and WEEX are not insured. Platform insolvency, "
         "hacking, or regulatory action can result in partial or total loss of funds."),
        ("Promotional Rate Duration",
         "The OKX 10 % PROMO is confirmed for 180 days only. MEXC and WEEX flexible "
         "rates may change without notice. Always verify current rates before investing."),
        ("NEXO Exclusion",
         "NEXO (11 % USDT) is excluded due to insufficient capital for its minimum. "
         "If capital grows sufficiently, consider adding NEXO for higher blended yield."),
        ("Tax Implications",
         "Interest income may be taxable in Sri Lanka. Consult a local tax adviser "
         "regarding FIRS and Inland Revenue reporting requirements."),
        ("Liquidity",
         "FDs have lock-in periods. Flexible exchange products allow redemption anytime "
         "but rate changes can affect total returns. Path A capital is illiquid."),
        ("Capital Preservation",
         "LKR FD principal is safe within SDBL limits (~LKR 600,000 per bank). "
         "Crypto exchange holdings are not covered by deposit insurance."),
        ("Interest Compounding Frequency",
         "Exchange earn products are modelled as daily-compounding flexible products. "
         "Actual compounding frequency may vary by platform."),
    ]

    for risk_title, risk_desc in risks:
        rp = doc.add_paragraph(style="List Bullet")
        r1 = rp.add_run(f"{risk_title}: ")
        r1.bold = True
        r1.font.color.rgb = ACCENT_RED
        r1.font.size = Pt(10)
        r2 = rp.add_run(risk_desc)
        r2.font.size = Pt(10)

    add_page_break(doc)

    # ══════════════════════════════════════════════════════════════════════
    # 10. APPENDIX – INTEREST RATE REFERENCE
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "10. Appendix – Interest Rate Reference", level=1, color=DARK_BLUE)

    heading(doc, "A. Selected LKR Bank FD Rates (Mar 2026)", level=2, color=MID_BLUE)

    lkr_data = [
        ("BOC",       "LKR FD", "1Y",   "6.75 %"),
        ("NTB",       "LKR FD", "1Y",   "8.00 %"),
        ("COM BANK",  "LKR FD", "1Y",   "8.00 %"),
        ("HNB",       "LKR FD", "1Y",   "8.00 %"),
        ("SAMPATH",   "LKR FD", "400D", "9.00 %"),
        ("SAMPATH",   "LKR FD", "2Y",   "10.00 %"),
        ("DFCC",      "LKR FD", "1Y",   "8.50 %"),
        ("PAN ASIA",  "LKR FD", "1Y",   "9.00 %"),
        ("NDB",       "LKR FD", "1Y",   "8.25 %"),
        ("PEOPLE'S",  "LKR FD", "1Y",   "6.75 %"),
    ]

    a_tbl = doc.add_table(rows=len(lkr_data) + 1, cols=4)
    a_tbl.style = "Table Grid"
    add_header_row(a_tbl, ["Bank", "Type", "Tenor", "Rate (p.a.)"])
    for ri, row in enumerate(lkr_data, start=1):
        add_data_row(a_tbl, ri, list(row), shade_alt=True)
    set_col_widths(a_tbl, [3.5, 3.5, 3.0, 3.0])

    heading(doc, "B. Selected Crypto Exchange Rates (Mar 2026)", level=2, color=MID_BLUE)

    ex_data = [
        ("MEXC",   "USDT", "Flexible", "15 %",   "Tier 1 ≤ 300 USDT"),
        ("WEEX",   "USDT", "Flexible", "13 %",   "Tier 1 ≤ 200 USDT"),
        ("COINEX", "USDT", "Flexible", "14.16 %","Tier 1 ≤ 1,000 USDT"),
        ("BITUNIX","USDT", "Flexible", "11.60 %","Tier 1 ≤ 200 USDT"),
        ("OKX",    "USDT", "Flex PROMO","10 %",  "≤ 500 USDT, 180D"),
        ("HTX",    "USDT", "Flexible", "10.00 %","Tier 1 ≤ 500 USDT"),
        ("BINANCE","USDT", "Flexible", "3.85 %", "Tier 1 ≤ 200 USDT"),
        ("BYBIT",  "USDT", "Flexible", "5.66 %", "Tier 1 ≤ 200 USDT"),
        ("NEXO",   "USDT", "Flexible", "11.00 %","Min capital needed – EXCLUDED"),
    ]

    b_tbl = doc.add_table(rows=len(ex_data) + 1, cols=5)
    b_tbl.style = "Table Grid"
    add_header_row(b_tbl, ["Exchange", "Asset", "Product", "Rate (APR)", "Notes"])
    for ri, row in enumerate(ex_data, start=1):
        add_data_row(b_tbl, ri, list(row), shade_alt=True)
        if "EXCLUDED" in row[4]:
            for cell in b_tbl.rows[ri].cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.color.rgb = ACCENT_RED
    set_col_widths(b_tbl, [2.5, 2.0, 3.0, 2.8, 5.0])

    doc.add_paragraph()
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer_p.add_run(
        "— End of Report —\n"
        "Compiled: March 2026 | Data: RAG-Verified Interest Rate Tracker | "
        "For personal financial planning purposes only. Not financial advice."
    )
    fr.italic = True
    fr.font.size = Pt(9)
    fr.font.color.rgb = MID_BLUE

    return doc


# ═══════════════════════════════════════════════════════════════════════════
#   MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    print("Calculating investment paths...")

    path_a_rows, summary_a = calculate_path_a()
    path_b_rows, summary_b = calculate_path_b()
    path_c_rows, summary_c = calculate_path_c()

    print(f"Path A ROI: {summary_a['roi_pct']:.2f} %  (LKR interest: {summary_a['total_interest_lkr']:,.2f})")
    print(f"Path B ROI: {summary_b['roi_pct']:.2f} %  (USD interest: {summary_b['total_interest_usd']:,.2f})")
    print(f"Path C ROI: {summary_c['roi_pct']:.2f} %  (USD interest: {summary_c['total_interest_usd']:,.2f})")

    print("\nBuilding Word document...")
    doc = build_document(path_a_rows, summary_a, path_b_rows, summary_b,
                         path_c_rows, summary_c)

    output_path = (
        "/home/runner/work/antigravity-awesome-skills/antigravity-awesome-skills"
        "/issues/244/Personal_Finance_ROI_Report_2026.docx"
    )
    doc.save(output_path)
    print(f"\n✅  Report saved to: {output_path}")


if __name__ == "__main__":
    main()
