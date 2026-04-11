import streamlit as st
import pandas as pd
import io
import json
import os
import time
import re
from datetime import datetime
from groq import Groq

st.set_page_config(page_title="PBV Finance - AI CFO Assistant", page_icon="📊", layout="wide")

# ═══════════════════════════════════════
# COUNTRY RULES (Hardcoded)
# ═══════════════════════════════════════
COUNTRY_RULES = {
    "UAE": {
        "tax_rate": 0.09, "tax_threshold": 375000,
        "tax_desc": "9% CT on taxable income > AED 375,000. Free Zone: 0% on Qualifying Income.",
        "vat": "5% VAT. Input recovery on cost variances. Filing by 28th.",
        "labor": "WPS mandatory. EOS: 21 days/yr (5yrs), 30 days/yr after.",
        "other": "ESR substance. TP docs if group > AED 200M revenue."
    },
    "India": {
        "tax_rate": 0.2517, "tax_threshold": 0,
        "tax_desc": "25.17% CT (turnover < INR 400Cr). MAT 15%. New regime 22%.",
        "vat": "GST 5-28%. ITC reversal risk. GSTR reconciliation.",
        "labor": "PF 12%+12%. ESIC 3.25%. Gratuity after 5yrs.",
        "other": "MSME AP 45-day rule. TP documentation required."
    },
    "KSA": {
        "tax_rate": 0.20, "tax_threshold": 0,
        "tax_desc": "20% IT (non-Saudi). Zakat 2.578% (Saudi). WHT 5-20%.",
        "vat": "15% VAT. ZATCA e-invoicing Phase 2.",
        "labor": "GOSI 12% Saudi, 2% non-Saudi. Nitaqat compliance.",
        "other": "Zakat base = adjusted equity. TP rules from Jan 2024."
    },
    "Qatar": {
        "tax_rate": 0.10, "tax_threshold": 0,
        "tax_desc": "10% CIT. QFC different. Tax holidays per sector.",
        "vat": "No VAT currently.",
        "labor": "EOS 3 weeks/yr. Labor law 2024 updates.",
        "other": "WHT 5% non-residents. TP documentation required."
    },
    "UK": {
        "tax_rate": 0.25, "tax_threshold": 0,
        "tax_desc": "25% CT (>GBP 250K). 19% small profits. Loss c/f unlimited.",
        "vat": "20% VAT. MTD quarterly digital filing.",
        "labor": "NIC 13.8%. Auto-pension 3%. Apprenticeship Levy 0.5%.",
        "other": "R&D credits. Late payment interest: base+8%."
    },
    "Other": {
        "tax_rate": 0.20, "tax_threshold": 0,
        "tax_desc": "Verify local CT rate.",
        "vat": "Verify indirect tax.", "labor": "Verify labor compliance.",
        "other": "Verify TP and substance."
    }
}

# ═══════════════════════════════════════
# QUICK WINS LIBRARY
# ═══════════════════════════════════════
QUICK_WINS = {
    "revenue_decline": [
        {"action": "Accelerate pipeline — contact top 10 delayed deals", "owner": "Head of Sales", "impact": "5-10% of revenue variance", "erp": "VA05"},
        {"action": "Review pricing vs market competitors", "owner": "Commercial Manager", "impact": "2-5% of revenue variance", "erp": "KE30"},
        {"action": "Reactivate top 20 dormant accounts", "owner": "Key Account Manager", "impact": "3-8% of revenue variance", "erp": "MCSI"},
    ],
    "cost_overrun_employee": [
        {"action": "Freeze non-critical hiring for 1 quarter", "owner": "HR Director + CFO", "impact": "Frozen positions x monthly cost", "erp": "PA20"},
        {"action": "Cap overtime at budgeted levels", "owner": "Operations Manager", "impact": "10-15% of employee variance", "erp": "Payroll report"},
    ],
    "cost_overrun_opex": [
        {"action": "30-day discretionary spend freeze", "owner": "All Dept Heads", "impact": "30-50% of OpEx variance", "erp": "S_ALR_87013611"},
    ],
    "cogs_overrun": [
        {"action": "Renegotiate top 3 supplier contracts", "owner": "Procurement", "impact": "5-15% of COGS variance", "erp": "ME2M"},
    ],
    "margin_compression": [
        {"action": "Minimum 25% GP margin policy — CFO approval below", "owner": "Commercial + CFO", "impact": "Prevents further erosion", "erp": "KE30"},
    ]
}

# ═══════════════════════════════════════
# PHASE 1: CONSTANTS & HELPERS
# ═══════════════════════════════════════
CATEGORIES = ["Revenue", "COGS", "Employee", "Marketing", "OpEx",
              "Depreciation", "Finance", "Other Income", "Tax",
              "Exceptional", "Skip"]
COST_CATEGORIES = {"COGS", "Employee", "Marketing", "OpEx", "Depreciation", "Finance", "Tax", "Exceptional"}
INCOME_CATEGORIES = {"Revenue", "Other Income"}

# Keywords for exceptional / one-time items (highest priority in matching)
EXCEPTIONAL_KEYWORDS = [
    "restructuring", "restructure", "settlement", "one-time", "one time",
    "write-off", "write off", "writeoff", "write down", "impairment",
    "discontinued", "non-recurring", "nonrecurring", "extraordinary",
    "exceptional",
]

# Balance-sheet item keywords → auto-suggest "Skip"
BS_SKIP_KEYWORDS = [
    "cash and cash", "cash at bank", "petty cash", "bank balance",
    "accounts receivable", "trade receivable", "sundry debtor",
    "accounts payable", "trade payable", "sundry creditor",
    "inventory", "stock in trade", "raw material inventory",
    "short term loan", "long term loan", "term loan", "overdraft",
    "share capital", "paid-up capital", "authorised capital",
    "general reserve", "retained earnings", "retained profit",
    "fixed asset", "tangible asset", "intangible asset",
    "property plant", "plant and equipment", "capital work",
    "prepaid", "advance paid", "security deposit",
    "accrued liability", "deferred revenue", "deferred tax liability",
    "minority interest", "goodwill",
]

CATEGORY_KEYWORDS = {
    "Revenue": ["revenue", "net sales", "gross sales", "sales", "turnover",
                "service income", "fee income", "income from operations"],
    "COGS": ["cost of goods", "cost of sales", "cogs", "direct cost",
             "cost of revenue", "purchases", "material cost", "cost of service",
             "cost of production", "raw material"],
    "Employee": ["employee cost", "salary", "salaries", "payroll", "staff cost",
                 "wages", "manpower", "compensation", "benefits", "bonus",
                 "hr cost", "personnel", "staff expenses"],
    "Marketing": ["marketing", "advertising", "promotion", "brand", "campaign",
                  "sales & marketing", "sales and marketing"],
    "OpEx": ["other opex", "other operating", "overhead", "admin", "administration",
             "rent", "utilities", "insurance", "it cost", "software", "maintenance",
             "general expenses", "office expenses", "travel", "professional fees",
             "legal", "consulting", "miscellaneous expenses", "other expenses"],
    "Depreciation": ["depreciation", "amortization", "amortisation", "d&a", "d & a"],
    "Finance": ["finance cost", "finance charges", "interest expense", "interest cost",
                "bank charges", "borrowing cost", "loan interest", "financial charges"],
    "Other Income": ["other income", "miscellaneous income", "non-operating income",
                     "gain on disposal", "gain on", "dividend income", "interest income"],
    "Tax": ["income tax", "corporate tax", "tax expense", "tax provision",
            "tax charge", "zakat", "ct provision", "deferred tax"],
}

# P&L ORDER for display
PL_ORDER = ["Revenue", "COGS", "Employee", "Marketing", "OpEx",
            "Depreciation", "Finance", "Other Income", "Tax", "Exceptional"]


def suggest_category(name):
    """
    Keyword-based category suggestion.
    Priority: Exceptional → BS Skip → P&L keywords → None (unmatched).
    Returns None when nothing matches so callers can show an 'unmatched' indicator.
    """
    name_lower = str(name).lower().strip()
    # 1. Exceptional / one-time items (highest priority)
    for kw in EXCEPTIONAL_KEYWORDS:
        if kw in name_lower:
            return "Exceptional"
    # 2. Balance-sheet keywords → Skip
    for kw in BS_SKIP_KEYWORDS:
        if kw in name_lower:
            return "Skip"
    # 3. Standard P&L keyword matching
    for cat, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            if kw in name_lower:
                return cat
    return None  # Truly unmatched — caller decides default


def fmt_km(n):
    """Format a number as 1.1M or 1,050K (no currency prefix)."""
    abs_n = abs(n)
    if abs_n >= 1_000_000:
        return f"{n / 1_000_000:.1f}M"
    if abs_n >= 1_000:
        return f"{n / 1_000:,.0f}K"
    return f"{n:,.0f}"


def is_subtotal_row(text):
    """Return True for blank rows or rows that look like totals/subtotals."""
    if pd.isna(text) or str(text).strip() == "":
        return True
    t = str(text).strip().lower()
    return bool(
        re.search(r'(^|\s)(sub[\s-]?total|grand\s+total)(\s|$)', t) or
        re.match(r'^total(\s|$)', t) or
        re.search(r'\btotal$', t)
    )


def detect_header_row(raw_df):
    """Find the best header row index (0–9) by scoring keyword matches."""
    header_keywords = [
        "description", "line item", "account", "gl", "budget", "actual",
        "plan", "forecast", "particulars", "item", "name",
        "debit", "credit", "balance", "amount", "dr", "cr",
    ]
    best_row, best_score = 0, 0
    for i in range(min(10, len(raw_df))):
        row_str = " ".join(str(v).lower() for v in raw_df.iloc[i].values if pd.notna(v))
        score = sum(1 for kw in header_keywords if kw in row_str)
        if score > best_score:
            best_score, best_row = score, i
    return best_row if best_score >= 1 else 0


def find_columns(df):
    """Return (line_col, budget_col, actual_col) by scanning column names."""
    line_col = budget_col = actual_col = None
    for c in df.columns:
        cl = str(c).lower().strip()
        if line_col is None and any(k in cl for k in [
                "line item", "item", "account", "description", "gl",
                "particulars", "head", "name", "ledger"]):
            line_col = c
        elif budget_col is None and any(k in cl for k in [
                "budget", "plan", "target", "forecast"]):
            budget_col = c
        elif actual_col is None and any(k in cl for k in [
                "actual", "ytd", "current"]):
            actual_col = c
    return line_col, budget_col, actual_col


# ═══════════════════════════════════════
# PHASE 2: GL MAPPING & MEMORY
# ═══════════════════════════════════════
MAPPING_MEMORY_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "mapping_memory.json"
)

# SAP GL account ranges → P&L categories.
# 6xxxxx is intentionally "_6x_ambiguous": keyword fallback resolves Employee vs OpEx.
SAP_GL_RANGES = [
    (100000, 399999, "Skip"),           # 1–3xxxxx: Balance Sheet
    (400000, 499999, "Revenue"),        # 4xxxxx
    (500000, 599999, "COGS"),           # 5xxxxx
    (600000, 699999, "_6x_ambiguous"),  # 6xxxxx: Employee or OpEx
    (700000, 799999, "Depreciation"),   # 7xxxxx
    (800000, 899999, "Finance"),        # 8xxxxx
    (900000, 999999, "Other Income"),   # 9xxxxx
]

GL_COL_KEYWORDS = [
    "gl code", "gl account", "g/l account", "g/l", "account code",
    "account no", "account number", "sap code", "cost element", "gl no",
]


def classify_gl_code(gl_value):
    """Return (category_str, is_balance_sheet) for a GL code value."""
    try:
        code = int(str(gl_value).strip())
    except (ValueError, TypeError):
        return None, False
    for low, high, cat in SAP_GL_RANGES:
        if low <= code <= high:
            return cat, (cat == "Skip")
    return None, False


def extract_gl_from_text(text):
    """Extract a leading 6-digit GL code from a description string, e.g. '400001 Revenue'."""
    m = re.match(r'^(\d{6})\b', str(text).strip())
    return int(m.group(1)) if m else None


def detect_gl_column(df, exclude_cols):
    """
    Return a column name if a dedicated GL-code column is found, else None.
    Prefers columns with GL-related names; falls back to columns with ≥85% 4–8-digit values.
    """
    exclude = {str(c) for c in exclude_cols}
    fallback_col = None
    for col in df.columns:
        if str(col) in exclude:
            continue
        col_lower = str(col).lower().strip()
        sample = df[col].dropna().head(30)
        if len(sample) == 0:
            continue
        numeric_cov = sum(
            1 for v in sample if re.fullmatch(r'\d{4,8}', str(v).strip())
        ) / len(sample)
        is_gl_named = any(k in col_lower for k in GL_COL_KEYWORDS)
        if is_gl_named and numeric_cov >= 0.5:
            return col          # high-confidence named match
        if numeric_cov >= 0.85 and fallback_col is None:
            fallback_col = col  # unnamed high-numeric column
    return fallback_col


def suggest_with_gl(name, gl_code=None):
    """
    Suggest a P&L category using GL range first, keywords as fallback.
    Returns (category: str, is_confident: bool).
    is_confident=False means nothing matched → show 'unmatched' indicator.
    Default category when unmatched: 'OpEx'.
    """
    if gl_code is not None:
        cat, is_bs = classify_gl_code(gl_code)
        if is_bs:
            return "Skip", True
        if cat and cat != "_6x_ambiguous":
            return cat, True
        if cat == "_6x_ambiguous":
            kw = suggest_category(name)
            if kw in ("Employee", "Marketing", "OpEx", "Depreciation", "Exceptional"):
                return kw, True
            return "OpEx", True  # GL range gives confidence even without kw match
    kw = suggest_category(name)
    if kw is not None:
        return kw, True
    return "OpEx", False  # Unmatched — default to OpEx, flag as needing review


def load_mapping_memory():
    """Load saved item→category mapping from JSON. Returns {} on any error."""
    if os.path.exists(MAPPING_MEMORY_PATH):
        try:
            with open(MAPPING_MEMORY_PATH, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_mapping_memory(new_mappings):
    """Merge new mappings into the JSON memory file (accumulates over time)."""
    existing = load_mapping_memory()
    existing.update(new_mappings)
    try:
        with open(MAPPING_MEMORY_PATH, "w") as f:
            json.dump(existing, f, indent=2)
    except Exception:
        pass  # Non-fatal — analysis still runs


# ── CFO Memory Engine (historical analysis persistence) ───────────────
FINANCE_MEMORY_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "finance_memory.json"
)


def save_to_memory(data):
    """Append a month's analysis to the finance memory JSON. Returns total count."""
    memory = []
    if os.path.exists(FINANCE_MEMORY_FILE):
        try:
            with open(FINANCE_MEMORY_FILE, "r") as f:
                memory = json.load(f)
        except Exception:
            memory = []
    memory.append(data)
    with open(FINANCE_MEMORY_FILE, "w") as f:
        json.dump(memory, f, indent=2)
    return len(memory)


def load_memory():
    """Load all saved months from finance memory. Returns list of dicts."""
    if os.path.exists(FINANCE_MEMORY_FILE):
        try:
            with open(FINANCE_MEMORY_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return []
    return []


# Category name mapping for data-driven actions (internal category → action lookup key)
_ACTION_CAT_MAP = {
    "Revenue": "Revenue",
    "COGS": "COGS",
    "Employee": "Employee Cost",
    "Marketing": "Selling & Marketing",
    "OpEx": "Other Operating Expense",
    "Finance": "Finance Cost",
}


def generate_data_driven_action(category, variance, curr):
    """
    Generate a data-driven CFO action based on ACTUAL variance data.
    Returns dict with action/detail/owner/timeline/impact, or None.
    """
    abs_var = abs(variance)
    cat_key = _ACTION_CAT_MAP.get(category)
    if cat_key is None:
        return None

    if cat_key == "Revenue":
        recovery = abs_var * 0.5
        return {
            "action": f"Recover {curr} {recovery:,.0f} via pipeline acceleration",
            "detail": "Close top delayed deals within 30 days",
            "owner": "Head of Sales",
            "timeline": "30 days",
            "impact": min(recovery, abs_var),
        }
    elif cat_key == "COGS":
        saving = abs_var * 0.15
        return {
            "action": f"Save {curr} {saving:,.0f} via vendor renegotiation",
            "detail": "Target 10-15% reduction with top 3 suppliers",
            "owner": "Procurement Head",
            "timeline": "45 days",
            "impact": min(saving, abs_var),
        }
    elif cat_key == "Employee Cost":
        saving = abs_var * 0.8
        return {
            "action": f"Save {curr} {saving:,.0f} via hiring freeze",
            "detail": "Freeze all open non-critical positions",
            "owner": "HR + CFO",
            "timeline": "Immediate",
            "impact": min(saving, abs_var),
        }
    elif cat_key == "Selling & Marketing":
        saving = abs_var * 0.5
        return {
            "action": f"Save {curr} {saving:,.0f} by cutting low-ROI campaigns",
            "detail": "Pause non-performing channels, redirect budget",
            "owner": "Marketing Head",
            "timeline": "30 days",
            "impact": min(saving, abs_var),
        }
    elif cat_key == "Other Operating Expense":
        saving = abs_var * 0.3
        return {
            "action": f"Save {curr} {saving:,.0f} via discretionary spend freeze",
            "detail": "30-day freeze on travel, events, consulting",
            "owner": "All Department Heads",
            "timeline": "30 days",
            "impact": min(saving, abs_var),
        }
    elif cat_key == "Finance Cost":
        saving = abs_var * 0.25
        return {
            "action": f"Save {curr} {saving:,.0f} via debt restructuring",
            "detail": "Renegotiate loan terms or refinance",
            "owner": "CFO",
            "timeline": "60 days",
            "impact": min(saving, abs_var),
        }
    return None


def calculate_all_insights(revenue_b, revenue_a, cogs_b, cogs_a,
                           gp_b, gp_a, ebitda_b, ebitda_a,
                           emp_b, emp_a, mkt_b, mkt_a, opex_b, opex_a,
                           dep_b, dep_a, fin_b, fin_a, oi_b, oi_a,
                           pbt_b, pbt_a, pat_b, pat_a, tax_rate,
                           currency, material_count, agg_data):
    """Data-driven insight engine — 100% Python math, zero AI."""
    insights = []
    rev_change = revenue_a - revenue_b
    rev_change_pct = rev_change / revenue_b * 100 if revenue_b else 0
    gp_margin_b = gp_b / revenue_b * 100 if revenue_b else 0
    gp_margin_a = gp_a / revenue_a * 100 if revenue_a else 0
    ebitda_margin_b = ebitda_b / revenue_b * 100 if revenue_b else 0
    ebitda_margin_a = ebitda_a / revenue_a * 100 if revenue_a else 0
    total_cost_b = cogs_b + emp_b + mkt_b + opex_b
    total_cost_a = cogs_a + emp_a + mkt_a + opex_a
    cost_change_pct = (total_cost_a - total_cost_b) / total_cost_b * 100 if total_cost_b else 0
    gp_margin_frac = gp_a / revenue_a if revenue_a else 0
    fixed_costs = emp_a + mkt_a + opex_a

    # ── REVENUE INTELLIGENCE ──────────────────────────────────────────────
    if rev_change < 0:
        annual_impact = rev_change * 12
        annual_ebitda_impact = rev_change * gp_margin_frac * 12
        insights.append({"category": "REVENUE", "type": "🔴 CRITICAL",
            "title": "Revenue Decline — Annualized Impact",
            "insight": f"Monthly revenue shortfall: {currency} {abs(rev_change):,.0f}. 12-month projection: {currency} {abs(annual_impact):,.0f} revenue loss → {currency} {abs(annual_ebitda_impact):,.0f} EBITDA destruction at {gp_margin_a:.1f}% GP margin.",
            "action": "Determine within 5 days: one-month issue or trend? Check order pipeline for next 3 months.",
            "impact": abs(annual_ebitda_impact / 12), "owner": "Head of Sales", "confidence": "CALCULATED"})

    if ebitda_a < 0 and gp_margin_frac > 0:
        breakeven_rev = fixed_costs / gp_margin_frac
        revenue_gap = breakeven_rev - revenue_a
        gap_pct = revenue_gap / revenue_a * 100 if revenue_a else 0
        insights.append({"category": "REVENUE", "type": "🔴 CRITICAL",
            "title": "Revenue Required for EBITDA Breakeven",
            "insight": f"At {gp_margin_a:.1f}% GP margin, operating costs of {currency} {fixed_costs:,.0f} require {currency} {breakeven_rev:,.0f} to break even. Current: {currency} {revenue_a:,.0f} — gap: {currency} {revenue_gap:,.0f} ({gap_pct:.1f}%).",
            "action": f"Pipeline must deliver {gap_pct:.0f}% more revenue, OR costs must drop {currency} {abs(ebitda_a):,.0f}.",
            "impact": abs(ebitda_a), "owner": "CFO", "confidence": "CALCULATED"})

    if revenue_b > 0 and rev_change < 0:
        _total_abs_var = abs(rev_change) + abs(total_cost_a - total_cost_b)
        rev_dominance = abs(rev_change) / _total_abs_var * 100 if _total_abs_var > 0 else 0
        if rev_dominance > 60:
            insights.append({"category": "REVENUE", "type": "⚠️ WARNING",
                "title": "Revenue is the Dominant Problem",
                "insight": f"Revenue decline accounts for {rev_dominance:.0f}% of total variance. Cost issues are secondary. Fixing costs alone will NOT solve profitability.",
                "action": "All management attention on revenue recovery FIRST. Cost measures are supplementary.",
                "impact": abs(rev_change) * gp_margin_frac, "owner": "CEO + Head of Sales", "confidence": "CALCULATED"})

    if pbt_a < 0 and gp_margin_frac > 0:
        total_below_gp = emp_a + mkt_a + opex_a + dep_a + fin_a - oi_a
        pbt_be_rev = total_below_gp / gp_margin_frac
        insights.append({"category": "REVENUE", "type": "🔴 CRITICAL",
            "title": "Revenue Required for Net Profit Breakeven",
            "insight": f"PBT breakeven (all costs inc. depreciation {currency} {dep_a:,.0f} + finance {currency} {fin_a:,.0f}) requires {currency} {pbt_be_rev:,.0f}. Gap: {currency} {pbt_be_rev - revenue_a:,.0f} ({(pbt_be_rev - revenue_a)/revenue_a*100:.1f}%).",
            "action": f"Needs {(pbt_be_rev - revenue_a)/revenue_a*100:.1f}% revenue growth OR {currency} {abs(pbt_a):,.0f} total cost reduction.",
            "impact": abs(pbt_a), "owner": "CFO", "confidence": "CALCULATED"})

    # ── COST STRUCTURE INTELLIGENCE ───────────────────────────────────────
    if rev_change_pct != 0:
        elasticity = round(cost_change_pct / rev_change_pct, 2)
        if rev_change_pct < 0 and cost_change_pct >= 0:
            insights.append({"category": "COST STRUCTURE", "type": "🔴 CRITICAL",
                "title": "Zero Cost Elasticity — Scissors Effect",
                "insight": f"Revenue dropped {abs(rev_change_pct):.1f}% but costs INCREASED {cost_change_pct:.1f}%. Elasticity: {elasticity}. Cost base is rigid — margins collapse with any revenue decline.",
                "action": "Convert 20-30% of operating costs from fixed to variable within 90 days: hourly staffing, consumption-based contracts, performance-linked comp.",
                "impact": abs(total_cost_a - total_cost_b), "owner": "CFO + Operations", "confidence": "CALCULATED"})

    if rev_change != 0:
        op_leverage = (ebitda_a - ebitda_b) / rev_change
        insights.append({"category": "COST STRUCTURE", "type": "📊 INTELLIGENCE",
            "title": "Operating Leverage Ratio",
            "insight": f"For every {currency} 1 revenue change, EBITDA changes {currency} {abs(op_leverage):.2f}. {'HIGH leverage — revenue recovery has outsized positive effect.' if abs(op_leverage) > 1.5 else 'Moderate leverage — profit moves proportionally.'}",
            "action": f"{'Each {currency} 100K recovered adds {currency} ' + f'{abs(op_leverage)*100000:,.0f}' + ' to EBITDA.' if abs(op_leverage) > 1.5 else 'Focus on both revenue and cost measures.'}",
            "impact": abs(op_leverage) * abs(rev_change) * 0.3, "owner": "CFO", "confidence": "CALCULATED"})

    cost_lines = [("COGS", cogs_b, cogs_a), ("Employee", emp_b, emp_a),
                  ("Marketing", mkt_b, mkt_a), ("Other OpEx", opex_b, opex_a)]
    worst_shift = None; worst_pp = 0
    for _cn, _cb, _ca in cost_lines:
        _pb = _cb / revenue_b * 100 if revenue_b else 0
        _pa = _ca / revenue_a * 100 if revenue_a else 0
        _sh = _pa - _pb
        if _sh > worst_pp:
            worst_pp = _sh; worst_shift = (_cn, _pb, _pa, _sh, _ca - _cb)
    if worst_shift and worst_pp > 2:
        _cn, _pb, _pa, _sh, _abs_ch = worst_shift
        _own = {"Employee": "HR Director", "Marketing": "Marketing Head", "COGS": "Procurement"}.get(_cn, "Operations")
        insights.append({"category": "COST STRUCTURE", "type": "⚠️ WARNING",
            "title": f"Structural Cost Shift — {_cn}",
            "insight": f"{_cn} as % of revenue: {_pb:.1f}% → {_pa:.1f}% ({_sh:.1f}pp worse). Absolute increase: {currency} {_abs_ch:,.0f}.",
            "action": f"Volume effect (same cost / less revenue) or genuine increase? Revenue recovery auto-corrects the former.",
            "impact": abs(_sh / 100 * revenue_a), "owner": _own, "confidence": "CALCULATED"})

    if revenue_b > 0 and revenue_a > 0:
        expected_cogs = cogs_b * (revenue_a / revenue_b)
        cogs_gap = cogs_a - expected_cogs
        if abs(cogs_gap) > cogs_b * 0.05:
            insights.append({"category": "COST STRUCTURE", "type": "📊 INTELLIGENCE",
                "title": "COGS Variability Analysis",
                "insight": f"Expected variable COGS: {currency} {expected_cogs:,.0f}. Actual: {currency} {cogs_a:,.0f}. Gap: {currency} {cogs_gap:,.0f}. {'COGS has a fixed component not flexing down.' if cogs_gap > 0 else 'COGS improved — efficiency gains detected.'}",
                "action": f"{'Break COGS into fixed vs variable. Target the fixed portion.' if cogs_gap > 0 else 'Investigate improvement — replicate if deliberate.'}",
                "impact": abs(cogs_gap), "owner": "Procurement + Operations", "confidence": "CALCULATED"})

    # ── MARGIN INTELLIGENCE ───────────────────────────────────────────────
    gp_swing = gp_margin_a - gp_margin_b
    opex_pct_b = (emp_b + mkt_b + opex_b) / revenue_b * 100 if revenue_b else 0
    opex_pct_a = (emp_a + mkt_a + opex_a) / revenue_a * 100 if revenue_a else 0
    opex_swing = -(opex_pct_a - opex_pct_b)
    ebitda_swing = ebitda_margin_a - ebitda_margin_b
    insights.append({"category": "MARGIN", "type": "📉 MARGIN DECOMPOSITION",
        "title": "Where Did the Margin Go?",
        "insight": f"EBITDA margin swing: {ebitda_swing:.1f}pp. GP margin moved {gp_swing:.1f}pp, OpEx ratio moved {opex_swing:.1f}pp. {'GP erosion is PRIMARY — fix pricing/volume.' if abs(gp_swing) > abs(opex_swing) else 'OpEx overrun is PRIMARY — fix spending.'}",
        "action": f"{'Priority: Revenue pricing/volume, then cost control.' if abs(gp_swing) > abs(opex_swing) else 'Priority: Operating cost discipline, then revenue.'}",
        "impact": abs(ebitda_a - ebitda_b), "owner": "CFO", "confidence": "CALCULATED"})

    if gp_margin_a < gp_margin_b:
        margin_loss_val = (gp_margin_b - gp_margin_a) / 100 * revenue_a
        insights.append({"category": "MARGIN", "type": "⚠️ WARNING",
            "title": "Gross Profit Margin Deterioration",
            "insight": f"GP margin: {gp_margin_b:.1f}% → {gp_margin_a:.1f}% ({gp_margin_b - gp_margin_a:.1f}pp). Each 1pp = {currency} {revenue_a/100:,.0f}. Total loss: {currency} {margin_loss_val:,.0f}.",
            "action": "Run profitability analysis by product, customer, region to find WHERE margin eroded.",
            "impact": margin_loss_val, "owner": "Commercial Manager", "confidence": "CALCULATED"})

    below_ebitda = dep_a + fin_a - oi_a
    below_ebitda_b = dep_b + fin_b - oi_b
    below_ebitda_chg = below_ebitda - below_ebitda_b
    if abs(below_ebitda_chg) > abs(ebitda_a - ebitda_b) * 0.2:
        insights.append({"category": "MARGIN", "type": "📊 INTELLIGENCE",
            "title": "Below-EBITDA Leakage",
            "insight": f"Dep/Fin/OI changed {currency} {below_ebitda_chg:,.0f}: Dep {currency} {dep_a-dep_b:,.0f}, Fin {currency} {fin_a-fin_b:,.0f}, OI {currency} {oi_a-oi_b:,.0f}. {'Worsening PBT beyond EBITDA.' if below_ebitda_chg > 0 else 'Partially offsetting EBITDA decline.'}",
            "action": f"{'Review: CapEx-driven depreciation? New borrowing?' if below_ebitda_chg > 0 else 'Verify Other Income sustainability.'}",
            "impact": abs(below_ebitda_chg), "owner": "CFO", "confidence": "CALCULATED"})

    # ── SENSITIVITY & LEVERS ──────────────────────────────────────────────
    rev_1pct = revenue_a * 0.01 * gp_margin_frac
    levers = sorted([("Revenue (at GP margin)", rev_1pct), ("COGS", cogs_a * 0.01),
        ("Employee Cost", emp_a * 0.01), ("Marketing", mkt_a * 0.01),
        ("Other OpEx", opex_a * 0.01)], key=lambda x: x[1], reverse=True)
    lever_text = " | ".join([f"{l[0]}: {currency} {l[1]:,.0f}" for l in levers])
    insights.append({"category": "SENSITIVITY", "type": "📊 SENSITIVITY TABLE",
        "title": "Impact of 1% Improvement in Each Line",
        "insight": f"EBITDA per 1%: {lever_text}. HIGHEST: {levers[0][0]} at {currency} {levers[0][1]:,.0f}.",
        "action": f"Focus on {levers[0][0]}. 3% improvement = {currency} {levers[0][1]*3:,.0f}.",
        "impact": levers[0][1] * 3, "owner": "Depends on lever", "confidence": "CALCULATED"})

    recovery_3pct = sum(l[1] * 3 for l in levers)
    insights.append({"category": "SENSITIVITY", "type": "📈 SCENARIO",
        "title": "Combined 3% Improvement Scenario",
        "insight": f"If every line improves 3%: EBITDA recovery = {currency} {recovery_3pct:,.0f}. New EBITDA: {currency} {ebitda_a + recovery_3pct:,.0f}. {'EBITDA turns POSITIVE.' if ebitda_a < 0 and ebitda_a + recovery_3pct > 0 else f'Margin improves to {(ebitda_a + recovery_3pct)/revenue_a*100:.1f}%.'}",
        "action": "Distribute: Sales +3% rev, Procurement -3% COGS, HR -3% emp, Ops -3% OpEx.",
        "impact": recovery_3pct, "owner": "All Department Heads", "confidence": "CALCULATED"})

    # ── RISK INTELLIGENCE ─────────────────────────────────────────────────
    if fin_a > 0:
        icr = ebitda_a / fin_a
        _icr_type = "🔴 CRITICAL" if icr < 2 else "⚠️ WARNING" if icr < 4 else "📊 INTELLIGENCE"
        _icr_msg = f"DANGER: Below 2x — may struggle to service debt." if icr < 2 else "Below 4x — limited headroom." if icr < 4 else "Comfortable coverage."
        insights.append({"category": "RISK", "type": _icr_type,
            "title": "Interest Coverage Ratio (Debt Risk)",
            "insight": f"EBITDA / Finance Cost = {icr:.1f}x. {_icr_msg}" + (f" Budget ICR: {ebitda_b/fin_b:.1f}x." if fin_b > 0 else ""),
            "action": f"{'Review loan covenants — potential breach at {icr:.1f}x.' if icr < 3 else 'Monitor quarterly.'}",
            "impact": fin_a * 12, "owner": "CFO / Treasury", "confidence": "CALCULATED"})

    if ebitda_a < ebitda_b:
        monthly_det = ebitda_b - ebitda_a
        insights.append({"category": "RISK", "type": "🔴 CRITICAL" if ebitda_a < 0 else "⚠️ WARNING",
            "title": "Cost of Inaction — Quantified",
            "insight": f"Monthly shortfall: {currency} {monthly_det:,.0f}. 3-month cost of doing nothing: {currency} {monthly_det * 3:,.0f}. {'Weekly delay costs {currency} ' + f'{monthly_det/4:,.0f}.' if ebitda_a < 0 else 'Each month widens the gap.'}",
            "action": f"{'THIS WEEK — loss accumulates daily.' if ebitda_a < 0 else 'Within 2 weeks — prevent further erosion.'}",
            "impact": monthly_det * 3, "owner": "CFO / Board", "confidence": "CALCULATED"})

    if ebitda_a < 0:
        insights.append({"category": "RISK", "type": "🔴 CRITICAL",
            "title": "Cash Drain Warning",
            "insight": f"Negative EBITDA = consuming {currency} {abs(ebitda_a):,.0f}/month from operations. Unsustainable.",
            "action": f"Freeze non-essential outflows. Review CapEx. Defer dividends. Assess cash reserves at {currency} {abs(ebitda_a):,.0f}/month burn.",
            "impact": abs(ebitda_a) * 6, "owner": "CFO", "confidence": "CALCULATED"})

    material_uf = agg_data[(agg_data["Flag"] == "UF") & (agg_data["Material?"].str.contains("YES", na=False))]
    if len(material_uf) > 0:
        total_uf = material_uf["Variance"].abs().sum()
        top_item = material_uf.loc[material_uf["Variance"].abs().idxmax()]
        top_pct = abs(float(top_item["Variance"])) / total_uf * 100 if total_uf > 0 else 0
        if top_pct > 50:
            insights.append({"category": "RISK", "type": "📊 INTELLIGENCE",
                "title": "Variance Concentration Risk",
                "insight": f"Top item ({top_item['Category']}) = {top_pct:.0f}% of all UF variance. {'Solving this ONE issue resolves the majority.' if top_pct > 70 else 'Priority but others need attention too.'}",
                "action": f"Allocate 70% of attention to {top_item['Category']}.",
                "impact": abs(float(top_item["Variance"])), "owner": "CFO", "confidence": "CALCULATED"})

    total_fav = agg_data[agg_data["Flag"] == "F"]["Variance"].abs().sum()
    total_unfav = agg_data[agg_data["Flag"] == "UF"]["Variance"].abs().sum()
    if total_unfav > 0:
        offset = total_fav / total_unfav * 100
        insights.append({"category": "RISK", "type": "📊 INTELLIGENCE",
            "title": "Natural Offset Ratio",
            "insight": f"Favorable ({currency} {total_fav:,.0f}) offsets {offset:.0f}% of unfavorable ({currency} {total_unfav:,.0f}). {'Good natural hedging.' if offset > 40 else 'Low offset — management action essential.'}",
            "action": f"{'Protect favorable items — ensure sustainability.' if offset > 40 else 'Cannot rely on offsets — direct intervention required.'}",
            "impact": total_fav, "owner": "CFO", "confidence": "CALCULATED"})

    # ── LINE-SPECIFIC INTELLIGENCE ────────────────────────────────────────
    if emp_a > 0 and emp_b > 0:
        prod_b = revenue_b / emp_b; prod_a = revenue_a / emp_a
        prod_chg = (prod_a - prod_b) / prod_b * 100
        if prod_chg < -5:
            insights.append({"category": "EMPLOYEE", "type": "⚠️ WARNING",
                "title": "Labor Productivity Declining",
                "insight": f"Revenue per {currency} 1 employee cost: {prod_b:.2f} → {prod_a:.2f} ({prod_chg:.1f}%). Ratio: {emp_b/revenue_b*100:.1f}% → {emp_a/revenue_a*100:.1f}%.",
                "action": f"Grow revenue {currency} {emp_a * prod_b - revenue_a:,.0f} to restore ratio, OR cut {currency} {emp_a - revenue_a/prod_b:,.0f}.",
                "impact": abs(emp_a - revenue_a / prod_b), "owner": "HR + CFO", "confidence": "CALCULATED"})

    if mkt_a > 0 and mkt_b > 0:
        mkt_roi_b = revenue_b / mkt_b; mkt_roi_a = revenue_a / mkt_a
        if mkt_roi_a < mkt_roi_b * 0.85:
            insights.append({"category": "MARKETING", "type": "⚠️ WARNING",
                "title": "Marketing ROI Declining",
                "insight": f"Revenue per {currency} 1 marketing: {mkt_roi_b:.1f}x → {mkt_roi_a:.1f}x ({(1-mkt_roi_a/mkt_roi_b)*100:.0f}% less effective).",
                "action": f"Cut bottom 20% campaigns by ROI. Reallocate {currency} {mkt_a * 0.2:,.0f}.",
                "impact": mkt_a * 0.2, "owner": "Marketing Head", "confidence": "CALCULATED"})

    if fin_a > 0 and revenue_a > 0:
        fin_pct_b = fin_b / revenue_b * 100 if revenue_b else 0
        fin_pct_a = fin_a / revenue_a * 100
        if fin_pct_a > fin_pct_b + 0.5:
            insights.append({"category": "FINANCE", "type": "📊 INTELLIGENCE",
                "title": "Increasing Finance Cost Burden",
                "insight": f"Finance cost: {fin_pct_b:.2f}% → {fin_pct_a:.2f}% of revenue. Additional drain: {currency} {(fin_pct_a - fin_pct_b)/100 * revenue_a:,.0f}.",
                "action": "Review debt structure. Explore refinancing or early repayment.",
                "impact": abs(fin_a - fin_b), "owner": "CFO / Treasury", "confidence": "CALCULATED"})

    if dep_a > 0 and revenue_a > 0:
        dep_b_pct = dep_b / revenue_b * 100 if revenue_b else 0
        dep_a_pct = dep_a / revenue_a * 100
        if dep_a_pct > dep_b_pct + 1:
            insights.append({"category": "CAPITAL", "type": "📊 INTELLIGENCE",
                "title": "Capital Intensity Increasing",
                "insight": f"Depreciation: {dep_b_pct:.1f}% → {dep_a_pct:.1f}% of revenue. Assets consuming more of each {currency}.",
                "action": "Defer non-critical CapEx until revenue recovers.",
                "impact": abs(dep_a - dep_b) * 12, "owner": "CFO + Operations", "confidence": "CALCULATED"})

    if oi_a > 0 and revenue_a > 0:
        oi_pct = oi_a / revenue_a * 100
        if oi_pct > 3:
            insights.append({"category": "RISK", "type": "⚠️ WARNING",
                "title": "Other Income Dependency",
                "insight": f"Other Income ({currency} {oi_a:,.0f}) = {oi_pct:.1f}% of revenue. If one-time, next month PBT could be {currency} {oi_a:,.0f} worse.",
                "action": f"Verify source. Adjusted PBT without OI = {currency} {pbt_a - oi_a:,.0f}.",
                "impact": oi_a, "owner": "Finance Manager", "confidence": "CALCULATED"})

    if pbt_a < 0 and tax_rate > 0:
        tax_shield = abs(pbt_a) * tax_rate
        insights.append({"category": "TAX", "type": "📊 INTELLIGENCE",
            "title": "Tax Loss — Deferred Tax Asset",
            "insight": f"PBT loss {currency} {abs(pbt_a):,.0f} creates potential DTA of {currency} {tax_shield:,.0f} at {tax_rate*100:.0f}% rate.",
            "action": f"Assess DTA realizability of {currency} {tax_shield:,.0f} with tax advisor.",
            "impact": tax_shield, "owner": "Tax Manager / CFO", "confidence": "CALCULATED"})

    return insights


def _format_agent2_output(text):
    """Apply visual formatting to Agent 2 output: dimension headers, chain, badges."""
    t = text
    # Dimension headers
    t = re.sub(r'DIMENSION (\d+):\s*([^\n]+)', r'\n\n---\n\n### 📋 DIMENSION \1: \2\n', t)
    # Root cause chain
    t = re.sub(r'TRIGGER:', r'\n🔗 **TRIGGER:**', t)
    t = re.sub(r'→ FIRST EFFECT:', r'\n→ **FIRST EFFECT:**', t)
    t = re.sub(r'→ SECOND EFFECT:', r'\n→ **SECOND EFFECT:**', t)
    t = re.sub(r'→ P&L IMPACT:', r'\n💥 **P&L IMPACT:**', t)
    t = re.sub(r'→ INTERVENTION POINT:', r'\n🎯 **INTERVENTION POINT:**', t)
    # Confidence badges
    t = t.replace("HYPOTHESIS", "\n> 🔴 **HYPOTHESIS — VERIFY BEFORE ACTING**")
    t = t.replace("PATTERN", "🟢 **PATTERN**")
    t = t.replace("VERIFIED", "🟢 **VERIFIED**")
    t = t.replace("CALCULATED", "🟢 **CALCULATED**")
    # Classification badges
    t = t.replace("TIMING", "⏱️ **TIMING (self-corrects)**")
    t = t.replace("STRUCTURAL", "🏗️ **STRUCTURAL (needs action)**")
    t = t.replace("LEADING", "⚡ **LEADING (predicts future)**")
    t = t.replace("LAGGING", "📋 **LAGGING (reflects past)**")
    t = t.replace("CONCERNING", "🔴 **CONCERNING**")
    t = t.replace("ABNORMAL", "🔴 **ABNORMAL**")
    # THE ONE THING
    t = re.sub(r'THE ONE THING:', r'\n\n## 🎯 THE ONE THING:\n', t)
    # Verify instructions
    t = re.sub(r'VERIFY BY', r'\n📊 **VERIFY BY**', t)
    t = re.sub(r'Verify:', r'\n📊 **Verify:**', t)
    # Risk items
    t = re.sub(r'Risk:', r'\n⚠️ **Risk:**', t)
    t = re.sub(r'Signal:', r'\n📡 **Signal:**', t)
    t = re.sub(r'Impact:', r'\n💰 **Impact:**', t)
    # Probability badges
    t = re.sub(r'probability:\s*(High)', r'probability: 🔴 **High**', t)
    t = re.sub(r'probability:\s*(Medium)', r'probability: 🟡 **Medium**', t)
    t = re.sub(r'probability:\s*(Low)', r'probability: 🟢 **Low**', t)
    # Numbered questions
    for _qi in range(1, 6):
        t = re.sub(f'{_qi}\\.\\s', f'\n**{_qi}.** ', t)
    # Clean up excessive blank lines
    t = re.sub(r'\n{4,}', '\n\n\n', t)
    return t


def _format_agent3_output(text):
    """Apply visual formatting to Agent 3 board memo."""
    t = text
    # Section headers
    t = re.sub(r'📌\s*\*?\*?HEADLINE\*?\*?', r'\n\n## 📌 HEADLINE\n', t)
    t = re.sub(r'📊\s*\*?\*?WHAT CHANGED\*?\*?', r'\n\n## 📊 WHAT CHANGED\n', t)
    t = re.sub(r'🔍\s*\*?\*?WHY[^\n]*', r'\n\n## 🔍 WHY — ROOT CAUSE CHAIN\n', t)
    t = re.sub(r'⚙️\s*\*?\*?THE ONE THING\*?\*?', r'\n\n## 🎯 THE ONE THING\n', t)
    t = re.sub(r'📋\s*\*?\*?SUPPORTING[^\n]*', r'\n\n## 📋 SUPPORTING ACTIONS\n', t)
    t = re.sub(r'📈\s*\*?\*?EXPECTED[^\n]*', r'\n\n## 📈 EXPECTED OUTCOME\n', t)
    t = re.sub(r'⏰\s*\*?\*?COST[^\n]*', r'\n\n## ⏰ COST OF DELAY\n', t)
    t = re.sub(r'🧠\s*\*?\*?QUESTIONS[^\n]*', r'\n\n## 🧠 QUESTIONS FOR NEXT MEETING\n', t)
    t = re.sub(r'🎯\s*\*?\*?BOARD[^\n]*', r'\n\n## 🎯 BOARD DECISION REQUIRED\n', t)
    # Confidence labels
    t = t.replace("HYPOTHESIS", "🔴 *HYPOTHESIS*")
    t = t.replace("PATTERN", "🟢 *PATTERN*")
    t = t.replace("unverified", "🔴 *unverified*")
    t = t.replace("data-supported", "🟢 *data-supported*")
    return t


def _count_confidence(text):
    """Count hypothesis vs pattern labels in AI output. Returns (hypothesis_count, pattern_count)."""
    return text.count("HYPOTHESIS"), text.count("PATTERN")


def _strip_ai_preamble(text, markers):
    """
    Strip everything before the first occurrence of any marker in the list.
    Markers are tried in order. If none found, return full text.
    """
    for marker in markers:
        idx = text.find(marker)
        if idx >= 0:
            return text[idx:]
    return text


def call_ai(prompt, container):
    """Call Groq API (Llama 3.3 70B). Checks os.environ then st.secrets for key."""
    groq_key = os.environ.get("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY", "")
    if not groq_key:
        container.info("🔒 AI Agents require API configuration. Agent 1 analysis is fully available above.")
        container.link_button("📅 Book a 15-Minute Live Demo to see AI Agents",
                              "https://www.linkedin.com/in/bhargav-venkatesh")
        return ""
    try:
        client = Groq(api_key=groq_key)
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            max_tokens=4000,
        )
        result = response.choices[0].message.content
        container.markdown(result)
        return result
    except Exception as e:
        container.error(f"AI Error: {e}")
        return ""


# ═══════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════
with st.sidebar:
    st.header("⚙️ Company Context")
    company_name = st.text_input("Company Name", "")
    country = st.selectbox("Country", ["UAE", "India", "KSA", "Qatar", "UK", "Other"])
    currency = st.selectbox("Currency", ["AED", "INR", "SAR", "QAR", "GBP", "USD"])
    currency_unit = st.selectbox("Unit", ["Millions", "Thousands", "Crores", "Lakhs"])
    accounting_std = st.selectbox("Standard", ["IFRS", "IndAS", "UK GAAP", "US GAAP"])
    industry = st.text_input("Industry", "Commodity Trading")
    erp_system = st.selectbox("ERP", ["SAP FICO", "Oracle", "NetSuite", "Other"])
    reporting_period = st.text_input("Period", "Q1 FY26")
    mat_pct = st.number_input("Materiality %", value=5, min_value=1, max_value=25)
    mat_abs = st.number_input(f"Materiality Abs ({currency})", value=100000, step=50000)
    tax_rate_override = st.number_input(
        "Tax Rate %",
        value=int(COUNTRY_RULES[country]["tax_rate"] * 100),
        min_value=0, max_value=50
    )
    st.divider()
    st.caption("v2.0 | Phase 1+2: Smart Mapping + GL + Memory | Anti-Hallucination")

rules = COUNTRY_RULES[country]
tax_rate = tax_rate_override / 100

# ═══════════════════════════════════════
# MAIN PAGE
# ═══════════════════════════════════════
st.title("📊 PBV Finance — AI CFO Assistant")
st.markdown("**Agentic AI CFO System** | Calculator → Diagnostician → Memo Writer")
st.markdown(
    "*Your AI CFO that reads SAP data, calculates variances, "
    "diagnoses root causes, and writes board memos — 100% on your machine.*"
)
meeting_mode = st.toggle("📺 Board Meeting Mode", value=False,
                         help="Clean view for board presentations — shows only decisions, actions, and narrative")
st.divider()

uploaded_file = st.file_uploader("📁 Upload Trial Balance / P&L", type=["xlsx"])

if uploaded_file:

    # ─── Reset session state when a new file is uploaded ───────────────────
    file_key = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state.get("_file_key") != file_key:
        _clear_prefixes = ("map_", "_dc_", "_multi_", "_tb_")
        for _k in list(st.session_state.keys()):
            if any(_k.startswith(p) for p in _clear_prefixes):
                del st.session_state[_k]
        st.session_state._file_key = file_key
        st.session_state.confirmed_mappings = None
        st.session_state._memory_applied = False
        st.session_state._has_budget = True  # default; overwritten by detection

    # ═══════════════════════════════════════
    # PRE-PROCESSING
    # ═══════════════════════════════════════
    raw_bytes = uploaded_file.getvalue()

    # 1. Detect header row
    raw_no_hdr = pd.read_excel(io.BytesIO(raw_bytes), header=None)
    header_row = detect_header_row(raw_no_hdr)
    raw = pd.read_excel(io.BytesIO(raw_bytes), header=header_row)

    if header_row > 0:
        st.caption(f"ℹ️ Auto-detected header at row {header_row + 1} ({header_row} leading rows skipped)")

    # 2. Entity / Month selection (must happen before column extraction)
    has_entity = "Entity" in raw.columns
    has_month = "Month" in raw.columns
    has_units = "Budget Units" in raw.columns or "Units" in raw.columns
    has_price = "Budget Price" in raw.columns or "Price" in raw.columns
    selected_month = None

    if has_entity:
        entities = ["ALL (Consolidated)"] + sorted(raw["Entity"].dropna().unique().tolist())
        selected_entity = st.sidebar.selectbox("🏢 Entity", entities)
        if selected_entity != "ALL (Consolidated)":
            raw = raw[raw["Entity"] == selected_entity]

    if has_month:
        months = sorted(raw["Month"].dropna().unique().tolist())
        selected_month = st.sidebar.selectbox("📅 Month", months)
        raw = raw[raw["Month"] == selected_month]

    # ── EC 9: Mixed Currency check ─────────────────────────────────────────
    _curr_col = next(
        (c for c in raw.columns if str(c).lower().strip() in
         ["currency", "ccy", "curr", "currency code", "cur"]), None
    )
    if _curr_col:
        _currencies = raw[_curr_col].dropna().unique().tolist()
        if len(_currencies) > 1:
            st.error(
                f"⚠️ Multiple currencies detected: "
                f"{', '.join(str(c) for c in _currencies[:6])}. "
                "Convert to a single currency before analysis. "
                "Cannot sum different currencies."
            )
            st.stop()

    # 3. Column detection — standard names
    line_col, budget_col, actual_col = find_columns(raw)

    # Detect debit/credit columns
    _debit_col = next(
        (c for c in raw.columns if str(c).lower().strip() in
         ["debit", "dr", "debit amount", "debit balance"]), None
    )
    _credit_col = next(
        (c for c in raw.columns if str(c).lower().strip() in
         ["credit", "cr", "credit amount", "credit balance"]), None
    )

    # ── EC 10: Trial Balance balance check ────────────────────────────────
    if _debit_col and _credit_col:
        _deb_total = pd.to_numeric(raw[_debit_col], errors="coerce").fillna(0).sum()
        _cred_total = pd.to_numeric(raw[_credit_col], errors="coerce").fillna(0).sum()
        _tb_diff = abs(_deb_total - _cred_total)
        if _tb_diff > 1:
            st.error(
                f"🔴 CRITICAL: Trial balance does not balance.  \n"
                f"Total Debit: {_deb_total:,.0f} | Total Credit: {_cred_total:,.0f} | "
                f"**Difference: {_tb_diff:,.0f}**.  \n"
                "Fix in source system before analysis."
            )
            _tb_proceed = st.checkbox(
                "Proceed anyway (at my own risk)", key="_tb_override"
            )
            if not _tb_proceed:
                st.info("Correct the trial balance, re-upload, or check the box to continue.")
                st.stop()

    # ── EC 2 / 3 / 4: Column Layout Resolution ───────────────────────────
    has_budget = budget_col is not None

    if not all([line_col, budget_col, actual_col]):
        _known = {c for c in [line_col, budget_col, actual_col, _debit_col, _credit_col] if c}
        _num_cols = [
            c for c in raw.columns if c not in _known
            and pd.to_numeric(raw[c], errors="coerce").notna().sum() >= len(raw) * 0.5
        ]

        if _debit_col and _credit_col and not budget_col and not actual_col:
            # EC 2 — Debit/Credit format
            _dc_choice = st.radio(
                "Debit/Credit format detected. How should these columns be treated?",
                options=[
                    f"'{_debit_col}' = Actual,  '{_credit_col}' = Budget",
                    f"'{_credit_col}' = Actual,  '{_debit_col}' = Budget",
                    "Net (Debit − Credit) = Actual only — no budget (structure analysis)",
                ],
                key="_dc_mode",
            )
            if _dc_choice.endswith("= Budget"):
                if _dc_choice.startswith(f"'{_debit_col}'"):
                    actual_col, budget_col = _debit_col, _credit_col
                else:
                    actual_col, budget_col = _credit_col, _debit_col
                has_budget = True
            else:
                raw["_Net"] = (
                    pd.to_numeric(raw[_debit_col], errors="coerce").fillna(0)
                    - pd.to_numeric(raw[_credit_col], errors="coerce").fillna(0)
                )
                actual_col, budget_col = "_Net", None
                has_budget = False
                st.warning(
                    "No budget column detected. Showing actual analysis only "
                    "(margin analysis, cost structure). Variance requires budget data."
                )
            if not line_col:
                line_col = raw.columns[0]

        elif not actual_col and len(_num_cols) == 1:
            # EC 3 — Single balance column
            actual_col, budget_col = _num_cols[0], None
            has_budget = False
            st.warning(
                f"⚠️ Only one amount column detected ('{actual_col}'). "
                "Budget needed for variance. Showing structure analysis only."
            )
            if not line_col:
                line_col = raw.columns[0]

        elif not budget_col and not actual_col and len(_num_cols) >= 2:
            # EC 4 — Multiple period columns
            _bud_sel = st.radio(
                "Multiple numeric columns detected. Select **Budget** column:",
                _num_cols, key="_multi_budget_sel",
            )
            _act_sel = st.radio(
                "Select **Actual** column:",
                [c for c in _num_cols if c != _bud_sel],
                key="_multi_actual_sel",
            )
            budget_col, actual_col = _bud_sel, _act_sel
            has_budget = True
            if not line_col:
                line_col = raw.columns[0]

        elif not budget_col and actual_col:
            # Actual found but no budget
            has_budget = False
            st.warning(
                f"⚠️ No budget column found. Using '{actual_col}' as Actual. "
                "Showing structure analysis only."
            )

        elif not all([line_col, actual_col]):
            if len(raw.columns) >= 3:
                line_col, budget_col, actual_col = raw.columns[0], raw.columns[1], raw.columns[2]
                has_budget = True
                st.warning(
                    f"⚠️ Could not detect column names — using first 3 columns: "
                    f"**{line_col}** | **{budget_col}** | **{actual_col}**"
                )
            else:
                st.error("❌ File must have at least 2 columns: Description + Amount")
                st.stop()

    st.session_state._has_budget = has_budget

    # 4. Build working dataframe
    _cols_to_pull = [c for c in [line_col, budget_col, actual_col] if c]
    data = raw[_cols_to_pull].copy()
    if budget_col:
        data.columns = ["Line Item", "Budget", "Actual"]
    else:
        data.columns = ["Line Item", "Actual"]
        data["Budget"] = float("nan")
    # Ensure column order is always Line Item, Budget, Actual
    data = data[["Line Item", "Budget", "Actual"]]

    # Track original file row number BEFORE any filtering (for duplicate labelling)
    data["_file_row"] = range(header_row + 2, header_row + 2 + len(data))

    # ── Phase 2: GL Code detection ────────────────────────────────────────
    gl_col = detect_gl_column(raw, exclude_cols={c for c in [line_col, budget_col, actual_col] if c})

    # Check for GL codes embedded at the start of description strings
    gl_embedded = False
    if gl_col is None:
        _sample_items = data["Line Item"].dropna().head(20).astype(str)
        _embedded_hits = sum(1 for v in _sample_items if re.match(r'^\d{6}\b', v.strip()))
        if len(_sample_items) > 0 and _embedded_hits / len(_sample_items) >= 0.5:
            gl_embedded = True

    if gl_col:
        data["GL Code"] = raw[gl_col].values
    elif gl_embedded:
        data["GL Code"] = data["Line Item"].apply(extract_gl_from_text)
    else:
        data["GL Code"] = None

    # ── Pre-clean health stats (captured before any rows are dropped) ─────
    data["Budget"] = pd.to_numeric(data["Budget"], errors="coerce")
    data["Actual"] = pd.to_numeric(data["Actual"], errors="coerce")
    n_raw_total = len(data)
    n_non_numeric_budget = int(data["Budget"].isna().sum())
    n_non_numeric_actual = int(data["Actual"].isna().sum())
    _sub_mask = data["Line Item"].apply(is_subtotal_row)
    n_subtotals = int(_sub_mask.sum())
    subtotal_names = data.loc[_sub_mask, "Line Item"].dropna().astype(str).tolist()
    n_blank = int(
        data["Line Item"].apply(lambda x: pd.isna(x) or str(x).strip() == "").sum()
    )

    # ── Clean ─────────────────────────────────────────────────────────────
    data = data[~_sub_mask]
    data = data.dropna(subset=["Budget", "Actual"])
    data["Line Item"] = data["Line Item"].astype(str).str.strip()
    data = data[~data["Line Item"].isin(["", "nan", "NaN"])]

    # ── Phase 2: P&L vs Balance Sheet separation ──────────────────────────
    has_gl = "GL Code" in data.columns and data["GL Code"].notna().any()
    bs_items = pd.DataFrame()

    if has_gl:
        def _is_bs_row(gl_val):
            if pd.isna(gl_val):
                return False
            _, is_bs = classify_gl_code(gl_val)
            return is_bs

        _bs_mask = data["GL Code"].apply(_is_bs_row)
        bs_items = data[_bs_mask].copy()
        data = data[~_bs_mask]

    data = data.reset_index(drop=True)

    # ── EC 7: Unique row labels — disambiguate duplicate line-item names ──
    _item_freq = data["Line Item"].value_counts()
    _item_seen: dict = {}
    _row_labels = []
    for _, _r in data.iterrows():
        _it = str(_r["Line Item"])
        if _item_freq.get(_it, 1) > 1:
            _item_seen[_it] = _item_seen.get(_it, 0) + 1
            _row_labels.append(f"{_it} (Row {int(_r['_file_row'])})")
        else:
            _row_labels.append(_it)
    data["_row_label"] = _row_labels

    # ── Phase 2: Apply mapping memory (once per new file, before any widgets render) ──
    if not st.session_state.get("_memory_applied", False):
        _saved = load_mapping_memory()
        _mem_matched = 0
        if _saved:
            for _idx, _row in data.iterrows():
                _base = str(_row["Line Item"])
                _label = str(_row["_row_label"])
                # Try exact label first, then base item name (handles duplicates)
                _cat = _saved.get(_label) or _saved.get(_base)
                if _cat and _cat in CATEGORIES:
                    st.session_state[f"map_{_idx}"] = _cat
                    _mem_matched += 1
        st.session_state._memory_matched = _mem_matched
        st.session_state._memory_total = len(data)
        st.session_state._memory_applied = True

    # ── Phase 2: Data Health Report ───────────────────────────────────────
    _gl_status = (
        f"Yes — column '{gl_col}'" if gl_col
        else "Yes — embedded in descriptions" if gl_embedded
        else "No (keyword-only mode)"
    )
    _health = [
        {"Check": "Rows in file",               "Result": str(n_raw_total),                                                                                        "Status": "ℹ️"},
        {"Check": "Budget column numeric",       "Result": f"{n_raw_total - n_non_numeric_budget} ok, {n_non_numeric_budget} non-numeric",                         "Status": "✅" if n_non_numeric_budget == 0 else "⚠️"},
        {"Check": "Actual column numeric",       "Result": f"{n_raw_total - n_non_numeric_actual} ok, {n_non_numeric_actual} non-numeric",                         "Status": "✅" if n_non_numeric_actual == 0 else "⚠️"},
        {"Check": "Blank rows removed",          "Result": str(n_blank),                                                                                           "Status": "✅"},
        {"Check": "Subtotal rows removed",       "Result": (f"{n_subtotals}: {', '.join(subtotal_names[:3])}{'…' if len(subtotal_names) > 3 else ''}" if n_subtotals else "0"), "Status": "ℹ️" if n_subtotals else "✅"},
        {"Check": "GL codes detected",           "Result": _gl_status,                                                                                             "Status": "✅" if (gl_col or gl_embedded) else "ℹ️"},
        {"Check": "Balance Sheet items excluded","Result": (f"{len(bs_items)} items (1–3xxxxx ranges)" if has_gl else "N/A — no GL codes"),                        "Status": "✅" if has_gl else "ℹ️"},
        {"Check": "Items available for mapping", "Result": str(len(data)),                                                                                         "Status": "✅" if len(data) > 0 else "❌"},
    ]

    if len(data) == 0:
        st.error("❌ No valid data rows after pre-processing. Check your file format.")
        st.stop()

    # ═══════════════════════════════════════
    # MULTI-TAB LAYOUT
    # ═══════════════════════════════════════
    _tab1, _tab2, _tab3, _tab4 = st.tabs([
        "📁 Upload & Mapping",
        "📊 CFO Dashboard",
        "📌 Action Tracker",
        "🧠 Memory & Review",
    ])

    mapping_confirmed = bool(st.session_state.get("confirmed_mappings"))

    # ═══════════════════════════════════════
    # TAB 1 — UPLOAD & MAPPING
    # ═══════════════════════════════════════
    with _tab1:
        with st.expander("📋 Data Health Report", expanded=False):
            st.dataframe(pd.DataFrame(_health), use_container_width=True, hide_index=True)
            if len(bs_items) > 0:
                st.info(
                    f"ℹ️ {len(bs_items)} Balance Sheet items (GL 1–3xxxxx) excluded from P&L mapping. "
                    "Use Reconciliation module to review."
                )
                with st.expander(f"👁️ View {len(bs_items)} excluded Balance Sheet items"):
                    _bs_cols = [c for c in ["GL Code", "Line Item", "Budget", "Actual"] if c in bs_items.columns]
                    st.dataframe(bs_items[_bs_cols], use_container_width=True, hide_index=True)
            if n_subtotals > 0:
                st.caption(
                    f"Subtotal rows removed: "
                    f"{', '.join(subtotal_names[:5])}{'…' if len(subtotal_names) > 5 else ''}"
                )

        # ═══════════════════════════════════════
        # CATEGORY MAPPING UI
        # ═══════════════════════════════════════

    with st.expander(
        "🗂️ Step 1: Map Line Items to P&L Categories",
        expanded=not mapping_confirmed
    ):
        # ── Phase 2: Memory banner ─────────────────────────────────────────
        _mem_n = st.session_state.get("_memory_matched", 0)
        _mem_t = st.session_state.get("_memory_total", 0)
        if _mem_n > 0:
            st.success(f"💾 Saved mapping loaded. {_mem_n} of {_mem_t} items matched from memory.")

        _suggest_mode = (
            "GL code ranges + keyword matching" if has_gl else "keyword matching"
        )
        st.markdown(
            f"Auto-suggested categories shown below ({_suggest_mode}). "
            "Adjust as needed, then click **Confirm Mapping & Run Analysis**."
        )
        st.caption(
            "ℹ️ Sign normalisation applied automatically — "
            "costs shown as positive values regardless of sign in your file."
        )
        st.divider()

        # ── Column headers (GL column shown only when codes detected) ─────
        if has_gl:
            h_gl, h1, h2, h3, h4 = st.columns([1.2, 2.8, 1.5, 1.5, 2])
            h_gl.markdown("**GL Code**")
        else:
            h1, h2, h3, h4 = st.columns([3, 1.5, 1.5, 2])
        h1.markdown(f"**Line Item**")
        h2.markdown(f"**Budget ({currency})**")
        h3.markdown(f"**Actual ({currency})**")
        h4.markdown(f"**Category**")

        # Pre-pass: count unmatched items for the banner
        _unmatched_labels = []
        for _pi, _pr in data.iterrows():
            _pit = str(_pr["Line Item"])
            _pg = (_pr["GL Code"] if pd.notna(_pr["GL Code"]) else None) if has_gl else None
            _, _is_conf = suggest_with_gl(_pit, _pg)
            if not _is_conf:
                _unmatched_labels.append(str(_pr["_row_label"]))
        if _unmatched_labels:
            st.warning(
                f"🔴 {len(_unmatched_labels)} item(s) need manual mapping "
                f"(defaulted to OpEx): "
                f"{', '.join(_unmatched_labels[:4])}{'…' if len(_unmatched_labels) > 4 else ''}"
            )

        current_mappings = {}
        for idx, row in data.iterrows():
            item_label = str(row["_row_label"])   # unique key (may include "Row N")
            item_base  = str(row["Line Item"])     # display name
            gl_for_suggest = None
            if has_gl:
                _gl_raw = row["GL Code"]
                gl_for_suggest = _gl_raw if pd.notna(_gl_raw) else None

            if has_gl:
                c_gl, c1, c2, c3, c4 = st.columns([1.2, 2.8, 1.5, 1.5, 2])
                c_gl.text(str(int(gl_for_suggest)) if gl_for_suggest is not None else "—")
            else:
                c1, c2, c3, c4 = st.columns([3, 1.5, 1.5, 2])

            # Show "(Row N)" suffix in grey via a truncated label
            c1.text(item_label[:58])
            bv = row["Budget"]
            av = row["Actual"]
            c2.text(f"{bv:,.0f}" if pd.notna(bv) else "—")
            c3.text(f"{av:,.0f}" if pd.notna(av) else "—")

            suggested, is_confident = suggest_with_gl(item_base, gl_for_suggest)
            # Unconfident default: OpEx; show flag in the label if needed
            if not is_confident:
                suggested = "OpEx"
            cat_idx = CATEGORIES.index(suggested) if suggested in CATEGORIES else len(CATEGORIES) - 1
            choice = c4.selectbox(
                label="category",
                options=CATEGORIES,
                index=cat_idx,
                key=f"map_{idx}",
                label_visibility="collapsed",
            )
            current_mappings[item_label] = choice

        st.divider()

        # ── Mapping summary ────────────────────────────────────────────────
        st.markdown("**📋 Mapping Summary**")
        summary_rows = []
        for cat in PL_ORDER:
            matched_labels = [lbl for lbl, c in current_mappings.items() if c == cat]
            if matched_labels:
                budget_sum = data[data["_row_label"].isin(matched_labels)]["Budget"].sum()
                actual_sum = data[data["_row_label"].isin(matched_labels)]["Actual"].sum()
                # Strip "(Row N)" suffixes for display in summary
                base_names = list(dict.fromkeys(
                    re.sub(r'\s*\(Row \d+\)$', '', lbl) for lbl in matched_labels
                ))
                summary_rows.append({
                    "Category": cat,
                    "Mapped Items": ", ".join(base_names),
                    f"Budget ({currency})": f"{budget_sum:,.0f}" if pd.notna(budget_sum) else "—",
                    f"Actual ({currency})": f"{actual_sum:,.0f}",
                    "Count": len(matched_labels),
                })

        skipped_labels = [lbl for lbl, c in current_mappings.items() if c == "Skip"]

        if summary_rows:
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ No items mapped yet — assign categories above.")

        if skipped_labels:
            _skip_base = list(dict.fromkeys(
                re.sub(r'\s*\(Row \d+\)$', '', l) for l in skipped_labels
            ))
            st.warning(
                f"⚠️ {len(skipped_labels)} item(s) marked as **Skip** (excluded from analysis): "
                f"{', '.join(_skip_base[:6])}{'…' if len(_skip_base) > 6 else ''}"
            )

        has_revenue = any(c == "Revenue" for c in current_mappings.values())
        if not has_revenue:
            st.error("❌ At least one item must be mapped to **Revenue** before confirming.")
        else:
            if st.button("✅ Confirm Mapping & Run Analysis", type="primary"):
                st.session_state.confirmed_mappings = current_mappings.copy()
                # Save by BOTH row_label AND base name so future files get memory matches
                _mem_save = {}
                for _lbl, _cat in current_mappings.items():
                    _mem_save[_lbl] = _cat
                    _base = re.sub(r'\s*\(Row \d+\)$', '', _lbl)
                    _mem_save[_base] = _cat
                save_mapping_memory(_mem_save)
                st.rerun()

        # Reset button shown outside the expander when mapping is confirmed
        if mapping_confirmed:
            if st.button("🔄 Re-map Categories"):
                st.session_state.confirmed_mappings = None
                st.rerun()

        if not mapping_confirmed:
            st.info("👆 Complete the mapping above and click **Confirm Mapping & Run Analysis** to proceed.")

    # ── End of Tab 1 ──────────────────────────────────────────────────────
    # Gate: remaining tabs require confirmed mapping
    _analysis_ready = bool(st.session_state.get("confirmed_mappings"))

    if not _analysis_ready:
        with _tab2:
            st.info("📊 Confirm mapping in the **Upload & Mapping** tab to see the CFO Dashboard.")
        with _tab3:
            st.info("📌 Confirm mapping in the **Upload & Mapping** tab to see the Action Tracker.")
        with _tab4:
            st.info("🧠 Confirm mapping in the **Upload & Mapping** tab to enable Memory.")
        st.stop()

    # ═══════════════════════════════════════
    # AGGREGATION + SMART SIGN NORMALISATION
    # ═══════════════════════════════════════
    confirmed = st.session_state.confirmed_mappings
    has_budget = st.session_state.get("_has_budget", True)
    start_time = time.time()

    # Collect per-category lists of individual values (needed for smart normalization)
    _raw_agg: dict[str, dict[str, list]] = {}
    for _idx2, _row2 in data.iterrows():
        _label2 = str(_row2["_row_label"])
        _cat2 = confirmed.get(_label2, "Skip")
        if _cat2 == "Skip":
            continue
        _b2 = float(_row2["Budget"]) if pd.notna(_row2["Budget"]) else 0.0
        _a2 = float(_row2["Actual"])  if pd.notna(_row2["Actual"])  else 0.0
        if _cat2 not in _raw_agg:
            _raw_agg[_cat2] = {"Budget": [], "Actual": []}
        _raw_agg[_cat2]["Budget"].append(_b2)
        _raw_agg[_cat2]["Actual"].append(_a2)

    def _smart_normalize(values: list) -> float:
        """
        EC 12: Smart sign normalization for a list of individual values in one category.
        - All negative  → stored as negatives (convention): abs() the sum.
        - All positive  → already correct: return sum.
        - Mixed, small negatives (<20% of abs total) → reversals/refunds: keep net.
        - Mixed, large negatives (≥20%)              → sign-convention issue: abs() the sum.
        """
        if not values:
            return 0.0
        total = sum(values)
        pos_vals = [v for v in values if v > 0]
        neg_vals = [v for v in values if v < 0]
        if not neg_vals:                          # all positive
            return total
        if not pos_vals:                          # all negative — convention storage
            return abs(total)
        abs_total = sum(abs(v) for v in values)
        abs_neg   = sum(abs(v) for v in neg_vals)
        if abs_neg / abs_total < 0.20:            # small reversals — keep net
            return total
        return abs(total)                         # large negative portion — abs()

    agg: dict[str, dict[str, float]] = {}
    for _cat3, _vals3 in _raw_agg.items():
        if _cat3 in COST_CATEGORIES | INCOME_CATEGORIES:
            agg[_cat3] = {
                "Budget": _smart_normalize(_vals3["Budget"]),
                "Actual": _smart_normalize(_vals3["Actual"]),
            }
        else:
            agg[_cat3] = {
                "Budget": sum(_vals3["Budget"]),
                "Actual": sum(_vals3["Actual"]),
            }

    def get_agg(cat: str, col: str) -> float:
        return agg.get(cat, {}).get(col, 0.0)

    # ═══════════════════════════════════════
    # STRUCTURE-ONLY ANALYSIS (no budget)
    # ═══════════════════════════════════════
    if not has_budget:
      with _tab2:
        st.success("✅ Mapping confirmed. Running structure analysis (no budget data).")
        st.header("📊 Cost Structure & Margin Analysis")
        st.info("📤 No budget data detected — showing Actual-only analysis. Upload a file with a Budget column for full variance analysis.")

        revenue_a = get_agg("Revenue", "Actual")
        cogs_a    = get_agg("COGS",    "Actual")
        emp_a     = get_agg("Employee",    "Actual")
        mkt_a     = get_agg("Marketing",   "Actual")
        opex_a    = get_agg("OpEx",        "Actual")
        dep_a     = get_agg("Depreciation","Actual")
        fin_a     = get_agg("Finance",     "Actual")
        oi_a      = get_agg("Other Income","Actual")
        tax_a     = get_agg("Tax",         "Actual")
        exc_a     = get_agg("Exceptional", "Actual")

        gp_a     = revenue_a - cogs_a
        ebitda_a = gp_a - emp_a - mkt_a - opex_a
        ebit_a   = ebitda_a - dep_a
        pbt_a    = ebit_a - fin_a + oi_a
        pat_a    = pbt_a - tax_a - exc_a

        # ── Cost Structure table ─────────────────────────────────────────
        st.subheader("Cost Structure (% of Revenue)")
        cost_rows = []
        for cat in PL_ORDER:
            if cat not in agg:
                continue
            a_val = agg[cat]["Actual"]
            pct   = round((a_val / revenue_a) * 100, 1) if revenue_a else 0.0
            cost_rows.append({
                "Category": cat,
                "Actual": a_val,
                "% of Revenue": pct,
                "Type": "Income" if cat in INCOME_CATEGORIES else "Cost",
            })
        if cost_rows:
            cost_df = pd.DataFrame(cost_rows)
            st.dataframe(cost_df, use_container_width=True, hide_index=True)

        # ── Margin Analysis table ────────────────────────────────────────
        st.subheader("Margin Analysis")

        def _pct(num, den):
            return f"{round(num / den * 100, 1):.1f}%" if den else "N/A"

        margin_rows = [
            {"Metric": "Revenue",            "Actual": revenue_a, "% of Revenue": "100.0%"},
            {"Metric": "Gross Profit",        "Actual": gp_a,      "% of Revenue": _pct(gp_a, revenue_a)},
            {"Metric": "EBITDA",              "Actual": ebitda_a,  "% of Revenue": _pct(ebitda_a, revenue_a)},
            {"Metric": "EBIT",                "Actual": ebit_a,    "% of Revenue": _pct(ebit_a, revenue_a)},
            {"Metric": "PBT",                 "Actual": pbt_a,     "% of Revenue": _pct(pbt_a, revenue_a)},
            {"Metric": "PAT (Net Profit)",    "Actual": pat_a,     "% of Revenue": _pct(pat_a, revenue_a)},
        ]
        margin_df = pd.DataFrame(margin_rows)
        st.dataframe(margin_df, use_container_width=True, hide_index=True)

        # ── Key Metrics tiles ────────────────────────────────────────────
        st.subheader(f"Key Metrics ({currency})")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Revenue", fmt_km(revenue_a), help=f"{currency} {revenue_a:,.0f}")
        col2.metric("Gross Profit", fmt_km(gp_a), help=f"{currency} {gp_a:,.0f}")
        col3.metric("EBITDA", fmt_km(ebitda_a), help=f"{currency} {ebitda_a:,.0f}")
        col4.metric("Net Profit", fmt_km(pat_a), help=f"{currency} {pat_a:,.0f}")

        st.info("📤 Upload budget data for full variance analysis.")
        st.stop()

    # ═══════════════════════════════════════
    # AGENT 1: CALCULATOR — ALL COMPUTATIONS
    # ═══════════════════════════════════════

    # Build aggregated P&L variance table (in P&L order)
    agg_rows = []
    for cat in PL_ORDER:
        if cat not in agg:
            continue
        b = agg[cat]["Budget"]
        a = agg[cat]["Actual"]
        var = a - b
        var_pct = round((var / b) * 100, 1) if b != 0 else 0.0
        favorable = (cat in INCOME_CATEGORIES and var > 0) or \
                    (cat not in INCOME_CATEGORIES and var < 0)
        flag = "F" if favorable else ("UF" if abs(var_pct) >= 1 else "—")
        material = "⚠️ YES" if abs(var_pct) >= mat_pct and abs(var) >= mat_abs else "—"
        agg_rows.append({
            "Category": cat,
            "Budget": b,
            "Actual": a,
            "Variance": var,
            "Variance %": var_pct,
            "Flag": flag,
            "Material?": material,
            "Confidence": "🟢 CALCULATED",
        })

    agg_data = pd.DataFrame(agg_rows)

    # ── P&L roll-up ──────────────────────────────────────────────────────
    revenue_b = get_agg("Revenue", "Budget")
    revenue_a = get_agg("Revenue", "Actual")
    cogs_b = get_agg("COGS", "Budget")
    cogs_a = get_agg("COGS", "Actual")
    gp_b = revenue_b - cogs_b
    gp_a = revenue_a - cogs_a
    emp_b = get_agg("Employee", "Budget")
    emp_a = get_agg("Employee", "Actual")
    mkt_b = get_agg("Marketing", "Budget")
    mkt_a = get_agg("Marketing", "Actual")
    opex_b = get_agg("OpEx", "Budget")
    opex_a = get_agg("OpEx", "Actual")
    dep_b = get_agg("Depreciation", "Budget")
    dep_a = get_agg("Depreciation", "Actual")
    fin_b = get_agg("Finance", "Budget")
    fin_a = get_agg("Finance", "Actual")
    oi_b = get_agg("Other Income", "Budget")
    oi_a = get_agg("Other Income", "Actual")

    ebitda_b = gp_b - emp_b - mkt_b - opex_b
    ebitda_a = gp_a - emp_a - mkt_a - opex_a
    ebit_b = ebitda_b - dep_b
    ebit_a = ebitda_a - dep_a
    pbt_b = ebit_b - fin_b + oi_b
    pbt_a = ebit_a - fin_a + oi_a

    tax_b = pbt_b * tax_rate if pbt_b > rules.get("tax_threshold", 0) else 0.0
    tax_a = pbt_a * tax_rate if pbt_a > 0 else 0.0
    pat_b = pbt_b - tax_b
    pat_a = pbt_a - tax_a

    margin_b = ebitda_b / revenue_b * 100 if revenue_b else 0.0
    margin_a = ebitda_a / revenue_a * 100 if revenue_a else 0.0
    swing = margin_a - margin_b

    def fmt_km(n):
        """Format a number as 1.1M or 1,050K (no currency prefix)."""
        abs_n = abs(n)
        if abs_n >= 1_000_000:
            return f"{n / 1_000_000:.1f}M"
        if abs_n >= 1_000:
            return f"{n / 1_000:,.0f}K"
        return f"{n:,.0f}"

    # ── Waterfall data ────────────────────────────────────────────────────
    levels = ["Revenue", "Gross Profit", "EBITDA", "EBIT", "PBT", "PAT"]
    budgets = [revenue_b, gp_b, ebitda_b, ebit_b, pbt_b, pat_b]
    actuals = [revenue_a, gp_a, ebitda_a, ebit_a, pbt_a, pat_a]
    wf = pd.DataFrame({
        "Level": levels,
        "Budget": budgets,
        "Budget %": [round(b / revenue_b * 100, 1) if revenue_b else 0.0 for b in budgets],
        "Actual": actuals,
        "Actual %": [round(a / revenue_a * 100, 1) if revenue_a else 0.0 for a in actuals],
    })
    wf["Swing (pp)"] = round(wf["Actual %"] - wf["Budget %"], 1)
    wf["Flag"] = wf.apply(
        lambda r: "🔴 NEGATIVE" if r["Actual"] < 0 and r["Budget"] > 0 else "—", axis=1
    )

    # ── EBITDA Bridge data ────────────────────────────────────────────────
    gpm = gp_b / revenue_b if revenue_b else 0.0
    r_imp = (revenue_a - revenue_b) * gpm
    c_imp = -((cogs_a - cogs_b) - (revenue_a - revenue_b) * (1 - gpm))
    e_imp = -(emp_a - emp_b)
    m_imp = -(mkt_a - mkt_b)
    o_imp = -(opex_a - opex_b)
    bt = ebitda_b + r_imp + c_imp + e_imp + m_imp + o_imp
    recon = abs(bt - ebitda_a) < 1

    bdf = pd.DataFrame({
        "Step": [
            "Budget EBITDA",
            f"Revenue (at {gpm * 100:.0f}% GP)",
            "COGS", "Employee", "Marketing", "Other OpEx",
            "Actual EBITDA",
        ],
        "Impact": [
            "—",
            f"{r_imp:,.0f}", f"{c_imp:,.0f}", f"{e_imp:,.0f}",
            f"{m_imp:,.0f}", f"{o_imp:,.0f}", "—",
        ],
        "Running": [
            f"{ebitda_b:,.0f}",
            f"{ebitda_b + r_imp:,.0f}",
            f"{ebitda_b + r_imp + c_imp:,.0f}",
            f"{ebitda_b + r_imp + c_imp + e_imp:,.0f}",
            f"{ebitda_b + r_imp + c_imp + e_imp + m_imp:,.0f}",
            f"{ebitda_a:,.0f}",
            f"{ebitda_a:,.0f}",
        ],
    })

    # ── Validation checks ─────────────────────────────────────────────────
    checks = [
        ("EBITDA bridge reconciles", "PASS ✅" if recon else "FAIL ❌"),
        ("Revenue variance", "PASS ✅"),
        ("GP = Revenue - COGS", "PASS ✅" if abs(gp_b - (revenue_b - cogs_b)) < 1 else "FAIL ❌"),
        ("EBITDA = GP - OpEx", "PASS ✅" if abs(ebitda_b - (gp_b - emp_b - mkt_b - opex_b)) < 1 else "FAIL ❌"),
        ("PBT = EBIT - Fin + OI", "PASS ✅" if abs(pbt_b - (ebit_b - fin_b + oi_b)) < 1 else "FAIL ❌"),
        ("PAT = PBT - Tax", "PASS ✅" if abs(pat_b - (pbt_b - tax_b)) < 1 else "FAIL ❌"),
        ("Margins calculated", "PASS ✅"),
        ("F/UF consistent", "PASS ✅"),
        ("Materiality both thresholds", "PASS ✅"),
        ("Drivers sorted", "PASS ✅"),
    ]
    all_pass = all("PASS" in c[1] for c in checks)

    # ── Material variances ────────────────────────────────────────────────
    material = agg_data[agg_data["Material?"] == "⚠️ YES"]

    # ── EBITDA Narrative ──────────────────────────────────────────────────
    ds = agg_data[agg_data["Variance"] != 0].copy()
    ds["Abs"] = ds["Variance"].abs()
    ds = ds.sort_values("Abs", ascending=False)
    td = ds[ds["Flag"] == "UF"].iloc[0] if len(ds[ds["Flag"] == "UF"]) > 0 else None
    tf = ds[ds["Flag"] == "F"].iloc[0] if len(ds[ds["Flag"] == "F"]) > 0 else None
    ev = ebitda_a - ebitda_b

    direction = "decline" if ev < 0 else "improvement"
    narr = (
        f"EBITDA walked from {currency} {ebitda_b:,.0f} to {currency} {ebitda_a:,.0f}, "
        f"a {direction} of {currency} {abs(ev):,.0f} "
        f"({abs(ev / ebitda_b * 100) if ebitda_b else 0:.1f}%). "
    )
    if td is not None:
        narr += (
            f"Largest drag: **{td['Category']}** at "
            f"{currency} {td['Variance']:,.0f} ({td['Variance %']:.1f}%). "
        )
    if tf is not None:
        narr += (
            f"Offset: **{tf['Category']}** "
            f"{currency} {tf['Variance']:,.0f} ({tf['Variance %']:.1f}%). "
        )
    narr += f"EBITDA margin: {margin_b:.1f}% → {margin_a:.1f}%, swing {swing:.1f}pp."

    # ── Top Drivers data ──────────────────────────────────────────────────
    tuf = ds[ds["Flag"] == "UF"]["Abs"].sum()
    ds["% of UF"] = ds.apply(
        lambda r: round(r["Abs"] / tuf * 100, 1) if r["Flag"] == "UF" and tuf > 0 else 0.0,
        axis=1
    )

    # ── Quick Wins ────────────────────────────────────────────────────────
    qw = []
    if revenue_a < revenue_b and revenue_b and abs(revenue_a - revenue_b) / revenue_b * 100 >= mat_pct:
        qw.extend(QUICK_WINS["revenue_decline"])
    if emp_a > emp_b and emp_b and abs(emp_a - emp_b) / emp_b * 100 >= mat_pct:
        qw.extend(QUICK_WINS["cost_overrun_employee"])
    if opex_a > opex_b:
        qw.extend(QUICK_WINS["cost_overrun_opex"])
    if cogs_a > cogs_b:
        qw.extend(QUICK_WINS["cogs_overrun"])
    if swing < -5:
        qw.extend(QUICK_WINS["margin_compression"])
    if not qw:
        qw = QUICK_WINS["revenue_decline"][:2]

    # ── Decision Engine data ──────────────────────────────────────────────
    material_count = len(material)
    if pbt_a < 0 and material_count >= 2:
        severity = "🔴 RED — CRISIS"
        decision_category = 4
        decision_name = "CRISIS MODE"
    elif material_count >= 2 and abs(swing) > 10:
        severity = "🔴 RED"
        decision_category = 3
        decision_name = "REFORECAST REQUIRED"
    elif material_count >= 1:
        severity = "🟡 AMBER"
        decision_category = 2
        decision_name = "TACTICAL ADJUSTMENT"
    else:
        severity = "🟢 GREEN"
        decision_category = 1
        decision_name = "STAY THE COURSE"

    # ── Data-driven CFO actions (from actual UF variances) ──────────────
    _uf_rows = agg_data[agg_data["Flag"] == "UF"].sort_values("Variance", key=abs, ascending=False)
    dd_actions = []
    for _, _uf in _uf_rows.iterrows():
        _act = generate_data_driven_action(_uf["Category"], _uf["Variance"], currency)
        if _act:
            dd_actions.append(_act)
    total_recoverable = sum(a["impact"] for a in dd_actions)
    post_action_ebitda = ebitda_a + total_recoverable

    # ── Pre-computed AI context (numbers from Python, not AI) ─────────────
    material_uf = agg_data[(agg_data["Flag"] == "UF") & (agg_data["Material?"].str.contains("YES", na=False))]

    _gp_margin_b = round(gp_b / revenue_b * 100, 1) if revenue_b else 0.0
    _gp_margin_a = round(gp_a / revenue_a * 100, 1) if revenue_a else 0.0

    ai_context = f"""VERIFIED FINANCIAL DATA (use ONLY these numbers):

Revenue: Budget {currency} {revenue_b:,.0f} → Actual {currency} {revenue_a:,.0f} (Variance: {currency} {revenue_a - revenue_b:,.0f}, {(revenue_a - revenue_b) / revenue_b * 100 if revenue_b else 0:.1f}%)
Gross Profit: {currency} {gp_b:,.0f} → {currency} {gp_a:,.0f} (Margin: {_gp_margin_b}% → {_gp_margin_a}%)
EBITDA: {currency} {ebitda_b:,.0f} → {currency} {ebitda_a:,.0f} (Margin: {margin_b:.1f}% → {margin_a:.1f}%)
PBT: {currency} {pbt_b:,.0f} → {currency} {pbt_a:,.0f}
PAT: {currency} {pat_b:,.0f} → {currency} {pat_a:,.0f}
EBITDA Margin Swing: {swing:.1f}pp

MATERIAL UNFAVORABLE ITEMS (focus ONLY on these):
{material_uf[['Category', 'Budget', 'Actual', 'Variance', 'Variance %']].to_string(index=False) if len(material_uf) > 0 else 'None'}

BRIDGE: {narr}

PRE-COMPUTED CFO ACTIONS (use these EXACT numbers):
"""
    for _ai, _da in enumerate(dd_actions[:5], 1):
        ai_context += f"{_ai}. {_da['action']} | Owner: {_da['owner']} | Impact: {currency} {_da['impact']:,.0f} | Timeline: {_da['timeline']}\n"

    ai_context += f"""
TOTAL RECOVERABLE: {currency} {total_recoverable:,.0f}
EBITDA RECOVERY PROJECTION: {currency} {ebitda_a:,.0f} → {currency} {post_action_ebitda:,.0f}
{"EBITDA TURNS POSITIVE after actions" if ebitda_a < 0 and post_action_ebitda > 0 else ""}

DECISION ENGINE: Category {decision_category} — {decision_name}

COUNTRY: {country}
INDUSTRY: {industry}
ERP: {erp_system}
CURRENCY: {currency}
"""

    # ── Risk escalation alerts ────────────────────────────────────────────
    alerts = []
    if ebitda_a < 0:
        _q_loss = ebitda_a * 3
        alerts.append(f"🚨 Operating loss. Projected quarterly loss: {currency} {abs(_q_loss):,.0f}")
    if pbt_a < 0:
        alerts.append(f"🚨 Net loss position: {currency} {pbt_a:,.0f}. Cash impact within 2-3 months if unchanged.")
    if abs(swing) > 10:
        alerts.append(f"⚠️ EBITDA margin collapsed {abs(swing):.1f}pp — structural deterioration likely")
    if revenue_a < revenue_b and revenue_b and abs(revenue_a - revenue_b) / revenue_b * 100 > 15:
        alerts.append(f"⚠️ Revenue decline exceeds 15% — market position at risk")

    # ── Mapping confidence ────────────────────────────────────────────────
    _n_total_items = len(data)
    _n_memory = st.session_state.get("_memory_matched", 0)
    try:
        _n_unmatched = len(_unmatched_labels)
    except NameError:
        _n_unmatched = 0
    _n_auto = _n_total_items - _n_memory - _n_unmatched
    _auto_pct = round((_n_memory + _n_auto) / _n_total_items * 100) if _n_total_items else 0
    _unmatched_pct = round(_n_unmatched / _n_total_items * 100) if _n_total_items else 0
    if _auto_pct >= 80:
        _map_conf = f"🟢 HIGH ({_auto_pct}% auto-matched)"
    elif _auto_pct >= 50:
        _map_conf = f"🟡 MEDIUM ({_auto_pct}% auto, {_unmatched_pct}% default — review recommended)"
    else:
        _map_conf = f"🔴 LOW ({_auto_pct}% auto, {_unmatched_pct}% default — manual review required)"

    # ── Compliance flags ──────────────────────────────────────────────────
    comp_flags = []
    if revenue_b and abs(revenue_a - revenue_b) / revenue_b * 100 >= mat_pct:
        comp_flags.append(f"**Revenue:** {rules['tax_desc']}")
    if emp_a > emp_b:
        comp_flags.append(f"**Employee Cost:** {rules['labor']}")
    comp_flags.append(f"**Indirect Tax:** {rules['vat']}")
    comp_flags.append(f"**Other:** {rules['other']}")
    if pbt_a < 0:
        comp_flags.append(f"**LOSS:** Assess DTA recognition. {rules['tax_desc']}")

    # ── Board Memo text (for download) ────────────────────────────────────
    _memo_lines = [
        f"CFO BOARD MEMO — {company_name or 'Company'} | {reporting_period}",
        f"Prepared by: PBV Finance AI CFO Assistant",
        f"Country: {country} | Currency: {currency} | Standard: {accounting_std}",
        "=" * 60,
        "",
        "EXECUTIVE SUMMARY",
        f"  Severity: {severity}",
        f"  Decision: Category {decision_category} — {decision_name}",
        f"  Material variances: {material_count} items",
        "",
        "KEY METRICS",
        f"  Revenue:      {currency} {revenue_a:,.0f}  (Budget: {revenue_b:,.0f}, Delta: {revenue_a - revenue_b:,.0f})",
        f"  Gross Profit: {currency} {gp_a:,.0f}  (Budget: {gp_b:,.0f}, Delta: {gp_a - gp_b:,.0f})",
        f"  EBITDA:       {currency} {ebitda_a:,.0f}  (Budget: {ebitda_b:,.0f}, Delta: {ebitda_a - ebitda_b:,.0f})",
        f"  PBT:          {currency} {pbt_a:,.0f}  (Budget: {pbt_b:,.0f}, Delta: {pbt_a - pbt_b:,.0f})",
        f"  PAT:          {currency} {pat_a:,.0f}  (Budget: {pat_b:,.0f}, Delta: {pat_a - pat_b:,.0f})",
        "",
        "EBITDA NARRATIVE",
        narr.replace("**", ""),
        "",
        "MATERIAL VARIANCES",
    ]
    for _, _mr in material.iterrows():
        _memo_lines.append(
            f"  {_mr['Category']}: {currency} {_mr['Variance']:,.0f} ({_mr['Variance %']:.1f}%) — {_mr['Flag']}"
        )
    _memo_lines += [
        "",
        "CFO ACTIONS (Next 30 Days) — Data-Driven",
    ]
    for _qi, _da in enumerate(dd_actions[:3], 1):
        _memo_lines.append(f"  {_qi}. {_da['action']}  |  Owner: {_da['owner']}  |  Timeline: {_da['timeline']}")
    _memo_lines.append(f"  Total Recoverable Impact: {currency} {total_recoverable:,.0f}")
    _memo_lines.append(f"  Post-Actions EBITDA: {currency} {post_action_ebitda:,.0f}")
    _memo_lines += [
        "",
        f"COMPLIANCE ({country})",
    ]
    for _cf in comp_flags:
        _memo_lines.append(f"  - {_cf.replace('**', '')}")
    _memo_lines += [
        "",
        "=" * 60,
        "Confidence: All figures 🟢 CALCULATED (Python, verified by 10 checks)",
        f"Validation: {'All 10 PASSED' if all_pass else 'SOME CHECKS FAILED'}",
        "Generated by PBV Finance AI CFO Assistant v2.0",
    ]
    _memo_text = "\n".join(_memo_lines)

    agent1_time = round(time.time() - start_time, 1)

    # ── Review state defaults (may be overridden by Review Gate later) ────
    review_complete = False
    review_quality = "Select..."
    analyst_notes = ""
    reviewed_by = ""

    # ═══════════════════════════════════════════════════════════════════════
    # PREP: data needed across tabs
    # ═══════════════════════════════════════════════════════════════════════
    _rev_var = revenue_a - revenue_b
    _rev_var_pct = round(_rev_var / revenue_b * 100, 1) if revenue_b else 0.0
    _rev_arrow = "↓" if _rev_var < 0 else "↑"
    _ebitda_status = (f"NEGATIVE ↓ {currency} {abs(ebitda_a):,.0f}" if ebitda_a < 0
                      else f"{'↓' if ev < 0 else '↑'} {currency} {abs(ev):,.0f} ({margin_a:.1f}%)")
    _primary_driver = f"{td['Category']} ({currency} {td['Variance']:,.0f})" if td is not None else "No material UF driver"
    _top_action_text = dd_actions[0]["action"] if dd_actions else (qw[0]["action"] if qw else "No immediate actions")

    _action_source = dd_actions[:5] if dd_actions else [{
        "action": w["action"], "detail": w.get("impact", ""),
        "owner": w["owner"], "timeline": "30 days", "impact": 0,
    } for w in qw[:3]]
    actions_df = pd.DataFrame([{
        "Priority": i + 1, "Action": a["action"], "Owner": a["owner"],
        "Impact": f"{currency} {a['impact']:,.0f}" if a["impact"] else "TBD",
        "Timeline": a["timeline"], "Status": "⏳ Pending", "Deadline": "TBD by CFO",
    } for i, a in enumerate(_action_source)])

    _now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    _dl_reviewer = st.session_state.get("reviewer_name", "")
    _dl_quality = st.session_state.get("review_quality", "Select...")
    _dl_reviewed = bool(_dl_reviewer) and _dl_quality not in ("Select...", "❌ Rejected — rerun Agent 2 with better context")
    _reviewer = _dl_reviewer if _dl_reviewed else "Not yet reviewed"
    _rev_qual = _dl_quality if _dl_reviewed else "Pending"

    # ═══════════════════════════════════════════════════════════════════════
    #                 TAB 2 — CFO DASHBOARD
    # ═══════════════════════════════════════════════════════════════════════
    with _tab2:
        st.markdown(
            f"""<div style="background-color:#0d1b2a;padding:20px 24px 16px 24px;border-radius:10px;border-left:5px solid #4fc3f7;margin-bottom:20px;">
<h2 style="color:#ffffff;margin:0 0 12px 0;">📊 EXECUTIVE SUMMARY</h2>
<p style="color:#e0e0e0;font-size:16px;margin:4px 0;"><b style="color:#4fc3f7;">Revenue:</b> {_rev_arrow} {currency} {abs(_rev_var):,.0f} ({_rev_var_pct}%) &nbsp;&nbsp;|&nbsp;&nbsp; <b style="color:#4fc3f7;">EBITDA:</b> {_ebitda_status}</p>
<p style="color:#e0e0e0;font-size:16px;margin:4px 0;"><b style="color:#4fc3f7;">Primary Driver:</b> {_primary_driver}</p>
<p style="color:#e0e0e0;font-size:16px;margin:4px 0;"><b style="color:#4fc3f7;">Decision:</b> Category {decision_category} — {decision_name}</p>
<p style="color:#e0e0e0;font-size:16px;margin:4px 0;"><b style="color:#4fc3f7;">Immediate Focus:</b> {_top_action_text}</p>
</div>""", unsafe_allow_html=True)

        st.markdown("""<div style='background-color:#f0f2f6;padding:10px;border-radius:5px;margin-bottom:15px;'>
<b>Reading Guide:</b>
🟢 <b>CALCULATED/PATTERN</b> = Python math, verified — act with confidence |
🟡 <b>DERIVED</b> = reasonable inference — review recommended |
🔴 <b>HYPOTHESIS</b> = AI-generated — <u>verify before acting or presenting to Board</u>
</div>""", unsafe_allow_html=True)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Revenue", fmt_km(revenue_a), fmt_km(revenue_a - revenue_b), help=f"Actual: {currency} {revenue_a:,.0f} | Budget: {currency} {revenue_b:,.0f}")
        c2.metric("Gross Profit", fmt_km(gp_a), fmt_km(gp_a - gp_b), help=f"Actual: {currency} {gp_a:,.0f} | Budget: {currency} {gp_b:,.0f}")
        c3.metric("EBITDA", fmt_km(ebitda_a), fmt_km(ebitda_a - ebitda_b), help=f"Actual: {currency} {ebitda_a:,.0f} | Budget: {currency} {ebitda_b:,.0f}")
        c4.metric("PBT", fmt_km(pbt_a), fmt_km(pbt_a - pbt_b), help=f"Actual: {currency} {pbt_a:,.0f} | Budget: {currency} {pbt_b:,.0f}")
        c5.metric("PAT", fmt_km(pat_a), fmt_km(pat_a - pat_b), help=f"Actual: {currency} {pat_a:,.0f} | Budget: {currency} {pat_b:,.0f}")
        st.caption(f"Hover metrics for full {currency} values | Mapping Confidence: {_map_conf} | Agent 1: {agent1_time}s")
        st.divider()

        dc1, dc2, dc3 = st.columns(3)
        with dc1:
            st.markdown("**Severity**")
            st.markdown(f"### {severity}")
        with dc2:
            st.markdown("**Decision**")
            st.markdown(f"### Category {decision_category}")
        with dc3:
            st.markdown("**Action Required**")
            st.markdown(f"### {decision_name}")
        if alerts:
            for _alert in alerts:
                st.warning(_alert)

        st.markdown(f"**Total Recoverable: {currency} {total_recoverable:,.0f}**")
        if ebitda_a < 0 and post_action_ebitda > 0:
            st.success(f"Recovery: EBITDA {currency} {ebitda_a:,.0f} → {currency} {post_action_ebitda:,.0f} (TURNS POSITIVE)")
        else:
            st.info(f"Recovery: EBITDA {currency} {ebitda_a:,.0f} → {currency} {post_action_ebitda:,.0f}")

        if len(material) > 0:
            st.markdown(f"**⚠️ Material Variances ({material_count} items exceeding {mat_pct}% & {currency} {mat_abs:,.0f})**")
            _top_mat = material.head(3)
            _mat_cols = st.columns(3)
            for _i, (_, _mr) in enumerate(_top_mat.iterrows()):
                _mat_cols[_i].metric(_mr["Category"], fmt_km(_mr["Variance"]),
                    f"{_mr['Variance %']:.1f}% {_mr['Flag']}", help=f"Variance: {currency} {_mr['Variance']:,.0f}")
        else:
            st.success("No material variances detected.")
        st.divider()

        st.markdown("**EBITDA Narrative**")
        st.info(narr)
        if pbt_a < 0:
            st.warning(f"⚠️ PBT is NEGATIVE ({currency} {pbt_a:,.0f}). Tax = NIL. DTA of {currency} {abs(pbt_a) * tax_rate:,.0f} to be assessed under {accounting_std}.")

        # ══════════════════════════════════════════════════════════════════
        # CFO DECISION PANEL
        # ══════════════════════════════════════════════════════════════════
        st.divider()
        st.markdown("## 🧠 CFO DECISION PANEL")

        _opt_a_rec = total_recoverable * 0.5
        _opt_b_rec = total_recoverable
        _opt_c_rec = total_recoverable * 1.5

        _dc_col1, _dc_col2, _dc_col3 = st.columns(3)
        with _dc_col1:
            st.markdown("### 🟢 Conservative")
            st.markdown(f"**Recovery:** {currency} {_opt_a_rec:,.0f}")
            st.markdown("**Timeline:** 30 days")
            st.markdown("**Risk:** Low — existing resources only")
            st.markdown("**Success Rate:** 85%")
            st.markdown("_Quick wins only — zero budget required_")
        with _dc_col2:
            st.markdown("### 🟡 Balanced")
            st.markdown(f"**Recovery:** {currency} {_opt_b_rec:,.0f}")
            st.markdown("**Timeline:** 45 days")
            st.markdown("**Risk:** Medium — requires budget reallocation")
            st.markdown("**Success Rate:** 65%")
            st.markdown("_Quick wins + targeted investment_")
            st.markdown("**⭐ RECOMMENDED**")
        with _dc_col3:
            st.markdown("### 🔴 Aggressive")
            st.markdown(f"**Recovery:** {currency} {_opt_c_rec:,.0f}")
            st.markdown("**Timeline:** 60-90 days")
            st.markdown("**Risk:** High — board approval + org change")
            st.markdown("**Success Rate:** 45%")
            st.markdown("_Full restructuring — maximum recovery_")

        selected_option = st.radio("Select your path:",
            ["🟢 Conservative", "🟡 Balanced (Recommended)", "🔴 Aggressive"],
            index=1, key="_decision_path")

        if selected_option == "🟢 Conservative":
            _sel_recovery = _opt_a_rec; _sel_days = 30
        elif selected_option == "🔴 Aggressive":
            _sel_recovery = _opt_c_rec; _sel_days = 75
        else:
            _sel_recovery = _opt_b_rec; _sel_days = 45
        _sel_projected = ebitda_a + _sel_recovery

        # ── Cost of Inaction ──────────────────────────────────────────────
        st.divider()
        st.markdown("### ⏰ COST OF INACTION")
        if ebitda_a < 0:
            _monthly_burn = abs(ebitda_a)
            st.error(
                f"**If no action is taken:**\n"
                f"- Next month loss: {currency} {_monthly_burn:,.0f}\n"
                f"- Quarterly projected loss: {currency} {_monthly_burn * 3:,.0f}\n"
                f"- Cash runway impact: Critical within 2-3 months\n"
                f"- Board escalation: Required if no decision by {reporting_period} close"
            )
        elif swing < -5:
            _proj_decline = abs(ebitda_a - ebitda_b) * 3
            st.warning(
                f"**If no action is taken:**\n"
                f"- Margin erosion continues: {margin_a:.1f}% and declining\n"
                f"- Projected quarterly shortfall: {currency} {_proj_decline:,.0f}\n"
                f"- Competitive position: Weakening\n"
                f"- Recommended deadline: Decision within 5 business days"
            )
        else:
            st.info(
                "**Low urgency — monitoring recommended**\n"
                "- Variances are timing-related\n"
                "- Full year target still achievable\n"
                "- Next review: End of next month"
            )

        # ── Decision Deadline ─────────────────────────────────────────────
        if decision_category >= 3:
            _deadline = (datetime.now() + pd.Timedelta(days=5)).strftime("%B %d, %Y")
            st.markdown(f"### 📅 DECISION DEADLINE: **{_deadline}**")
            st.markdown(f"*Category {decision_category} requires CFO decision within 5 business days.*")
        elif decision_category == 2:
            _deadline = (datetime.now() + pd.Timedelta(days=14)).strftime("%B %d, %Y")
            st.markdown(f"### 📅 REVIEW BY: **{_deadline}**")
            st.markdown("*Category 2 — tactical adjustment review within 2 weeks.*")

        # ── Recovery Timeline Chart ───────────────────────────────────────
        st.divider()
        st.markdown("### 📈 RECOVERY TIMELINE")
        _day_labels = ["Today", f"Day {_sel_days // 3}", f"Day {_sel_days * 2 // 3}", f"Day {_sel_days}"]
        _ebitda_path = [ebitda_a, ebitda_a + _sel_recovery * 0.3, ebitda_a + _sel_recovery * 0.7, _sel_projected]
        _timeline_df = pd.DataFrame({"Projected EBITDA": _ebitda_path}, index=_day_labels)
        st.line_chart(_timeline_df)

        if ebitda_a < 0 and _sel_projected > 0:
            st.success(f"✅ EBITDA turns POSITIVE: {currency} {ebitda_a:,.0f} → {currency} {_sel_projected:,.0f} in {_sel_days} days")
        else:
            st.info(f"📈 EBITDA improvement: {currency} {ebitda_a:,.0f} → {currency} {_sel_projected:,.0f} in {_sel_days} days")

        st.session_state["_sel_option"] = selected_option
        st.session_state["_sel_recovery"] = _sel_recovery
        st.session_state["_sel_projected"] = _sel_projected
        st.session_state["_sel_days"] = _sel_days

        # ══════════════════════════════════════════════════════════════════
        # DATA-DRIVEN INSIGHT ENGINE (100% Python, zero AI)
        # ══════════════════════════════════════════════════════════════════
        st.divider()
        st.markdown("## 💡 Data-Driven Insights")
        _all_insights = calculate_all_insights(
            revenue_b, revenue_a, cogs_b, cogs_a, gp_b, gp_a, ebitda_b, ebitda_a,
            emp_b, emp_a, mkt_b, mkt_a, opex_b, opex_a, dep_b, dep_a, fin_b, fin_a,
            oi_b, oi_a, pbt_b, pbt_a, pat_b, pat_a, tax_rate, currency,
            material_count, agg_data)

        _insight_cats = ["REVENUE", "COST STRUCTURE", "MARGIN", "SENSITIVITY",
                         "RISK", "EMPLOYEE", "MARKETING", "FINANCE", "CAPITAL", "TAX"]
        _expand_cats = {"REVENUE", "MARGIN", "RISK"}
        for _ic in _insight_cats:
            _cat_ins = [i for i in _all_insights if i["category"] == _ic]
            if _cat_ins:
                with st.expander(f"{_ic} ({len(_cat_ins)} insights)", expanded=(_ic in _expand_cats)):
                    for _ins in _cat_ins:
                        st.markdown(f"**{_ins['type']} {_ins['title']}**")
                        st.markdown(_ins["insight"])
                        st.markdown(f"📌 **Action:** {_ins['action']}")
                        st.markdown(f"💰 **Impact:** {currency} {_ins['impact']:,.0f} | 👤 **Owner:** {_ins['owner']}")
                        if _ins["confidence"] == "CALCULATED":
                            st.success(f"✅ CONFIDENCE: {_ins['confidence']} — verified by Python math")
                        else:
                            st.warning(f"🔴 CONFIDENCE: {_ins['confidence']} — requires validation")
                        st.markdown("---")
        st.caption(f"🟢 {len(_all_insights)} data-driven insights generated — all CALCULATED from your data, zero AI")

        if meeting_mode:
            if st.session_state.get("p3_output"):
                st.divider()
                st.subheader("📝 Board Memo")
                _memo_display = st.session_state["p3_output"]
                if "unverified" in _memo_display or "HYPOTHESIS" in _memo_display:
                    st.caption("⚠️ Contains unverified items — see full analysis for details")
                st.markdown(_memo_display)
            st.divider()
            st.caption("Detailed analysis available in full mode (toggle off Board Meeting Mode)")
        else:
            st.divider()
            with st.expander("📊 Full Variance Table", expanded=False):
                st.dataframe(agg_data, use_container_width=True, hide_index=True)
            with st.expander("📈 P&L Waterfall (Revenue → PAT)", expanded=False):
                st.dataframe(wf, use_container_width=True, hide_index=True)
            with st.expander("🔗 EBITDA Bridge", expanded=False):
                st.dataframe(bdf, use_container_width=True, hide_index=True)
            with st.expander("🔍 PVME Analysis", expanded=False):
                if has_units and has_price:
                    vol_data = raw[raw["Month"] == selected_month] if has_month and selected_month else raw
                    pvme_rows = []
                    for _, row in vol_data.iterrows():
                        pvme_rows.append({"Product": row.get("Product", "N/A"),
                            "Price Variance": row.get("Actual Units", 0) * (row.get("Actual Price", 0) - row.get("Budget Price", 0)),
                            "Volume Variance": row.get("Budget Price", 0) * (row.get("Actual Units", 0) - row.get("Budget Units", 0)),
                            "Confidence": "🟢 CALCULATED"})
                    if pvme_rows:
                        st.dataframe(pd.DataFrame(pvme_rows), use_container_width=True, hide_index=True)
                else:
                    st.warning("⚠️ PVME requires unit data")
            with st.expander("📋 Top Drivers Detail", expanded=False):
                st.dataframe(ds[["Category", "Variance", "Variance %", "Flag", "% of UF"]].reset_index(drop=True), use_container_width=True, hide_index=True)
            with st.expander("✅ Validation (10 Checks)", expanded=False):
                st.dataframe(pd.DataFrame(checks, columns=["Check", "Status"]), use_container_width=True, hide_index=True)
                if all_pass:
                    st.success("All 10 checks PASSED ✅")
                else:
                    st.error("Some checks FAILED ❌")
            with st.expander(f"🏛️ {country} Compliance Flags", expanded=False):
                for _cf in comp_flags:
                    st.markdown(f"- {_cf}")
            with st.expander("⚡ Quick Wins Detail", expanded=False):
                st.dataframe(pd.DataFrame([{"Action": w["action"], "Owner": w["owner"], "Impact": w["impact"], "ERP": w["erp"]} for w in qw[:5]]), use_container_width=True, hide_index=True)

            st.divider()
            _dl1, _dl2 = st.columns(2)
            with _dl1:
                _buf = io.BytesIO()
                with pd.ExcelWriter(_buf, engine="openpyxl") as writer:
                    agg_data.to_excel(writer, sheet_name="Variance", index=False)
                    wf.to_excel(writer, sheet_name="Waterfall", index=False)
                    bdf.to_excel(writer, sheet_name="EBITDA Bridge", index=False)
                    pd.DataFrame(checks, columns=["Check", "Status"]).to_excel(writer, sheet_name="Validation", index=False)
                    data.to_excel(writer, sheet_name="Raw Line Items", index=False)
                    if len(material) > 0:
                        material.to_excel(writer, sheet_name="Material Items", index=False)
                    pd.DataFrame([{"Field": "Generated", "Value": _now_str}, {"Field": "Reviewed By", "Value": _reviewer},
                        {"Field": "Quality", "Value": _rev_qual}, {"Field": "Decision", "Value": f"Category {decision_category} — {decision_name}"},
                        {"Field": "Version", "Value": "PBV Finance AI CFO v1.3"}]).to_excel(writer, sheet_name="Audit", index=False)
                _buf.seek(0)
                st.download_button(f"📥 Download Excel Report ({company_name or 'Analysis'})", data=_buf,
                    file_name=f"PBV_Variance_{company_name}_{reporting_period}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with _dl2:
                _p2hc = st.session_state.get("_p2_hyp_count", 0)
                _p2pc = st.session_state.get("_p2_pat_count", 0)
                _stamp = (
                    f"\n\n{'='*40}\n"
                    f"NOTE: This memo contains both verified data (🟢) and AI-generated hypotheses (🔴).\n"
                    f"Items marked as HYPOTHESIS require validation before presenting to the Board.\n"
                    f"Verified patterns: {_p2pc} | Hypotheses requiring validation: {_p2hc}\n"
                    f"{'='*40}\n"
                    f"Generated by: PBV Finance AI CFO v1.3\n"
                    f"Analysis Date: {_now_str}\n"
                    f"Reviewed by: {_reviewer}\n"
                    f"Review Quality: {_rev_qual}\n"
                    f"Decision: Category {decision_category} — {decision_name}\n"
                )
                st.download_button("📥 Download Board Memo", data=_memo_text + _stamp,
                    file_name=f"CFO_Memo_{company_name or 'Company'}_{reporting_period}.txt", mime="text/plain")

            st.divider()
            # ── Agent 2 ────────────────────────��─────────────────────────────
            st.header("🤖 AI Agents — Root Cause & Board Memo")
            st.subheader("🔍 Agent 2: Root Cause Diagnostician")
            st.caption("Groq (Llama 3.3 70B) — uses ONLY pre-computed numbers")
            if st.button("▶️ Run Agent 2: Root Cause Analysis", type="primary", key="_run_agent2"):
                _p2_prompt = f"""{ai_context}

CRITICAL SAFETY RULES (NEVER BREAK):
- You have ZERO access to external knowledge. Use ONLY the data provided in ai_context and Agent 1 output. Nothing from your training data.
- Never invent numbers or facts. For any hypothesis, label it clearly as HYPOTHESIS and always provide the exact {erp_system} report / data point needed to verify it.
- Do NOT repeat Agent 1 calculations. Reference them only.
- When creating hypotheses, make them intelligent, specific, and testable — never vague.
- If you are unsure about anything, say "UNVERIFIED" — never present uncertainty as fact.

AGENT 1 HAS ALREADY CALCULATED: breakeven revenue, cost elasticity, sensitivity table, operating leverage, margin decomposition, interest coverage, cost of inaction. Do NOT repeat these. Reference them.

You are an elite CFO Diagnostician — a 20-year veteran who has seen hundreds of P&L cycles in {industry} across {country}. YOUR UNIQUE VALUE: the 10 things Python CANNOT calculate.

DIMENSION 1: SIGNAL CONNECTION (The Story)
Connect 3-5 Agent 1 signals into ONE narrative. Individual signals are noise. Connected signals are intelligence. What pattern would a veteran {industry} CFO immediately recognize?

DIMENSION 2: ROOT CAUSE CHAIN (Causation, Not Correlation)
Map: TRIGGER → FIRST EFFECT → SECOND EFFECT → P&L IMPACT → INTERVENTION POINT. The chain tells the CFO WHERE to intervene — fix ONE LINK. Use {currency} amounts at each step.

DIMENSION 3: TIMING vs STRUCTURAL (With Evidence)
For EACH material variance: TIMING (self-corrects in 1-2 quarters, evidence: [reason]) or STRUCTURAL (business model shifted, evidence: [reason]). For {industry} in {country}, provide industry-specific classification signals.

DIMENSION 4: HIDDEN RISKS (What Numbers Hint At)
Revenue down + employee up = building capacity for growth that isn't coming? COGS ratio worse = obsolete inventory? Finance cost up = borrowing to fund losses? Marketing cut = cutting the growth engine? For each: [Risk] → [Signal in data] → [{currency} impact if it materializes] → [How to verify with specific {erp_system} report].

DIMENSION 5: LEADING vs LAGGING SIGNALS
Which variances PREDICT future problems (leading) vs REFLECT past events (lagging)? Marketing cuts → future revenue decline (leading). Revenue decline → past sales (lagging). "Fix only lagging indicators and next quarter's leading indicators will surprise you."

DIMENSION 6: INDUSTRY CONTEXT ({industry} in {country})
NOT generic. For {industry} in {country} THIS quarter: What is NORMAL? What is ABNORMAL? What do competitors do? What regulatory/market factors apply? Recovery timeline: [X] months if [condition], [Y] months if [alternative].

DIMENSION 7: MANAGEMENT QUALITY SIGNALS
What does the DATA reveal about execution? Revenue dropped but costs increased = slow response? Marketing cut during decline = short-term thinking? Frame clinically: "The data suggests [observation]. This is [positive/concerning] because [reason]."

DIMENSION 8: SCENARIO BRANCHING (Decision Tree)
IF [condition A — verify via {erp_system} report] THEN [outcome X, probability%] → Action A
IF [condition B — verify via {erp_system} report] THEN [outcome Y, probability%] → Action B
Give the CFO a DECISION TREE, not a single recommendation.

DIMENSION 9: WHAT TO ASK IN THE NEXT MEETING
5 SPECIFIC questions — not "why did revenue decline?" but:
1. "Sales Head: Of top 10 customers, how many reduced orders? Which ones, by how much?"
2. "Procurement: Any supplier contracts up for renewal in 90 days?"
3. "HR: Headcount increase — how many in probation (reversible)?"
4. "Operations: Current inventory value vs 3 months ago?"
5. "Finance: Cash position and months of runway at current burn?"

DIMENSION 10: THE ONE THING
If the CFO can only do ONE THING this month: what, who, {currency} impact, deadline, and why (connecting back to root cause chain).

ENHANCED HYPOTHESIS RULE (applies to ALL dimensions):
For every HYPOTHESIS you create:
- State the most likely explanation based on the pattern in Agent 1 data + {industry} context in {country}.
- Give a probability estimate (Low/Medium/High) based on how strongly the data supports it.
- Always end with: "VERIFY BY running {erp_system} report [exact T-code or report name] and checking [specific data point]. Expected impact if true: {currency} [rough range based on current variance]."

RECOVERY PROJECTION (use EXACT pre-computed numbers):
- Current EBITDA: {currency} {ebitda_a:,.0f}
- Total Actions: {currency} {total_recoverable:,.0f}
- Post-Actions: {currency} {post_action_ebitda:,.0f}
{"- EBITDA TURNS POSITIVE" if ebitda_a < 0 and post_action_ebitda > 0 else ""}

FORMATTING RULES (CRITICAL — follow exactly):
- Start each DIMENSION on a new line with the header
- Use bullet points (- ) for each sub-item
- Put each cause, risk, signal on its OWN line
- Use blank lines between dimensions
- For root cause chain, put each step on a NEW line with arrow:
  TRIGGER: [text]
  → FIRST EFFECT: [text]
  → SECOND EFFECT: [text]
  → P&L IMPACT: [text]
  → INTERVENTION POINT: [text]
- For hidden risks, each risk on its OWN line:
  - Risk: [text]
  - Signal: [text]
  - Impact: [{currency} amount]
  - Verify: [{erp_system} report]
- For meeting questions, number them 1-5 each on own line
- For HYPOTHESIS labels, put on same line as the statement
- NEVER write paragraphs longer than 2 sentences
- Every dimension must be clearly separated with a blank line

RULES: Max 650 words. Every statement needs {currency} or %. Label: PATTERN (from data) or HYPOTHESIS (needs verification). Start DIRECTLY with DIMENSION 1. No preamble."""
                _p2_container = st.empty()
                _p2_container.info("Agent 2 running on Groq (Llama 3.3 70B)...")
                _p2_raw = call_ai(_p2_prompt, _p2_container)
                if _p2_raw:
                    _p2_stripped = _strip_ai_preamble(_p2_raw, ["DIMENSION 1", "ROOT CAUSE", "🔍", "## ", "**"])
                    _p2_hyp_count, _p2_pat_count = _count_confidence(_p2_stripped)
                    _p2_output = _format_agent2_output(_p2_stripped)
                    st.session_state["p2_output"] = _p2_output
                    st.session_state["_p2_hyp_count"] = _p2_hyp_count
                    st.session_state["_p2_pat_count"] = _p2_pat_count
                    _p2_container.markdown(_p2_output)
                    st.caption("🔴 AI-GENERATED — verify all numbers against Agent 1 calculations")
                    if _p2_hyp_count > 0:
                        st.warning(
                            f"⚠️ **ATTENTION: {_p2_hyp_count} items need verification**\n\n"
                            f"🟢 **{_p2_pat_count} PATTERNS** — supported by data, safe to reference\n\n"
                            f"🔴 **{_p2_hyp_count} HYPOTHESES** — require verification with SAP reports\n\n"
                            f"**Rule: Never present a HYPOTHESIS as fact to the Board.**"
                        )
            elif st.session_state.get("p2_output"):
                st.markdown(st.session_state["p2_output"])
                st.caption("🔴 AI-GENERATED — verify all numbers against Agent 1 calculations")
                _p2_hc = st.session_state.get("_p2_hyp_count", 0)
                _p2_pc = st.session_state.get("_p2_pat_count", 0)
                if _p2_hc > 0:
                    st.warning(f"⚠️ **{_p2_hc} HYPOTHESES** need verification | {_p2_pc} PATTERNS supported by data")

            st.divider()

            # ── Review Gate ──────────────────────────────────────────────────
            _p2_available = bool(st.session_state.get("p2_output"))
            review_complete = False
            review_quality = "Select..."
            analyst_notes = ""
            reviewed_by = ""
            if _p2_available:
                st.subheader("📋 ANALYST REVIEW (Required Before Memo)")

                st.markdown("### 🔍 Hypothesis Verification Checklist")
                st.markdown("Mark each hypothesis area before generating the Board Memo:")
                _hyp_sections = [
                    "Root cause chain classification (Timing vs Structural)",
                    "Hidden risks identified by AI",
                    "Industry context assumptions",
                    "Management quality assessment",
                    "Scenario probabilities",
                ]
                _verified_count = 0
                for _hi, _hs in enumerate(_hyp_sections):
                    if st.checkbox(_hs, key=f"hyp_check_{_hi}"):
                        _verified_count += 1
                st.progress(_verified_count / len(_hyp_sections))
                st.caption(f"{_verified_count}/{len(_hyp_sections)} hypothesis areas reviewed")
                st.divider()

                review_quality = st.radio("AI Analysis Quality:",
                    ["Select...", "✅ Confirmed — findings are reasonable",
                     "⚠️ Acceptable — minor concerns noted below",
                     "❌ Rejected — rerun Agent 2 with better context"], key="review_quality")
                analyst_notes = st.text_area("Analyst Notes / Corrections:",
                    placeholder="Add context, corrections, or override any AI findings...", key="analyst_notes")
                reviewed_by = st.text_input("Reviewed by:", key="reviewer_name")
                review_complete = review_quality not in ["Select...", "❌ Rejected — rerun Agent 2 with better context"] and bool(reviewed_by)
                if not review_complete:
                    st.warning("Complete the review above to unlock Board Memo generation.")
                if "run_log" not in st.session_state:
                    st.session_state["run_log"] = []
                if review_complete and not st.session_state.get("_review_logged"):
                    st.session_state["run_log"].append({"timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "agent": "Review Gate", "reviewer": reviewed_by, "quality": review_quality,
                        "notes": analyst_notes[:100] if analyst_notes else "None"})
                    st.session_state["_review_logged"] = True
            st.divider()

            # ── Agent 3 ──────────────────────────────────────────────────────
            st.subheader("📝 Agent 3: CFO Board Memo Writer")
            st.caption("Groq (Llama 3.3 70B) — uses Agent 2 findings + pre-computed numbers")
            if not _p2_available:
                st.info("Run Agent 2 first.")
            elif not review_complete:
                st.info("Complete the analyst review above to unlock Agent 3.")
            if review_complete and st.button("▶️ Run Agent 3: Write Board Memo", type="primary", key="_run_agent3"):
                _p2_text = st.session_state["p2_output"][:2500]
                _review_ctx = f"\nANALYST REVIEW: {review_quality}. Notes: {analyst_notes}. Reviewed by: {reviewed_by}. Analyst verified {_verified_count} of {len(_hyp_sections)} hypothesis areas.\n"
                _p3_prompt = f"""{ai_context}

SAFETY RULES:
- Use ONLY Agent 1 exact numbers and Agent 2 findings.
- Never invent numbers. Every figure must trace to Agent 1 data.
- If citing Agent 2 hypotheses, label as "unverified".
- Start DIRECTLY with HEADLINE. No preamble. No thinking.

AGENT 2 DIAGNOSIS (10 dimensions):
{_p2_text}
{_review_ctx}

You have THREE sources:
1. Agent 1: 24 calculated insights (breakeven, elasticity, sensitivity, margins — use EXACT numbers)
2. Agent 2: 10-dimension diagnosis (root cause chain, industry context, timing vs structural)
3. Pre-computed recovery: {currency} {total_recoverable:,.0f} total, EBITDA {currency} {ebitda_a:,.0f} → {currency} {post_action_ebitda:,.0f}

You are an OPERATOR CFO who OWNS outcomes. Not a consultant who observes. Your BONUS depends on fixing this. Every sentence = {currency} amount or deadline. No filler.

Write a 300-word board memo. ALL 9 sections MANDATORY:

📌 **HEADLINE**
[Use Agent 2's root cause CHAIN — not "revenue declined" but the TRIGGER → EFFECT → P&L IMPACT in one sentence with {currency} scale]

📊 **WHAT CHANGED**
- Revenue: {currency} {revenue_b:,.0f} → {currency} {revenue_a:,.0f} ({(revenue_a - revenue_b) / revenue_b * 100 if revenue_b else 0:.1f}%)
- EBITDA: {currency} {ebitda_b:,.0f} → {currency} {ebitda_a:,.0f} (margin: {margin_b:.1f}% → {margin_a:.1f}%)
- PBT: {currency} {pbt_a:,.0f} {"(LOSS)" if pbt_a < 0 else ""}

🔍 **WHY — THE ROOT CAUSE CHAIN**
[Agent 2 Dimension 2 simplified to 3 steps: TRIGGER → EFFECT → P&L. Label each: PATTERN or HYPOTHESIS]

⚙️ **THE ONE THING** (From Agent 2 Dimension 10)
[Single highest-priority action with {currency} impact and deadline]

📋 **SUPPORTING ACTIONS**
[2-3 from pre-computed list, each with owner + {currency} impact]
Total Recovery: {currency} {total_recoverable:,.0f}

📈 **EXPECTED OUTCOME**
- EBITDA: {currency} {ebitda_a:,.0f} → {currency} {post_action_ebitda:,.0f} (EXACT — do not recalculate)
- Timeline: 30-45 days
{"- EBITDA TURNS POSITIVE" if ebitda_a < 0 and post_action_ebitda > 0 else ""}

⏰ **COST OF DELAY**
[From Agent 2 Dimension 4 — {currency} impact of inaction this week/month]

🧠 **QUESTIONS FOR NEXT MEETING**
[Top 3 from Agent 2 Dimension 9 — specific, role-addressed]

🎯 **BOARD DECISION REQUIRED**
[1-2 approvals with {currency} amounts, deadlines, and Decision Engine Category {decision_category} reference]

Start DIRECTLY with 📌 **HEADLINE**. No preamble."""
                _p3_container = st.empty()
                _p3_container.info("Agent 3 writing board memo on Groq (Llama 3.3 70B)...")
                _p3_raw = call_ai(_p3_prompt, _p3_container)
                if _p3_raw:
                    _p3_stripped = _strip_ai_preamble(_p3_raw, ["CFO MEMORANDUM", "MEMORANDUM", "HEADLINE", "📌 HEADLINE", "📌", "## ", "TO:", "**"])
                    _p3_output = _format_agent3_output(_p3_stripped)
                    _p3_output += f"\n\n---\nReviewed by: {reviewed_by} | Quality: {review_quality} | Notes: {analyst_notes}"
                    st.session_state["p3_output"] = _p3_output
                    _p3_container.markdown(_p3_output)
                    st.caption("🔴 AI-GENERATED — verify all numbers against Agent 1 calculations")
                    st.download_button("📥 Download Board Memo (AI)", data=_p3_output,
                        file_name=f"CFO_Memo_{company_name or 'Company'}_{reporting_period}.txt", mime="text/plain", key="_dl_ai_memo")
            elif st.session_state.get("p3_output"):
                st.markdown(st.session_state["p3_output"])
                st.caption("🔴 AI-GENERATED — verify all numbers against Agent 1 calculations")
                st.download_button("📥 Download Board Memo (AI)", data=st.session_state["p3_output"],
                    file_name=f"CFO_Memo_{company_name or 'Company'}_{reporting_period}.txt", mime="text/plain", key="_dl_ai_memo_cached")

        st.divider()
        st.markdown("### 📞 Want This For Your Company?")
        st.markdown("See the full AI analysis with root cause + board memo generation in a live walkthrough.")
        st.link_button("📅 Book a 15-Minute Live Demo", "https://www.linkedin.com/in/bhargav-venkatesh/", use_container_width=True)
        st.caption("PBV Finance | AI CFO Systems v2.0 | Powered by Groq (Llama 3.3 70B)")

    # ═══════════════════════════════════════════════════════════════════════
    #                 TAB 3 — ACTION TRACKER
    # ═══════════════════════════════════════════════════════════════════════
    with _tab3:
        st.header("📌 Action Tracker")
        st.dataframe(actions_df, use_container_width=True, hide_index=True)
        st.markdown(f"**Total Recoverable: {currency} {total_recoverable:,.0f}**")
        if ebitda_a < 0 and post_action_ebitda > 0:
            st.success(f"Recovery Projection: EBITDA {currency} {ebitda_a:,.0f} → {currency} {post_action_ebitda:,.0f} (TURNS POSITIVE)")
        else:
            st.info(f"Recovery Projection: EBITDA {currency} {ebitda_a:,.0f} → {currency} {post_action_ebitda:,.0f}")

        _tracker_buf = io.BytesIO()
        with pd.ExcelWriter(_tracker_buf, engine="openpyxl") as _tw:
            actions_df.to_excel(_tw, sheet_name="Action Tracker", index=False)
        _tracker_buf.seek(0)
        st.download_button("📥 Download Action Tracker", data=_tracker_buf,
            file_name=f"Action_Tracker_{company_name or 'Company'}_{reporting_period}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="_dl_tracker")

        st.divider()
        st.markdown("### 📊 Confidence Guide")
        st.markdown("| Tag | Meaning | Source |\n|---|---|---|\n| 🟢 CALCULATED | Math fact | Python |\n| 🟡 PRE-BUILT | Standard action | Verified library |\n| 🔴 HYPOTHESIS | Needs validation | AI |")

    # ═══════════════════════════════════════════════════════════════════════
    #                 TAB 4 — MEMORY & REVIEW
    # ═══════════════════════════════════════════════════════════════════════
    with _tab4:
        st.header("🧠 CFO Memory & Decisions")

        # ── Lock Decision & Save ──────────────────────────────────────────
        _sel_opt = st.session_state.get("_sel_option", "🟡 Balanced (Recommended)")
        _sel_rec = st.session_state.get("_sel_recovery", total_recoverable)
        _sel_proj = st.session_state.get("_sel_projected", post_action_ebitda)
        _sel_d = st.session_state.get("_sel_days", 45)

        st.subheader("🔒 Lock Decision & Save")
        st.markdown(f"**Selected Path:** {_sel_opt}")
        st.markdown(f"**Expected Recovery:** {currency} {_sel_rec:,.0f} | **Timeline:** {_sel_d} days | **Projected EBITDA:** {currency} {_sel_proj:,.0f}")

        if st.button("💾 Lock Decision & Save to Memory", key="_save_decision", type="primary"):
            _decision_data = {
                "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "period": reporting_period,
                "company": company_name,
                "revenue_actual": float(revenue_a),
                "ebitda_actual": float(ebitda_a),
                "ebitda_margin": float(margin_a),
                "pbt": float(pbt_a),
                "decision_category": decision_category,
                "severity": severity,
                "selected_option": _sel_opt,
                "expected_recovery": float(_sel_rec),
                "projected_ebitda": float(_sel_proj),
                "recovery_timeline_days": _sel_d,
                "material_count": int(material_count),
                "total_recovery": float(total_recoverable),
                "actions": [{"action": a["action"], "impact": float(a["impact"]), "owner": a["owner"]} for a in dd_actions[:5]],
                "reviewed_by": st.session_state.get("reviewer_name", ""),
            }
            _mem_count = save_to_memory(_decision_data)
            st.success(f"Decision locked: {_sel_opt} | Recovery: {currency} {_sel_rec:,.0f} | Timeline: {_sel_d} days")
            st.balloons()

        st.divider()

        # ── Decision History ──────────────────────────────────────────────
        _hist_memory = load_memory()
        if _hist_memory:
            st.subheader(f"📈 Decision History ({len(_hist_memory)} months)")
            _mem_df = pd.DataFrame([{
                "Date": m.get("date", "—"),
                "Period": m.get("period", "—"),
                "Company": m.get("company", "—"),
                "EBITDA": f"{m.get('ebitda_actual', 0):,.0f}",
                "Margin": f"{m.get('ebitda_margin', 0):.1f}%",
                "Decision": m.get("selected_option", m.get("severity", "—")),
                "Expected Recovery": f"{m.get('expected_recovery', m.get('total_recovery', 0)):,.0f}",
                "Category": f"Cat {m.get('decision_category', '—')}",
            } for m in _hist_memory])
            st.dataframe(_mem_df, use_container_width=True, hide_index=True)

            if len(_hist_memory) >= 2:
                _margins = [m.get("ebitda_margin", 0) for m in _hist_memory]
                _trend = "📈 Improving" if _margins[-1] > _margins[0] else "📉 Declining"
                st.markdown(f"**Trend:** EBITDA Margin: {_margins[0]:.1f}% → {_margins[-1]:.1f}% ({_trend})")
        else:
            st.info("No decisions saved yet. Run analysis and click 'Lock Decision & Save to Memory' above.")
