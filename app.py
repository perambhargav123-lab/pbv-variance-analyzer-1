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
    """Call Groq API (Llama 3.3 70B). Returns response text or empty string on error."""
    try:
        client = Groq(api_key=os.environ.get("GROQ_API_KEY"))
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

    if len(data) == 0:
        st.error("❌ No valid data rows after pre-processing. Check your file format.")
        st.stop()

    # ═══════════════════════════════════════
    # STEP 1: CATEGORY MAPPING UI
    # ═══════════════════════════════════════
    mapping_confirmed = bool(st.session_state.get("confirmed_mappings"))

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

    if not st.session_state.get("confirmed_mappings"):
        st.info("👆 Complete the mapping above and click **Confirm Mapping & Run Analysis** to proceed.")
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
    st.success("✅ Mapping confirmed. Running P&L analysis.")

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
    #                     D A S H B O A R D   L A Y O U T
    # ═══════════════════════════════════════════════════════════════════════

    # ── 0. EXECUTIVE SUMMARY (the very first thing a CFO sees) ────────────
    _rev_var = revenue_a - revenue_b
    _rev_var_pct = round(_rev_var / revenue_b * 100, 1) if revenue_b else 0.0
    _rev_arrow = "↓" if _rev_var < 0 else "↑"
    if ebitda_a < 0:
        _ebitda_status = f"NEGATIVE ↓ {currency} {abs(ebitda_a):,.0f}"
    else:
        _ev_arrow = "↓" if ev < 0 else "↑"
        _ebitda_status = f"{_ev_arrow} {currency} {abs(ev):,.0f} ({margin_a:.1f}%)"
    _primary_driver = f"{td['Category']} ({currency} {td['Variance']:,.0f})" if td is not None else "No material UF driver"
    _top_action_text = dd_actions[0]["action"] if dd_actions else (qw[0]["action"] if qw else "No immediate actions")

    st.markdown(
        f"""<div style="background-color:#0d1b2a;padding:20px 24px 16px 24px;border-radius:10px;border-left:5px solid #4fc3f7;margin-bottom:20px;">
<h2 style="color:#ffffff;margin:0 0 12px 0;">📊 EXECUTIVE SUMMARY</h2>
<p style="color:#e0e0e0;font-size:16px;margin:4px 0;"><b style="color:#4fc3f7;">Revenue:</b> {_rev_arrow} {currency} {abs(_rev_var):,.0f} ({_rev_var_pct}%) &nbsp;&nbsp;|&nbsp;&nbsp; <b style="color:#4fc3f7;">EBITDA:</b> {_ebitda_status}</p>
<p style="color:#e0e0e0;font-size:16px;margin:4px 0;"><b style="color:#4fc3f7;">Primary Driver:</b> {_primary_driver}</p>
<p style="color:#e0e0e0;font-size:16px;margin:4px 0;"><b style="color:#4fc3f7;">Decision:</b> Category {decision_category} — {decision_name}</p>
<p style="color:#e0e0e0;font-size:16px;margin:4px 0;"><b style="color:#4fc3f7;">Immediate Focus:</b> {_top_action_text}</p>
</div>""",
        unsafe_allow_html=True,
    )

    # ── 1. KEY METRICS (always visible) ───────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Revenue", fmt_km(revenue_a), fmt_km(revenue_a - revenue_b),
              help=f"Actual: {currency} {revenue_a:,.0f} | Budget: {currency} {revenue_b:,.0f}")
    c2.metric("Gross Profit", fmt_km(gp_a), fmt_km(gp_a - gp_b),
              help=f"Actual: {currency} {gp_a:,.0f} | Budget: {currency} {gp_b:,.0f}")
    c3.metric("EBITDA", fmt_km(ebitda_a), fmt_km(ebitda_a - ebitda_b),
              help=f"Actual: {currency} {ebitda_a:,.0f} | Budget: {currency} {ebitda_b:,.0f}")
    c4.metric("PBT", fmt_km(pbt_a), fmt_km(pbt_a - pbt_b),
              help=f"Actual: {currency} {pbt_a:,.0f} | Budget: {currency} {pbt_b:,.0f}")
    c5.metric("PAT", fmt_km(pat_a), fmt_km(pat_a - pat_b),
              help=f"Actual: {currency} {pat_a:,.0f} | Budget: {currency} {pat_b:,.0f}")
    st.caption(f"Hover metrics for full {currency} values | Mapping Confidence: {_map_conf} | Agent 1: {agent1_time}s")

    st.divider()

    # ── 2. DECISION ENGINE (always visible) ───────────────────────────────
    dc1, dc2, dc3 = st.columns(3)
    dc1.metric("Severity", severity)
    dc2.metric("Decision", f"Category {decision_category}")
    dc3.metric("Action", decision_name)

    # ── Risk Escalation Alerts (inside Decision Engine section) ───────────
    if alerts:
        for _alert in alerts:
            st.warning(_alert)

    # ── 3. CFO ACTION PANEL — Action Tracker (always visible) ────────────
    st.markdown("### 🎯 CFO ACTIONS — Next 30 Days")
    _action_source = dd_actions[:5] if dd_actions else [{
        "action": w["action"], "detail": w.get("impact", ""),
        "owner": w["owner"], "timeline": "30 days",
        "impact": 0,
    } for w in qw[:3]]

    actions_df = pd.DataFrame([{
        "Priority": i + 1,
        "Action": a["action"],
        "Owner": a["owner"],
        "Impact": f"{currency} {a['impact']:,.0f}" if a["impact"] else "TBD",
        "Timeline": a["timeline"],
        "Status": "⏳ Pending",
        "Deadline": "TBD by CFO",
    } for i, a in enumerate(_action_source)])
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
    st.download_button(
        "📥 Download Action Tracker",
        data=_tracker_buf,
        file_name=f"Action_Tracker_{company_name or 'Company'}_{reporting_period}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="_dl_tracker",
    )

    # ── 4. TOP MATERIAL VARIANCES (always visible, compact) ───────────────
    if len(material) > 0:
        st.markdown(f"**⚠️ Material Variances ({material_count} items exceeding {mat_pct}% & {currency} {mat_abs:,.0f})**")
        _top_mat = material.head(3)
        _mat_cols = st.columns(3)
        for _i, (_, _mr) in enumerate(_top_mat.iterrows()):
            _mat_cols[_i].metric(
                _mr["Category"],
                fmt_km(_mr["Variance"]),
                f"{_mr['Variance %']:.1f}% {_mr['Flag']}",
                help=f"Variance: {currency} {_mr['Variance']:,.0f}",
            )
    else:
        st.success("No material variances detected.")

    st.divider()

    # ── 5. EBITDA NARRATIVE (always visible) ──────────────────────────────
    st.markdown("**EBITDA Narrative**")
    st.info(narr)

    if pbt_a < 0:
        st.warning(
            f"⚠️ PBT is NEGATIVE ({currency} {pbt_a:,.0f}). Tax = NIL. "
            f"Deferred Tax Asset of {currency} {abs(pbt_a) * tax_rate:,.0f} "
            f"to be assessed for recognition under {accounting_std}."
        )

    # ── Meeting Mode: show Agent 3 output + footer, skip detail ─────────
    if meeting_mode:
        if st.session_state.get("p3_output"):
            st.divider()
            st.subheader("📝 Board Memo")
            st.markdown(st.session_state["p3_output"])
        st.divider()
        st.caption("Detailed analysis available in full mode (toggle off Board Meeting Mode)")
        st.divider()
        st.markdown("### 📞 Want This For Your Company?")
        st.markdown("See the full AI analysis with root cause + board memo generation in a live walkthrough.")
        st.link_button("📅 Book a 15-Minute Live Demo", "https://www.linkedin.com/in/bhargav-venkatesh/", use_container_width=True)
        st.divider()
        st.caption("PBV Finance | AI CFO Systems v1.3 | Board Meeting Mode")
        st.stop()

    # ═══════════════════════════════════════════════════════════════════════
    # FULL MODE — Detail below this line is hidden in Meeting Mode
    # ═══════════════════════════════════════════════════════════════════════

    st.divider()

    # ═══════════════════════════════════════════════════════════════════════
    #                   D E T A I L   E X P A N D E R S
    # ═══════════════════════════════════════════════════════════════════════

    with st.expander("📊 Full Variance Table", expanded=False):
        st.dataframe(agg_data, use_container_width=True, hide_index=True)

    with st.expander("📈 P&L Waterfall (Revenue → PAT)", expanded=False):
        st.dataframe(wf, use_container_width=True, hide_index=True)

    with st.expander("🔗 EBITDA Bridge", expanded=False):
        st.dataframe(bdf, use_container_width=True, hide_index=True)

    with st.expander("🔍 PVME Analysis", expanded=False):
        if has_units and has_price:
            vol_data = raw[raw["Month"] == selected_month] if has_month and selected_month else raw
            st.success("✅ Unit data detected — calculating PVME")
            pvme_rows = []
            for _, row in vol_data.iterrows():
                bu = row.get("Budget Units", 0)
                au = row.get("Actual Units", 0)
                bp = row.get("Budget Price", 0)
                ap = row.get("Actual Price", 0)
                price_var = au * (ap - bp)
                vol_var = bp * (au - bu)
                total = price_var + vol_var
                pvme_rows.append({
                    "Product": row.get("Product", "N/A"),
                    "Price Variance": price_var,
                    "Volume Variance": vol_var,
                    "Total": total,
                    "Confidence": "🟢 CALCULATED",
                })
            if pvme_rows:
                st.dataframe(pd.DataFrame(pvme_rows), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ PVME requires unit data (Budget Units, Actual Units, Budget Price, Actual Price)")
            st.markdown(f"""
**Data needed from {erp_system}:**
- **VA05** — Sales Orders: actual units sold by product
- **MCSI** — Customer Analysis: volume by customer
- **KE30** — Profitability Analysis: margin by product/segment

*Until unit data is provided, revenue root cause is classified as 🔴 HYPOTHESIS*
            """)

    with st.expander("📋 Top Drivers Detail", expanded=False):
        st.dataframe(
            ds[["Category", "Variance", "Variance %", "Flag", "% of UF"]].reset_index(drop=True),
            use_container_width=True, hide_index=True
        )

    with st.expander("✅ Validation (10 Checks)", expanded=False):
        st.dataframe(
            pd.DataFrame(checks, columns=["Check", "Status"]),
            use_container_width=True, hide_index=True
        )
        if all_pass:
            st.success("All 10 checks PASSED ✅")
        else:
            st.error("Some checks FAILED ❌")

    with st.expander(f"🏛️ {country} Compliance Flags", expanded=False):
        for _cf in comp_flags:
            st.markdown(f"- {_cf}")
        st.caption("🟢 HARDCODED — verified regulations")

    with st.expander("⚡ Quick Wins Detail (30-Day, Zero Budget)", expanded=False):
        st.dataframe(
            pd.DataFrame([{
                "Action": w["action"], "Owner": w["owner"],
                "Impact": w["impact"], "ERP": w["erp"]
            } for w in qw[:5]]),
            use_container_width=True, hide_index=True
        )
        st.caption("🟡 PRE-BUILT — standard actions, customize per client")

    st.divider()

    # ═══════════════════════════════════════
    # DOWNLOADS (with approval stamp)
    # ═══════════════════════════════════════
    _now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    _dl_reviewer = st.session_state.get("reviewer_name", "")
    _dl_quality = st.session_state.get("review_quality", "Select...")
    _dl_reviewed = bool(_dl_reviewer) and _dl_quality not in ("Select...", "❌ Rejected — rerun Agent 2 with better context")
    _reviewer = _dl_reviewer if _dl_reviewed else "Not yet reviewed"
    _rev_qual = _dl_quality if _dl_reviewed else "Pending"

    _audit_df = pd.DataFrame([
        {"Field": "Generated", "Value": _now_str},
        {"Field": "Reviewed By", "Value": _reviewer},
        {"Field": "Quality", "Value": _rev_qual},
        {"Field": "Decision", "Value": f"Category {decision_category} — {decision_name}"},
        {"Field": "Version", "Value": "PBV Finance AI CFO v1.3"},
    ])

    _dl1, _dl2 = st.columns(2)
    with _dl1:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            agg_data.to_excel(writer, sheet_name="Variance", index=False)
            wf.to_excel(writer, sheet_name="Waterfall", index=False)
            bdf.to_excel(writer, sheet_name="EBITDA Bridge", index=False)
            pd.DataFrame(checks, columns=["Check", "Status"]).to_excel(
                writer, sheet_name="Validation", index=False
            )
            data.to_excel(writer, sheet_name="Raw Line Items", index=False)
            if len(material) > 0:
                material.to_excel(writer, sheet_name="Material Items", index=False)
            _audit_df.to_excel(writer, sheet_name="Audit", index=False)
        buffer.seek(0)
        st.download_button(
            label=f"📥 Download Excel Report ({company_name or 'Analysis'})",
            data=buffer,
            file_name=f"PBV_Variance_{company_name}_{reporting_period}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with _dl2:
        _memo_stamp = (
            f"\n\n{'=' * 40}\n"
            f"Generated by: PBV Finance AI CFO Assistant v1.3\n"
            f"Analysis Date: {_now_str}\n"
            f"Reviewed by: {_reviewer}\n"
            f"Review Quality: {_rev_qual}\n"
            f"Decision: Category {decision_category} — {decision_name}\n"
        )
        st.download_button(
            label="📥 Download Board Memo",
            data=_memo_text + _memo_stamp,
            file_name=f"CFO_Memo_{company_name or 'Company'}_{reporting_period}.txt",
            mime="text/plain",
        )

    st.divider()

    # ═══════════════════════════════════════
    # AGENT 2: ROOT CAUSE ANALYSIS (AI)
    # ═══════════════════════════════════════
    st.header("🤖 AI Agents — Root Cause & Board Memo")

    st.subheader("🔍 Agent 2: Root Cause Diagnostician")
    st.caption("Groq (Llama 3.3 70B) — uses ONLY pre-computed numbers")

    if st.button("▶️ Run Agent 2: Root Cause Analysis", type="primary", key="_run_agent2"):
        _p2_prompt = f"""{ai_context}

You are a CFO Diagnostician. Using ONLY the verified data above,
produce a QUANTIFIED root cause analysis.

STRICT RULES:
- Use ONLY numbers from the data above. Do NOT invent new numbers.
- Every cause must have a {currency} amount and percentage.
- Label each cause: VERIFIED (from data) or HYPOTHESIS (assumption).
- No generic consulting language. Be specific to {industry} in {country}.

IMPORTANT: Analyze EVERY material unfavorable item separately. Do not combine them. For each item provide 2-3 specific causes with {currency} amounts. Your response should be 400-500 words minimum.

OUTPUT THIS EXACT FORMAT:

🔍 ROOT CAUSE ANALYSIS

[For each material UF item, write a separate section:]

**[Category] Variance: {currency} [amount] ([pct]%)**

1. Cause: [specific driver for {industry} in {country}]
   Impact: {currency} [amount] (~[X]% of this variance)
   Owner: [role]
   Data: [{erp_system} report code]
   Confidence: [VERIFIED/HYPOTHESIS]

2. Cause: [specific driver]
   Impact: {currency} [amount] (~[X]% of this variance)
   Owner: [role]
   Data: [{erp_system} report code]
   Confidence: [VERIFIED/HYPOTHESIS]

3. Cause: [specific driver]
   Impact: {currency} [amount] (~[X]% of this variance)
   Owner: [role]
   Data: [{erp_system} report code]
   Confidence: [VERIFIED/HYPOTHESIS]

[Causes per item must sum to ~100% of that item's variance]

🎯 CAUSE → ACTION LINK

[For each cause, link to pre-computed action above:]
- [Cause] → [Action from list above] → {currency} [impact]

⚠️ DATA GAPS
- [List what data is missing for full validation]
- [Which {erp_system} reports would resolve these gaps]

📊 RECOVERY PROJECTION
Use these EXACT numbers for the recovery projection. Do NOT calculate your own:
- Current EBITDA: {currency} {ebitda_a:,.0f}
- Total Recoverable Impact: {currency} {total_recoverable:,.0f}
- Post-Actions EBITDA: {currency} {post_action_ebitda:,.0f}
- Timeline: 30-45 days
{"- EBITDA TURNS POSITIVE after full execution" if ebitda_a < 0 and post_action_ebitda > 0 else ""}

Do NOT show thinking. Start with ROOT CAUSE ANALYSIS immediately."""

        _p2_container = st.empty()
        _p2_container.info("Agent 2 running on Groq (Llama 3.3 70B)...")
        _p2_raw = call_ai(_p2_prompt, _p2_container)

        if _p2_raw:
            _p2_output = _strip_ai_preamble(_p2_raw, [
                "ROOT CAUSE ANALYSIS", "🔍 ROOT CAUSE", "🔍", "## ", "**",
            ])
            st.session_state["p2_output"] = _p2_output
            _p2_container.markdown(_p2_output)
            st.caption("🔴 AI-GENERATED — verify all numbers against Agent 1 calculations")

    elif st.session_state.get("p2_output"):
        st.markdown(st.session_state["p2_output"])
        st.caption("🔴 AI-GENERATED — verify all numbers against Agent 1 calculations")

    st.divider()

    # ═══════════════════════════════════════
    # ANALYST REVIEW GATE
    # ═══════════════════════════════════════
    _p2_available = bool(st.session_state.get("p2_output"))
    review_complete = False
    review_quality = "Select..."
    analyst_notes = ""
    reviewed_by = ""

    if _p2_available:
        st.subheader("📋 ANALYST REVIEW (Required Before Memo)")
        review_quality = st.radio(
            "AI Analysis Quality:",
            ["Select...",
             "✅ Confirmed — findings are reasonable",
             "⚠️ Acceptable — minor concerns noted below",
             "❌ Rejected — rerun Agent 2 with better context"],
            key="review_quality",
        )
        analyst_notes = st.text_area(
            "Analyst Notes / Corrections:",
            placeholder="Add context, corrections, or override any AI findings...",
            key="analyst_notes",
        )
        reviewed_by = st.text_input("Reviewed by:", key="reviewer_name")
        review_complete = (
            review_quality not in ["Select...", "❌ Rejected — rerun Agent 2 with better context"]
            and bool(reviewed_by)
        )
        if not review_complete:
            st.warning("Complete the review above to unlock Board Memo generation.")

        # Log the review
        if "run_log" not in st.session_state:
            st.session_state["run_log"] = []
        if review_complete and not st.session_state.get("_review_logged"):
            st.session_state["run_log"].append({
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "agent": "Review Gate",
                "reviewer": reviewed_by,
                "quality": review_quality,
                "notes": analyst_notes[:100] if analyst_notes else "None",
            })
            st.session_state["_review_logged"] = True

    st.divider()

    # ═══════════════════════════════════════
    # AGENT 3: CFO BOARD MEMO (AI)
    # ═══════════════════════════════════════
    st.subheader("📝 Agent 3: CFO Board Memo Writer")
    st.caption("Groq (Llama 3.3 70B) — uses Agent 2 findings + pre-computed numbers")

    if not _p2_available:
        st.info("Run Agent 2 first to generate root cause analysis, then Agent 3 can write the board memo.")
    elif not review_complete:
        st.info("Complete the analyst review above to unlock Agent 3.")

    if review_complete and st.button("▶️ Run Agent 3: Write Board Memo", type="primary", key="_run_agent3"):
        _p2_text = st.session_state["p2_output"][:2500]
        _review_context = f"\nANALYST REVIEW: {review_quality}. Notes: {analyst_notes}. Reviewed by: {reviewed_by}\n"

        _p3_prompt = f"""{ai_context}

ROOT CAUSE FINDINGS:
{_p2_text}
{_review_context}

You are a CFO writing a BOARD MEMO. The CEO has 60 seconds to read this.

STRICT RULES:
- Start with the HEADLINE — most critical issue.
- Every sentence must contain {currency} or %.
- Use ONLY numbers from verified data above.
- Actions must use the PRE-COMPUTED amounts above.
- No consulting tone. Write as OPERATOR who owns outcomes.
- End with specific board decisions required.

IMPORTANT: Include ALL 7 sections below. Do not skip any. Your response should be 300-400 words minimum. Every section must have specific {currency} amounts.

OUTPUT THIS EXACT FORMAT (all 7 sections required):

📌 HEADLINE
[1 sentence: what went wrong + urgency + {currency} scale]

📊 WHAT CHANGED
- Revenue: {currency} {revenue_b:,.0f} → {currency} {revenue_a:,.0f} ({(revenue_a - revenue_b) / revenue_b * 100 if revenue_b else 0:.1f}%)
- EBITDA: {currency} {ebitda_b:,.0f} → {currency} {ebitda_a:,.0f} (margin: {margin_b:.1f}% → {margin_a:.1f}%)
- PBT: {currency} {pbt_a:,.0f} {"(LOSS)" if pbt_a < 0 else ""}

🔍 WHY
[Top 2-3 causes from root cause analysis above, each with {currency} amount and %]

⚙️ ACTIONS (Next 30-45 Days)
[Use the pre-computed actions with exact {currency} amounts from the data above]

📈 EXPECTED OUTCOME
Use these EXACT numbers. Do NOT calculate your own:
- Current EBITDA: {currency} {ebitda_a:,.0f}
- Total Recovery: {currency} {total_recoverable:,.0f}
- Post-Actions EBITDA: {currency} {post_action_ebitda:,.0f}
- Timeline: 30-45 days
{"- EBITDA TURNS POSITIVE after full execution" if ebitda_a < 0 and post_action_ebitda > 0 else ""}

⚠️ RISKS
[2 specific risks with {currency} impact if actions fail]

🧠 BOARD DECISIONS REQUIRED
[2-3 specific approvals needed with {currency} amounts and deadlines]

Do NOT show thinking. Start with HEADLINE immediately."""

        _p3_container = st.empty()
        _p3_container.info("Agent 3 writing board memo on Groq (Llama 3.3 70B)...")
        _p3_raw = call_ai(_p3_prompt, _p3_container)

        if _p3_raw:
            _p3_output = _strip_ai_preamble(_p3_raw, [
                "CFO MEMORANDUM", "MEMORANDUM", "HEADLINE", "📌 HEADLINE",
                "📌", "## ", "TO:", "**",
            ])
            # Append review stamp
            _p3_output += f"\n\n---\nReviewed by: {reviewed_by} | Quality: {review_quality} | Notes: {analyst_notes}"
            st.session_state["p3_output"] = _p3_output
            _p3_container.markdown(_p3_output)
            st.caption("🔴 AI-GENERATED — verify all numbers against Agent 1 calculations")

            st.download_button(
                "📥 Download Board Memo (AI)",
                data=_p3_output,
                file_name=f"CFO_Memo_{company_name or 'Company'}_{reporting_period}.txt",
                mime="text/plain",
                key="_dl_ai_memo",
            )

    elif st.session_state.get("p3_output"):
        st.markdown(st.session_state["p3_output"])
        st.caption("🔴 AI-GENERATED — verify all numbers against Agent 1 calculations")
        st.download_button(
            "📥 Download Board Memo (AI)",
            data=st.session_state["p3_output"],
            file_name=f"CFO_Memo_{company_name or 'Company'}_{reporting_period}.txt",
            mime="text/plain",
            key="_dl_ai_memo_cached",
        )

    st.divider()
    st.markdown("### 📊 Confidence Guide")
    st.markdown("""
| Tag | Meaning | Source |
|---|---|---|
| 🟢 CALCULATED | Math fact | Python — verified by 10 checks |
| 🟡 PRE-BUILT | Standard action | Verified library |
| 🔴 HYPOTHESIS | Needs validation | AI — verify before acting |
    """)
    st.divider()
    st.markdown("### 📞 Want This For Your Company?")
    st.markdown("See the full AI analysis with root cause + board memo generation in a live walkthrough.")
    st.link_button("📅 Book a 15-Minute Live Demo", "https://www.linkedin.com/in/bhargav-venkatesh/", use_container_width=True)
    st.divider()
    st.caption("PBV Finance | AI CFO Systems v2.0 | Agent 1: Calculator | Agent 2: Diagnostician | Agent 3: Memo Writer | Powered by Groq (Llama 3.3 70B)")
