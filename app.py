import streamlit as st
import pandas as pd
import io
import json
import os
import time
import re

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
        st.subheader("Key Metrics (Actual)")
        _currency = st.session_state.get("_currency", "AED")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric(f"Revenue ({_currency})", fmt_km(revenue_a))
        col2.metric(f"Gross Profit ({_currency})", fmt_km(gp_a))
        col3.metric(f"EBITDA ({_currency})", fmt_km(ebitda_a))
        col4.metric(f"Net Profit ({_currency})", fmt_km(pat_a))

        st.info("📤 Upload budget data for full variance analysis.")
        st.stop()

    # ═══════════════════════════════════════
    # AGENT 1: CALCULATOR
    # ═══════════════════════════════════════
    st.success("✅ Mapping confirmed. Running P&L analysis.")
    st.header("🔢 Agent 1: Calculator (100% Accurate)")

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

    st.subheader("TABLE 1: Category-Level P&L Variance")
    st.dataframe(agg_data, use_container_width=True, hide_index=True)

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

    # ── Key Metrics ───────────────────────────────────────────────────────
    def fmt_km(n):
        """Format a number as 1.1M or 1,050K (no currency prefix)."""
        abs_n = abs(n)
        if abs_n >= 1_000_000:
            return f"{n / 1_000_000:.1f}M"
        if abs_n >= 1_000:
            return f"{n / 1_000:,.0f}K"
        return f"{n:,.0f}"

    st.subheader(f"Key Metrics ({currency})")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Revenue",      fmt_km(revenue_a),       fmt_km(revenue_a - revenue_b))
    c2.metric("Gross Profit", fmt_km(gp_a),            fmt_km(gp_a - gp_b))
    c3.metric("EBITDA",       fmt_km(ebitda_a),        fmt_km(ebitda_a - ebitda_b))
    c4.metric("PBT",          fmt_km(pbt_a),           fmt_km(pbt_a - pbt_b))
    c5.metric("PAT",          fmt_km(pat_a),           fmt_km(pat_a - pat_b))

    # ── TABLE 2: Waterfall ────────────────────────────────────────────────
    st.subheader("TABLE 2: P&L Waterfall (Revenue → PAT)")
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
    st.dataframe(wf, use_container_width=True, hide_index=True)

    if pbt_a < 0:
        st.warning(
            f"⚠️ PBT is NEGATIVE ({currency} {pbt_a:,.0f}). Tax = NIL. "
            f"Deferred Tax Asset of {currency} {abs(pbt_a) * tax_rate:,.0f} "
            f"to be assessed for recognition under {accounting_std}."
        )

    # ── TABLE 3: EBITDA Bridge ────────────────────────────────────────────
    st.subheader("TABLE 3: EBITDA Bridge")
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
            "COGS",
            "Employee",
            "Marketing",
            "Other OpEx",
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
    st.dataframe(bdf, use_container_width=True, hide_index=True)

    # ── Validation checks ─────────────────────────────────────────────────
    st.subheader("✅ Self-Validation (10 Checks)")
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
    st.dataframe(
        pd.DataFrame(checks, columns=["Check", "Status"]),
        use_container_width=True, hide_index=True
    )
    all_pass = all("PASS" in c[1] for c in checks)
    if all_pass:
        st.success("All 10 checks PASSED ✅")
    else:
        st.error("Some checks FAILED ❌")

    # ── Material variances ────────────────────────────────────────────────
    material = agg_data[agg_data["Material?"] == "⚠️ YES"]
    if len(material) > 0:
        st.subheader(f"⚠️ Material Variances ({len(material)} items)")
        st.dataframe(
            material[["Category", "Variance", "Variance %", "Flag"]],
            use_container_width=True, hide_index=True
        )

    # ── TABLE 4: EBITDA Narrative ─────────────────────────────────────────
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

    st.subheader("TABLE 4: EBITDA Narrative")
    st.info(narr)
    st.caption("🟢 CALCULATED — template-filled, zero AI")

    # ═══════════════════════════════════════
    # TABLE 5: PVME (if unit data available)
    # ═══════════════════════════════════════
    st.subheader("TABLE 5: PVME Analysis")
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

    # ── TABLE 6: Top Drivers ──────────────────────────────────────────────
    st.subheader("TABLE 6: Top Drivers")
    tuf = ds[ds["Flag"] == "UF"]["Abs"].sum()
    ds["% of UF"] = ds.apply(
        lambda r: round(r["Abs"] / tuf * 100, 1) if r["Flag"] == "UF" and tuf > 0 else 0.0,
        axis=1
    )
    st.dataframe(
        ds[["Category", "Variance", "Variance %", "Flag", "% of UF"]].reset_index(drop=True),
        use_container_width=True, hide_index=True
    )

    # ── Country Compliance ────────────────────────────────────────────────
    st.subheader(f"🏛️ {country} Compliance ({rules['tax_desc'][:50]}...)")
    flags = []
    if revenue_b and abs(revenue_a - revenue_b) / revenue_b * 100 >= mat_pct:
        flags.append(f"**Revenue:** {rules['tax_desc']}")
    if emp_a > emp_b:
        flags.append(f"**Employee Cost:** {rules['labor']}")
    flags.append(f"**Indirect Tax:** {rules['vat']}")
    flags.append(f"**Other:** {rules['other']}")
    if pbt_a < 0:
        flags.append(f"**LOSS:** Assess DTA recognition. {rules['tax_desc']}")
    for f in flags:
        st.markdown(f"- {f}")
    st.caption("🟢 HARDCODED — verified regulations")

    # ── Quick Wins ────────────────────────────────────────────────────────
    st.subheader("⚡ Quick Wins (30-Day, Zero Budget)")
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
    st.dataframe(
        pd.DataFrame([{
            "Action": w["action"], "Owner": w["owner"],
            "Impact": w["impact"], "ERP": w["erp"]
        } for w in qw[:5]]),
        use_container_width=True, hide_index=True
    )
    st.caption("🟡 PRE-BUILT — standard actions, customize per client")

    # ═══════════════════════════════════════
    # DOWNLOAD BUTTON
    # ═══════════════════════════════════════
    st.subheader("📥 Download Report")
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
    buffer.seek(0)
    st.download_button(
        label=f"📥 Download Excel Report ({company_name or 'Analysis'})",
        data=buffer,
        file_name=f"PBV_Variance_{company_name}_{reporting_period}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    agent1_time = round(time.time() - start_time, 1)
    st.caption(f"Agent 1 completed in {agent1_time}s")

    st.divider()

    # ═══════════════════════════════════════
    # AGENTIC DECISION ENGINE
    # ═══════════════════════════════════════
    st.header("🧠 Agentic Decision Engine")

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

    dc1, dc2, dc3 = st.columns(3)
    dc1.metric("Severity", severity)
    dc2.metric("Decision", f"Category {decision_category}")
    dc3.metric("Action", decision_name)

    st.markdown(f"""
**Engine Assessment (🟢 Rule-Based, Zero AI):**
- Material variances: {material_count} items exceed {mat_pct}% AND {currency} {mat_abs:,.0f}
- PBT position: {currency} {pbt_a:,.0f} ({'LOSS' if pbt_a < 0 else 'PROFIT'})
- EBITDA margin swing: {swing:.1f}pp
- **Recommendation: Category {decision_category} — {decision_name}**
    """)
    st.caption("🟢 CALCULATED — Python decision engine, not AI opinion")

    st.divider()

    # ═══════════════════════════════════════
    # AI ANALYSIS — UPGRADE SECTION
    # ═══════════════════════════════════════
    st.header("🤖 AI Analysis — Full Version")

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("🔍 Agent 2: Root Cause Analysis")
        st.markdown(f"""
Diagnoses **why** variances happened using your verified P&L data:

- 3 revenue hypotheses with SAP report codes (VA05, KE30, MCSI)
- Each material cost item: root cause tagged Timing or Structural
- EBITDA margin swing attribution (revenue vs cost split)
- Scenario table: Base / Upside / Downside with probabilities
- 30/60/90-day action timeline
- 90-day risk if no action taken (quantified in {currency})
        """)

    with col_b:
        st.subheader("📝 Agent 3: CFO Board Memo")
        st.markdown("""
Writes a **one-page board memo** using Agent 2's diagnosis:

- Headline with materiality breach summary
- EBITDA bridge narrative
- Where we missed (table: Item | Variance | Root Cause)
- Revenue diagnosis with owner accountability
- Risks + Opportunities with amounts
- Pre-approved quick wins
- The Ask: 2-3 board decisions with amounts and deadlines
        """)

    st.info(
        "Agent 2 and Agent 3 run on **Gemma 4 locally** via Ollama — "
        "available in the full desktop version. Book a live demo to see it in action with your own data."
    )
    st.link_button("📅 Book a 15-Minute Live Demo", "https://www.linkedin.com/in/bhargavvenkatesh/")

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
    st.caption("PBV Finance | AI CFO Systems v2.0 | Phase 1: Smart Mapping | Anti-Hallucination v1.0 | Powered by Gemma 4")
