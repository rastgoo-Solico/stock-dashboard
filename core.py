"""core.py — پارس اکسل کدال، پیکربندی صنایع، محاسبه نسبت‌ها."""
import re
import pandas as pd
import numpy as np

# ═══════════════════════════ پارس اعداد فارسی ═══════════════════════════
PERSIAN = "۰۱۲۳۴۵۶۷۸۹"
ARABIC  = "٠١٢٣٤٥٦٧٨٩"

def to_en(s: str) -> str:
    for i, ch in enumerate(PERSIAN): s = s.replace(ch, str(i))
    for i, ch in enumerate(ARABIC):  s = s.replace(ch, str(i))
    return s

def normalize_label(s: str) -> str:
    s = to_en(str(s))
    s = s.replace("ي", "ی").replace("ك", "ک").replace("\u200c", "")
    return re.sub(r"\s+", "", s).strip()

def parse_num(x):
    if x is None or (isinstance(x, float) and np.isnan(x)): return np.nan
    if isinstance(x, (int, float)): return float(x)
    s = to_en(str(x)).strip()
    if not s or s in {"-", "*", ""}: return np.nan
    neg = s.startswith("(") and s.endswith(")")
    s = s.replace(",", "").replace("،", "").replace(" ", "").replace("%", "").strip("()")
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return np.nan

def find_row(df: pd.DataFrame, candidates):
    """ابتدا تطبیق دقیق، سپس substring."""
    if df.shape[1] < 2 or df.empty: return None
    col_a = df.iloc[:, 0].astype(str).map(normalize_label)
    for cand in candidates:
        cn = normalize_label(cand)
        mask = col_a == cn
        if mask.any(): return df[mask].iloc[0]
    for cand in candidates:
        cn = normalize_label(cand)
        if len(cn) < 4: continue  # از تطبیق خیلی کوتاه پرهیز کن
        mask = col_a.str.contains(re.escape(cn), na=False)
        if mask.any(): return df[mask].iloc[0]
    return None

def extract_from_excel(file_like, labels_dict: dict) -> dict:
    """پارس بهترین‌تلاش. برمی‌گرداند {key: value_in_million_rials}."""
    out = {}
    try:
        xls = pd.ExcelFile(file_like)
    except Exception:
        return out
    for sheet in xls.sheet_names:
        try:
            raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        except Exception:
            continue
        for hdr in range(min(6, len(raw))):
            df = raw.iloc[hdr + 1:].reset_index(drop=True)
            if df.shape[1] < 2 or df.empty: continue
            for key, cands in labels_dict.items():
                if key in out and not pd.isna(out[key]): continue
                row = find_row(df, cands)
                if row is None: continue
                nums = []
                for c in df.columns[1:6]:
                    v = parse_num(row.get(c))
                    if not pd.isna(v): nums.append(v)
                    if len(nums) >= 2: break
                if nums:
                    out[key] = nums[0]
                    if len(nums) > 1 and f"{key}_prev" not in out:
                        out[f"{key}_prev"] = nums[1]
    return out

# ═══════════════════════════ برچسب‌های صنعتی ═══════════════════════════
BANK_LABELS = {
    "rev_loans":     ["درآمد تسهیلات اعطایی", "درآمد تسهيلات اعطايي"],
    "rev_fees":      ["درآمد کارمزد"],
    "deposit_cost":  ["هزینه سود سپرده‌ها", "هزينه سود سپرده‌ها"],
    "admin_cost":    ["هزینه های اداری و عمومی", "هزينه هاي اداري و عمومي"],
    "provision":     ["هزینه مطالبات مشکوک الوصول", "هزينه مطالبات مشکوک الوصول"],
    "op_income":     ["درآمد عملیاتی"],
    "op_profit":     ["سود(زیان) عملیاتی", "سود(زیان)عملیاتی"],
    "net_income":    ["سود(زیان)خالص", "سود (زیان) خالص"],
    "eps":           ["سود (زیان) خالص هر سهم– ریال", "سود (زیان) خالص هر سهم"],
    "cash":          ["موجودی نقد", "موجودي نقد"],
    "total_assets":  ["جمع دارایی‌ها", "جمع دارایی ها"],
    "total_liab":    ["جمع بدهی‌ها", "جمع بدهی ها"],
    "equity":        ["جمع حقوق مالکانه"],
    "deposits":      ["سپرده‌های مشتریان", "سپرده‌هاي مشتریان"],
    "loans_gov":     ["تسهیلات اعطایی و مطالبات از اشخاص دولتی", "تسهيلات اعطايي و مطالبات از اشخاص دولتي"],
    "loans_nongov":  ["تسهیلات اعطایی و مطالبات از اشخاص غیر دولتی- ریالی",
                      "تسهیلات اعطایی و مطالبات از اشخاص غیر دولتی"],
    "total_loans":   ["مانده تسهیلات پایان دوره"],
    "npl_doubtful":  ["ذخیره اختصاصی مطالبات مشکوک الوصول"],
}

GENERIC_LABELS = {
    "revenue":        ["فروش", "درآمدهای عملیاتی", "درآمد عملیاتی"],
    "cogs":           ["بهای تمام شده کالای فروش رفته", "بهای تمام شده"],
    "gross_profit":   ["سود (زیان) ناخالص", "سود(زیان) ناخالص"],
    "sga":            ["هزینه های فروش، اداری و عمومی", "هزینه های اداری و عمومی"],
    "op_profit":      ["سود(زیان) عملیاتی", "سود(زیان)عملیاتی"],
    "fin_cost":       ["هزینه های مالی", "هزینه‌های مالی"],
    "pretax":         ["سود(زیان)قبل از مالیات", "سود (زیان) عملیات در حال تداوم قبل از مالیات"],
    "net_income":     ["سود(زیان)خالص", "سود (زیان) خالص"],
    "eps":            ["سود (زیان) خالص هر سهم– ریال", "سود (زیان) خالص هر سهم", "سود هر سهم"],
    "cash":           ["موجودی نقد", "موجودي نقد"],
    "receivables":    ["حساب ها و اسناد دریافتنی تجاری", "حساب‌های دریافتنی تجاری"],
    "inventory":      ["موجودی مواد و کالا", "موجودي مواد و كالا"],
    "current_assets": ["جمع دارایی‌های جاری", "جمع دارایی های جاری"],
    "total_assets":   ["جمع دارایی‌ها", "جمع دارایی ها"],
    "payables":       ["حساب ها و اسناد پرداختنی تجاری", "حساب‌های پرداختنی تجاری"],
    "current_liab":   ["جمع بدهی‌های جاری", "جمع بدهی های جاری"],
    "total_debt":     ["تسهیلات مالی دریافتی", "جمع بدهی‌ها", "جمع بدهی ها"],
    "equity":         ["جمع حقوق مالکانه"],
    "cfo":            ["جریان خالص ورود (خروج) نقد حاصل از فعالیت های عملیاتی",
                       "جريان خالص ورود (خروج) نقد حاصل از فعاليت هاي عملياتي"],
    "capex":          ["پرداخت های نقدی برای تحصیل دارایی ثابت مشهود",
                       "پرداخت های نقدی برای تحصیل دارایی ثابت"],
}

PHARMA_LABELS = {**GENERIC_LABELS,
                 "rd": ["هزینه تحقیق و توسعه", "هزینه های تحقیق و توسعه"]}

# ═══════════════════════════ توابع کمکی ═══════════════════════════
def safe_div(a, b):
    try:
        if b in (0, None) or pd.isna(a) or pd.isna(b): return np.nan
        return a / b
    except Exception:
        return np.nan

def safe_add(*vals):
    if any(pd.isna(v) for v in vals): return np.nan
    return float(sum(vals))

def safe_mul(a, b):
    if pd.isna(a) or pd.isna(b): return np.nan
    return a * b

# ═══════════════════════════ موتور نسبت‌ها ═══════════════════════════
def bank_ratios(d: dict, price_rial: float, shares: float, period_months: int = 6) -> dict:
    f = 12 / period_months
    ni = d.get("net_income", np.nan)
    ni_a = safe_mul(ni, f)
    eq = d.get("equity", np.nan); ta = d.get("total_assets", np.nan)
    eps = d.get("eps", np.nan)
    eps_a = safe_mul(eps, f) if not pd.isna(eps) else safe_div(safe_mul(ni_a, 1e6), shares)

    loans = d.get("total_loans", np.nan)
    if pd.isna(loans) or loans == 0:
        loans = safe_add(d.get("loans_gov", 0) or 0, d.get("loans_nongov", 0) or 0)
    bvps = safe_div(safe_mul(eq, 1e6), shares)

    r = {}
    r["EPS سالانه (ریال)"]      = eps_a
    r["BVPS (ریال)"]            = bvps
    r["ROE سالانه (%)"]         = safe_mul(safe_div(ni_a, eq), 100)
    r["ROA سالانه (%)"]         = safe_mul(safe_div(ni_a, ta), 100)
    r["Cost-to-Income (%)"]     = safe_mul(safe_div(abs(d.get("admin_cost", np.nan)), d.get("op_income", np.nan)), 100)
    r["حاشیه سود خالص (%)"]     = safe_mul(safe_div(ni, d.get("op_income", np.nan)), 100)
    r["LDR (%)"]                = safe_mul(safe_div(loans, d.get("deposits", np.nan)), 100)
    r["پوشش ذخیره (%)"]         = safe_mul(safe_div(d.get("npl_doubtful", np.nan), loans), 100)
    r["نقد/دارایی (%)"]         = safe_mul(safe_div(d.get("cash", np.nan), ta), 100)
    r["بدهی/دارایی (%)"]        = safe_mul(safe_div(d.get("total_liab", np.nan), ta), 100)
    r["P/E"]                    = safe_div(price_rial, eps_a)
    r["P/B"]                    = safe_div(price_rial, bvps)
    r["Earnings Yield (%)"]     = safe_mul(safe_div(eps_a, price_rial), 100)
    r["Market Cap (همت)"]       = safe_div(safe_mul(price_rial, shares), 1e12)
    return r

def generic_ratios(d: dict, price_rial: float, shares: float, period_months: int = 6) -> dict:
    f = 12 / period_months
    rev  = d.get("revenue", np.nan)
    cogs = abs(d["cogs"]) if not pd.isna(d.get("cogs", np.nan)) else np.nan
    ni   = d.get("net_income", np.nan)
    ni_a = safe_mul(ni, f)
    eq   = d.get("equity", np.nan); ta = d.get("total_assets", np.nan)
    ar   = d.get("receivables", np.nan); inv = d.get("inventory", np.nan)
    ap   = d.get("payables", np.nan); ca = d.get("current_assets", np.nan); cl = d.get("current_liab", np.nan)
    debt = d.get("total_debt", np.nan)
    cfo  = d.get("cfo", np.nan)
    capex = abs(d["capex"]) if not pd.isna(d.get("capex", np.nan)) else 0
    eps_a = safe_div(safe_mul(ni_a, 1e6), shares)
    bvps  = safe_div(safe_mul(eq, 1e6), shares)
    mcap  = safe_mul(price_rial, shares)
    days  = 30 * period_months

    dso = safe_mul(safe_div(ar, rev), days)
    dio = safe_mul(safe_div(inv, cogs), days)
    dpo = safe_mul(safe_div(ap, cogs), days)
    ccc = (dso + dio - dpo) if not (pd.isna(dso) or pd.isna(dio) or pd.isna(dpo)) else np.nan

    r = {}
    r["حاشیه سود ناخالص (%)"]   = safe_mul(safe_div(safe_add(rev, -(cogs or 0)) if not pd.isna(rev) else np.nan, rev), 100)
    r["حاشیه سود عملیاتی (%)"]  = safe_mul(safe_div(d.get("op_profit", np.nan), rev), 100)
    r["حاشیه سود خالص (%)"]     = safe_mul(safe_div(ni, rev), 100)
    r["ROE (%)"]                = safe_mul(safe_div(ni_a, eq), 100)
    r["ROA (%)"]                = safe_mul(safe_div(ni_a, ta), 100)
    r["نسبت جاری"]              = safe_div(ca, cl)
    r["Debt/Equity"]            = safe_div(debt, eq)
    r["Interest coverage"]      = safe_div(d.get("op_profit", np.nan), abs(d["fin_cost"]) if not pd.isna(d.get("fin_cost", np.nan)) else np.nan)
    r["گردش موجودی (بار)"]      = safe_div(safe_mul(cogs, f), inv)
    r["DSO (روز)"]              = dso
    r["DIO (روز)"]              = dio
    r["DPO (روز)"]              = dpo
    r["CCC (روز)"]              = ccc
    r["FCF Yield (%)"]          = safe_mul(safe_div(safe_mul(safe_add(safe_mul(cfo, f), -abs(safe_mul(capex, f))), 1e6), mcap), 100)
    r["EPS سالانه (ریال)"]      = eps_a
    r["BVPS (ریال)"]            = bvps
    r["P/E"]                    = safe_div(price_rial, eps_a) if (eps_a and eps_a > 0) else np.nan
    r["P/B"]                    = safe_div(price_rial, bvps)
    r["P/S"]                    = safe_div(mcap, safe_mul(rev, safe_mul(f, 1e6)))
    r["Market Cap (همت)"]       = safe_div(mcap, 1e12)
    return r

def petro_ratios(d, price_rial, shares, period_months=6):
    r = generic_ratios(d, price_rial, shares, period_months)
    f = 12 / period_months
    # تقریب EBITDA = سود عملیاتی + استهلاک (تقریبی ۵٪ فروش)
    ebitda = safe_add(d.get("op_profit", np.nan), safe_mul(d.get("revenue", np.nan), 0.05))
    ebitda_a = safe_mul(ebitda, f)
    net_debt = safe_add(d.get("total_debt", 0) or 0, -(d.get("cash", 0) or 0))
    mcap = safe_mul(price_rial, shares)
    ev = safe_add(mcap, safe_mul(net_debt, 1e6))
    r["EV/EBITDA"] = safe_div(ev, safe_mul(ebitda_a, 1e6))
    return r

def pharma_ratios(d, price_rial, shares, period_months=6):
    r = generic_ratios(d, price_rial, shares, period_months)
    r["R&D / فروش (%)"] = safe_mul(safe_div(d.get("rd", np.nan), d.get("revenue", np.nan)), 100)
    return r

def auto_ratios(d, price_rial, shares, period_months=6):
    return generic_ratios(d, price_rial, shares, period_months)

# ═══════════════════════════ پیکربندی صنایع ═══════════════════════════
SECTOR_CONFIG = {
    "bank": {
        "name": "بانک و مؤسسه مالی",
        "labels": BANK_LABELS, "ratios_fn": bank_ratios,
        "primary": ["P/B", "P/E", "ROE سالانه (%)", "Cost-to-Income (%)", "LDR (%)"],
        "secondary": ["ROA سالانه (%)", "پوشش ذخیره (%)", "حاشیه سود خالص (%)", "Earnings Yield (%)", "Market Cap (همت)"],
        "benchmark": {"P/E": 5.8, "P/B": 1.2, "ROE سالانه (%)": 25, "Cost-to-Income (%)": 40},
        "warning": None,
    },
    "petro": {
        "name": "پتروشیمی و شیمیایی",
        "labels": GENERIC_LABELS, "ratios_fn": petro_ratios,
        "primary": ["P/E", "EV/EBITDA", "حاشیه سود ناخالص (%)", "ROE (%)", "FCF Yield (%)"],
        "secondary": ["P/B", "P/S", "ROA (%)", "Debt/Equity", "Interest coverage"],
        "benchmark": {"P/E": 6.5, "P/B": 1.5, "ROE (%)": 20},
        "warning": "⚠️ پتروشیمی چرخه‌ای است — P/E فعلی گمراه‌کننده؛ میانگین ۵ساله ملاک است.",
    },
    "pharma": {
        "name": "دارویی",
        "labels": PHARMA_LABELS, "ratios_fn": pharma_ratios,
        "primary": ["P/E", "P/B", "حاشیه سود ناخالص (%)", "R&D / فروش (%)", "ROE (%)"],
        "secondary": ["ROA (%)", "DSO (روز)", "DIO (روز)", "CCC (روز)", "FCF Yield (%)"],
        "benchmark": {"P/E": 8.0, "P/B": 2.5, "ROE (%)": 25},
        "warning": "⚠️ مواد اولیه وارداتی — نرخ ارز بر حاشیه سود اثر مستقیم دارد.",
    },
    "auto": {
        "name": "خودرو و قطعات",
        "labels": GENERIC_LABELS, "ratios_fn": auto_ratios,
        "primary": ["P/E", "حاشیه سود ناخالص (%)", "Debt/Equity", "Interest coverage", "ROE (%)"],
        "secondary": ["گردش موجودی (بار)", "DSO (روز)", "FCF Yield (%)", "P/B", "ROA (%)"],
        "benchmark": {"P/E": 5.0, "P/B": 1.0},
        "warning": "⚠️ نرخ‌گذاری دستوری حاشیه‌ها را distort می‌کند؛ P/E در زیان‌دهی بی‌معناست.",
    },
    "generic": {
        "name": "عمومی (فولاد، سیمان، غذایی، ...)",
        "labels": GENERIC_LABELS, "ratios_fn": generic_ratios,
        "primary": ["P/E", "P/B", "ROE (%)", "ROA (%)", "حاشیه سود ناخالص (%)"],
        "secondary": ["P/S", "Debt/Equity", "نسبت جاری", "CCC (روز)", "FCF Yield (%)"],
        "benchmark": {"P/E": 5.5, "P/B": 1.3, "ROE (%)": 20},
        "warning": None,
    },
}

# ═══════════════════════════ داده دمو (وبملت) ═══════════════════════════
DEMO_BANK = {
    "rev_loans": 1_835_436, "rev_loans_prev": 1_259_107,
    "deposit_cost": 1_220_160, "admin_cost": 707_453, "admin_cost_prev": 333_720,
    "provision": 43_432, "op_income": 2_150_762, "op_income_prev": 1_325_688,
    "op_profit": 1_271_937, "net_income": 1_025_278, "net_income_prev": 525_849,
    "eps": 303,
    "cash": 3_536_996, "total_assets": 59_110_305, "total_liab": 52_628_589,
    "equity": 6_481_715, "deposits": 27_590_264, "deposits_prev": 22_270_416,
    "loans_gov": 28_271_151, "loans_nongov": 16_938_084,
    "total_loans": 47_324_584, "npl_doubtful": 1_438_867,
}
