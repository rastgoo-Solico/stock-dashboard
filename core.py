"""core.py — پارس اکسل/HTML کدال + هشدار سلامت نسبت‌ها"""
import re
import io
import pandas as pd
import numpy as np

PERSIAN = "۰۱۲۳۴۵۶۷۸۹"
ARABIC  = "٠١٢٣٤٥٦٧٨٩"

def to_en(s):
    for i, ch in enumerate(PERSIAN): s = s.replace(ch, str(i))
    for i, ch in enumerate(ARABIC):  s = s.replace(ch, str(i))
    return s

def normalize_chars(s):
    return (s.replace("ي", "ی").replace("ى", "ی").replace("ك", "ک")
             .replace("ة", "ه").replace("\u200c", "")
             .replace("\u200f", "").replace("\u200e", "").replace("\u0640", ""))

def norm_label(s):
    s = to_en(str(s)); s = normalize_chars(s)
    return re.sub(r"\s+", "", s).strip()

def parse_num(x):
    if x is None or (isinstance(x, float) and np.isnan(x)): return np.nan
    if isinstance(x, (int, float)): return float(x)
    s = to_en(str(x)).strip()
    if not s or s in {"-", "*", ""}: return np.nan
    neg = s.startswith("(") and s.endswith(")")
    s = s.replace(",", "").replace("،", "").replace(" ", "").replace("%", "").strip("()")
    try:
        v = float(s); return -v if neg else v
    except: return np.nan

def safe_div(a, b):
    try:
        if b in (0, None) or pd.isna(a) or pd.isna(b): return np.nan
        return a / b
    except: return np.nan

def safe_add(*vs): return np.nan if any(pd.isna(v) for v in vs) else float(sum(vs))
def safe_mul(a, b): return np.nan if pd.isna(a) or pd.isna(b) else a * b


# ═══════════════ پیشنهاد قیمت (فقط جهت اطلاع) ═══════════════
def suggest_price_toman(ratios: dict, benchmark: dict) -> tuple:
    eps = ratios.get("EPS سالانه (ریال)", np.nan)
    bvps = ratios.get("BVPS (ریال)", np.nan)
    peer_pe = benchmark.get("P/E", 6.0)
    peer_pb = benchmark.get("P/B", 1.2)
    suggestions, labels = [], []
    if eps and not pd.isna(eps) and eps > 0:
        suggestions.append(eps * peer_pe / 10); labels.append(f"P/E={peer_pe}×EPS")
    if bvps and not pd.isna(bvps) and bvps > 0:
        suggestions.append(bvps * peer_pb / 10); labels.append(f"P/B={peer_pb}×BVPS")
    if eps and bvps and eps > 0 and bvps > 0:
        suggestions.append(np.sqrt(22.5 * eps * bvps) / 10); labels.append("Graham")
    if suggestions:
        return float(np.median(suggestions)), list(zip(labels, suggestions))
    return np.nan, []


# ═══════════════ بررسی سلامت ═══════════════
def health_check(ratios: dict, price_toman: float, shares_bn: float,
                 data: dict) -> list:
    warnings = []
    pe = ratios.get("P/E", np.nan)
    pb = ratios.get("P/B", np.nan)
    eq = data.get("equity", np.nan)

    if not pd.isna(pe):
        if pe < 0.3:
            warnings.append(
                f"🚨 P/E = {pe:.2f} غیرممکن است — قیمت (تومان/ریال) یا تعداد سهام را چک کنید."
            )
        elif pe < 0.5:
            warnings.append(
                f"🟡 P/E = {pe:.2f} بسیار پایین — آیا تعداد سهام درست است؟"
            )
        elif pe > 100:
            warnings.append(
                f"🚨 P/E = {pe:.0f} غیرممکن است — قیمت یا EPS را چک کنید."
            )
        elif pe > 30:
            warnings.append(f"🟡 P/E = {pe:.0f} بالا است — آیا قیمت درست است؟")

    if not pd.isna(pb):
        if pb < 0.02:
            warnings.append(
                f"🚨 P/B = {pb:.3f} غیرممکن است — تعداد سهام یا قیمت اشتباه است."
            )
        elif pb > 50:
            warnings.append(f"🚨 P/B = {pb:.1f} غیرممکن است — قیمت را چک کنید.")

    if not pd.isna(eq) and eq > 0 and shares_bn > 0:
        mcap_million = price_toman * 10 * shares_bn * 1e9 / 1e6
        ratio = mcap_million / eq
        if ratio < 0.02:
            warnings.append(
                f"🚨 ارزش بازار ({mcap_million:,.0f} م.ریال) تنها "
                f"{ratio*100:.1f}٪ حقوق مالکانه ({eq:,.0f}) است — قیمت احتمالاً خیلی پایین است."
            )
        elif ratio > 50:
            warnings.append(
                f"🚨 ارزش بازار {ratio:.0f}× حقوق مالکانه — قیمت احتمالاً خیلی بالا است."
            )

    cap = data.get("capital", np.nan)
    if not pd.isna(cap) and cap > 10_000 and shares_bn < 10:
        n = data.get("nominal_value", np.nan)
        if pd.isna(n) or n <= 0: n = 1000
        expected = (cap * 1e6) / (n * 1e9)
        warnings.append(
            f"🟡 سرمایه {cap:,.0f} م.ریال است، با اسمی {n:,.0f} "
            f"انتظار {expected:.2f} میلیارد سهم داریم — نه {shares_bn:.2f}."
        )

    return warnings


# ═══════════════ پارس فایل ═══════════════
def parse_any_format(content_bytes):
    if not content_bytes or len(content_bytes) < 8:
        return None, "empty", ["فایل خالی"]
    magic = content_bytes[:8]
    if magic[:2] == b'PK': kind = "xlsx"
    elif magic[:4] == b'\xd0\xcf\x11\xe0': kind = "xls"
    elif (b'<html' in content_bytes[:3000].lower()
          or b'<table' in content_bytes[:3000].lower()
          or b'<?xml' in content_bytes[:200]): kind = "html"
    else: kind = "unknown"
    errors = []
    if kind in ("xlsx", "xls", "unknown"):
        engines = ['xlrd', 'openpyxl'] if kind == "xls" else ['openpyxl', 'xlrd']
        for e in engines:
            try: return pd.ExcelFile(io.BytesIO(content_bytes), engine=e), kind, errors
            except Exception as ex: errors.append(f"[excel-{e}] {str(ex)[:100]}")
    if kind in ("html", "unknown"):
        for flavor in ['html5lib', 'bs4']:
            try:
                dfs = pd.read_html(io.BytesIO(content_bytes), flavor=flavor)
                if dfs: return dfs, f"html-{flavor}", errors
            except Exception as ex: errors.append(f"[html-{flavor}] {str(ex)[:100]}")
    for enc in ['utf-8', 'utf-8-sig', 'windows-1256', 'latin-1']:
        try:
            text = content_bytes.decode(enc, errors='ignore')
            for sep in ['\t', ';', ',']:
                try:
                    df = pd.read_csv(io.StringIO(text), sep=sep, header=None,
                                     on_bad_lines='skip', engine='python')
                    if df.shape[1] > 1 and df.shape[0] > 2:
                        return [df], f"csv-{enc}", errors
                except: continue
        except: continue
    return None, kind, errors


def extract_excel(file_like, labels):
    try:
        if hasattr(file_like, 'getvalue'): content = file_like.getvalue()
        elif hasattr(file_like, 'read'):
            file_like.seek(0); content = file_like.read()
        else: content = file_like
    except: return {}
    result, kind, errors = parse_any_format(content)
    if result is None: return {}
    out = {}
    if isinstance(result, pd.ExcelFile):
        for sh in result.sheet_names:
            try:
                raw = pd.read_excel(result, sheet_name=sh, header=None)
                _extract_from_raw(raw, labels, out)
            except: continue
    elif isinstance(result, list):
        for df in result:
            if df is not None and not df.empty and df.shape[0] >= 2:
                _extract_from_raw(df, labels, out)
    elif isinstance(result, pd.DataFrame):
        _extract_from_raw(result, labels, out)
    return out


def _extract_from_raw(raw, labels, out):
    """استخراج با ترجیح بزرگ‌ترین قدرمطلق (سود سالانه > فصلی > درصد)."""
    if raw is None or raw.empty: return

    if raw.shape[1] >= 2:
        for hdr in range(min(len(raw), 30)):
            df = raw.iloc[hdr + 1:].reset_index(drop=True)
            if df.empty or df.shape[1] < 2: continue
            col_a_norm = df.iloc[:, 0].astype(str).map(norm_label)
            for k, cands in labels.items():
                for cand in cands:
                    cn = norm_label(cand)
                    if not cn: continue
                    mask = col_a_norm == cn
                    if not mask.any() and len(cn) >= 4:
                        mask = col_a_norm.str.contains(re.escape(cn), na=False)
                    if not mask.any(): continue
                    row = df[mask].iloc[0]
                    nums = []
                    for c in df.columns[1:20]:
                        v = parse_num(row.get(c))
                        if not pd.isna(v): nums.append(v)
                        if len(nums) >= 6: break
                    if nums:
                        best = max(nums, key=abs)
                        existing = out.get(k, np.nan)
                        if pd.isna(existing) or abs(best) > abs(existing):
                            out[k] = best
                            others = [x for x in nums if x != best]
                            if others:
                                out[f"{k}_prev"] = max(others, key=abs)
                    break

    missing = [k for k in labels if k not in out or pd.isna(out.get(k))]
    if not missing: return

    for idx in range(len(raw)):
        for col in raw.columns:
            cell = raw.iloc[idx][col]
            if pd.isna(cell): continue
            text = str(cell)
            if len(text) < 12: continue
            for k in list(missing):
                if k in out and not pd.isna(out[k]):
                    if k in missing: missing.remove(k)
                    continue
                for cand in labels[k]:
                    nums = numbers_after_label(text, cand)
                    if nums:
                        best = max(nums, key=abs)
                        out[k] = best
                        others = [x for x in nums if x != best]
                        if others: out[f"{k}_prev"] = max(others, key=abs)
                        if k in missing: missing.remove(k)
                        break


def numbers_after_label(text, label, max_nums=6):
    if not text or not label: return []
    text_norm = normalize_chars(to_en(str(text)))
    label_norm = normalize_chars(label)
    if not label_norm: return []
    idx = text_norm.find(label_norm)
    if idx < 0: return []
    start = idx + len(label_norm)
    if start >= len(text_norm): return []
    preview = text_norm[start:start + 40]
    nm = re.search(r'[\d]', preview)
    if not nm: return []
    if re.search(r'[\u0600-\u06FF]', preview[:nm.start()]): return []
    after = text_norm[start:start + 500]
    nums = []
    for m in re.finditer(r'\(?[\d][\d,]*(?:\.\d+)?\)?', after):
        n = parse_num(m.group())
        if not pd.isna(n):
            nums.append(n)
            if len(nums) >= max_nums: break
    return nums


# ═══════════════ برچسب‌ها ═══════════════
BANK_LABELS = {
    "rev_loans":["درآمد تسهیلات اعطایی"], "rev_fees":["درآمد کارمزد"],
    "deposit_cost":["هزینه سود سپرده‌ها"], "admin_cost":["هزینه های اداری و عمومی"],
    "provision":["هزینه مطالبات مشکوک الوصول"], "op_income":["درآمد عملیاتی"],
    "op_profit":["سود (زیان) عملیاتی"], "net_income":["سود (زیان) خالص"],
    "eps":["سود (زیان) خالص هر سهم","سود (زیان) پایه هر سهم"],
    "cash":["موجودی نقد"], "total_assets":["جمع دارایی‌ها"],
    "total_liab":["جمع بدهی‌ها"], "equity":["جمع حقوق مالکانه"],
    "deposits":["سپرده‌های مشتریان"],
    "loans_gov":["تسهیلات اعطایی و مطالبات از اشخاص دولتی"],
    "loans_nongov":["تسهیلات اعطایی و مطالبات از اشخاص غیر دولتی- ریالی","تسهیلات اعطایی و مطالبات از اشخاص غیر دولتی"],
    "total_loans":["مانده تسهیلات پایان دوره"],
    "npl_doubtful":["ذخیره اختصاصی مطالبات مشکوک الوصول"],
    "capital":["سرمایه"],
    "nominal_value":["ارزش اسمی هر سهم (ریال)","ارزش اسمی هر سهم"],
}
GENERIC_LABELS = {
    "revenue":["فروش","درآمدهای عملیاتی"], "cogs":["بهای تمام شده کالای فروش رفته","بهای تمام شده"],
    "sga":["هزینه های فروش، اداری و عمومی","هزینه های اداری و عمومی"],
    "op_profit":["سود (زیان) عملیاتی"], "fin_cost":["هزینه های مالی"],
    "net_income":["سود (زیان) خالص"],
    "eps":["سود (زیان) خالص هر سهم","سود (زیان) پایه هر سهم","سود هر سهم"],
    "cash":["موجودی نقد"], "receivables":["حساب ها و اسناد دریافتنی تجاری"],
    "inventory":["موجودی مواد و کالا"], "current_assets":["جمع دارایی‌های جاری"],
    "total_assets":["جمع دارایی‌ها"], "payables":["حساب ها و اسناد پرداختنی تجاری"],
    "current_liab":["جمع بدهی‌های جاری"],
    "total_debt":["تسهیلات مالی دریافتی","جمع بدهی‌ها"], "equity":["جمع حقوق مالکانه"],
    "cfo":["جریان خالص ورود (خروج) نقد حاصل از فعالیت های عملیاتی"],
    "capex":["پرداخت های نقدی برای تحصیل دارایی ثابت مشهود"],
    "capital":["سرمایه"], "nominal_value":["ارزش اسمی هر سهم (ریال)","ارزش اسمی هر سهم"],
}
PHARMA_LABELS = {**GENERIC_LABELS, "rd":["هزینه تحقیق و توسعه","هزینه های تحقیق و توسعه"]}
HOLDING_LABELS = {
    "div_income":["درآمد سود سهام","درآمد حاصل از سود سهام","سود سهام"],
    "interest_income":["درآمد سود تضمین شده","درآمد سود تضمین‌شده"],
    "inv_sale_gain":["سود (زیان) فروش سرمایه گذاریها"],
    "inv_reval":["سود (زیان) تغییر ارزش سرمایه گذاریها"],
    "other_rev":["سایر درآمدها"], "rev_total":["جمع درآمدهای عملیاتی"],
    "salary":["هزینه های حقوق، دستمزد و مزایا"], "other_exp":["سایر هزینه‌ها"],
    "op_exp_total":["جمع هزینه های عملیاتی"], "op_profit":["سود (زیان) عملیاتی"],
    "fin_cost":["هزینه های مالی"], "non_op":["سایر درآمدها و هزینه‌های غیرعملیاتی"],
    "pretax":["سود (زیان) عملیات در حال تداوم قبل از مالیات","سود (زیان) قبل از احتساب سهم گروه"],
    "net_income":["سود (زیان) خالص"],
    "eps":["سود (زیان) خالص هر سهم","سود (زیان) پایه هر سهم"],
    "capital":["سرمایه"], "nominal_value":["ارزش اسمی هر سهم (ریال)","ارزش اسمی هر سهم"],
    "cash":["موجودی نقد"], "short_inv":["سرمایه‌گذاری‌های کوتاه‌مدت"],
    "long_inv":["سرمایه‌گذاری‌های بلندمدت"],
    "associates":["سرمایه گذاری در شرکتهای وابسته"],
    "total_assets":["جمع دارایی‌ها"], "total_liab":["جمع بدهی‌ها"],
    "equity":["جمع حقوق مالکانه","حقوق مالکانه قابل انتساب به مالکان شرکت اصلی"],
    "current_assets":["جمع دارایی‌های جاری"], "current_liab":["جمع بدهی‌های جاری"],
    "receivables":["دریافتنی‌های تجاری و سایر دریافتنی‌ها"],
    "cfo":["جریان خالص ورود (خروج) نقد حاصل از فعالیت های عملیاتی"],
    "dps_approved":["سود سهام مصوب"],
    "dps_proposed":["سود سهام پیشنهادی هیئت مدیره"],
    "ni_last_year":["مبلغ سود (زیان) خالص سال مالی گذشته"],
}


# ═══════════════ نسبت‌ها (EPS از NI محاسبه می‌شود) ═══════════════
def _compute_eps(d, ni_a, sh, f, m):
    """EPS را از NI محاسبه کن — قابل اطمینان‌ترین."""
    if not pd.isna(ni_a) and sh and sh > 0:
        return safe_div(ni_a * 1e6, sh)
    eps_file = d.get("eps", np.nan)
    return eps_file * f if not pd.isna(eps_file) else np.nan


def bank_ratios(d, p, sh, m=6):
    f = 12/m
    ni = d.get("net_income", np.nan); ni_a = ni*f if not pd.isna(ni) else np.nan
    eq = d.get("equity", np.nan); ta = d.get("total_assets", np.nan)
    eps_a = _compute_eps(d, ni_a, sh, f, m)
    loans = d.get("total_loans", np.nan)
    if pd.isna(loans) or loans == 0:
        loans = safe_add(d.get("loans_gov",0) or 0, d.get("loans_nongov",0) or 0)
    bvps = safe_div(eq*1e6, sh)
    return {
        "EPS سالانه (ریال)": eps_a, "BVPS (ریال)": bvps,
        "ROE سالانه (%)": safe_div(ni_a, eq)*100, "ROA سالانه (%)": safe_div(ni_a, ta)*100,
        "Cost-to-Income (%)": safe_div(abs(d.get("admin_cost", np.nan)), d.get("op_income", np.nan))*100,
        "حاشیه سود خالص (%)": safe_div(ni, d.get("op_income", np.nan))*100,
        "LDR (%)": safe_div(loans, d.get("deposits", np.nan))*100,
        "پوشش ذخیره (%)": safe_div(d.get("npl_doubtful", np.nan), loans)*100,
        "نقد/دارایی (%)": safe_div(d.get("cash", np.nan), ta)*100,
        "بدهی/دارایی (%)": safe_div(d.get("total_liab", np.nan), ta)*100,
        "P/E": safe_div(p, eps_a), "P/B": safe_div(p, bvps),
        "Earnings Yield (%)": safe_div(eps_a, p)*100,
        "Market Cap (همت)": safe_div(p*sh, 1e12),
    }

def generic_ratios(d, p, sh, m=6):
    f = 12/m; rev = d.get("revenue", np.nan)
    cogs = abs(d["cogs"]) if not pd.isna(d.get("cogs", np.nan)) else np.nan
    ni = d.get("net_income", np.nan); ni_a = ni*f if not pd.isna(ni) else np.nan
    eq = d.get("equity", np.nan); ta = d.get("total_assets", np.nan)
    ar = d.get("receivables", np.nan); inv = d.get("inventory", np.nan)
    ap = d.get("payables", np.nan); ca = d.get("current_assets", np.nan); cl = d.get("current_liab", np.nan)
    debt = d.get("total_debt", np.nan); cfo = d.get("cfo", np.nan)
    capex = abs(d["capex"]) if not pd.isna(d.get("capex", np.nan)) else 0
    eps_a = _compute_eps(d, ni_a, sh, f, m)
    bvps = safe_div(eq*1e6, sh); mcap = p*sh; days = 30*m
    dso = safe_div(ar, rev)*days; dio = safe_div(inv, cogs)*days; dpo = safe_div(ap, cogs)*days
    ccc = (dso+dio-dpo) if not (pd.isna(dso) or pd.isna(dio) or pd.isna(dpo)) else np.nan
    return {
        "حاشیه سود ناخالص (%)": safe_div(safe_add(rev, -(cogs or 0)), rev)*100 if not pd.isna(rev) else np.nan,
        "حاشیه سود عملیاتی (%)": safe_div(d.get("op_profit", np.nan), rev)*100,
        "حاشیه سود خالص (%)": safe_div(ni, rev)*100,
        "ROE (%)": safe_div(ni_a, eq)*100, "ROA (%)": safe_div(ni_a, ta)*100,
        "نسبت جاری": safe_div(ca, cl), "Debt/Equity":
