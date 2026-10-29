"""داشبورد تحلیل سهام — فارسی، راست‌چین، تحلیل سهامدار."""
import re
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="داشبورد تحلیل سهام", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

# ═══════════════════ RTL + IranYekan ═══════════════════
st.markdown("""
<style>
@font-face {
  font-family: 'IranYekan';
  src: url('https://cdn.jsdelivr.net/gh/font-store/font-iranYekan/IRANYekanWeb-Regular.woff2') format('woff2');
  font-weight: 400; font-display: swap;
}
@font-face {
  font-family: 'B Nazanin';
  src: url('https://cdn.jsdelivr.net/gh/font-store/font-bnazanin/BNazanin.woff2') format('woff2');
  font-display: swap;
}
@import url('https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css');

html, body, [class*="css"], .stApp, .stMarkdown, .stMetric, button, input, select, textarea, table {
  font-family: 'IranYekan', 'Vazirmatn', 'B Nazanin', Tahoma, sans-serif !important;
}
.stApp { direction: rtl; text-align: right; }
[data-testid="stSidebar"], [data-testid="stSidebar"] * { direction: rtl; text-align: right; }
[data-testid="stMetricValue"], [data-testid="stMetricLabel"] { direction: rtl; }
[data-testid="metric-container"] {
  background:#f8fafc; border-radius:12px; padding:14px;
  border-right:4px solid #2563eb;
}
h1, h2, h3, h4 { color:#0f172a; direction: rtl; }
.warn-box {
  background:#fef3c7; border-right:4px solid #f59e0b;
  padding:12px; border-radius:8px; margin:8px 0; direction: rtl;
}
.good-box {
  background:#dcfce7; border-right:4px solid #16a34a;
  padding:12px; border-radius:8px; margin:8px 0; direction: rtl;
}
.bad-box {
  background:#fee2e2; border-right:4px solid #dc2626;
  padding:12px; border-radius:8px; margin:8px 0; direction: rtl;
}
/* اعداد راست‌به‌چپ شکن */
[data-testid="stMetricValue"] { font-size: 1.5rem; font-weight: 700; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════ توابع کمکی ═══════════════════
PERSIAN = "۰۱۲۳۴۵۶۷۸۹"; ARABIC = "٠١٢٣٤٥٦٧٨٩"
def to_en(s):
    for i, ch in enumerate(PERSIAN): s = s.replace(ch, str(i))
    for i, ch in enumerate(ARABIC):  s = s.replace(ch, str(i))
    return s
def norm_label(s):
    s = to_en(str(s)).replace("ي","ی").replace("ك","ک").replace("\u200c","")
    return re.sub(r"\s+","",s).strip()
def parse_num(x):
    if x is None or (isinstance(x,float) and np.isnan(x)): return np.nan
    if isinstance(x,(int,float)): return float(x)
    s = to_en(str(x)).strip()
    if not s or s in {"-","*"}: return np.nan
    neg = s.startswith("(") and s.endswith(")")
    s = s.replace(",","").replace("،","").replace(" ","").replace("%","").strip("()")
    try:
        v = float(s); return -v if neg else v
    except: return np.nan
def safe_div(a,b):
    try:
        if b in (0,None) or pd.isna(a) or pd.isna(b): return np.nan
        return a/b
    except: return np.nan
def safe_add(*vs):
    return np.nan if any(pd.isna(v) for v in vs) else float(sum(vs))
def fmt(v, pct=False, dec=1, suffix=""):
    if v is None or pd.isna(v): return "—"
    return f"{v:,.{dec}f}٪{suffix}" if pct else f"{v:,.{dec}f}{suffix}"
def find_row(df, cands):
    if df.shape[1] < 2 or df.empty: return None
    col_a = df.iloc[:,0].astype(str).map(norm_label)
    for c in cands:
        m = col_a == norm_label(c)
        if m.any(): return df[m].iloc[0]
    for c in cands:
        cn = norm_label(c)
        if len(cn) < 4: continue
        m = col_a.str.contains(re.escape(cn), na=False)
        if m.any(): return df[m].iloc[0]
    return None
def extract_excel(file, labels):
    out = {}
    try: xls = pd.ExcelFile(file)
    except: return out
    for sh in xls.sheet_names:
        try: raw = pd.read_excel(xls, sheet_name=sh, header=None)
        except: continue
        for hdr in range(min(6, len(raw))):
            df = raw.iloc[hdr+1:].reset_index(drop=True)
            if df.shape[1] < 2 or df.empty: continue
            for k, cands in labels.items():
                if k in out and not pd.isna(out[k]): continue
                row = find_row(df, cands)
                if row is None: continue
                nums = []
                for c in df.columns[1:6]:
                    v = parse_num(row.get(c))
                    if not pd.isna(v): nums.append(v)
                    if len(nums) >= 2: break
                if nums:
                    out[k] = nums[0]
                    if len(nums) > 1: out[f"{k}_prev"] = nums[1]
    return out

# ═══════════════════ برچسب‌های صنعتی ═══════════════════
BANK_LABELS = {
    "rev_loans":    ["درآمد تسهیلات اعطایی", "درآمد تسهيلات اعطايي"],
    "rev_fees":     ["درآمد کارمزد"],
    "deposit_cost": ["هزینه سود سپرده‌ها", "هزينه سود سپرده‌ها"],
    "admin_cost":   ["هزینه های اداری و عمومی", "هزينه هاي اداري و عمومي"],
    "provision":    ["هزینه مطالبات مشکوک الوصول"],
    "op_income":    ["درآمد عملیاتی"],
    "op_profit":    ["سود(زیان) عملیاتی"],
    "net_income":   ["سود(زیان)خالص", "سود (زیان) خالص"],
    "eps":          ["سود (زیان) خالص هر سهم– ریال", "سود (زیان) خالص هر سهم"],
    "cash":         ["موجودی نقد", "موجودي نقد"],
    "total_assets": ["جمع دارایی‌ها", "جمع دارایی ها"],
    "total_liab":   ["جمع بدهی‌ها", "جمع بدهی ها"],
    "equity":       ["جمع حقوق مالکانه"],
    "deposits":     ["سپرده‌های مشتریان"],
    "loans_gov":    ["تسهیلات اعطایی و مطالبات از اشخاص دولتی"],
    "loans_nongov": ["تسهیلات اعطایی و مطالبات از اشخاص غیر دولتی- ریالی",
                     "تسهیلات اعطایی و مطالبات از اشخاص غیر دولتی"],
    "total_loans":  ["مانده تسهیلات پایان دوره"],
    "npl_doubtful": ["ذخیره اختصاصی مطالبات مشکوک الوصول"],
    "dps_approved": ["سود سهام مصوب مجمع سال مالی گذشته"],
    "dps_proposed": ["سود سهام پیشنهادی هیئت مدیره در سال مالی گذشته"],
    "ni_last_year": ["مبلغ سود (زیان) خالص سال مالی گذشته"],
}
GENERIC_LABELS = {
    "revenue":        ["فروش", "درآمدهای عملیاتی"],
    "cogs":           ["بهای تمام شده کالای فروش رفته", "بهای تمام شده"],
    "sga":            ["هزینه های فروش، اداری و عمومی", "هزینه های اداری و عمومی"],
    "op_profit":      ["سود(زیان) عملیاتی"],
    "fin_cost":       ["هزینه های مالی", "هزینه‌های مالی"],
    "net_income":     ["سود(زیان)خالص", "سود (زیان) خالص"],
    "eps":            ["سود (زیان) خالص هر سهم– ریال", "سود (زیان) خالص هر سهم", "سود هر سهم"],
    "cash":           ["موجودی نقد", "موجودي نقد"],
    "receivables":    ["حساب ها و اسناد دریافتنی تجاری"],
    "inventory":      ["موجودی مواد و کالا"],
    "current_assets": ["جمع دارایی‌های جاری"],
    "total_assets":   ["جمع دارایی‌ها", "جمع دارایی ها"],
    "payables":       ["حساب ها و اسناد پرداختنی تجاری"],
    "current_liab":   ["جمع بدهی‌های جاری"],
    "total_debt":     ["تسهیلات مالی دریافتی", "جمع بدهی‌ها"],
    "equity":         ["جمع حقوق مالکانه"],
    "cfo":            ["جریان خالص ورود (خروج) نقد حاصل از فعالیت های عملیاتی",
                       "جريان خالص ورود (خروج) نقد حاصل از فعاليت هاي عملياتي"],
    "capex":          ["پرداخت های نقدی برای تحصیل دارایی ثابت مشهود"],
    "dps_approved":   ["سود سهام مصوب مجمع سال مالی گذشته"],
    "ni_last_year":   ["مبلغ سود (زیان) خالص سال مالی گذشته"],
}
PHARMA_LABELS = {**GENERIC_LABELS, "rd": ["هزینه تحقیق و توسعه", "هزینه های تحقیق و توسعه"]}

# ═══════════════════ نسبت‌ها ═══════════════════
def bank_ratios(d, p, sh, m=6):
    f = 12/m; ni = d.get("net_income", np.nan); ni_a = ni*f if not pd.isna(ni) else np.nan
    eq = d.get("equity", np.nan); ta = d.get("total_assets", np.nan)
    eps = d.get("eps", np.nan); eps_a = eps*f if not pd.isna(eps) else safe_div(ni_a*1e6, sh)
    loans = d.get("total_loans", np.nan)
    if pd.isna(loans) or loans == 0:
        loans = safe_add(d.get("loans_gov",0) or 0, d.get("loans_nongov",0) or 0)
    bvps = safe_div(eq*1e6, sh)
    r = {}
    r["EPS سالانه (ریال)"]   = eps_a
    r["BVPS (ریال)"]         = bvps
    r["ROE سالانه (%)"]      = safe_div(ni_a, eq)*100
    r["ROA سالانه (%)"]      = safe_div(ni_a, ta)*100
    r["Cost-to-Income (%)"]  = safe_div(abs(d.get("admin_cost",np.nan)), d.get("op_income",np.nan))*100
    r["حاشیه سود خالص (%)"]  = safe_div(ni, d.get("op_income",np.nan))*100
    r["LDR (%)"]             = safe_div(loans, d.get("deposits",np.nan))*100
    r["پوشش ذخیره (%)"]      = safe_div(d.get("npl_doubtful",np.nan), loans)*100
    r["نقد/دارایی (%)"]      = safe_div(d.get("cash",np.nan), ta)*100
    r["بدهی/دارایی (%)"]     = safe_div(d.get("total_liab",np.nan), ta)*100
    r["P/E"]                 = safe_div(p, eps_a)
    r["P/B"]                 = safe_div(p, bvps)
    r["Earnings Yield (%)"]  = safe_div(eps_a, p)*100
    r["Market Cap (همت)"]    = safe_div(p*sh, 1e12)
    return r

def generic_ratios(d, p, sh, m=6):
    f = 12/m; rev = d.get("revenue", np.nan)
    cogs = abs(d["cogs"]) if not pd.isna(d.get("cogs",np.nan)) else np.nan
    ni = d.get("net_income", np.nan); ni_a = ni*f if not pd.isna(ni) else np.nan
    eq = d.get("equity", np.nan); ta = d.get("total_assets", np.nan)
    ar = d.get("receivables", np.nan); inv = d.get("inventory", np.nan)
    ap = d.get("payables", np.nan); ca = d.get("current_assets", np.nan); cl = d.get("current_liab", np.nan)
    debt = d.get("total_debt", np.nan); cfo = d.get("cfo", np.nan)
    capex = abs(d["capex"]) if not pd.isna(d.get("capex",np.nan)) else 0
    eps_a = safe_div(ni_a*1e6, sh); bvps = safe_div(eq*1e6, sh)
    mcap = p*sh; days = 30*m
    dso = safe_div(ar,rev)*days; dio = safe_div(inv,cogs)*days; dpo = safe_div(ap,cogs)*days
    ccc = dso+dio-dpo if not (pd.isna(dso) or pd.isna(dio) or pd.isna(dpo)) else np.nan
    r = {}
    r["حاشیه سود ناخالص (%)"]  = safe_div(safe_add(rev, -(cogs or 0)), rev)*100 if not pd.isna(rev) else np.nan
    r["حاشیه سود عملیاتی (%)"] = safe_div(d.get("op_profit",np.nan), rev)*100
    r["حاشیه سود خالص (%)"]    = safe_div(ni, rev)*100
    r["ROE (%)"]               = safe_div(ni_a, eq)*100
    r["ROA (%)"]               = safe_div(ni_a, ta)*100
    r["نسبت جاری"]             = safe_div(ca, cl)
    r["Debt/Equity"]           = safe_div(debt, eq)
    r["Interest coverage"]     = safe_div(d.get("op_profit",np.nan), abs(d["fin_cost"]) if not pd.isna(d.get("fin_cost",np.nan)) else np.nan)
    r["گردش موجودی (بار)"]     = safe_div(cogs*f, inv)
    r["DSO (روز)"]             = dso
    r["DIO (روز)"]             = dio
    r["DPO (روز)"]             = dpo
    r["CCC (روز)"]             = ccc
    r["FCF Yield (%)"]         = safe_div((cfo*f - abs(capex*f))*1e6, mcap)*100
    r["EPS سالانه (ریال)"]     = eps_a
    r["BVPS (ریال)"]           = bvps
    r["P/E"]                   = safe_div(p, eps_a) if eps_a and eps_a > 0 else np.nan
    r["P/B"]                   = safe_div(p, bvps)
    r["P/S"]                   = safe_div(mcap, rev*f*1e6)
    r["Market Cap (همت)"]      = safe_div(mcap, 1e12)
    return r

def petro_ratios(d, p, sh, m=6):
    r = generic_ratios(d, p, sh, m); f = 12/m
    ebitda = safe_add(d.get("op_profit",np.nan), safe_mul(d.get("revenue",np.nan), 0.05)) if not pd.isna(d.get("op_profit",np.nan)) else np.nan
    net_debt = safe_add(d.get("total_debt",0) or 0, -(d.get("cash",0) or 0))
    mcap = p*sh; ev = safe_add(mcap, net_debt*1e6)
    r["EV/EBITDA"] = safe_div(ev, ebitda*f*1e6)
    return r

def pharma_ratios(d, p, sh, m=6):
    r = generic_ratios(d, p, sh, m)
    r["R&D / فروش (%)"] = safe_div(d.get("rd",np.nan), d.get("revenue",np.nan))*100
    return r

def safe_mul(a,b):
    return np.nan if pd.isna(a) or pd.isna(b) else a*b

SECTOR_CONFIG = {
    "bank": {"name":"بانک و مؤسسه مالی", "labels":BANK_LABELS, "fn":bank_ratios,
             "primary":["P/B","P/E","ROE سالانه (%)","Cost-to-Income (%)","LDR (%)"],
             "secondary":["ROA سالانه (%)","پوشش ذخیره (%)","حاشیه سود خالص (%)","Earnings Yield (%)","Market Cap (همت)"],
             "benchmark":{"P/E":5.8,"P/B":1.2,"ROE سالانه (%)":25}},
    "petro": {"name":"پتروشیمی و شیمیایی", "labels":GENERIC_LABELS, "fn":petro_ratios,
              "primary":["P/E","EV/EBITDA","حاشیه سود ناخالص (%)","ROE (%)","FCF Yield (%)"],
              "secondary":["P/B","P/S","ROA (%)","Debt/Equity","Interest coverage"],
              "benchmark":{"P/E":6.5,"P/B":1.5,"ROE (%)":20}},
    "pharma": {"name":"دارویی", "labels":PHARMA_LABELS, "fn":pharma_ratios,
               "primary":["P/E","P/B","حاشیه سود ناخالص (%)","R&D / فروش (%)","ROE (%)"],
               "secondary":["ROA (%)","DSO (روز)","DIO (روز)","CCC (روز)","FCF Yield (%)"],
               "benchmark":{"P/E":8.0,"P/B":2.5,"ROE (%)":25}},
    "auto": {"name":"خودرو و قطعات", "labels":GENERIC_LABELS, "fn":generic_ratios,
             "primary":["P/E","حاشیه سود ناخالص (%)","Debt/Equity","Interest coverage","ROE (%)"],
             "secondary":["گردش موجودی (بار)","DSO (روز)","FCF Yield (%)","P/B","ROA (%)"],
             "benchmark":{"P/E":5.0,"P/B":1.0}},
    "generic": {"name":"عمومی (فولاد، سیمان، غذایی...)","labels":GENERIC_LABELS,"fn":generic_ratios,
                "primary":["P/E","P/B","ROE (%)","ROA (%)","حاشیه سود ناخالص (%)"],
                "secondary":["P/S","Debt/Equity","نسبت جاری","CCC (روز)","FCF Yield (%)"],
                "benchmark":{"P/E":5.5,"P/B":1.3,"ROE (%)":20}},
}

DEMO_BANK = {
    "rev_loans":1_835_436,"rev_loans_prev":1_259_107,
    "deposit_cost":1_220_160,"admin_cost":707_453,"admin_cost_prev":333_720,
    "provision":43_432,"op_income":2_150_762,"op_income_prev":1_325_688,
    "op_profit":1_271_937,"net_income":1_025_278,"net_income_prev":525_849,"eps":303,
    "cash":3_536_996,"total_assets":59_110_305,"total_liab":52_628_589,
    "equity":6_481_715,"deposits":27_590_264,"deposits_prev":22_270_416,
    "loans_gov":28_271_151,"loans_nongov":16_938_084,
    "total_loans":47_324_584,"npl_doubtful":1_438_867,
    "dps_approved":155_572,"dps_proposed":153_538,"ni_last_year":1_535_376,
}

# ═══════════════════ واژه‌نامه ═══════════════════
GLOSSARY = [
    ("P/E", "Price to Earnings — نسبت قیمت هر سهم به سود سالانه هر سهم. زیر ۵ = ارزان، بالای ۱۰ = گران"),
    ("P/B", "Price to Book — نسبت قیمت هر سهم به ارزش دفتری هر سهم. زیر ۱ = زیر ارزش دفتری"),
    ("ROA", "Return on Assets — بازده دارایی؛ سود خالص تقسیم بر کل دارایی. بالای ۳٪ در بانک خوب است"),
    ("ROE", "Return on Equity — بازده حقوق مالکانه؛ سود خالص تقسیم بر حقوق مالکانه. بالای ۲۰٪ عالی"),
    ("EPS", "Earnings per Share — سود هر سهم؛ سود خالص تقسیم بر تعداد سهام"),
    ("BVPS", "Book Value per Share — ارزش دفتری هر سهم؛ حقوق مالکانه تقسیم بر تعداد سهام"),
    ("DPS", "Dividend per Share — سود نقدی هر سهم؛ بخشی از سود که نقداً به سهامدار پرداخت می‌شود"),
    ("TSR", "Total Shareholder Return — بازده کل سهامدار؛ سود سرمایه + سود نقدی"),
    ("Payout Ratio", "نسبت تقسیم سود = DPS ÷ EPS؛ نشان می‌دهد چه سهمی از سود نقد می‌شود"),
    ("Dividend Yield", "بازده سود نقدی = DPS ÷ قیمت سهم؛ بازده نقدی سالانه سهامدار"),
    ("LDR", "Loan to Deposit Ratio — نسبت تسهیلات به سپرده. بالای ۱۰۰٪ = فشار نقدینگی بانک"),
    ("NPL", "Non-Performing Loans — مطالبات غیرجاری (معوق + مشکوک‌الوصول)"),
    ("NIM", "Net Interest Margin — حاشیه سود خالص تسهیلات؛ تفاوت نرخ سود دریافتی و پرداختی"),
    ("CASA", "Current Account Savings Account — سپرده جاری و پس‌انداز (ارزان‌ترین منبع تأمین مالی)"),
    ("Cost-to-Income", "نسبت هزینه عملیاتی به درآمد عملیاتی؛ پایین‌تر = کارایی بهتر"),
    ("EBITDA", "Earnings Before Interest, Taxes, Depreciation & Amortization — سود قبل از بهره، مالیات و استهلاک"),
    ("EV", "Enterprise Value — ارزش شرکت = ارزش بازار + بدهی خالص"),
    ("FCF", "Free Cash Flow — جریان نقد آزاد = نقد عملیاتی منهای سرمایه‌گذاری"),
    ("DSO", "Days Sales Outstanding — دوره وصول مطالبات (میانگین روز تا دریافت پول از مشتری)"),
    ("DIO", "Days Inventory Outstanding — دوره گردش موجودی (میانگین روز نگهداری کالا در انبار)"),
    ("DPO", "Days Payable Outstanding — دوره پرداخت بدهی (میانگین روز تا پرداخت به تأمین‌کننده)"),
    ("CCC", "Cash Conversion Cycle — چرخه تبدیل نقد = DSO + DIO − DPO"),
    ("CAGR", "Compound Annual Growth Rate — نرخ رشد مرکب سالانه"),
    ("Target Price", "قیمت هدف — پیش‌بینی قیمت سهم بر اساس نسبت‌های صنعت و سود"),
    ("Graham Formula", "فرمول بنجامین گراهام برای ارزش ذاتی = √(22.5 × EPS × BVPS)"),
    ("Bull/Base/Bear", "سناریوهای خوش‌بینانه / پایه / بدبینانه برای آینده سهم"),
    ("Margin", "حاشیه سود — درصد سود نسبت به فروش"),
    ("Fair Value", "ارزش منصفانه — قیمتی که منطقاً باید سهم داشته باشد"),
]

# ═══════════════════ سایدبار ═══════════════════
st.sidebar.header("⚙️ تنظیمات")
sk = st.sidebar.selectbox("صنعت", list(SECTOR_CONFIG.keys()), format_func=lambda k: SECTOR_CONFIG[k]["name"])
cfg = SECTOR_CONFIG[sk]
uploaded = st.sidebar.file_uploader("صورت‌های مالی (اکسل کدال)", type=["xlsx","xls"])

st.sidebar.subheader("💰 پارامترهای بازار")
price = st.sidebar.number_input("قیمت روز سهم (تومان)", min_value=1.0, value=143.0 if sk=="bank" else 100.0, step=1.0)
shares = st.sidebar.number_input("تعداد سهام (میلیارد)", min_value=0.01, value=3.382 if sk=="bank" else 1.0, step=0.01, format="%.3f")
months = st.sidebar.selectbox("دوره صورت مالی (ماه)", [3,6,9,12], index=1)
dollar = st.sidebar.number_input("نرخ دلار (ریال) — اختیاری", min_value=0.0, value=700000.0, step=1000.0)

# ═══════════════════ پارس ═══════════════════
data = {}
if uploaded:
    with st.spinner("در حال پارس فایل..."):
        data = extract_excel(uploaded, cfg["labels"])
if sk == "bank" and not data:
    data = dict(DEMO_BANK)
    st.info("ℹ️ داده دمو (وبملت) بارگذاری شد. برای سهم دیگر، فایل را آپلود یا در تب «داده» وارد کنید.")

ratios = cfg["fn"](data, p=price*10, sh=shares*1e9, m=months)

# ═══════════════════ هدر ═══════════════════
st.title(f"📊 داشبورد تحلیل — {cfg['name']}")
st.caption(f"قیمت: **{price:,.0f} تومان** | سهام: **{shares:.2f} میلیارد** | دوره: **{months} ماه** | دلار: **{dollar:,.0f} ریال**")

# پرچم‌های هشدار
if not pd.isna(data.get("net_income")) and data.get("net_income",0) < 0:
    st.markdown("<div class='bad-box'>🔴 شرکت در دوره جاری زیان‌ده است — P/E معنا ندارد.</div>", unsafe_allow_html=True)

# نوار پوشش
req = [k for k in cfg["labels"] if not k.endswith("_prev")]
found = [k for k in req if not pd.isna(data.get(k))]
miss = [k for k in req if pd.isna(data.get(k))]
cov = int(len(found)/len(req)*100) if req else 0
c1, c2 = st.columns([3,1])
c1.progress(cov, text=f"پوشش داده: {len(found)}/{len(req)} فیلد ({cov}٪)")
c2.caption("⚠️ برخی فیلدها گم — تب «داده» را چک کنید." if miss else "✅ همه فیلدها موجود.")

# ═══════════════════ تب‌ها ═══════════════════
T = st.tabs(["🏠 خلاصه", "📈 سود و زیان", "🏦 ترازنامه", "💰 ارزش‌گذاری",
             "👤 تحلیل سهامدار", "🔧 داده", "📖 واژه‌نامه"])

# ── ۱. خلاصه ──
with T[0]:
    st.subheader("شاخص‌های کلیدی عملکرد")
    cols = st.columns(5)
    for i,k in enumerate(cfg["primary"]):
        cols[i%5].metric(k, fmt(ratios.get(k), pct="%" in k or "Yield" in k, dec=2))
    st.divider()
    cols = st.columns(5)
    for i,k in enumerate(cfg["secondary"]):
        cols[i%5].metric(k, fmt(ratios.get(k), pct="%" in k or "Yield" in k, dec=2))

    # پرچم‌های کیفیت
    flags = []
    if ratios.get("Cost-to-Income (%)", 0) > 45:
        flags.append("🟡 نسبت هزینه به درآمد بالاست — کارایی پایین")
    if ratios.get("LDR (%)", 0) > 100:
        flags.append("🟡 LDR بالای ۱۰۰٪ — فشار نقدینگی")
    if sk == "bank" and ratios.get("پوشش ذخیره (%)", 0) < 60:
        flags.append("🟡 پوشش ذخیره مطالبات پایین")
    if not pd.isna(ratios.get("P/E")) and 0 < ratios.get("P/E", 99) < 4:
        flags.append("🟢 P/E پایین — احتمال ارزندگی")
    if not pd.isna(ratios.get("P/B")) and 0 < ratios.get("P/B", 99) < 1:
        flags.append("🟢 P/B زیر ۱ — معامله زیر ارزش دفتری")
    if flags:
        st.subheader("🚩 پرچم‌های تحلیلی")
        for f in flags: st.write(f)

# ── ۲. سود و زیان ──
with T[1]:
    if sk == "bank":
        oi = data.get("op_income",0)/1e3; dc = abs(data.get("deposit_cost",0))/1e3
        ac = abs(data.get("admin_cost",0))/1e3; pv = abs(data.get("provision",0))/1e3
        ni = data.get("net_income",0)/1e3; other = ni-(oi-dc-ac-pv)
        x = ["درآمد عملیاتی","هزینه سود سپرده","هزینه اداری","ذخیره مطالبات","سایر اقلام خالص","سود خالص"]
        y = [oi,-dc,-ac,-pv,other,ni]
        meas = ["relative","relative","relative","relative","relative","total"]
    else:
        rv = data.get("revenue",0)/1e3; cs = abs(data.get("cogs",0))/1e3
        sg = abs(data.get("sga",0))/1e3; fc = abs(data.get("fin_cost",0))/1e3
        ni = data.get("net_income",0)/1e3; gp = rv-cs; other = ni-gp+sg+fc
        x = ["فروش","بهای تمام شده","سود ناخالص","هزینه فروش/اداری","هزینه مالی","سایر/مالیات","سود خالص"]
        y = [rv,-cs,gp,-sg,-fc,other,ni]
        meas = ["relative","relative","total","relative","relative","relative","total"]
    fig = go.Figure(go.Waterfall(measure=meas, x=x, y=y,
        text=[f"{v:,.0f}" for v in y], textposition="outside",
        connector={"line":{"color":"#94a3b8"}},
        increasing={"marker":{"color":"#16a34a"}},
        decreasing={"marker":{"color":"#dc2626"}},
        totals={"marker":{"color":"#2563eb"}}))
    fig.update_layout(height=440, showlegend=False, yaxis_title="همت",
                      title="پل سود (همت) — «سایر اقلام» = اقلام باقی‌مانده تا تطبیق")
    st.plotly_chart(fig, use_container_width=True)
    st.caption("💡 اگر «سایر اقلام» بزرگ باشد، سود از فعالیت اصلی نمی‌آید — پرچم کیفیت سود.")

# ── ۳. ترازنامه ──
with T[2]:
    ta = data.get("total_assets",0)/1e3; cash = data.get("cash",0)/1e3
    if sk == "bank":
        loans = data.get("total_loans",0)/1e3
        labels_a = ["نقد","تسهیلات","سایر"]
        vals_a = [cash, loans, max(0,ta-cash-loans)]
    else:
        inv = data.get("inventory",0)/1e3; ar = data.get("receivables",0)/1e3
        labels_a = ["نقد","موجودی","مطالبات","سایر"]
        vals_a = [cash, inv, ar, max(0,ta-cash-inv-ar)]
    c1, c2 = st.columns(2)
    with c1:
        fig = px.pie(values=vals_a, names=labels_a, hole=0.55,
                     title="ترکیب دارایی‌ها (همت)",
                     color_discrete_sequence=px.colors.sequential.Blues_r)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        eq = data.get("equity",0)/1e3; liab = data.get("total_liab",0)/1e3
        fig = px.pie(values=[liab,eq], names=["بدهی","حقوق مالکانه"], hole=0.55,
                     title="ساختار سرمایه (همت)", color_discrete_sequence=["#dc2626","#16a34a"])
        st.plotly_chart(fig, use_container_width=True)

# ── ۴. ارزش‌گذاری ──
with T[3]:
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("P/E", fmt(ratios.get("P/E"), dec=2))
    c2.metric("P/B", fmt(ratios.get("P/B"), dec=2))
    if "EV/EBITDA" in ratios: c3.metric("EV/EBITDA", fmt(ratios.get("EV/EBITDA"), dec=2))
    else: c3.metric("Earnings Yield", fmt(ratios.get("Earnings Yield (%)"), pct=True))
    c4.metric("Market Cap", fmt(ratios.get("Market Cap (همت)"), dec=0)+" همت")

    st.divider()
    st.subheader("🎯 حساسیت P/E به قیمت")
    eps = ratios.get("EPS سالانه (ریال)")
    if eps and not pd.isna(eps) and eps > 0:
        prices = np.linspace(price*0.5, price*2.0, 40)
        sens = pd.DataFrame({"قیمت (تومان)":prices, "P/E":[(p*10)/eps for p in prices]})
        fig = px.line(sens, x="قیمت (تومان)", y="P/E", markers=True)
        fig.add_hline(y=5, line_dash="dash", line_color="orange", annotation_text="P/E=5")
        fig.add_hline(y=7, line_dash="dash", line_color="red", annotation_text="P/E=7")
        fig.add_vline(x=price, line_dash="dot", line_color="green",
                      annotation_text=f"قیمت {price:.0f}")
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("EPS منفی/صفر — منحنی P/E معنا ندارد.")

    if cfg["benchmark"]:
        st.subheader("📊 مقایسه با بنچمارک صنعت")
        rows = []
        for k,bv in cfg["benchmark"].items():
            sv = ratios.get(k, np.nan)
            if not pd.isna(sv): rows.append({"شاخص":k, "این سهم":sv, "میانه صنعت":bv})
        if rows:
            df = pd.DataFrame(rows)
            fig = px.bar(df.melt(id_vars="شاخص", var_name="مبنا", value_name="مقدار"),
                         x="شاخص", y="مقدار", color="مبنا", barmode="group", text_auto=".2f")
            fig.update_layout(height=380)
            st.plotly_chart(fig, use_container_width=True)

# ── ۵. تحلیل سهامدار (جدید) ──
with T[4]:
    st.header("👤 تحلیل از دید سهامدار")
    st.caption("این تحلیل مبتنی بر داده‌های تاریخی و مفروضات است — **توصیه سرمایه‌گذاری نیست**.")

    # ─── DPS و بازده نقدی ───
    eps_rial = ratios.get("EPS سالانه (ریال)", np.nan)
    bvps = ratios.get("BVPS (ریال)", np.nan)
    price_rial = price * 10
    shares_n = shares * 1e9

    dps_hist = np.nan
    if data.get("dps_approved") and not pd.isna(data.get("dps_approved")):
        dps_hist = data["dps_approved"] * 1e6 / shares_n  # rial per share
    payout = np.nan
    if data.get("dps_approved") and data.get("ni_last_year"):
        payout = safe_div(data["dps_approved"], data["ni_last_year"]) * 100
    div_yield = safe_div(dps_hist, price_rial) * 100 if not pd.isna(dps_hist) else 0

    st.subheader("💵 سود نقدی و بازده سهامدار")
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("DPS تاریخی (ریال)", fmt(dps_hist, dec=0))
    c2.metric("بازده سود نقدی", fmt(div_yield, pct=True))
    c3.metric("Payout Ratio", fmt(payout, pct=True))
    c4.metric("EPS سالانه", fmt(eps_rial, dec=0))
    c5.metric("BVPS", fmt(bvps, dec=0))

    if div_yield < 2:
        st.markdown("<div class='warn-box'>⚠️ بازده سود نقدی زیر ۲٪ — بازده اصلی سهامدار فقط از رشد قیمت است.</div>", unsafe_allow_html=True)
    elif div_yield > 5:
        st.markdown("<div class='good-box'>✅ بازده سود نقدی بالای ۵٪ — جریان نقدی مناسب برای سهامدار.</div>", unsafe_allow_html=True)

    # ─── قیمت هدف ───
    st.subheader("🎯 قیمت هدف (سه روش)")
    peer_pe = cfg["benchmark"].get("P/E", 5.5)
    peer_pb = cfg["benchmark"].get("P/B", 1.3)
    tp_pe = (eps_rial * peer_pe / 10) if eps_rial and eps_rial > 0 else np.nan
    tp_pb = (bvps * peer_pb / 10) if bvps and bvps > 0 else np.nan
    tp_graham = np.sqrt(22.5 * eps_rial * bvps) / 10 if eps_rial and bvps and eps_rial > 0 and bvps > 0 else np.nan
    targets = [t for t in [tp_pe, tp_pb, tp_graham] if not pd.isna(t)]
    target_avg = np.mean(targets) if targets else np.nan

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("P/E محور", fmt(tp_pe, dec=0)+" تومان" if not pd.isna(tp_pe) else "—")
    c2.metric("P/B محور", fmt(tp_pb, dec=0)+" تومان" if not pd.isna(tp_pb) else "—")
    c3.metric("Graham", fmt(tp_graham, dec=0)+" تومان" if not pd.isna(tp_graham) else "—")
    if not pd.isna(target_avg):
        upside = (target_avg - price) / price * 100
        c4.metric("میانگین هدف", f"{target_avg:,.0f} تومان",
                  delta=f"{upside:+.1f}%", delta_color="normal" if upside > 0 else "inverse")

    # ─── TSR ۱ساله ───
    if not pd.isna(target_avg):
        cap_gain = (target_avg - price) / price * 100
        tsr_1y = cap_gain + (div_yield or 0)
        st.subheader("💎 بازده کل ۱ساله سهامدار (TSR)")
        c1,c2,c3 = st.columns(3)
        c1.metric("سود سرمایه", f"{cap_gain:+.1f}%")
        c2.metric("سود نقدی", f"{div_yield or 0:+.1f}%")
        c3.metric("TSR کل", f"{tsr_1y:+.1f}%",
                  delta_color="normal" if tsr_1y > 0 else "inverse")

    # ─── سناریوها ───
    st.subheader("🔮 سناریوهای ۱۲ماهه")
    eps_prev = data.get("eps_prev", np.nan)
    if eps_prev and not pd.isna(eps_prev) and eps_prev > 0 and eps_rial:
        eps_growth = (eps_rial - eps_prev) / eps_prev * 100
    else:
        eps_growth = 15  # پیش‌فرض محافظه‌کارانه

    scenarios = []
    for name, gf, pf in [("🐂 خوش‌بینانه", 1.2, 1.2), ("📊 پایه", 1.0, 1.0), ("🐻 بدبینانه", 0.5, 0.8)]:
        eps_new = eps_rial * (1 + eps_growth * gf / 100) if eps_rial else np.nan
        pe_target = peer_pe * pf
        tp = eps_new * pe_target / 10 if eps_new else np.nan
        ret = (tp - price) / price * 100 + (div_yield or 0) if tp else np.nan
        scenarios.append({"سناریو": name, "رشد EPS%": round(eps_growth*gf,1),
                          "P/E هدف": round(pe_target,1), "EPS جدید": eps_new,
                          "قیمت هدف (تومان)": tp, "TSR%": ret})
    df_scen = pd.DataFrame(scenarios)
    st.dataframe(df_scen, use_container_width=True, hide_index=True)

    fig = px.bar(df_scen, x="سناریو", y="قیمت هدف (تومان)", color="سناریو", text_auto=".0f",
                 color_discrete_sequence=["#16a34a","#2563eb","#dc2626"])
    fig.add_hline(y=price, line_dash="dash", line_color="gray",
                  annotation_text=f"قیمت فعلی {price:.0f}")
    fig.update_layout(height=380, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

    # ─── چشم‌انداز ۳ساله ───
    st.subheader("📅 چشم‌انداز ۳ساله (اگر روند فعلی ادامه یابد)")
    if eps_rial and not pd.isna(eps_rial) and eps_rial > 0:
        eps_3y = eps_rial * (1 + eps_growth/100) ** 3
        tp_3y = eps_3y * peer_pe / 10
        cap_gain_3y = (tp_3y - price) / price * 100
        div_total_3y = (div_yield or 0) * 3
        tsr_3y = cap_gain_3y + div_total_3y
        c1,c2,c3 = st.columns(3)
        c1.metric("EPS سال ۳", fmt(eps_3y, dec=0)+" ریال")
        c2.metric("قیمت هدف سال ۳", fmt(tp_3y, dec=0)+" تومان")
        c3.metric("TSR تجمعی ۳ساله", f"{tsr_3y:+.1f}%",
                  delta_color="normal" if tsr_3y > 0 else "inverse")
        cagr_3y = ((tp_3y + sum([div_yield*price/100]*3)) / price) ** (1/3) - 1 if tp_3y else np.nan
        if not pd.isna(cagr_3y):
            st.caption(f"معادل بازده مرکب سالانه: **{cagr_3y*100:.1f}%** — این را با نرخ سود بانکی و تورم مقایسه کنید.")

    # ─── ریسک‌ها ───
    st.subheader("⚠️ ریسک‌های کلیدی سهامدار")
    risks = []
    if div_yield < 2: risks.append("🔴 سود نقدی ناچیز — بازده فقط از رشد قیمت")
    if payout and not pd.isna(payout) and payout < 15: risks.append("🟡 نسبت تقسیم سود پایین (Payout کمتر از ۱۵٪)")
    if ratios.get("LDR (%)", 0) > 100: risks.append("🟡 LDR بالای ۱۰۰٪ — فشار نقدینگی")
    if ratios.get("Cost-to-Income (%)", 0) > 45: risks.append("🟡 هزینه اداری بالا — فشردگی حاشیه سود")
    if sk == "bank" and ratios.get("پوشش ذخیره (%)", 0) < 60: risks.append("🟡 پوشش ذخیره مطالبات کم")
    if not pd.isna(ratios.get("P/E")) and ratios.get("P/E", 0) > 10: risks.append("🟡 P/E بالا — احتمال گران بودن")
    if not risks: risks.append("✅ ریسک قابل‌توجهی از داده‌ها شناسایی نشد")
    for r in risks: st.write(r)

    # ─── جمع‌بندی ───
    st.subheader("📋 جمع‌بندی تحلیلی")
    score = 0
    if not pd.isna(target_avg):
        if target_avg > price * 1.2: score += 2
        elif target_avg > price: score += 1
    if not pd.isna(ratios.get("P/E")) and ratios.get("P/E") < 5: score += 1
    if ratios.get("ROE سالانه (%)", ratios.get("ROE (%)", 0)) > 20: score += 1
    if (div_yield or 0) > 3: score += 1

    if score >= 4: verdict, color = "🟢 جذاب برای بررسی بیشتر", "good-box"
    elif score >= 2: verdict, color = "🟡 نیازمند بررسی دقیق‌تر", "warn-box"
    else: verdict, color = "🔴 پرریسک", "bad-box"
    st.markdown(f"<div class='{color}'><b>سیگنال تحلیلی (نه توصیه):</b> {verdict} — امتیاز {score}/6</div>", unsafe_allow_html=True)
    st.caption("⚠️ این تحلیل خودکار است و جایگزین مشاوره سرمایه‌گذاری حرفه‌ای نیست. همیشه با منبع اصلی کدال و مشاور رسمی چک کنید.")

# ── ۶. داده ──
with T[5]:
    st.subheader("✏️ ورود / ویرایش دستی (میلیون ریال)")
    rows = []
    for k, cands in cfg["labels"].items():
        cur = data.get(k, np.nan)
        prev = data.get(f"{k}_prev", np.nan)
        rows.append({"کلید":k, "شرح":cands[0],
                     "دوره جاری": None if pd.isna(cur) else float(cur),
                     "دوره قبل": None if pd.isna(prev) else float(prev)})
    df_edit = pd.DataFrame(rows)
    edited = st.data_editor(df_edit, use_container_width=True, height=450, num_rows="fixed",
        column_config={
            "کلید": st.column_config.TextColumn(disabled=True),
            "شرح": st.column_config.TextColumn(disabled=True),
            "دوره جاری": st.column_config.NumberColumn(format="%.0f"),
            "دوره قبل": st.column_config.NumberColumn(format="%.0f"),
        }, key="editor")
    if st.button("✅ اعمال تغییرات"):
        new = dict(data)
        for _, row in edited.iterrows():
            k = row["کلید"]
            if pd.notna(row["دوره جاری"]): new[k] = float(row["دوره جاری"])
            if pd.notna(row["دوره قبل"]): new[f"{k}_prev"] = float(row["دوره قبل"])
        st.session_state["data"] = new
        st.rerun()

    st.divider()
    st.subheader("📥 نسبت‌های محاسبه‌شده")
    df_r = pd.DataFrame([{"نسبت":k, "مقدار":round(v,3) if not pd.isna(v) else None}
                         for k,v in ratios.items()])
    st.dataframe(df_r, use_container_width=True, height=400)
    st.download_button("⬇️ دانلود CSV", df_r.to_csv(index=False).encode("utf-8-sig"),
                       f"ratios_{sk}.csv", "text/csv")

# ── ۷. واژه‌نامه ──
with T[6]:
    st.header("📖 واژه‌نامه اصطلاحات")
    st.caption("همه اصطلاحات انگلیسی به‌کاررفته در این داشبورد با توضیح فارسی.")
    for en, fa in GLOSSARY:
        st.markdown(f"**{en}** — {fa}")

# ── واژه‌نامه پایین صفحه (همیشه دیده می‌شود) ──
st.divider()
with st.expander("📖 راهنمای سریع اصطلاحات (کلیک کنید)", expanded=False):
    for en, fa in GLOSSARY:
        st.markdown(f"**{en}**: {fa}")

st.caption("⚠️ این ابزار تحلیلی است و توصیه سرمایه‌گذاری نیست. ارقام ممکن است حسابرسی‌نشده باشند.")
