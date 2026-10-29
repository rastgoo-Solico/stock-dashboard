"""app.py — ورود دستی تعداد سهام و قیمت (بدون محاسبه خودکار)"""
import io
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from core import (SECTOR_CONFIG, extract_excel, parse_any_format,
                  suggest_price_toman, health_check)

st.set_page_config(page_title="داشبورد تحلیل سهام", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css');
html, body, [class*="css"], .stApp, button, input, select, textarea, table {
  font-family: 'Vazirmatn', Tahoma, sans-serif !important;
}
.stApp, [data-testid="stSidebar"], [data-testid="stSidebar"] * { direction: rtl; text-align: right; }
[data-testid="metric-container"] { background:#f8fafc; border-radius:12px; padding:14px;
  border-right:4px solid #2563eb; }
h1,h2,h3 { color:#0f172a; }
.warn-box { background:#fef3c7; border-right:4px solid #f59e0b; padding:12px;
  border-radius:8px; margin:8px 0; direction:rtl; }
.good-box { background:#dcfce7; border-right:4px solid #16a34a; padding:12px;
  border-radius:8px; margin:8px 0; direction:rtl; }
.bad-box { background:#fee2e2; border-right:4px solid #dc2626; padding:12px;
  border-radius:8px; margin:8px 0; direction:rtl; }
.info-box { background:#dbeafe; border-right:4px solid #2563eb; padding:12px;
  border-radius:8px; margin:8px 0; direction:rtl; }
</style>
""", unsafe_allow_html=True)

def fmt(v, pct=False, dec=1, suffix=""):
    if v is None or pd.isna(v): return "—"
    return f"{v:,.{dec}f}٪{suffix}" if pct else f"{v:,.{dec}f}{suffix}"

def safe_pct(a, b):
    try:
        if b in (0, None) or pd.isna(a) or pd.isna(b): return np.nan
        return a / b * 100
    except: return np.nan

def merge_prefer_larger(target, source):
    """ادغام — مقدار با قدرمطلق بزرگ‌تر (سالانه > فصلی) برنده می‌شود."""
    for k, v in source.items():
        if v is None or pd.isna(v): continue
        if k not in target or pd.isna(target.get(k)):
            target[k] = v
        else:
            try:
                if abs(float(v)) > abs(float(target.get(k))):
                    target[k] = v
            except (ValueError, TypeError):
                pass

GLOSSARY = [
    ("P/E","قیمت به سود هر سهم"), ("P/B","قیمت به ارزش دفتری"),
    ("P/NAV","قیمت به خالص ارزش دارایی‌ها"), ("ROA","سود خالص ÷ کل دارایی"),
    ("ROE","سود خالص ÷ حقوق مالکانه"), ("EPS","سود هر سهم"),
    ("BVPS","ارزش دفتری هر سهم"), ("DPS","سود نقدی هر سهم"),
    ("Dividend Yield","DPS ÷ قیمت"), ("Payout Ratio","DPS ÷ EPS"),
    ("LDR","تسهیلات ÷ سپرده (بانک)"), ("NPL","مطالبات غیرجاری"),
    ("Cost-to-Income","هزینه ÷ درآمد عملیاتی"),
    ("EBITDA","سود قبل از بهره، مالیات، استهلاک"),
    ("EV","ارزش شرکت"), ("FCF","جریان نقد آزاد"),
    ("DSO","دوره وصول مطالبات"), ("DIO","دوره گردش موجودی"),
    ("DPO","دوره پرداخت"), ("CCC","DSO + DIO − DPO"),
    ("TSR","بازده کل سهامدار"), ("Graham","√(22.5 × EPS × BVPS)"),
]

# ═══════════════ سایدبار ═══════════════
st.sidebar.header("⚙️ تنظیمات")
sk = st.sidebar.selectbox("صنعت", list(SECTOR_CONFIG.keys()),
                          format_func=lambda k: SECTOR_CONFIG[k]["name"], index=0)
cfg = SECTOR_CONFIG[sk]

uploaded = st.sidebar.file_uploader("صورت‌های مالی (چند فایل مجاز)",
                                     type=["xlsx", "xls"], accept_multiple_files=True)

# ═══════════════ پارس فایل‌ها ═══════════════
data = {}
file_report = []
previews = {}

if uploaded:
    with st.spinner(f"در حال پارس {len(uploaded)} فایل..."):
        for f in uploaded:
            content = b""
            try:
                if hasattr(f, 'getvalue'): content = f.getvalue()
                elif hasattr(f, 'read'):
                    try: f.seek(0)
                    except: pass
                    content = f.read()
            except Exception as e:
                file_report.append((f.name, 0, len(cfg["labels"]), f"❌ {str(e)[:60]}", 0))
                continue
            if not content:
                file_report.append((f.name, 0, len(cfg["labels"]), "❌ خالی", 0))
                continue

            size = len(content)
            result, kind, errors = parse_any_format(content)
            if result is None:
                file_report.append((f.name, 0, len(cfg["labels"]),
                                    f"❌ {kind}، {size:,} بایت", 0))
                continue

            try:
                if isinstance(result, pd.ExcelFile):
                    sample = pd.read_excel(result, sheet_name=result.sheet_names[0],
                                           header=None, nrows=40)
                    n_tables = len(result.sheet_names)
                elif isinstance(result, list):
                    biggest = max(result, key=lambda d: d.shape[0] * d.shape[1])
                    sample = biggest.head(40); n_tables = len(result)
                else:
                    sample = result.head(40); n_tables = 1
                previews[f.name] = sample
            except: continue

            ncols = sample.shape[1] if not sample.empty else 0

            try:
                extracted = extract_excel(io.BytesIO(content), cfg["labels"])
            except: continue

            merge_prefer_larger(data, extracted)
            found = len([k for k, v in extracted.items() if not pd.isna(v)])
            total = len(cfg["labels"])
            file_report.append((f.name, found, total,
                                f"{kind} | {n_tables} جدول | {size:,} بایت", ncols))

# ═══════════════ تعداد سهام — دستی ═══════════════
st.sidebar.divider()
st.sidebar.subheader("📊 تعداد سهام (دستی)")
st.sidebar.caption("⚠️ این عدد را از صورت‌های مالی یا سایت کدال بخوانید")
st.sidebar.caption("فرمول: سرمایه(م.ریال) ÷ اسمی(ریال) ÷ ۱۰۰۰")

if data:
    with st.sidebar.expander("🔍 مقادیر یافت‌شده در فایل (فقط اطلاع)"):
        cap = data.get("capital", np.nan)
        nom = data.get("nominal_value", np.nan)
        if not pd.isna(cap) and cap > 0:
            n = nom if not pd.isna(nom) and nom > 0 else 1000
            calc = (cap * 1e6) / (n * 1e9)
            st.write(f"**سرمایه**: {cap:,.0f} م.ریال")
            st.write(f"**اسمی**: {n:,.0f} ریال")
            st.write(f"**محاسبه**: {calc:.2f} میلیارد سهم")
            st.warning("⚠️ ممکن است اشتباه باشد — چند «سرمایه» در فایل هست")
        else:
            st.caption("سرمایه یافت نشد")

with st.sidebar.expander("🧮 محاسبه‌گر تعداد سهام"):
    st.caption("اگر نمی‌دانید، اینجا محاسبه کنید")
    cap_in = st.number_input("سرمایه (میلیون ریال)", min_value=0.0,
                              value=300000000.0, step=1000000.0, key="calc_cap",
                              format="%.0f")
    nom_in = st.number_input("اسمی (ریال)", min_value=0.0,
                              value=1000.0, step=100.0, key="calc_nom",
                              format="%.0f")
    if cap_in > 0 and nom_in > 0:
        result_sh = (cap_in * 1e6) / (nom_in * 1e9)
        st.success(f"**{result_sh:.2f} میلیارد سهم**")
        if st.button("📋 استفاده در فیلد", key="use_calc"):
            st.session_state.shares_input = float(result_sh)
            st.rerun()

if "shares_input" not in st.session_state:
    st.session_state.shares_input = 1.0

shares = st.sidebar.number_input(
    "تعداد سهام (میلیارد)",
    min_value=0.01, step=0.1, format="%.2f",
    key="shares_input",
    help="مثال: سرمایه ۳۰۰,۰۰۰,۰۰۰ م.ریال و اسمی ۱۰۰۰ → ۳۰۰ میلیارد سهم",
)

# ═══════════════ قیمت — دستی ═══════════════
st.sidebar.divider()
st.sidebar.subheader("💰 قیمت سهم (دستی)")
st.sidebar.caption("آخرین قیمت معامله در بازار (به تومان)")

months = st.sidebar.selectbox("دوره صورت مالی (ماه)", [3, 6, 9, 12], index=3)

if data:
    with st.sidebar.expander("💡 قیمت پیشنهادی (فقط اطلاع)"):
        temp_r = cfg["fn"](data, p=1.0, sh=shares*1e9, m=months)
        sug, sug_list = suggest_price_toman(temp_r, cfg["benchmark"])
        if sug_list:
            for lbl, val in sug_list:
                st.write(f"{lbl}: **{val:,.0f} ت**")
            st.caption("⚠️ بر اساس داده‌های تاریخی — قیمت واقعی بازار را وارد کنید")
        else:
            st.caption("داده کافی برای پیشنهاد نیست")

if "price_input" not in st.session_state:
    st.session_state.price_input = 1000.0

price = st.sidebar.number_input(
    "قیمت روز (تومان)",
    min_value=1.0, step=10.0, format="%.0f",
    key="price_input",
    help="قیمت را به تومان وارد کنید — نه ریال. اگر ریال می‌دانید بر ۱۰ تقسیم کنید.",
)

dollar = st.sidebar.number_input("نرخ دلار (ریال) — اختیاری",
                                  min_value=0.0, value=700000.0, step=1000.0)

# ═══════════════ محاسبه ═══════════════
ratios = cfg["fn"](data, p=price * 10, sh=shares * 1e9, m=months)
warnings = health_check(ratios, price, shares, data)

st.title(f"📊 داشبورد تحلیل — {cfg['name']}")
st.caption(f"قیمت: **{price:,.0f} تومان** | سهام: **{shares:.2f} میلیارد** | دوره: **{months} ماه** | فایل‌ها: **{len(uploaded) if uploaded else 0}**")

# هشدارهای سلامت
if warnings:
    st.markdown("### ⚠️ بررسی سلامت داده")
    for w in warnings:
        st.markdown(f"<div class='bad-box'>{w}</div>", unsafe_allow_html=True)

# اطلاعات شفاف — بدون باگ
if uploaded:
    st.markdown("### 🔍 اطلاعات محاسبه")
    c1, c2, c3, c4 = st.columns(4)
    eps_val = ratios.get("EPS سالانه (ریال)")
    bvps_val = ratios.get("BVPS (ریال)")
    pe_val = ratios.get("P/E")
    pb_val = ratios.get("P/B")
    mcap_val = ratios.get("Market Cap (همت)")
    
    c1.metric("تعداد سهام (دستی)", fmt(shares, dec=2) + " م.سهم")
    c2.metric("قیمت (دستی)", fmt(price, dec=0) + " تومان")
    c3.metric("EPS سالانه", fmt(eps_val, dec=0) + " ریال")
    c4.metric("Market Cap", fmt(mcap_val, dec=0) + " همت")
    
    st.caption(
        f"📐 **P/E** = قیمت(ریال) ÷ EPS = {price*10:,.0f} ÷ {fmt(eps_val, dec=0)} = "
        f"**{fmt(pe_val, dec=2)}**"
    )
    st.caption(
        f"📐 **P/B** = قیمت(ریال) ÷ BVPS = {price*10:,.0f} ÷ {fmt(bvps_val, dec=0)} = "
        f"**{fmt(pb_val, dec=2)}**"
    )

# پوشش
req = [k for k in cfg["labels"] if not k.endswith("_prev")]
found = [k for k in req if not pd.isna(data.get(k))]
miss = [k for k in req if pd.isna(data.get(k))]
cov = int(len(found)/len(req)*100) if req else 0
c1, c2 = st.columns([3, 1])
c1.progress(cov, text=f"پوشش: {len(found)}/{len(req)} ({cov}٪)")
c2.caption("⚠️ فیلد گم" if miss else "✅ کامل")

# گزارش فایل
if file_report:
    with st.expander("📋 گزارش پارس هر فایل", expanded=False):
        for row in file_report:
            fname, found_n, total, note = row[0], row[1], row[2], row[3]
            pct = int(found_n / total * 100) if total else 0
            icon = "🟢" if pct >= 40 else ("🟡" if pct >= 10 else "🔴")
            st.write(f"{icon} **{fname}** — {found_n}/{total} ({pct}٪) — `{note}`")
    with st.expander("🔬 نمای خام (Debug)", expanded=False):
        for fname, *_ in file_report:
            st.markdown(f"### `{fname}`")
            pv = previews.get(fname)
            if isinstance(pv, pd.DataFrame):
                st.write(f"ابعاد: {pv.shape[0]}×{pv.shape[1]}")
                st.dataframe(pv, use_container_width=True, height=300)

# ═══════════════ تب‌ها ═══════════════
T = st.tabs(["🏠 خلاصه", "📈 سود و زیان", "🏦 ترازنامه", "💰 ارزش‌گذاری",
             "👤 تحلیل سهامدار", "🔧 داده", "📖 واژه‌نامه"])

with T[0]:
    st.subheader("شاخص‌های کلیدی")
    cols = st.columns(5)
    for i, k in enumerate(cfg["primary"]):
        cols[i % 5].metric(k, fmt(ratios.get(k), pct="%" in k or "Yield" in k, dec=2))
    st.divider()
    cols = st.columns(5)
    for i, k in enumerate(cfg["secondary"]):
        cols[i % 5].metric(k, fmt(ratios.get(k), pct="%" in k or "Yield" in k, dec=2))

    flags = []
    if not pd.isna(ratios.get("P/E")) and 0 < ratios.get("P/E", 99) < 4: flags.append("🟢 P/E پایین")
    if not pd.isna(ratios.get("P/B")) and 0 < ratios.get("P/B", 99) < 1: flags.append("🟢 P/B زیر ۱")
    if ratios.get("ROE (%)", 0) > 20: flags.append("🟢 ROE بالای ۲۰٪")
    if ratios.get("Debt/Equity", 0) > 2: flags.append("🟡 اهرم بالا")
    if flags:
        st.subheader("🚩 پرچم‌ها")
        for f in flags: st.write(f)

with T[1]:
    ni = data.get("net_income", 0) / 1e3
    if sk == "holding":
        div = data.get("div_income", 0) / 1e3
        itr = data.get("interest_income", 0) / 1e3
        sal = abs(data.get("salary", 0)) / 1e3
        oex = abs(data.get("other_exp", 0)) / 1e3
        fin = abs(data.get("fin_cost", 0)) / 1e3
        oth = ni - (div + itr - sal - oex - fin)
        x = ["سود سهام","سود تضمین","هزینه حقوق","سایر هزینه","هزینه مالی","سایر","سود خالص"]
        y = [div, itr, -sal, -oex, -fin, oth, ni]
        meas = ["relative"] * 6 + ["total"]
    else:
        rev = data.get("revenue", 0) / 1e3
        cs = abs(data.get("cogs", 0)) / 1e3
        sg = abs(data.get("sga", 0)) / 1e3
        fc = abs(data.get("fin_cost", 0)) / 1e3
        gp = rev - cs
        oth = ni - gp + sg + fc
        x = ["فروش","بهای تمام","سود ناخالص","هزینه فروش","هزینه مالی","سایر","سود خالص"]
        y = [rev, -cs, gp, -sg, -fc, oth, ni]
        meas = ["relative","relative","total","relative","relative","relative","total"]
    fig = go.Figure(go.Waterfall(measure=meas, x=x, y=y,
        text=[f"{v:,.0f}" for v in y], textposition="outside",
        connector={"line": {"color": "#94a3b8"}},
        increasing={"marker": {"color": "#16a34a"}},
        decreasing={"marker": {"color": "#dc2626"}},
        totals={"marker": {"color": "#2563eb"}}))
    fig.update_layout(height=440, showlegend=False, yaxis_title="همت")
    st.plotly_chart(fig, use_container_width=True)

with T[2]:
    ta = data.get("total_assets", 0) / 1e3; cash = data.get("cash", 0) / 1e3
    eq = data.get("equity", 0) / 1e3; liab = data.get("total_liab", 0) / 1e3
    if sk == "holding":
        li = data.get("long_inv", 0) / 1e3; asc = data.get("associates", 0) / 1e3
        labels = ["نقد","سرمایه‌گذاری بلندمدت","وابسته","سایر"]
        vals = [cash, li, asc, max(0, ta - cash - li - asc)]
    else:
        inv = data.get("inventory", 0) / 1e3; ar = data.get("receivables", 0) / 1e3
        labels = ["نقد","موجودی","مطالبات","سایر"]
        vals = [cash, inv, ar, max(0, ta - cash - inv - ar)]
    c1, c2 = st.columns(2)
    with c1:
        fig = px.pie(values=vals, names=labels, hole=0.55, title="ترکیب دارایی (همت)",
                     color_discrete_sequence=px.colors.sequential.Blues_r)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.pie(values=[liab, eq], names=["بدهی","حقوق مالکانه"], hole=0.55,
                     title="ساختار سرمایه (همت)",
                     color_discrete_sequence=["#dc2626","#16a34a"])
        st.plotly_chart(fig, use_container_width=True)
    st.subheader("📊 نسبت‌های ساختاری")
    c1, c2, c3 = st.columns(3)
    c1.metric("بدهی/دارایی", fmt(safe_pct(liab, ta), pct=True))
    c2.metric("حقوق/دارایی", fmt(safe_pct(eq, ta), pct=True))
    c3.metric("نقد/دارایی", fmt(safe_pct(cash, ta), pct=True))

with T[3]:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("P/E", fmt(ratios.get("P/E"), dec=2))
    c2.metric("P/B", fmt(ratios.get("P/B"), dec=2))
    if "EV/EBITDA" in ratios: c3.metric("EV/EBITDA", fmt(ratios.get("EV/EBITDA"), dec=2))
    elif "P/NAV (تخمینی)" in ratios: c3.metric("P/NAV", fmt(ratios.get("P/NAV (تخمینی)"), dec=2))
    else: c3.metric("Earnings Yield", fmt(ratios.get("Earnings Yield (%)"), pct=True))
    c4.metric("Market Cap", fmt(ratios.get("Market Cap (همت)"), dec=0) + " همت")

    eps = ratios.get("EPS سالانه (ریال)")
    if eps and not pd.isna(eps) and eps > 0:
        st.subheader("🎯 حساسیت P/E")
        prices = np.linspace(max(1, price*0.2), price * 2.0, 40)
        sens = pd.DataFrame({"قیمت": prices, "P/E": [(p * 10) / eps for p in prices]})
        fig = px.line(sens, x="قیمت", y="P/E", markers=True)
        fig.add_hline(y=5, line_dash="dash", line_color="orange", annotation_text="P/E=5")
        fig.add_hline(y=10, line_dash="dash", line_color="red", annotation_text="P/E=10")
        fig.add_vline(x=price, line_dash="dot", line_color="green", annotation_text=f"فعلی {price:,.0f}")
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)

    if cfg["benchmark"]:
        st.subheader("📊 مقایسه با صنعت")
        rows = []
        for k, bv in cfg["benchmark"].items():
            sv = ratios.get(k, np.nan)
            if not pd.isna(sv): rows.append({"شاخص": k, "این سهم": sv, "صنعت": bv})
        if rows:
            df = pd.DataFrame(rows)
            fig = px.bar(df.melt(id_vars="شاخص", var_name="مبنا", value_name="مقدار"),
                         x="شاخص", y="مقدار", color="مبنا", barmode="group", text_auto=".2f")
            fig.update_layout(height=380)
            st.plotly_chart(fig, use_container_width=True)

with T[4]:
    st.header("👤 تحلیل سهامدار")
    eps_a = ratios.get("EPS سالانه (ریال)", np.nan)
    bvps = ratios.get("BVPS (ریال)", np.nan)
    dps = ratios.get("DPS (ریال)", np.nan)
    div_yield = ratios.get("بازده سود نقدی (%)", 0)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("EPS", fmt(eps_a, dec=0))
    c2.metric("BVPS", fmt(bvps, dec=0))
    c3.metric("DPS", fmt(dps, dec=0))
    c4.metric("بازده نقدی", fmt(div_yield, pct=True))

    st.subheader("🎯 قیمت هدف")
    peer_pe = cfg["benchmark"].get("P/E", 6.0)
    peer_pb = cfg["benchmark"].get("P/B", 1.2)
    tp_pe = (eps_a * peer_pe / 10) if eps_a and eps_a > 0 else np.nan
    tp_pb = (bvps * peer_pb / 10) if bvps and bvps > 0 else np.nan
    tp_g = np.sqrt(22.5 * eps_a * bvps) / 10 if eps_a and bvps and eps_a > 0 and bvps > 0 else np.nan
    targets = [t for t in [tp_pe, tp_pb, tp_g] if not pd.isna(t)]
    target = np.mean(targets) if targets else np.nan

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("P/E محور", fmt(tp_pe, dec=0) + " ت" if not pd.isna(tp_pe) else "—")
    c2.metric("P/B محور", fmt(tp_pb, dec=0) + " ت" if not pd.isna(tp_pb) else "—")
    c3.metric("Graham", fmt(tp_g, dec=0) + " ت" if not pd.isna(tp_g) else "—")
    if not pd.isna(target):
        up = (target - price) / price * 100
        c4.metric("هدف", f"{target:,.0f} ت", delta=f"{up:+.1f}%",
                  delta_color="normal" if up > 0 else "inverse")

    if not pd.isna(target):
        cap_gain = (target - price) / price * 100
        tsr = cap_gain + (div_yield or 0)
        st.subheader("💎 TSR ۱ساله")
        c1, c2, c3 = st.columns(3)
        c1.metric("سرمایه", f"{cap_gain:+.1f}%")
        c2.metric("نقدی", f"{div_yield or 0:+.1f}%")
        c3.metric("TSR", f"{tsr:+.1f}%")

    st.subheader("🔮 سناریوها")
    eps_prev = data.get("eps_prev", np.nan)
    eps_g = (eps_a - eps_prev) / eps_prev * 100 if (eps_prev and not pd.isna(eps_prev) and eps_prev > 0 and eps_a and not pd.isna(eps_a)) else 15

    scen = []
    for nm, gf, pf in [("🐂 خوش‌بینانه", 1.2, 1.2), ("📊 پایه", 1.0, 1.0), ("🐻 بدبینانه", 0.5, 0.8)]:
        en = eps_a * (1 + eps_g * gf / 100) if eps_a and not pd.isna(eps_a) else np.nan
        pe_t = peer_pe * pf
        tp = en * pe_t / 10 if en and not pd.isna(en) else np.nan
        ret = (tp - price) / price * 100 + (div_yield or 0) if tp and not pd.isna(tp) else np.nan
        scen.append({"سناریو": nm, "رشد EPS%": round(eps_g * gf, 1),
                     "P/E هدف": round(pe_t, 1), "EPS": en,
                     "قیمت هدف": tp, "TSR%": ret})
    st.dataframe(pd.DataFrame(scen), use_container_width=True, hide_index=True)

    st.subheader("⚠️ ریسک‌ها")
    risks = []
    if not pd.isna(div_yield) and div_yield < 2: risks.append("🔴 سود نقدی ناچیز")
    if ratios.get("Debt/Equity", 0) > 2: risks.append("🟡 اهرم بالا")
    if not pd.isna(ratios.get("P/E")) and ratios.get("P/E", 0) > 10: risks.append("🟡 P/E بالا")
    if not risks: risks.append("✅ ریسکی شناسایی نشد")
    for r in risks: st.write(r)

    st.subheader("📋 جمع‌بندی")
    score = 0
    if not pd.isna(target):
        if target > price * 1.2: score += 2
        elif target > price: score += 1
    if not pd.isna(ratios.get("P/E")) and 0 < ratios.get("P/E") < 5: score += 1
    if ratios.get("ROE (%)", ratios.get("ROE سالانه (%)", 0)) > 20: score += 1
    if (div_yield or 0) > 3: score += 1
    if score >= 4: v, c = "🟢 جذاب", "good-box"
    elif score >= 2: v, c = "🟡 نیازمند بررسی", "warn-box"
    else: v, c = "🔴 پرریسک", "bad-box"
    st.markdown(f"<div class='{c}'><b>سیگنال:</b> {v} — {score}/6</div>", unsafe_allow_html=True)

with T[5]:
    st.subheader("✏️ داده (میلیون ریال)")
    rows = []
    for k, cands in cfg["labels"].items():
        cur = data.get(k, np.nan); prev = data.get(f"{k}_prev", np.nan)
        rows.append({"کلید": k, "شرح": cands[0],
                     "جاری": None if pd.isna(cur) else float(cur),
                     "قبل": None if pd.isna(prev) else float(prev)})
    df_edit = pd.DataFrame(rows)
    edited = st.data_editor(df_edit, use_container_width=True, height=500, num_rows="fixed",
        column_config={
            "کلید": st.column_config.TextColumn(disabled=True),
            "شرح": st.column_config.TextColumn(disabled=True),
            "جاری": st.column_config.NumberColumn(format="%.0f"),
            "قبل": st.column_config.NumberColumn(format="%.0f"),
        }, key="editor")
    if st.button("✅ اعمال"):
        for _, row in edited.iterrows():
            k = row["کلید"]
            if pd.notna(row["جاری"]): data[k] = float(row["جاری"])
            if pd.notna(row["قبل"]): data[f"{k}_prev"] = float(row["قبل"])
        st.rerun()

    st.divider()
    st.subheader("📥 نسبت‌ها")
    df_r = pd.DataFrame([{"نسبت": k, "مقدار": round(v, 3) if not pd.isna(v) else None}
                         for k, v in ratios.items()])
    st.dataframe(df_r, use_container_width=True, height=400)
    st.download_button("⬇️ CSV", df_r.to_csv(index=False).encode("utf-8-sig"),
                       f"ratios_{sk}.csv", "text/csv")

    if miss:
        st.subheader("⚠️ فیلدهای گم")
        for m in miss: st.write(f"🔴 `{m}`")

with T[6]:
    st.header("📖 واژه‌نامه")
    for en, fa in GLOSSARY: st.markdown(f"**{en}** — {fa}")

st.divider()
st.caption("⚠️ ابزار تحلیلی است — توصیه سرمایه‌گذاری نیست.")
