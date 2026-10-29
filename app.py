"""app.py — داشبورد چند-صنعتی تحلیل سهام."""
import io
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from core import SECTOR_CONFIG, extract_from_excel, DEMO_BANK

st.set_page_config(page_title="داشبورد تحلیل سهام", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css');
html, body, [class*="css"] { font-family: 'Vazirmatn', sans-serif; }
.stApp { direction: rtl; text-align: right; }
[data-testid="stSidebar"] { direction: rtl; }
[data-testid="metric-container"] { background:#f8fafc; border-radius:12px; padding:12px;
  border-right:4px solid #2563eb; }
h1, h2, h3 { color:#0f172a; }
.warn-box { background:#fef3c7; border-right:4px solid #f59e0b; padding:12px;
  border-radius:8px; margin:8px 0; }
</style>
""", unsafe_allow_html=True)

# ─────────────── سایدبار ───────────────
st.sidebar.header("⚙️ تنظیمات")
sector_key = st.sidebar.selectbox(
    "صنعت",
    list(SECTOR_CONFIG.keys()),
    format_func=lambda k: SECTOR_CONFIG[k]["name"],
)
cfg = SECTOR_CONFIG[sector_key]

uploaded = st.sidebar.file_uploader(
    "صورت‌های مالی (اکسل کدال)",
    type=["xlsx", "xls"],
    help="فایل صورت‌های مالی یا گزارش تفسیری کدال را آپلود کنید.",
)

st.sidebar.subheader("💰 پارامترهای بازار")
price_toman = st.sidebar.number_input("قیمت روز سهم (تومان)", min_value=1.0,
                                      value=143.0 if sector_key == "bank" else 100.0, step=1.0)
shares_bn = st.sidebar.number_input("تعداد سهام (میلیارد)", min_value=0.01,
                                     value=3.382 if sector_key == "bank" else 1.0,
                                     step=0.01, format="%.3f")
period_months = st.sidebar.selectbox("دوره صورت مالی (ماه)", [3, 6, 9, 12], index=1)

st.sidebar.divider()
if sector_key == "bank" and st.sidebar.button("🔄 بارگذاری داده دمو (وبملت)"):
    st.session_state["force_demo"] = True

# ─────────────── پارس داده ───────────────
data = {}
if uploaded:
    with st.spinner("در حال پارس فایل..."):
        data = extract_from_excel(uploaded, cfg["labels"])

# دمو فقط برای بانک اگر هیچ داده‌ای نیست
if sector_key == "bank" and not data:
    data = dict(DEMO_BANK)
    st.info("ℹ️ داده دمو (وبملت، ۶ماهه منتهی به ۱۴۰۵/۰۶/۳۱) بارگذاری شد. برای سهم دیگر، فایل را آپلود یا دستی وارد کنید.")

# محاسبه نسبت‌ها
ratios = cfg["ratios_fn"](data, price_rial=price_toman * 10, shares=shares_bn * 1e9,
                          period_months=period_months)

# ─────────────── هدر ───────────────
st.title(f"📊 داشبورد تحلیل — {cfg['name']}")
st.caption(f"قیمت: **{price_toman:,.0f} تومان** | سهام: **{shares_bn:.2f} میلیارد** | دوره: **{period_months} ماه**")

if cfg["warning"]:
    st.markdown(f"<div class='warn-box'>{cfg['warning']}</div>", unsafe_allow_html=True)
if not pd.isna(data.get("net_income")) and data.get("net_income", 0) < 0:
    st.markdown("<div class='warn-box'>🔴 شرکت در دوره جاری زیان‌ده است — P/E معنا ندارد.</div>", unsafe_allow_html=True)

# نوار پوشش داده
req = [k for k in cfg["labels"] if not k.endswith("_prev")]
found = [k for k in req if not pd.isna(data.get(k))]
missing = [k for k in req if pd.isna(data.get(k))]
cov = len(found) / len(req) * 100 if req else 0
c1, c2 = st.columns([3, 1])
c1.progress(int(cov), text=f"پوشش داده: {len(found)}/{len(req)} فیلد ({cov:.0f}٪)")
c2.caption("⚠️ فیلدهای گم‌شده را در تب «داده و ورود دستی» پر کنید." if missing else "✅ همه فیلدها موجود است.")

# ─────────────── تب‌ها ───────────────
t1, t2, t3, t4, t5 = st.tabs(["🏠 خلاصه", "📈 سود و زیان", "🏦 ترازنامه", "💰 ارزش‌گذاری", "🔧 داده و ورود دستی"])

def fmt(x, pct=False, dec=1):
    if x is None or pd.isna(x): return "—"
    return f"{x:,.{dec}f}٪" if pct else f"{x:,.{dec}f}"

# ═══ تب ۱: خلاصه ═══
with t1:
    cols = st.columns(5)
    for i, k in enumerate(cfg["primary"]):
        v = ratios.get(k, np.nan)
        cols[i % 5].metric(k, fmt(v, pct="%" in k or "Yield" in k, dec=1 if "%" in k else 2))
    st.divider()
    cols = st.columns(5)
    for i, k in enumerate(cfg["secondary"]):
        v = ratios.get(k, np.nan)
        cols[i % 5].metric(k, fmt(v, pct="%" in k or "Yield" in k, dec=1 if "%" in k else 2))

# ═══ تب ۲: سود و زیان ═══
with t2:
    if sector_key == "bank":
        oi = data.get("op_income", 0) / 1e3
        dc = abs(data.get("deposit_cost", 0)) / 1e3
        ac = abs(data.get("admin_cost", 0)) / 1e3
        pv = abs(data.get("provision", 0)) / 1e3
        ni = data.get("net_income", 0) / 1e3
        other = ni - (oi - dc - ac - pv)
        x = ["درآمد عملیاتی", "هزینه سود سپرده", "هزینه اداری", "ذخیره مطالبات", "سایر اقلام خالص", "سود خالص"]
        y = [oi, -dc, -ac, -pv, other, ni]
        meas = ["relative", "relative", "relative", "relative", "relative", "total"]
    else:
        rv = data.get("revenue", 0) / 1e3
        cs = abs(data.get("cogs", 0)) / 1e3
        sg = abs(data.get("sga", 0)) / 1e3
        fc = abs(data.get("fin_cost", 0)) / 1e3
        ni = data.get("net_income", 0) / 1e3
        gp = rv - cs
        other = ni - gp + sg + fc
        x = ["فروش", "بهای تمام شده", "سود ناخالص", "هزینه فروش/اداری", "هزینه مالی", "سایر/مالیات", "سود خالص"]
        y = [rv, -cs, gp, -sg, -fc, other, ni]
        meas = ["relative", "relative", "total", "relative", "relative", "relative", "total"]
    fig = go.Figure(go.Waterfall(
        measure=meas, x=x, y=y, text=[f"{v:,.0f}" for v in y], textposition="outside",
        connector={"line": {"color": "#94a3b8"}},
        increasing={"marker": {"color": "#16a34a"}},
        decreasing={"marker": {"color": "#dc2626"}},
        totals={"marker": {"color": "#2563eb"}},
    ))
    fig.update_layout(height=440, showlegend=False, yaxis_title="همت",
                      title="پل سود (همت) — «سایر اقلام» = اقلام باقی‌مانده تا تطبیق با سود خالص")
    st.plotly_chart(fig, use_container_width=True)

    st.caption("💡 اگر «سایر اقلام خالص» بزرگ باشد، سود از فعالیت اصلی نمی‌آید — پرچم کیفیت سود.")

# ═══ تب ۳: ترازنامه ═══
with t3:
    ta = data.get("total_assets", 0) / 1e3
    cash = data.get("cash", 0) / 1e3
    if sector_key == "bank":
        loans = data.get("total_loans", 0) / 1e3
        other_a = max(0, ta - cash - loans)
        labels_a = ["نقد", "تسهیلات", "سایر دارایی‌ها"]
        vals_a = [cash, loans, other_a]
    else:
        inv = data.get("inventory", 0) / 1e3
        ar = data.get("receivables", 0) / 1e3
        other_a = max(0, ta - cash - inv - ar)
        labels_a = ["نقد", "موجودی کالا", "مطالبات", "سایر"]
        vals_a = [cash, inv, ar, other_a]
    c1, c2 = st.columns(2)
    with c1:
        fig = px.pie(values=vals_a, names=labels_a, hole=0.55,
                     title="ترکیب دارایی‌ها (همت)",
                     color_discrete_sequence=px.colors.sequential.Blues_r)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        eq = data.get("equity", 0) / 1e3
        liab = data.get("total_liab", 0) / 1e3
        fig = px.pie(values=[liab, eq], names=["بدهی", "حقوق مالکانه"], hole=0.55,
                     title="ساختار سرمایه (همت)",
                     color_discrete_sequence=["#dc2626", "#16a34a"])
        st.plotly_chart(fig, use_container_width=True)

# ═══ تب ۴: ارزش‌گذاری ═══
with t4:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("P/E", fmt(ratios.get("P/E"), dec=2))
    c2.metric("P/B", fmt(ratios.get("P/B"), dec=2))
    if "EV/EBITDA" in ratios:
        c3.metric("EV/EBITDA", fmt(ratios.get("EV/EBITDA"), dec=2))
    else:
        c3.metric("Earnings Yield", fmt(ratios.get("Earnings Yield (%)"), pct=True))
    c4.metric("Market Cap", fmt(ratios.get("Market Cap (همت)"), dec=0) + " همت")

    st.divider()
    st.subheader("🎯 حساسیت P/E به قیمت")
    eps = ratios.get("EPS سالانه (ریال)")
    if eps and not pd.isna(eps) and eps > 0:
        prices = np.linspace(price_toman * 0.5, price_toman * 2.0, 40)
        sens = pd.DataFrame({"قیمت (تومان)": prices, "P/E": [(p * 10) / eps for p in prices]})
        fig = px.line(sens, x="قیمت (تومان)", y="P/E", markers=True)
        fig.add_hline(y=5, line_dash="dash", line_color="orange", annotation_text="P/E=5")
        fig.add_hline(y=7, line_dash="dash", line_color="red", annotation_text="P/E=7")
        fig.add_vline(x=price_toman, line_dash="dot", line_color="green",
                      annotation_text=f"قیمت {price_toman:.0f}")
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("EPS منفی یا صفر — منحنی P/E معنا ندارد.")

    if cfg["benchmark"]:
        st.subheader("📊 مقایسه با بنچمارک صنعت")
        rows = []
        for k, bv in cfg["benchmark"].items():
            sv = ratios.get(k, np.nan)
            if not pd.isna(sv):
                rows.append({"شاخص": k, "این سهم": sv, "میانه صنعت": bv})
        if rows:
            df = pd.DataFrame(rows)
            fig = px.bar(df.melt(id_vars="شاخص", var_name="مبنا", value_name="مقدار"),
                         x="شاخص", y="مقدار", color="مبنا", barmode="group", text_auto=".2f")
            fig.update_layout(height=380)
            st.plotly_chart(fig, use_container_width=True)

# ═══ تب ۵: داده و ورود دستی ═══
with t5:
    st.subheader("✏️ ورود / ویرایش دستی داده‌ها (میلیون ریال)")
    st.caption("مقادیر یافت‌شده از اکسل را می‌توانید اصلاح کنید یا فیلدهای گم‌شده را پر کنید.")

    # ساخت دیتافریم قابل ویرایش
    rows = []
    for k, cands in cfg["labels"].items():
        cur = data.get(k, np.nan)
        prev = data.get(f"{k}_prev", np.nan)
        rows.append({
            "کلید": k,
            "شرح": cands[0],
            "دوره جاری": None if pd.isna(cur) else float(cur),
            "دوره قبل": None if pd.isna(prev) else float(prev),
        })
    df_edit = pd.DataFrame(rows)

    edited = st.data_editor(
        df_edit, use_container_width=True, height=500, num_rows="fixed",
        column_config={
            "کلید": st.column_config.TextColumn(disabled=True),
            "شرح": st.column_config.TextColumn(disabled=True),
            "دوره جاری": st.column_config.NumberColumn(format="%.0f"),
            "دوره قبل": st.column_config.NumberColumn(format="%.0f"),
        },
        key="manual_edit",
    )

    # به‌روزرسانی دیکشنری داده
    new_data = dict(data)
    for _, row in edited.iterrows():
        k = row["کلید"]
        if pd.notna(row["دوره جاری"]): new_data[k] = float(row["دوره جاری"])
        if pd.notna(row["دوره قبل"]):  new_data[f"{k}_prev"] = float(row["دوره قبل"])
    if st.button("✅ اعمال تغییرات"):
        st.session_state["data"] = new_data
        st.rerun()

    st.divider()
    st.subheader("📥 نسبت‌های محاسبه‌شده")
    df_r = pd.DataFrame([{"نسبت": k, "مقدار": round(v, 3) if not pd.isna(v) else None}
                         for k, v in ratios.items()])
    st.dataframe(df_r, use_container_width=True, height=400)
    st.download_button("⬇️ دانلود CSV نسبت‌ها",
                       df_r.to_csv(index=False).encode("utf-8-sig"),
                       f"ratios_{sector_key}.csv", "text/csv")

# فوتر
st.divider()
st.caption("⚠️ این ابزار تحلیلی است و توصیه سرمایه‌گذاری نیست. ارقام ممکن است حسابرسی‌نشده باشند.")
