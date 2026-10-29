"""app.py — داشبورد چند-صنعتی با آپلود چند فایل"""
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from core import SECTOR_CONFIG, extract_excel

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
  border-radius:8px; margin:8px 0; }
.good-box { background:#dcfce7; border-right:4px solid #16a34a; padding:12px;
  border-radius:8px; margin:8px 0; }
.bad-box { background:#fee2e2; border-right:4px solid #dc2626; padding:12px;
  border-radius:8px; margin:8px 0; }
</style>
""", unsafe_allow_html=True)

def fmt(v,pct=False,dec=1,suffix=""):
    if v is None or pd.isna(v): return "—"
    return f"{v:,.{dec}f}٪{suffix}" if pct else f"{v:,.{dec}f}{suffix}"

GLOSSARY = [
    ("P/E","نسبت قیمت به سود هر سهم. زیر ۵ = ارزان"),
    ("P/B","نسبت قیمت به ارزش دفتری. زیر ۱ = زیر ارزش دفتری"),
    ("P/NAV","نسبت قیمت به خالص ارزش دارایی‌ها. برای هلدینگ حیاتی — زیر ۱ = تخفیف"),
    ("NAV","Net Asset Value — ارزش روز دارایی‌ها منهای بدهی‌ها"),
    ("ROA","بازده دارایی = سود خالص ÷ کل دارایی"),
    ("ROE","بازده حقوق مالکانه = سود خالص ÷ حقوق مالکانه"),
    ("EPS","سود هر سهم = سود خالص ÷ تعداد سهام"),
    ("BVPS","ارزش دفتری هر سهم = حقوق مالکانه ÷ تعداد سهام"),
    ("DPS","سود نقدی هر سهم"),
    ("Dividend Yield","بازده سود نقدی = DPS ÷ قیمت"),
    ("Payout Ratio","نسبت تقسیم سود = DPS ÷ EPS"),
    ("LDR","نسبت تسهیلات به سپرده (بانکی)"),
    ("NPL","مطالبات غیرجاری (معوق + مشکوک)"),
    ("Cost-to-Income","هزینه عملیاتی ÷ درآمد عملیاتی"),
    ("EBITDA","سود قبل از بهره، مالیات و استهلاک"),
    ("EV","ارزش شرکت = ارزش بازار + بدهی خالص"),
    ("FCF","جریان نقد آزاد"),
    ("DSO/DIO/DPO","دوره وصول/گردش/پرداخت (روز)"),
    ("CCC","چرخه تبدیل نقد = DSO + DIO − DPO"),
    ("Margin","حاشیه سود (درصد)"),
]

st.sidebar.header("⚙️ تنظیمات")
sk = st.sidebar.selectbox("صنعت", list(SECTOR_CONFIG.keys()),
                          format_func=lambda k: SECTOR_CONFIG[k]["name"],
                          index=0)  # پیش‌فرض: هلدینگ
cfg = SECTOR_CONFIG[sk]

uploaded = st.sidebar.file_uploader(
    "صورت‌های مالی (چند فایل مجاز)",
    type=["xlsx","xls"],
    accept_multiple_files=True,
    help="می‌توانید چند اکسل کدال را همزمان آپلود کنید — داده‌ها ادغام می‌شوند."
)

st.sidebar.subheader("💰 پارامترهای بازار")
price = st.sidebar.number_input("قیمت روز سهم (تومان)", min_value=1.0, value=1000.0, step=10.0)
shares = st.sidebar.number_input("تعداد سهام (میلیارد)", min_value=0.01, value=30.0, step=0.1, format="%.2f")
months = st.sidebar.selectbox("دوره صورت مالی (ماه)", [3,6,9,12], index=3)

# ═══ پارس چند فایل ═══
data = {}
file_report = []  # (filename, sheet_count, fields_found, fields_total)

if uploaded:
    with st.spinner(f"در حال پارس {len(uploaded)} فایل..."):
        for f in uploaded:
            extracted = extract_excel(f, cfg["labels"])
            found = len([k for k,v in extracted.items() if not pd.isna(v)])
            total = len(cfg["labels"])
            file_report.append((f.name, found, total))
            # ادغام — فایل‌های بعدی فقط فیلدهای گم‌شده را پر می‌کنند
            for k,v in extracted.items():
                if k not in data or pd.isna(data.get(k)):
                    data[k] = v

# ═══ هشدار صریح اگر داده نیامد ═══
if uploaded and len([k for k,v in data.items() if not pd.isna(v)]) == 0:
    st.markdown(
        f"<div class='bad-box'>🔴 <b>هیچ داده‌ای از {len(uploaded)} فایل استخراج نشد.</b><br>"
        f"احتمالاً صنعت انتخابی («{cfg['name']}») با ساختار فایل هم‌خوان نیست. "
        f"صنعت را در سایدبار به <b>«هلدینگ/سرمایه‌گذاری»</b> تغییر دهید.</div>",
        unsafe_allow_html=True
    )

ratios = cfg["fn"](data, p=price*10, sh=shares*1e9, m=months)

# ═══ هدر ═══
st.title(f"📊 داشبورد تحلیل — {cfg['name']}")
st.caption(f"قیمت: **{price:,.0f} تومان** | سهام: **{shares:.2f} میلیارد** | دوره: **{months} ماه** | فایل‌ها: **{len(uploaded) if uploaded else 0}**")

# گزارش هر فایل
if file_report:
    with st.expander("📋 گزارش پارس هر فایل", expanded=True):
        for fname, found, total in file_report:
            pct = int(found/total*100) if total else 0
            color = "🟢" if pct>=50 else ("🟡" if pct>=20 else "🔴")
            st.write(f"{color} **{fname}** — {found}/{total} فیلد ({pct}٪)")

# نوار پوشش کل
req = [k for k in cfg["labels"] if not k.endswith("_prev")]
found_all = [k for k in req if not pd.isna(data.get(k))]
miss = [k for k in req if pd.isna(data.get(k))]
cov = int(len(found_all)/len(req)*100) if req else 0
c1,c2 = st.columns([3,1])
c1.progress(cov, text=f"پوشش کل: {len(found_all)}/{len(req)} ({cov}٪)")
c2.caption("⚠️ فیلد گم — تب «داده»" if miss else "✅ کامل")

# هشدار زیان
if not pd.isna(data.get("net_income")) and data.get("net_income",0) < 0:
    st.markdown("<div class='bad-box'>🔴 شرکت زیان‌ده است.</div>", unsafe_allow_html=True)

# ═══ تب‌ها ═══
T = st.tabs(["🏠 خلاصه","📈 سود و زیان","🏦 ترازنامه","💰 ارزش‌گذاری","👤 تحلیل سهامدار","🔧 داده","📖 واژه‌نامه"])

# ── خلاصه ──
with T[0]:
    st.subheader("شاخص‌های کلیدی")
    cols = st.columns(5)
    for i,k in enumerate(cfg["primary"]):
        cols[i%5].metric(k, fmt(ratios.get(k), pct="%" in k or "Yield" in k, dec=2))
    st.divider()
    cols = st.columns(5)
    for i,k in enumerate(cfg["secondary"]):
        cols[i%5].metric(k, fmt(ratios.get(k), pct="%" in k or "Yield" in k, dec=2))

    flags = []
    if not pd.isna(ratios.get("P/E")) and 0 < ratios.get("P/E",99) < 4:
        flags.append("🟢 P/E پایین — ارزندگی")
    if not pd.isna(ratios.get("P/B")) and 0 < ratios.get("P/B",99) < 1:
        flags.append("🟢 P/B زیر ۱ — زیر ارزش دفتری")
    if ratios.get("ROE (%)",0) > 20:
        flags.append("🟢 ROE بالای ۲۰٪ — سودآوری قوی")
    if ratios.get("Debt/Equity",0) > 2:
        flags.append("🟡 Debt/Equity بالا — ریسک بدهی")
    if flags:
        st.subheader("🚩 پرچم‌های تحلیلی")
        for f in flags: st.write(f)

# ── سود و زیان ──
with T[1]:
    ni = data.get("net_income",0)/1e3
    rt = data.get("rev_total",0)/1e3
    if sk == "holding":
        div = data.get("div_income",0)/1e3
        itr = data.get("interest_income",0)/1e3
        sal = data.get("salary",0)/1e3
        oex = data.get("other_exp",0)/1e3
        fin = data.get("fin_cost",0)/1e3
        oth = ni - (div+itr - sal - oex - fin)
        x=["درآمد سود سهام","درآمد سود تضمین‌شده","هزینه حقوق","سایر هزینه‌ها","هزینه مالی","سایر","سود خالص"]
        y=[div,itr,-sal,-oex,-fin,oth,ni]
        meas=["relative","relative","relative","relative","relative","relative","total"]
    else:
        rev = data.get("revenue",0)/1e3
        cs = abs(data.get("cogs",0))/1e3
        sg = abs(data.get("sga",0))/1e3
        fc = abs(data.get("fin_cost",0))/1e3
        gp = rev-cs; oth = ni-gp+sg+fc
        x=["فروش","بهای تمام","سود ناخالص","هزینه فروش/اداری","هزینه مالی","سایر/مالیات","سود خالص"]
        y=[rev,-cs,gp,-sg,-fc,oth,ni]
        meas=["relative","relative","total","relative","relative","relative","total"]
    fig = go.Figure(go.Waterfall(measure=meas, x=x, y=y,
        text=[f"{v:,.0f}" for v in y], textposition="outside",
        connector={"line":{"color":"#94a3b8"}},
        increasing={"marker":{"color":"#16a34a"}},
        decreasing={"marker":{"color":"#dc2626"}},
        totals={"marker":{"color":"#2563eb"}}))
    fig.update_layout(height=440, showlegend=False, yaxis_title="همت")
    st.plotly_chart(fig, use_container_width=True)

# ── ترازنامه ──
with T[2]:
    ta = data.get("total_assets",0)/1e3; cash = data.get("cash",0)/1e3
    eq = data.get("equity",0)/1e3; liab = data.get("total_liab",0)/1e3
    if sk == "holding":
        li = data.get("long_inv",0)/1e3; si = data.get("short_inv",0)/1e3; as_ = data.get("associates",0)/1e3
        labels = ["نقد","سرمایه‌گذاری بلندمدت","شرکت‌های وابسته","سایر"]
        vals = [cash, li, as_, max(0, ta-cash-li-as_) ]
    else:
        inv = data.get("inventory",0)/1e3; ar = data.get("receivables",0)/1e3
        labels = ["نقد","موجودی","مطالبات","سایر"]
        vals = [cash, inv, ar, max(0, ta-cash-inv-ar)]
    c1,c2 = st.columns(2)
    with c1:
        fig = px.pie(values=vals, names=labels, hole=0.55, title="ترکیب دارایی‌ها (همت)",
                     color_discrete_sequence=px.colors.sequential.Blues_r)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.pie(values=[liab,eq], names=["بدهی","حقوق مالکانه"], hole=0.55,
                     title="ساختار سرمایه (همت)", color_discrete_sequence=["#dc2626","#16a34a"])
        st.plotly_chart(fig, use_container_width=True)

# ── ارزش‌گذاری ──
with T[3]:
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("P/E", fmt(ratios.get("P/E"), dec=2))
    c2.metric("P/B", fmt(ratios.get("P/B"), dec=2))
    if "EV/EBITDA" in ratios: c3.metric("EV/EBITDA", fmt(ratios.get("EV/EBITDA"), dec=2))
    elif "P/NAV (تخمینی)" in ratios: c3.metric("P/NAV", fmt(ratios.get("P/NAV (تخمینی)"), dec=2))
    else: c3.metric("Earnings Yield", fmt(ratios.get("Earnings Yield (%)"), pct=True))
    c4.metric("Market Cap", fmt(ratios.get("Market Cap (همت)"), dec=0)+" همت")

    eps = ratios.get("EPS سالانه (ریال)")
    if eps and not pd.isna(eps) and eps > 0:
        st.subheader("🎯 حساسیت P/E به قیمت")
        prices = np.linspace(price*0.5, price*2.0, 40)
        sens = pd.DataFrame({"قیمت (تومان)":prices, "P/E":[(p*10)/eps for p in prices]})
        fig = px.line(sens, x="قیمت (تومان)", y="P/E", markers=True)
        fig.add_hline(y=5, line_dash="dash", line_color="orange", annotation_text="P/E=5")
        fig.add_vline(x=price, line_dash="dot", line_color="green", annotation_text=f"{price:.0f}")
        fig.update_layout(height=400); st.plotly_chart(fig, use_container_width=True)

    if cfg["benchmark"]:
        st.subheader("📊 مقایسه با صنعت")
        rows=[]
        for k,bv in cfg["benchmark"].items():
            sv = ratios.get(k, np.nan)
            if not pd.isna(sv): rows.append({"شاخص":k,"این سهم":sv,"صنعت":bv})
        if rows:
            df = pd.DataFrame(rows)
            fig = px.bar(df.melt(id_vars="شاخص",var_name="مبنا",value_name="مقدار"),
                         x="شاخص",y="مقدار",color="مبنا",barmode="group",text_auto=".2f")
            fig.update_layout(height=380); st.plotly_chart(fig, use_container_width=True)

# ── تحلیل سهامدار ──
with T[4]:
    st.header("👤 تحلیل از دید سهامدار")
    eps_a = ratios.get("EPS سالانه (ریال)", np.nan)
    bvps = ratios.get("BVPS (ریال)", np.nan)
    dps = ratios.get("DPS (ریال)", np.nan)
    div_yield = ratios.get("بازده سود نقدی (%)", 0)

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("EPS", fmt(eps_a, dec=0))
    c2.metric("BVPS", fmt(bvps, dec=0))
    c3.metric("DPS", fmt(dps, dec=0))
    c4.metric("بازده نقدی", fmt(div_yield, pct=True))

    # قیمت هدف
    st.subheader("🎯 قیمت هدف")
    peer_pe = cfg["benchmark"].get("P/E", 6.0)
    peer_pb = cfg["benchmark"].get("P/B", 1.2)
    tp_pe = (eps_a*peer_pe/10) if eps_a and eps_a>0 else np.nan
    tp_pb = (bvps*peer_pb/10) if bvps and bvps>0 else np.nan
    tp_g = np.sqrt(22.5*eps_a*bvps)/10 if eps_a and bvps and eps_a>0 and bvps>0 else np.nan
    targets = [t for t in [tp_pe,tp_pb,tp_g] if not pd.isna(t)]
    target = np.mean(targets) if targets else np.nan

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("P/E محور", fmt(tp_pe,dec=0)+" ت" if not pd.isna(tp_pe) else "—")
    c2.metric("P/B محور", fmt(tp_pb,dec=0)+" ت" if not pd.isna(tp_pb) else "—")
    c3.metric("Graham", fmt(tp_g,dec=0)+" ت" if not pd.isna(tp_g) else "—")
    if not pd.isna(target):
        up = (target-price)/price*100
        c4.metric("میانگین هدف", f"{target:,.0f} ت", delta=f"{up:+.1f}%")

    # TSR
    if not pd.isna(target):
        cap_gain = (target-price)/price*100
        tsr = cap_gain + (div_yield or 0)
        st.subheader("💎 TSR ۱ساله")
        c1,c2,c3 = st.columns(3)
        c1.metric("سود سرمایه", f"{cap_gain:+.1f}%")
        c2.metric("سود نقدی", f"{div_yield or 0:+.1f}%")
        c3.metric("TSR کل", f"{tsr:+.1f}%")

    # سناریوها
    st.subheader("🔮 سناریوهای ۱۲ماهه")
    eps_prev = data.get("eps_prev", np.nan)
    eps_g = ((eps_a-eps_prev)/eps_prev*100) if (eps_prev and not pd.isna(eps_prev) and eps_prev>0 and eps_a) else 15
    scen = []
    for nm,gf,pf in [("🐂 خوش‌بینانه",1.2,1.2),("📊 پایه",1.0,1.0),("🐻 بدبینانه",0.5,0.8)]:
        en = eps_a*(1+eps_g*gf/100) if eps_a else np.nan
        pe_t = peer_pe*pf
        tp = en*pe_t/10 if en else np.nan
        ret = (tp-price)/price*100+(div_yield or 0) if tp else np.nan
        scen.append({"سناریو":nm,"رشد EPS%":round(eps_g*gf,1),"P/E هدف":round(pe_t,1),"EPS":en,"قیمت هدف (ت)":tp,"TSR%":ret})
    st.dataframe(pd.DataFrame(scen), use_container_width=True, hide_index=True)

    # ریسک
    st.subheader("⚠️ ریسک‌ها")
    risks = []
    if (div_yield or 0) < 2: risks.append("🔴 سود نقدی ناچیز")
    if ratios.get("Debt/Equity",0) > 2: risks.append("🟡 اهرم بالا")
    if not pd.isna(ratios.get("P/E")) and ratios.get("P/E",0) > 10: risks.append("🟡 P/E بالا")
    if not risks: risks.append("✅ ریسک خاصی شناسایی نشد")
    for r in risks: st.write(r)

    # جمع‌بندی
    score = 0
    if not pd.isna(target) and target > price*1.2: score += 2
    elif not pd.isna(target) and target > price: score += 1
    if not pd.isna(ratios.get("P/E")) and ratios.get("P/E") < 5: score += 1
    if ratios.get("ROE (%)",0) > 20: score += 1
    if (div_yield or 0) > 3: score += 1
    if score >= 4: v,c = "🟢 جذاب","good-box"
    elif score >= 2: v,c = "🟡 نیازمند بررسی","warn-box"
    else: v,c = "🔴 پرریسک","bad-box"
    st.markdown(f"<div class='{c}'><b>سیگنال (نه توصیه):</b> {v} — {score}/6</div>", unsafe_allow_html=True)

# ── داده ──
with T[5]:
    st.subheader("✏️ داده استخراج‌شده (قابل ویرایش)")
    rows = []
    for k,cands in cfg["labels"].items():
        cur = data.get(k,np.nan); prev = data.get(f"{k}_prev",np.nan)
        rows.append({"کلید":k,"شرح":cands[0],
                     "جاری": None if pd.isna(cur) else float(cur),
                     "قبل": None if pd.isna(prev) else float(prev)})
    df_edit = pd.DataFrame(rows)
    edited = st.data_editor(df_edit, use_container_width=True, height=450, num_rows="fixed",
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
    st.subheader("📥 نسبت‌های محاسبه‌شده")
    df_r = pd.DataFrame([{"نسبت":k,"مقدار":round(v,3) if not pd.isna(v) else None} for k,v in ratios.items()])
    st.dataframe(df_r, use_container_width=True, height=400)
    st.download_button("⬇️ CSV", df_r.to_csv(index=False).encode("utf-8-sig"), f"ratios_{sk}.csv","text/csv")

# ── واژه‌نامه ──
with T[6]:
    st.header("📖 واژه‌نامه")
    for en,fa in GLOSSARY:
        st.markdown(f"**{en}** — {fa}")

st.divider()
with st.expander("📖 راهنمای سریع اصطلاحات", expanded=False):
    for en,fa in GLOSSARY:
        st.markdown(f"**{en}**: {fa}")
st.caption("⚠️ تحلیلی است — توصیه سرمایه‌گذاری نیست.")
