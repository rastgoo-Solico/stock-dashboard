"""core.py — پارس اکسل کدال + نسبت‌های ۶ صنعت (بانک، پترو، دارو، خودرو، هلدینگ، عمومی)"""
import re
import pandas as pd
import numpy as np

PERSIAN="۰۱۲۳۴۵۶۷۸۹"; ARABIC="٠١٢٣٤٥٦٧٨٩"
def to_en(s):
    for i,ch in enumerate(PERSIAN): s=s.replace(ch,str(i))
    for i,ch in enumerate(ARABIC):  s=s.replace(ch,str(i))
    return s
def norm_label(s):
    s=to_en(str(s)).replace("ي","ی").replace("ك","ک").replace("\u200c","")
    return re.sub(r"\s+","",s).strip()
def parse_num(x):
    if x is None or (isinstance(x,float) and np.isnan(x)): return np.nan
    if isinstance(x,(int,float)): return float(x)
    s=to_en(str(x)).strip()
    if not s or s in {"-","*",""}: return np.nan
    neg=s.startswith("(") and s.endswith(")")
    s=s.replace(",","").replace("،","").replace(" ","").replace("%","").strip("()")
    try:
        v=float(s); return -v if neg else v
    except: return np.nan
def safe_div(a,b):
    try:
        if b in (0,None) or pd.isna(a) or pd.isna(b): return np.nan
        return a/b
    except: return np.nan
def safe_add(*vs): return np.nan if any(pd.isna(v) for v in vs) else float(sum(vs))
def safe_mul(a,b): return np.nan if pd.isna(a) or pd.isna(b) else a*b

def find_row(df, cands):
    if df.shape[1] < 2 or df.empty: return None
    col_a=df.iloc[:,0].astype(str).map(norm_label)
    for c in cands:
        m=col_a==norm_label(c)
        if m.any(): return df[m].iloc[0]
    for c in cands:
        cn=norm_label(c)
        if len(cn)<4: continue
        m=col_a.str.contains(re.escape(cn),na=False)
        if m.any(): return df[m].iloc[0]
    return None

def extract_excel(file_like, labels):
    out={}
    try: xls=pd.ExcelFile(file_like)
    except Exception: return out
    for sh in xls.sheet_names:
        try: raw=pd.read_excel(xls,sheet_name=sh,header=None)
        except Exception: continue
        for hdr in range(min(6,len(raw))):
            df=raw.iloc[hdr+1:].reset_index(drop=True)
            if df.shape[1]<2 or df.empty: continue
            for k,cands in labels.items():
                if k in out and not pd.isna(out[k]): continue
                row=find_row(df,cands)
                if row is None: continue
                nums=[]
                for c in df.columns[1:6]:
                    v=parse_num(row.get(c))
                    if not pd.isna(v): nums.append(v)
                    if len(nums)>=2: break
                if nums:
                    out[k]=nums[0]
                    if len(nums)>1: out[f"{k}_prev"]=nums[1]
    return out

# ═══════════════ برچسب‌ها ═══════════════
BANK_LABELS = {
    "rev_loans":["درآمد تسهیلات اعطایی","درآمد تسهيلات اعطايي"],
    "rev_fees":["درآمد کارمزد"],
    "deposit_cost":["هزینه سود سپرده‌ها"],
    "admin_cost":["هزینه های اداری و عمومی"],
    "provision":["هزینه مطالبات مشکوک الوصول"],
    "op_income":["درآمد عملیاتی"],
    "op_profit":["سود(زیان) عملیاتی"],
    "net_income":["سود(زیان)خالص","سود (زیان) خالص"],
    "eps":["سود (زیان) خالص هر سهم– ریال","سود (زیان) خالص هر سهم"],
    "cash":["موجودی نقد"],
    "total_assets":["جمع دارایی‌ها"],
    "total_liab":["جمع بدهی‌ها"],
    "equity":["جمع حقوق مالکانه"],
    "deposits":["سپرده‌های مشتریان"],
    "loans_gov":["تسهیلات اعطایی و مطالبات از اشخاص دولتی"],
    "loans_nongov":["تسهیلات اعطایی و مطالبات از اشخاص غیر دولتی- ریالی","تسهیلات اعطایی و مطالبات از اشخاص غیر دولتی"],
    "total_loans":["مانده تسهیلات پایان دوره"],
    "npl_doubtful":["ذخیره اختصاصی مطالبات مشکوک الوصول"],
}
GENERIC_LABELS = {
    "revenue":["فروش","درآمدهای عملیاتی"],
    "cogs":["بهای تمام شده کالای فروش رفته","بهای تمام شده"],
    "sga":["هزینه های فروش، اداری و عمومی","هزینه های اداری و عمومی"],
    "op_profit":["سود(زیان) عملیاتی"],
    "fin_cost":["هزینه های مالی"],
    "net_income":["سود(زیان)خالص","سود (زیان) خالص"],
    "eps":["سود (زیان) خالص هر سهم– ریال","سود (زیان) خالص هر سهم","سود هر سهم"],
    "cash":["موجودی نقد"],
    "receivables":["حساب ها و اسناد دریافتنی تجاری"],
    "inventory":["موجودی مواد و کالا"],
    "current_assets":["جمع دارایی‌های جاری"],
    "total_assets":["جمع دارایی‌ها"],
    "payables":["حساب ها و اسناد پرداختنی تجاری"],
    "current_liab":["جمع بدهی‌های جاری"],
    "total_debt":["تسهیلات مالی دریافتی","جمع بدهی‌ها"],
    "equity":["جمع حقوق مالکانه"],
    "cfo":["جریان خالص ورود (خروج) نقد حاصل از فعالیت های عملیاتی"],
    "capex":["پرداخت های نقدی برای تحصیل دارایی ثابت مشهود"],
}
PHARMA_LABELS = {**GENERIC_LABELS,"rd":["هزینه تحقیق و توسعه","هزینه های تحقیق و توسعه"]}

# ═══ هلدینگ / سرمایه‌گذاری — کلید فایل شما ═══
HOLDING_LABELS = {
    "div_income":["درآمد سود سهام","درآمد حاصل از سود سهام"],
    "interest_income":["درآمد سود تضمین شده","درآمد سود تضمین‌شده"],
    "inv_sale_gain":["سود (زیان) فروش سرمایه گذاریها","سود (زیان) فروش سرمایه‌گذاریها","سود (زیان) فروش سرمایه گذاری"],
    "inv_reval":["سود (زیان) تغییر ارزش سرمایه گذاریها","سود (زیان) تعدیل ذخیره کاهش ارزش سرمایه گذاریهای کوتاه مدت"],
    "other_rev":["سایر درآمدها"],
    "rev_total":["جمع درآمدهای عملیاتی"],
    "salary":["هزینه هاى حقوق، دستمزد و مزایا","هزینه‌های حقوق، دستمزد و مزایا"],
    "other_exp":["سایر هزینه‌ها","ساير هزینه‌ها"],
    "op_exp_total":["جمع هزینه های عملیاتی"],
    "op_profit":["سود (زیان) عملياتى","سود (زیان) عملیاتی"],
    "fin_cost":["هزینه های مالی","هزینه‌های مالی","هزينه‌ هاى مالى"],
    "non_op":["سایر درآمدها و هزینه‌های غیرعملیاتی","ساير درآمدها و هزينه ‏هاى غيرعملياتى"],
    "pretax":["سود (زیان) عملیات در حال تداوم قبل از ماليات","سود (زیان) قبل از احتساب سهم گروه"],
    "net_income":["سود (زیان) خالص","سود(زیان)خالص","سود (زیان) خالص عملیات در حال تداوم"],
    "eps":["سود (زیان) خالص هر سهم (ریال)","سود (زیان) پایه هر سهم","سود (زیان) خالص هر سهم"],
    "capital":["سرمايه","سرمایه"],
    "cash":["موجودی نقد","موجودی نقد"],
    "short_inv":["سرمایه‌گذاری‌های کوتاه‌مدت","سرمایه گذاریهای کوتاه مدت"],
    "long_inv":["سرمایه‌گذاری‌های بلندمدت","سرمایه گذاریهای بلندمدت"],
    "associates":["سرمایه گذاری در شرکتهای وابسته","سرمایه‌گذاری در شرکت‌های وابسته"],
    "total_assets":["جمع دارایی‌ها"],
    "total_liab":["جمع بدهی‌ها"],
    "equity":["جمع حقوق مالکانه","حقوق مالکانه قابل انتساب به مالکان شرکت اصلی"],
    "current_assets":["جمع دارایی‌های جاری"],
    "current_liab":["جمع بدهی‌های جاری"],
    "receivables":["دریافتنی‌های تجاری و سایر دریافتنی‌ها"],
    "cfo":["جریان خالص ورود (خروج) نقد حاصل از فعالیت های عملیاتی","جريان ‌خالص ‌ورود‌ (خروج) ‌نقد حاصل از فعاليت‌هاي ‌عملياتي"],
    "dps_approved":["سود سهام مصوب"],
    "dps_proposed":["سود سهام پیشنهادی هیئت مدیره"],
    "ni_last_year":["مبلغ سود (زیان) خالص سال مالی گذشته"],
}

# ═══════════════ نسبت‌ها ═══════════════
def bank_ratios(d,p,sh,m=6):
    f=12/m; ni=d.get("net_income",np.nan); ni_a=ni*f if not pd.isna(ni) else np.nan
    eq=d.get("equity",np.nan); ta=d.get("total_assets",np.nan)
    eps=d.get("eps",np.nan); eps_a=eps*f if not pd.isna(eps) else safe_div(ni_a*1e6,sh)
    loans=d.get("total_loans",np.nan)
    if pd.isna(loans) or loans==0:
        loans=safe_add(d.get("loans_gov",0) or 0,d.get("loans_nongov",0) or 0)
    bvps=safe_div(eq*1e6,sh)
    r={}
    r["EPS سالانه (ریال)"]=eps_a; r["BVPS (ریال)"]=bvps
    r["ROE سالانه (%)"]=safe_div(ni_a,eq)*100; r["ROA سالانه (%)"]=safe_div(ni_a,ta)*100
    r["Cost-to-Income (%)"]=safe_div(abs(d.get("admin_cost",np.nan)),d.get("op_income",np.nan))*100
    r["حاشیه سود خالص (%)"]=safe_div(ni,d.get("op_income",np.nan))*100
    r["LDR (%)"]=safe_div(loans,d.get("deposits",np.nan))*100
    r["پوشش ذخیره (%)"]=safe_div(d.get("npl_doubtful",np.nan),loans)*100
    r["نقد/دارایی (%)"]=safe_div(d.get("cash",np.nan),ta)*100
    r["بدهی/دارایی (%)"]=safe_div(d.get("total_liab",np.nan),ta)*100
    r["P/E"]=safe_div(p,eps_a); r["P/B"]=safe_div(p,bvps)
    r["Earnings Yield (%)"]=safe_div(eps_a,p)*100
    r["Market Cap (همت)"]=safe_div(p*sh,1e12)
    return r

def generic_ratios(d,p,sh,m=6):
    f=12/m; rev=d.get("revenue",np.nan)
    cogs=abs(d["cogs"]) if not pd.isna(d.get("cogs",np.nan)) else np.nan
    ni=d.get("net_income",np.nan); ni_a=ni*f if not pd.isna(ni) else np.nan
    eq=d.get("equity",np.nan); ta=d.get("total_assets",np.nan)
    ar=d.get("receivables",np.nan); inv=d.get("inventory",np.nan)
    ap=d.get("payables",np.nan); ca=d.get("current_assets",np.nan); cl=d.get("current_liab",np.nan)
    debt=d.get("total_debt",np.nan); cfo=d.get("cfo",np.nan)
    capex=abs(d["capex"]) if not pd.isna(d.get("capex",np.nan)) else 0
    eps_a=safe_div(ni_a*1e6,sh); bvps=safe_div(eq*1e6,sh); mcap=p*sh; days=30*m
    dso=safe_div(ar,rev)*days; dio=safe_div(inv,cogs)*days; dpo=safe_div(ap,cogs)*days
    ccc=(dso+dio-dpo) if not (pd.isna(dso) or pd.isna(dio) or pd.isna(dpo)) else np.nan
    r={}
    r["حاشیه سود ناخالص (%)"]=safe_div(safe_add(rev,-(cogs or 0)),rev)*100 if not pd.isna(rev) else np.nan
    r["حاشیه سود عملیاتی (%)"]=safe_div(d.get("op_profit",np.nan),rev)*100
    r["حاشیه سود خالص (%)"]=safe_div(ni,rev)*100
    r["ROE (%)"]=safe_div(ni_a,eq)*100; r["ROA (%)"]=safe_div(ni_a,ta)*100
    r["نسبت جاری"]=safe_div(ca,cl); r["Debt/Equity"]=safe_div(debt,eq)
    r["Interest coverage"]=safe_div(d.get("op_profit",np.nan),abs(d["fin_cost"]) if not pd.isna(d.get("fin_cost",np.nan)) else np.nan)
    r["گردش موجودی (بار)"]=safe_div(cogs*f,inv)
    r["DSO (روز)"]=dso; r["DIO (روز)"]=dio; r["DPO (روز)"]=dpo; r["CCC (روز)"]=ccc
    r["FCF Yield (%)"]=safe_div((cfo*f-abs(capex*f))*1e6,mcap)*100
    r["EPS سالانه (ریال)"]=eps_a; r["BVPS (ریال)"]=bvps
    r["P/E"]=safe_div(p,eps_a) if eps_a and eps_a>0 else np.nan
    r["P/B"]=safe_div(p,bvps); r["P/S"]=safe_div(mcap,rev*f*1e6)
    r["Market Cap (همت)"]=safe_div(mcap,1e12)
    return r

def petro_ratios(d,p,sh,m=6):
    r=generic_ratios(d,p,sh,m); f=12/m
    op=d.get("op_profit",np.nan); rev=d.get("revenue",np.nan)
    ebitda=safe_add(op,safe_mul(rev,0.05)) if not pd.isna(op) else np.nan
    net_debt=safe_add(d.get("total_debt",0) or 0,-(d.get("cash",0) or 0))
    mcap=p*sh; ev=safe_add(mcap,net_debt*1e6)
    r["EV/EBITDA"]=safe_div(ev,ebitda*f*1e6)
    return r

def pharma_ratios(d,p,sh,m=6):
    r=generic_ratios(d,p,sh,m)
    r["R&D / فروش (%)"]=safe_div(d.get("rd",np.nan),d.get("revenue",np.nan))*100
    return r

def holding_ratios(d,p,sh,m=12):
    f=12/m; ni=d.get("net_income",np.nan); ni_a=ni*f if not pd.isna(ni) else np.nan
    eq=d.get("equity",np.nan); ta=d.get("total_assets",np.nan)
    eps=d.get("eps",np.nan); eps_a=eps*f if not pd.isna(eps) else safe_div(ni_a*1e6,sh)
    bvps=safe_div(eq*1e6,sh); mcap=p*sh
    # دارایی‌های سرمایه‌گذاری
    long_inv=d.get("long_inv",np.nan); assoc=d.get("associates",np.nan); short_inv=d.get("short_inv",np.nan)
    total_inv=safe_add(*[v for v in [long_inv,assoc,short_inv] if not pd.isna(v)])
    div_inc=d.get("div_income",np.nan)
    rev_tot=d.get("rev_total",np.nan)
    r={}
    r["EPS سالانه (ریال)"]=eps_a; r["BVPS (ریال)"]=bvps
    r["ROE (%)"]=safe_div(ni_a,eq)*100
    r["ROA (%)"]=safe_div(ni_a,ta)*100
    r["حاشیه سود خالص (%)"]=safe_div(ni,rev_tot)*100
    r["سهم سود سهام از درآمد (%)"]=safe_div(div_inc,rev_tot)*100
    r["درصد سرمایه‌گذاری از دارایی (%)"]=safe_div(total_inv,ta)*100
    r["نسبت جاری"]=safe_div(d.get("current_assets",np.nan),d.get("current_liab",np.nan))
    r["Debt/Equity"]=safe_div(d.get("total_liab",np.nan),eq)
    r["P/E"]=safe_div(p,eps_a) if eps_a and eps_a>0 else np.nan
    r["P/B"]=safe_div(p,bvps)
    r["P/NAV (تخمینی)"]=safe_div(p,bvps)
    r["Market Cap (همت)"]=safe_div(mcap,1e12)
    # DPS
    dps=np.nan
    if d.get("dps_approved") and d.get("capital"):
        dps=d["dps_approved"]*1e6/(d["capital"]*1e3)  # capital in million rial → shares = capital/1000
    r["DPS (ریال)"]=dps
    r["بازده سود نقدی (%)"]=safe_div(dps,p)*100 if not pd.isna(dps) else np.nan
    return r

SECTOR_CONFIG = {
    "holding":{"name":"هلدینگ / سرمایه‌گذاری","labels":HOLDING_LABELS,"fn":holding_ratios,
        "primary":["P/E","P/B","P/NAV (تخمینی)","ROE (%)","حاشیه سود خالص (%)"],
        "secondary":["EPS سالانه (ریال)","BVPS (ریال)","ROA (%)","Debt/Equity","بازده سود نقدی (%)"],
        "benchmark":{"P/E":6.0,"P/B":1.0,"ROE (%)":20}},
    "bank":{"name":"بانک","labels":BANK_LABELS,"fn":bank_ratios,
        "primary":["P/B","P/E","ROE سالانه (%)","Cost-to-Income (%)","LDR (%)"],
        "secondary":["ROA سالانه (%)","پوشش ذخیره (%)","حاشیه سود خالص (%)","Earnings Yield (%)","Market Cap (همت)"],
        "benchmark":{"P/E":5.8,"P/B":1.2,"ROE سالانه (%)":25}},
    "petro":{"name":"پتروشیمی","labels":GENERIC_LABELS,"fn":petro_ratios,
        "primary":["P/E","EV/EBITDA","حاشیه سود ناخالص (%)","ROE (%)","FCF Yield (%)"],
        "secondary":["P/B","P/S","ROA (%)","Debt/Equity","Interest coverage"],
        "benchmark":{"P/E":6.5,"P/B":1.5,"ROE (%)":20}},
    "pharma":{"name":"دارویی","labels":PHARMA_LABELS,"fn":pharma_ratios,
        "primary":["P/E","P/B","حاشیه سود ناخالص (%)","R&D / فروش (%)","ROE (%)"],
        "secondary":["ROA (%)","DSO (روز)","DIO (روز)","CCC (روز)","FCF Yield (%)"],
        "benchmark":{"P/E":8.0,"P/B":2.5,"ROE (%)":25}},
    "auto":{"name":"خودرو","labels":GENERIC_LABELS,"fn":generic_ratios,
        "primary":["P/E","حاشیه سود ناخالص (%)","Debt/Equity","Interest coverage","ROE (%)"],
        "secondary":["گردش موجودی (بار)","DSO (روز)","FCF Yield (%)","P/B","ROA (%)"],
        "benchmark":{"P/E":5.0,"P/B":1.0}},
    "generic":{"name":"عمومی","labels":GENERIC_LABELS,"fn":generic_ratios,
        "primary":["P/E","P/B","ROE (%)","ROA (%)","حاشیه سود ناخالص (%)"],
        "secondary":["P/S","Debt/Equity","نسبت جاری","CCC (روز)","FCF Yield (%)"],
        "benchmark":{"P/E":5.5,"P/B":1.3,"ROE (%)":20}},
}
