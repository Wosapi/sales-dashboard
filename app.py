# -*- coding: utf-8 -*-
"""
لوحة تحكم مبيعات تفاعلية – Streamlit + Plotly
التشغيل:  pip install -r requirements.txt  ثم  streamlit run app.py
ضع ملف Total_202611.xlsb بجانب app.py (أو ارفعه من الشريط الجانبي).
"""
import io
import os

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from analytics import (AR_MONTHS, PRODUCTS, area_table, build_insights, build_recommendations,
                       clean, fmt, inactive_areas, mlabel, month_status, monthly, quarterly,
                       seasonal_index, top_orders, yoy)

DEFAULT_FILE = os.environ.get("DATA_FILE", os.path.join(os.path.dirname(os.path.abspath(__file__)), "Total_202611.xlsb"))
SHEET = "Data"

# ألوان (باليت مرجعية مُتحقَّق منها): المنتجات بترتيب ثابت، والسنوات بترتيب ثابت
C_PROD = dict(zip(PRODUCTS, ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6"]))
C_YEARS = ["#2a78d6", "#eb6834", "#1baf7a", "#6250d6"]
C_MAIN, C_GOOD, C_BAD = "#2a78d6", "#008300", "#c0392b"
GRID = "rgba(128,128,128,0.18)"

st.set_page_config(page_title="لوحة مبيعات الزيوت", page_icon="📊", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');
html, body, .stApp, [data-testid="stSidebar"] {font-family:'Tajawal',sans-serif;}
.stApp, [data-testid="stSidebar"] {direction: rtl; text-align: right;}
.stPlotlyChart, [data-testid="stDataFrame"] {direction: ltr;}
.block-container {padding-top: 1.6rem; max-width: 1400px;}
.kpi {border:1px solid rgba(128,128,128,.25); background:rgba(128,128,128,.06); border-radius:14px; padding:16px 20px;}
.kpi .l {font-size:.95rem; opacity:.75;} .kpi .v {font-size:2rem; font-weight:700; line-height:1.3;}
.kpi .s {font-size:.82rem; opacity:.7;}
.card {border:1px solid rgba(128,128,128,.25); border-radius:12px; padding:14px 18px; margin-bottom:10px; background:rgba(128,128,128,.04);}
.card b.t {display:block; margin-bottom:4px;}
.card.good {border-right:5px solid #008300;} .card.bad {border-right:5px solid #c0392b;}
.card.ins {border-right:5px solid #2a78d6;} .card.rec {border-right:5px solid #eb6834;}
h2, h3 {font-weight:700 !important;}
</style>
""", unsafe_allow_html=True)


# ------------------------------------------------------------------ تحميل البيانات
@st.cache_data(show_spinner="جاري قراءة الملف وتنظيف البيانات…")
def load(file_bytes: bytes, name: str):
    engine = "pyxlsb" if name.lower().endswith(".xlsb") else "openpyxl"
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=SHEET, engine=engine)
    return clean(raw)


with st.sidebar:
    st.header("⚙️ الفلاتر")
    up = st.file_uploader("ملف المبيعات (xlsb / xlsx)", type=["xlsb", "xlsx", "xlsm"])

if up is not None:
    data_bytes, data_name = up.getvalue(), up.name
elif os.path.exists(DEFAULT_FILE):
    with open(DEFAULT_FILE, "rb") as fh:
        data_bytes, data_name = fh.read(), DEFAULT_FILE
else:
    st.info("ارفع ملف المبيعات من الشريط الجانبي أو ضعه بجانب app.py باسم Total_202611.xlsb")
    st.stop()

try:
    df_full, meta = load(data_bytes, data_name)
except Exception as e:  # noqa
    st.error(f"تعذّرت قراءة الملف: {e}")
    st.stop()

status = month_status(df_full)
complete = status["complete"]
all_years = sorted(df_full.Year.unique())
ycol = {y: C_YEARS[i % len(C_YEARS)] for i, y in enumerate(all_years)}

with st.sidebar:
    years_sel = st.multiselect("السنة (Year)", all_years, default=all_years)
    areas_all = list(df_full.groupby("Area").Amount.sum().sort_values(ascending=False).index)
    areas_sel = st.multiselect("المنطقة (Region)", areas_all, default=areas_all)
    q_sel = st.multiselect("الربع السنوي (Quarter)", ["Q1", "Q2", "Q3", "Q4"], default=["Q1", "Q2", "Q3", "Q4"])
    st.caption("الأشهر الناقصة/الجزئية تُعرض بتمييز خاص وتُستبعد من المقارنات (النمو، الموسمية، أقوى/أضعف شهر).")

if not (years_sel and areas_sel and q_sel):
    st.warning("اختر قيمة واحدة على الأقل في كل فلتر.")
    st.stop()

base = df_full[df_full.Area.isin(areas_sel) & df_full.Quarter.isin(q_sel)]   # بدون فلتر السنة (لحساب YoY)
df = base[base.Year.isin(years_sel)]
if df.empty:
    st.warning("لا توجد بيانات ضمن هذه الفلاتر.")
    st.stop()
ref_year = max(years_sel)

# ------------------------------------------------------------------ العنوان و KPIs
st.title("📊 لوحة تحكم المبيعات")
st.caption(f"الفترة: {df.Date.min():%Y-%m-%d} → {df.Date.max():%Y-%m-%d}  |  العملة: ريال سعودي (ر.س)  |  عدد الفواتير: {len(df):,}")

yo = yoy(base, complete, ref_year)
k1, k2, k3 = st.columns(3)
k1.markdown(f'<div class="kpi"><div class="l">إجمالي المبيعات (Amount)</div><div class="v">{fmt(df.Amount.sum())} ر.س</div>'
            f'<div class="s">متوسط الفاتورة {fmt(df.Amount.mean())} ر.س</div></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi"><div class="l">إجمالي الكميات (Total Quantities)</div><div class="v">{fmt(df["Total Quantities"].sum())} وحدة</div>'
            f'<div class="s">متوسط الطلب {df["Total Quantities"].mean():,.1f} وحدة</div></div>', unsafe_allow_html=True)
if yo and yo["growth"] is not None:
    g = yo["growth"]; col = C_GOOD if g >= 0 else C_BAD; arrow = "▲" if g >= 0 else "▼"
    ms = "، ".join(AR_MONTHS[m - 1] for m in yo["months"])
    k3.markdown(f'<div class="kpi"><div class="l">نمو المبيعات YoY ({ref_year} مقابل {ref_year - 1})</div>'
                f'<div class="v" style="color:{col}">{arrow} {g:+.1%}</div>'
                f'<div class="s">{fmt(yo["cur"])} مقابل {fmt(yo["prev"])} ر.س — مقارنة عادلة على أشهر مكتملة في السنتين: {ms}</div></div>',
                unsafe_allow_html=True)
else:
    k3.markdown(f'<div class="kpi"><div class="l">نمو المبيعات YoY</div><div class="v">غير متاح</div>'
                f'<div class="s">لا توجد أشهر مكتملة مشتركة بين {ref_year} و{ref_year - 1} ضمن الفلاتر الحالية</div></div>',
                unsafe_allow_html=True)

# جاهز: بيانات شهرية
m = monthly(df, status)
m = m[m.index.year.isin(years_sel) & m.index.quarter.isin([int(q[1]) for q in q_sel])]
labels = [mlabel(t) for t in m.index]
is_full = [t in complete for t in m.index]


def style(fig, h=400, legend=True):
    fig.update_layout(height=h, margin=dict(l=10, r=10, t=30, b=10), showlegend=legend,
                      legend=dict(orientation="h", y=1.12, x=0), font=dict(family="Tajawal"),
                      hoverlabel=dict(font_family="Tajawal"))
    fig.update_xaxes(gridcolor=GRID, showline=False)
    fig.update_yaxes(gridcolor=GRID, tickformat=",.0f", zeroline=False)
    return fig


# ------------------------------------------------------------------ 1) شهري + منتجات
st.header("1) المبيعات الشهرية وتفاصيل المنتجات")
t1, t2, t3 = st.tabs(["إجمالي المبيعات الشهرية", "الكميات حسب فئة المنتج", "الجدول الشهري"])
with t1:
    fig = go.Figure(go.Bar(
        x=labels, y=m.Sales, marker_color=[C_MAIN if f else "rgba(42,120,214,0.35)" for f in is_full],
        customdata=["مكتمل" if f else "ناقص/جزئي" for f in is_full],
        hovertemplate="%{x}<br>المبيعات: %{y:,.0f} ر.س<br>الحالة: %{customdata}<extra></extra>", name="المبيعات"))
    fig.update_xaxes(type="category", categoryorder="array", categoryarray=labels)
    st.plotly_chart(style(fig, legend=False), use_container_width=True)
    st.caption("الأعمدة الباهتة = شهر ناقص أو جزئي. الأشهر بلا عمود = لا توجد بيانات مسجلة في الملف.")
with t2:
    fig = go.Figure()
    for p in PRODUCTS:
        fig.add_bar(x=labels, y=m[p], name=p, marker_color=C_PROD[p], marker_line_width=1,
                    marker_line_color="rgba(255,255,255,0.7)",
                    hovertemplate=f"%{{x}}<br>{p}: %{{y:,.0f}} وحدة<extra></extra>")
    fig.update_layout(barmode="stack")
    fig.update_xaxes(type="category", categoryorder="array", categoryarray=labels)
    st.plotly_chart(style(fig, 440), use_container_width=True)
    st.caption("ملاحظة: الملف يسجل المبلغ (Amount) للفاتورة ككل، وليس لكل منتج؛ لذا يُقاس أداء الفئات بالكميات (وحدات).")
    tot = df[PRODUCTS].sum(); ptab = pd.DataFrame({"المنتج": tot.index, "الكمية": tot.values, "النسبة": tot.values / tot.sum()})
    ptab = ptab.sort_values("الكمية", ascending=False)
    ptab["الكمية"] = ptab["الكمية"].map(fmt); ptab["النسبة"] = ptab["النسبة"].map(lambda v: f"{v:.1%}")
    st.dataframe(ptab, hide_index=True, use_container_width=True)
with t3:
    tb = pd.DataFrame({"الشهر": labels, "الحالة": ["مكتمل" if f else "ناقص/جزئي" for f in is_full]})
    tb["المبيعات (ر.س)"] = [fmt(v) if pd.notna(v) else "—" for v in m.Sales]
    tb["الكميات"] = [fmt(v) if pd.notna(v) else "—" for v in m.Qty]
    tb["عدد الفواتير"] = [fmt(v) if pd.notna(v) else "—" for v in m.Orders]
    for p in PRODUCTS:
        tb[p] = [fmt(v) if pd.notna(v) else "—" for v in m[p]]
    st.dataframe(tb, hide_index=True, use_container_width=True)

# ------------------------------------------------------------------ 2) الاتجاه والموسمية
st.header("2) اتجاه المبيعات والموسمية")
c1, c2 = st.columns([3, 2])
with c1:
    fig = go.Figure()
    fig.add_scatter(x=labels, y=m.Sales, mode="lines+markers", name="المبيعات الشهرية", line=dict(color=C_MAIN, width=2),
                    marker=dict(size=8, color=C_MAIN), connectgaps=False,
                    hovertemplate="%{x}<br>%{y:,.0f} ر.س<extra></extra>")
    inc = [(l, v) for l, v, f in zip(labels, m.Sales, is_full) if not f and pd.notna(v)]
    if inc:
        fig.add_scatter(x=[a for a, _ in inc], y=[b for _, b in inc], mode="markers", name="شهر جزئي (غير قابل للمقارنة)",
                        marker=dict(size=11, symbol="circle-open", line=dict(width=2, color=C_BAD)), hoverinfo="skip")
    ma = m.Sales.where(pd.Series(is_full, index=m.index)).rolling(3, min_periods=2).mean()
    fig.add_scatter(x=labels, y=ma, mode="lines", name="متوسط متحرك 3 أشهر", line=dict(color="#eb6834", width=2, dash="dot"),
                    hoverinfo="skip")
    fig.update_xaxes(type="category", categoryorder="array", categoryarray=labels, tickangle=-45)
    st.plotly_chart(style(fig, 420), use_container_width=True)
with c2:
    mm = df.groupby("YM").Amount.sum()
    mm = mm[mm.index.isin(complete)]
    fig = go.Figure()
    for y in sorted(df.Year.unique()):
        s = mm[mm.index.year == y]
        ser = pd.Series({AR_MONTHS[i - 1]: s.get(pd.Timestamp(y, i, 1), None) for i in range(1, 13)}, dtype="float")
        fig.add_scatter(x=list(ser.index), y=ser.values, mode="lines+markers", name=str(y),
                        line=dict(color=ycol[y], width=2), marker=dict(size=7), connectgaps=False,
                        hovertemplate=f"{y} – %{{x}}<br>%{{y:,.0f}} ر.س<extra></extra>")
    fig.update_xaxes(type="category", categoryorder="array", categoryarray=AR_MONTHS, tickangle=-45)
    fig.update_layout(title=dict(text="الموسمية: كل سنة مقابل أشهر السنة (أشهر مكتملة فقط)", font=dict(size=13)))
    st.plotly_chart(style(fig, 420), use_container_width=True)
si = seasonal_index(df, complete)
if len(si):
    best = si[si.N >= 2].Index.nlargest(3)
    if len(best):
        st.caption("المؤشر الموسمي (1.00× = متوسط السنة، بعد إزالة أثر النمو): الأعلى — " +
                   "، ".join(f"{AR_MONTHS[i - 1]} {v:.2f}×" for i, v in best.items()))

# ------------------------------------------------------------------ 3) المناطق
st.header("3) تحليل المناطق")
at = area_table(df, complete, ref_year)
dead = inactive_areas(df_full)
a1, a2 = st.columns([3, 2])
with a1:
    a = at.sort_values("Sales")
    fig = go.Figure(go.Bar(y=a.index, x=a.Sales, orientation="h", marker_color=C_MAIN,
                           text=[f"{fmt(s)}  ({p:.1%})" for s, p in zip(a.Sales, a.Share)], textposition="outside",
                           hovertemplate="%{y}<br>%{x:,.0f} ر.س<extra></extra>"))
    fig.update_xaxes(range=[0, a.Sales.max() * 1.35])
    fig.update_layout(yaxis=dict(type="category"))
    st.plotly_chart(style(fig, max(300, 55 * len(a) + 80), legend=False), use_container_width=True)
with a2:
    gv = at.Growth.dropna()
    if len(gv):
        fast, slow = gv.idxmax(), gv.idxmin()
        st.markdown(f'<div class="card good"><b class="t">▲ المنطقة الأسرع نمواً ({ref_year} مقابل {ref_year - 1})</b>'
                    f'{fast}: {gv[fast]:+.1%}</div>', unsafe_allow_html=True)
        title = "▼ المنطقة الأكثر تراجعاً" if gv[slow] < 0 else "▼ المنطقة الأقل نمواً"
        st.markdown(f'<div class="card bad"><b class="t">{title}</b>{slow}: {gv[slow]:+.1%}</div>', unsafe_allow_html=True)
    else:
        st.info("لا تتوفر أشهر مكتملة مشتركة لحساب نمو المناطق ضمن الفلاتر الحالية.")
    dd = [d for d in dead if d in at.index]
    if dd:
        st.markdown(f'<div class="card bad"><b class="t">⚠ مناطق متوقفة (لا مبيعات منذ أكثر من 90 يوماً)</b>{"، ".join(dd)}</div>',
                    unsafe_allow_html=True)
    st.caption("النمو = مقارنة على الأشهر المكتملة في السنتين فقط، ويُحسب للمناطق التي تجاوزت مبيعاتها 5,000 ر.س في السنة السابقة. "
               "النمو −100% يعني أن المنطقة لم تُسجَّل لها مبيعات في أشهر المقارنة.")
atab = pd.DataFrame({"المنطقة": at.index, "المبيعات (ر.س)": at.Sales.map(fmt).values, "المساهمة %": at.Share.map(lambda v: f"{v:.1%}").values,
                     "الكميات": at.Qty.map(fmt).values, "الفواتير": at.Orders.map(fmt).values,
                     "النمو YoY": [f"{v:+.1%}" if pd.notna(v) else "—" for v in at.Growth],
                     "الحالة": ["متوقفة" if x in dead else "نشطة" for x in at.index]})
st.dataframe(atab, hide_index=True, use_container_width=True)

# ------------------------------------------------------------------ 4) الأرباع
st.header("4) تحليل الأرباع السنوية")
qd = quarterly(df, complete)
fig = go.Figure()
for y in sorted(qd.Year.unique()):
    s = qd[qd.Year == y].set_index("Quarter").reindex(["Q1", "Q2", "Q3", "Q4"])
    ok = s.Complete.fillna(True).astype(bool).tolist()
    fig.add_bar(x=s.index, y=s.Sales, name=str(y), marker_color=ycol[y], marker_opacity=[1 if o else 0.4 for o in ok],
                text=[f"{v:,.0f}" + ("" if o else "*") if pd.notna(v) else "" for v, o in zip(s.Sales, ok)],
                textposition="outside", cliponaxis=False,
                customdata=["مكتمل" if o else "ناقص/جزئي" for o in ok],
                hovertemplate=f"{y} – %{{x}}<br>%{{y:,.0f}} ر.س<br>%{{customdata}}<extra></extra>")
fig.update_layout(barmode="group")
st.plotly_chart(style(fig, 420), use_container_width=True)
st.caption("* ربع غير مكتمل (يحتوي شهراً ناقصاً أو جزئياً) — الأعمدة الباهتة غير قابلة للمقارنة المباشرة مع بقية الأرباع.")

# ------------------------------------------------------------------ 5) Top 10
st.header("5) أعلى 10 طلبات من حيث الكمية")
tp = top_orders(df).copy()
tp.insert(0, "#", range(1, len(tp) + 1))
tp["Date"] = tp.Date.dt.strftime("%Y-%m-%d")
tp["Total Quantities"] = tp["Total Quantities"].map(fmt); tp["Amount"] = tp.Amount.map(fmt)
tp.columns = ["#", "التاريخ", "العميل", "المنطقة", "الكمية", "المبلغ (ر.س)"]
st.dataframe(tp, hide_index=True, use_container_width=True)

# ------------------------------------------------------------------ Insights & Recommendations
st.header("💡 الرؤى التحليلية (Insights)")
for i, (t, txt) in enumerate(build_insights(df, df_full, status, ref_year), 1):
    st.markdown(f'<div class="card ins"><b class="t">{i}. {t}</b>{txt}</div>', unsafe_allow_html=True)
st.header("🎯 التوصيات الاستراتيجية (Recommendations)")
for i, (t, txt) in enumerate(build_recommendations(df, df_full, status, ref_year), 1):
    st.markdown(f'<div class="card rec"><b class="t">{i}. {t}</b>{txt}</div>', unsafe_allow_html=True)

with st.expander("جودة البيانات وما تم تنظيفه"):
    st.markdown(f"""
- صفوف الورقة الخام: **{meta['raw_rows']:,}** — منها **{meta['non_sales_rows']:,}** صف سداد/إيداع (بلا منطقة) استُبعدت لأنها ليست مبيعات، والمتبقي **{meta['sales_rows']:,}** فاتورة مبيعات.
- عمود Date حُوِّل إلى Datetime (تواريخ غير صالحة: {meta['bad_dates']}).
- قيم المبلغ المفقودة ({meta['missing_amount']}) والكميات المفقودة للمنتجات عوملت كصفر.
- أشهر بلا أي بيانات: {'، '.join(mlabel(x) for x in status['missing']) or 'لا يوجد'}.
- أشهر جزئية: {'، '.join(mlabel(x) for x in status['partial']) or 'لا يوجد'} (عدد فواتيرها أقل من 40% من الوسيط).
- فئة DEO 1L (12 وحدة) غير مدرجة ضمن الفئات السبع المطلوبة، لذا مجموع الفئات قد يقل قليلاً عن إجمالي الكميات.
""")
