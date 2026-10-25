"""منطق التحليل (pandas فقط) – مفصول عن واجهة Streamlit ليسهل اختباره."""
import re
import numpy as np
import pandas as pd

PRODUCTS = ["PCMO-Min 24L", "PCMO-Min 12L", "PCMO-Syn", "DEO 5L", "DEO 4L", "Fluids 4L", "Fluids 12L"]
REQUIRED = ["Amount", "Area", "Customer - name", "Total Quantities", "Date"] + PRODUCTS
AR_MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
             "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
PARTIAL_RATIO = 0.4  # شهر عدد فواتيره أقل من 40% من الوسيط يُعتبر جزئياً/ناقصاً


def _norm(c):
    return re.sub(r"\s+", " ", str(c)).strip()


def clean(raw: pd.DataFrame):
    """تنظيف الورقة الخام. يعيد (df, meta)."""
    df = raw.copy()
    df.columns = [_norm(c) for c in df.columns]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"أعمدة مفقودة في الورقة: {missing}")
    df = df[REQUIRED].copy()
    meta = {"raw_rows": len(df)}

    # التاريخ → datetime (يدعم الأرقام التسلسلية لإكسل القادمة من xlsb)
    d = df["Date"]
    if pd.api.types.is_numeric_dtype(d):
        df["Date"] = pd.to_datetime(d, unit="D", origin="1899-12-30", errors="coerce")
    else:
        df["Date"] = pd.to_datetime(d, errors="coerce")
    meta["bad_dates"] = int(df["Date"].isna().sum())

    # صفوف المبيعات فقط: لها منطقة. (صفوف السداد/الإيداعات بلا منطقة)
    df["Area"] = df["Area"].astype("string").str.strip()
    is_sale = df["Area"].notna() & (df["Area"] != "")
    meta["non_sales_rows"] = int((~is_sale).sum())
    df = df[is_sale & df["Date"].notna()].copy()

    # القيم المفقودة
    meta["missing_amount"] = int(df["Amount"].isna().sum())
    for c in ["Amount", "Total Quantities"] + PRODUCTS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["Customer - name"] = df["Customer - name"].astype("string").fillna("غير معروف").str.strip()
    df["Area"] = df["Area"].astype(str)
    df["Customer - name"] = df["Customer - name"].astype(str)

    df["Year"] = df["Date"].dt.year
    df["Month"] = df["Date"].dt.month
    df["YM"] = df["Date"].dt.to_period("M").dt.to_timestamp()
    df["Quarter"] = "Q" + df["Date"].dt.quarter.astype(str)
    df = df.sort_values("Date").reset_index(drop=True)
    meta["sales_rows"] = len(df)
    return df, meta


def month_status(df):
    """يحدد الأشهر المكتملة/الجزئية/المفقودة من البيانات الكاملة (قبل الفلاتر)."""
    cnt = df.groupby("YM").size()
    full_idx = pd.date_range(cnt.index.min(), cnt.index.max(), freq="MS")
    cnt = cnt.reindex(full_idx, fill_value=0)
    med = cnt[cnt > 0].median()
    complete = set(cnt[cnt >= PARTIAL_RATIO * med].index)
    missing = [m for m in cnt.index if cnt[m] == 0]
    partial = [m for m in cnt.index if 0 < cnt[m] < PARTIAL_RATIO * med]
    return {"complete": complete, "missing": missing, "partial": partial, "counts": cnt}


def mlabel(ts):
    return f"{AR_MONTHS[ts.month - 1]} {ts.year}"


def fmt(x, d=0):
    return f"{x:,.{d}f}"


def monthly(df, status):
    idx = pd.date_range(status["counts"].index.min(), status["counts"].index.max(), freq="MS")
    g = df.groupby("YM").agg(Sales=("Amount", "sum"), Qty=("Total Quantities", "sum"), Orders=("Amount", "size"))
    for p in PRODUCTS:
        g[p] = df.groupby("YM")[p].sum()
    g = g.reindex(idx)
    # أشهر بلا مبيعات ضمن الفلتر = 0 ، أما الأشهر المفقودة فعلياً من الملف فتبقى فارغة (فجوة)
    real_gap = g.index.isin(status["missing"])
    g.loc[~real_gap] = g.loc[~real_gap].fillna(0)
    g.index.name = "YM"
    return g


def _cmp_months(ref_year, complete, df_months):
    """أشهر (رقم الشهر) المكتملة في سنة المرجع والسنة السابقة معاً."""
    out = []
    for m in range(1, 13):
        a, b = pd.Timestamp(ref_year, m, 1), pd.Timestamp(ref_year - 1, m, 1)
        if a in complete and b in complete:
            out.append(m)
    return out


def yoy(df, complete, ref_year):
    """نمو سنوي على أشهر مكتملة في السنتين (مقارنة عادلة). يعيد dict أو None."""
    months = _cmp_months(ref_year, complete, None)
    if not months:
        return None
    cur = df[(df.Year == ref_year) & df.Month.isin(months)]
    prv = df[(df.Year == ref_year - 1) & df.Month.isin(months)]
    c, p = cur.Amount.sum(), prv.Amount.sum()
    cq, pq = cur["Total Quantities"].sum(), prv["Total Quantities"].sum()
    return {"months": months, "cur": c, "prev": p,
            "growth": (c / p - 1) if p > 0 else None,
            "cur_q": cq, "prev_q": pq,
            "growth_q": (cq / pq - 1) if pq > 0 else None}


def area_table(df, complete, ref_year, min_prev=5000):
    t = df.groupby("Area").agg(Sales=("Amount", "sum"), Qty=("Total Quantities", "sum"), Orders=("Amount", "size"))
    t["Share"] = t.Sales / t.Sales.sum() if t.Sales.sum() else 0
    t = t.sort_values("Sales", ascending=False)
    gr = {}
    for a in t.index:
        r = yoy(df[df.Area == a], complete, ref_year) if ref_year else None
        gr[a] = r["growth"] if r and r["prev"] >= min_prev else np.nan
    t["Growth"] = pd.Series(gr)
    return t


def quarterly(df, complete):
    g = df.groupby(["Year", "Quarter"]).agg(Sales=("Amount", "sum"), Qty=("Total Quantities", "sum")).reset_index()
    def ok(r):
        q = int(r.Quarter[1])
        return all(pd.Timestamp(int(r.Year), m, 1) in complete for m in range(3 * q - 2, 3 * q + 1))
    g["Complete"] = g.apply(ok, axis=1) if len(g) else []
    return g


def seasonal_index(df, complete):
    """مؤشر موسمي: مبيعات الشهر ÷ متوسط الأشهر المكتملة في سنته، ثم متوسط كل شهر عبر السنوات."""
    m = df.groupby("YM").Amount.sum()
    m = m[m.index.isin(complete)]
    if m.empty:
        return pd.DataFrame(columns=["Index", "N"])
    s = pd.DataFrame({"v": m})
    s["Year"], s["Month"] = s.index.year, s.index.month
    s["rel"] = s.v / s.groupby("Year").v.transform("mean")
    return s.groupby("Month").agg(Index=("rel", "mean"), N=("rel", "size"))


def top_orders(df, n=10):
    t = df.sort_values(["Total Quantities", "Amount"], ascending=False).head(n)
    return t[["Date", "Customer - name", "Area", "Total Quantities", "Amount"]]


# ---------------------------------------------------------------- الرؤى والتوصيات
def inactive_areas(df_full, days=90, min_orders=20):
    last = df_full.groupby("Area").Date.max()
    cnt = df_full.groupby("Area").size()
    end = df_full.Date.max()
    return [a for a in last.index if (end - last[a]).days > days and cnt[a] >= min_orders]


def _product_share_shift(df):
    """تغير حصة كل منتج (من إجمالي وحدات المنتجات السبعة) بين آخر سنتين."""
    yrs = sorted(df.Year.unique())
    if len(yrs) < 2:
        return None
    q = df.groupby("Year")[PRODUCTS].sum()
    sh = q.div(q.sum(axis=1), axis=0)
    return yrs[-2], yrs[-1], (sh.loc[yrs[-1]] - sh.loc[yrs[-2]]), sh


def build_insights(df, df_full, status, ref_year):
    """5 رؤى تحليلية (عنوان، نص). تعتمد على البيانات المفلترة."""
    out = []
    complete = status["complete"]
    # 1) أقوى/أضعف شهر (أشهر مكتملة فقط)
    try:
        m = df.groupby("YM").Amount.sum()
        m = m[m.index.isin(complete)]
        hi, lo = m.idxmax(), m.idxmin()
        out.append(("أقوى وأضعف شهر",
                    f"أقوى شهر مكتمل هو {mlabel(hi)} بمبيعات {fmt(m[hi])} ر.س، وأضعفها {mlabel(lo)} بمبيعات {fmt(m[lo])} ر.س "
                    f"(الفارق {m[hi] / m[lo]:.1f} ضعف). استُبعدت الأشهر الناقصة/الجزئية من المقارنة حتى لا تُشوَّه النتيجة."))
    except Exception:
        out.append(("أقوى وأضعف شهر", "لا تكفي البيانات المكتملة ضمن الفلاتر الحالية."))
    # 2) الموسمية
    try:
        si = seasonal_index(df, complete)
        si = si[si.N >= 2]
        if si.empty:
            raise ValueError
        top = si.Index.nlargest(3); bot = si.Index.nsmallest(2)
        t = "، ".join(f"{AR_MONTHS[i - 1]} ({v:.2f}×)" for i, v in top.items())
        b = "، ".join(f"{AR_MONTHS[i - 1]} ({v:.2f}×)" for i, v in bot.items())
        out.append(("النمط الموسمي",
                    f"بعد إزالة أثر النمو السنوي، الأشهر الأقوى موسمياً: {t}. والأضعف: {b} (المؤشر 1.00× = المتوسط الشهري للسنة). "
                    f"المؤشر مبني على سنتين فقط لأغلب الأشهر، فهو اتجاه إرشادي لا حكم نهائي."))
    except Exception:
        out.append(("النمط الموسمي", "لا تكفي البيانات لحساب مؤشر موسمي ضمن الفلاتر الحالية."))
    # 3) المنتجات
    try:
        tot = df[PRODUCTS].sum()
        sh = tot / tot.sum()
        l1, l2 = sh.nlargest(2).index
        txt = (f"{l1} هو الأعلى بحصة {sh[l1]:.1%} من الوحدات، يليه {l2} بنسبة {sh[l2]:.1%}؛ "
               f"ومعاً يمثلان {sh[l1] + sh[l2]:.1%} من الكميات.")
        s = _product_share_shift(df)
        if s:
            y0, y1, d, _ = s
            g, l = d.idxmax(), d.idxmin()
            txt += (f" الحصة الأسرع صعوداً بين {y0} و{y1}: {g} ({d[g] * 100:+.1f} نقطة مئوية)، "
                    f"والأكثر تراجعاً: {l} ({d[l] * 100:+.1f} نقطة).")
        out.append(("تفوق المنتجات", txt))
    except Exception:
        out.append(("تفوق المنتجات", "لا توجد كميات كافية ضمن الفلاتر الحالية."))
    # 4) المناطق
    try:
        at = area_table(df, complete, ref_year)
        top = at.index[0]
        txt = f"{top} تتصدر بحصة {at.Share.iloc[0]:.1%} من المبيعات"
        if len(at) > 1:
            txt += f"، تليها {at.index[1]} ({at.Share.iloc[1]:.1%})"
        txt += "."
        dead = [a for a in inactive_areas(df_full) if a in at.index]
        if dead:
            txt += f" مناطق بلا أي مبيعات منذ أكثر من 90 يوماً: {'، '.join(dead)}."
        out.append(("تفوق المناطق", txt))
    except Exception:
        out.append(("تفوق المناطق", "لا توجد بيانات مناطق ضمن الفلاتر الحالية."))
    # 5) الشذوذ
    try:
        parts = []
        if status["missing"]:
            parts.append("لا توجد أي مبيعات مسجلة في: " + "، ".join(mlabel(x) for x in status["missing"]))
        if status["partial"]:
            parts.append("أشهر جزئية (بيانات ناقصة): " + "، ".join(mlabel(x) for x in status["partial"]))
        m = df.groupby("YM").Amount.sum()
        m = m[m.index.isin(complete)]
        ch = m.pct_change()
        # القفزة تُحسب فقط بين شهرين متتاليين فعلاً
        ch = ch[[(i - pd.DateOffset(months=1)) in m.index for i in ch.index]]
        jumps = ch[ch > 0.45]
        if len(jumps):
            parts.append("قفزات شهرية حادة: " + "، ".join(f"{mlabel(i)} ({v:+.0%})" for i, v in jumps.items()))
        big = df.loc[df["Total Quantities"].idxmax()]
        parts.append(f"أكبر طلب منفرد: {fmt(big['Total Quantities'])} وحدة لعميل «{big['Customer - name']}» بقيمة {fmt(big.Amount)} ر.س ({big.Date:%Y-%m-%d})")
        unit = (df.Amount / df["Total Quantities"].replace(0, np.nan))
        med = unit.median()
        odd = df[(unit < 0.3 * med) & (df["Total Quantities"] >= 10)]
        if len(odd):
            r = odd.loc[odd["Total Quantities"].idxmax()]
            parts.append(f"{len(odd)} طلب كمية كبيرة بسعر وحدة أقل من 30% من الوسيط (مثلاً «{r['Customer - name']}»: {fmt(r['Total Quantities'])} وحدة بـ {fmt(r.Amount)} ر.س) – يستحق مراجعة الإدخال")
        out.append(("شذوذ وقفزات غير عادية", "؛ ".join(parts) + "."))
    except Exception:
        out.append(("شذوذ وقفزات غير عادية", "لا توجد بيانات كافية ضمن الفلاتر الحالية."))
    return out


def build_recommendations(df, df_full, status, ref_year):
    recs = []
    complete = status["complete"]
    # 1) المناطق
    try:
        dead = inactive_areas(df_full)
        at_full = df_full.groupby("Area").Amount.sum()
        tot = df_full.Amount.sum()
        top = at_full.idxmax()
        if dead:
            s = at_full[dead].sum()
            recs.append(("إعادة تفعيل المناطق المتوقفة وتحصين المناطق القائدة",
                         f"مناطق {'، '.join(dead)} توقفت مبيعاتها، وهي ساهمت تاريخياً بـ {fmt(s)} ر.س ({s / tot:.1%} من الإجمالي). "
                         f"خصّصوا خطة زيارات/عروض إعادة تنشيط لها وراجعوا سبب التوقف (توزيع، منافس، مندوب)، "
                         f"مع حماية {top} التي تحقق {at_full[top] / tot:.1%} من المبيعات عبر مخزون كافٍ وخدمة أسرع."))
        else:
            recs.append(("توسيع التغطية في المناطق الأضعف",
                         f"{top} تحقق {at_full[top] / tot:.1%} من المبيعات؛ ضعوا أهداف نمو للمناطق الأصغر وراجعوا تغطية المندوبين."))
    except Exception:
        recs.append(("تحسين التوزيع الجغرافي", "راجعوا توزيع المبيعات على المناطق وركّزوا على الأقل تغطية."))
    # 2) المنتجات
    try:
        s = _product_share_shift(df_full)
        tot = df_full[PRODUCTS].sum(); sh = tot / tot.sum()
        if s:
            y0, y1, d, _ = s
            g, l = d.idxmax(), d.idxmin()
            recs.append(("تحسين مزيج المنتجات بحزم بيع مترابطة",
                         f"{g} يكسب حصة ({d[g] * 100:+.1f} نقطة بين {y0} و{y1}) بينما يتراجع {l} ({d[l] * 100:+.1f} نقطة). "
                         f"اربطوا المنتجات الصاعدة بالمنتج القائد ({sh.idxmax()}، {sh.max():.1%} من الوحدات) في حزم/خصومات كمية، "
                         f"وضعوا خطة لاستعادة {l} عند عملاء الورش ومحطات الخدمة."))
        else:
            recs.append(("تحسين مزيج المنتجات", f"{sh.idxmax()} يقود المبيعات؛ اربطوا بقية المنتجات به في حزم عروض."))
    except Exception:
        recs.append(("تحسين مزيج المنتجات", "ادعموا المنتجات الأقل مبيعاً بحزم مع المنتج القائد."))
    # 3) الموسمية والعملاء الكبار
    try:
        si = seasonal_index(df_full, complete); si = si[si.N >= 2]
        top = "، ".join(AR_MONTHS[i - 1] for i in si.Index.nlargest(3).index)
        low = AR_MONTHS[si.Index.idxmin() - 1]
        cust = df_full.groupby("Customer - name").Amount.sum().nlargest(5)
        recs.append(("التخطيط الموسمي وبرنامج للعملاء الكبار",
                     f"ارفعوا المخزون وجهّزوا الحملات قبل الأشهر الأقوى ({top}) وأطلقوا عروض تنشيط في {low} الأضعف. "
                     f"أكبر 5 عملاء يمثلون {cust.sum() / df_full.Amount.sum():.1%} فقط من المبيعات (تنوع جيد)، "
                     f"لذا أنشئوا شرائح خصم كمية للطلبات الكبيرة (30+ وحدة) لتشجيع تكرارها، وعالجوا فجوات التسجيل في الأشهر الناقصة."))
    except Exception:
        recs.append(("التخطيط الموسمي", "خطّطوا المخزون والعروض وفق أشهر الذروة والركود."))
    return recs
