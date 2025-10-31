# -*- coding: utf-8 -*-
# mms_disconnected_deeptracker.py
import re
import math
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="MMS | Disconnected Deep Tracker", layout="wide")

# =========================
# Helpers
# =========================
def norm_col(c: str) -> str:
    return re.sub(r"\s+", " ", str(c).strip()).lower()

def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    nm = {norm_col(c): c for c in df.columns}
    for c in candidates:
        if norm_col(c) in nm:
            return nm[norm_col(c)]
    # partial fallback
    for c in df.columns:
        for cand in candidates:
            if norm_col(cand) in norm_col(c):
                return c
    return None

def infer_bucket_from_name(name: str) -> str:
    name = (name or "").lower()
    hints = {
        "استبدال": ["استبدال"],
        "تحسين": ["تحسين", "استخراج", "تحسين واستخراج"],
        "صيانة": ["صيانة"],
        "كشف": ["كشف", "معاينة", "كشف ومعاينة"],
        "فحص": ["فحص", "inspection", "بور اوف", "power off"]
    }
    for b, kws in hints.items():
        for kw in kws:
            if kw in name:
                return b
    return "غير محدد"

# ---- Column candidate sets (AR/EN) ----
# Disconnected
DISC_METER_CANDS = ["HES Device Id", "Meter Number", "Meter No", "رقم العداد"]
DISC_SITE_CANDS  = ["Utility Site Id", "Functional Location", "الموقع الوظيفي"]
DISC_LAST_CANDS  = ["Last Daily", "Last Communication", "آخر قراءة", "آخر اتصال"]
DISC_LAT_CANDS   = ["Latitude"]
DISC_LON_CANDS   = ["Longitude"]

# Shared tasks (inspection / maintenance)
METER_CANDS   = ["Meter No", "Meter Number", "HES Device Id", "رقم العداد"]
FUNCLOC_CANDS = ["Functional Location", "Utility Site Id", "الموقع الوظيفي"]
ACC_CANDS     = ["SAP Account Number", "Account Number", "رقم الحساب"]

REG_CANDS     = ["Task Registration Date Time", "Request Registration Date Time", "تاريخ تسجيل المهمة", "تاريخ تسجيل الطلب"]
CLOSE_CANDS   = ["Task Closed Time", "Task Completed Time", "وقت إقفال المهمة", "تاريخ إقفال المهمة"]
STATUS_CANDS  = ["Task Status", "Request Status", "حالة المهمة", "حالة الطلب"]
RESULT_CANDS  = ["Technician's Result", "Final Result", "Final Result (Dispatcher's Result)", "نتيجة الفني", "النتيجة النهائية"]
TYPE_CANDS    = ["Request Type", "نوع الطلب"]
VIP_CANDS     = ["Subscription VIP", "VIP", "نوع المشترك", "Account Type", "تصنيف المشترك"]

# =========================
# Readers
# =========================
def read_disconnected(file):
    xl = pd.ExcelFile(file)
    sheet = xl.sheet_names[0]
    df = xl.parse(sheet)

    d_meter = pick_col(df, DISC_METER_CANDS)
    d_site  = pick_col(df, DISC_SITE_CANDS)
    d_last  = pick_col(df, DISC_LAST_CANDS)
    d_lat   = pick_col(df, DISC_LAT_CANDS)
    d_lon   = pick_col(df, DISC_LON_CANDS)

    if d_last:
        df[d_last] = pd.to_datetime(df[d_last], errors="coerce")

    # Build keys
    df["_KEY_METER"] = df[d_meter].astype(str).str.strip() if d_meter else ""
    if d_site:
        df["_KEY_SITE"]  = df[d_site].astype(str).str.strip()

    meta = dict(sheet=sheet, d_meter=d_meter, d_site=d_site, d_last=d_last, d_lat=d_lat, d_lon=d_lon, columns=list(df.columns))
    return df, meta

def read_tasks(file, is_inspection=False):
    xl = pd.ExcelFile(file)
    sheet = xl.sheet_names[0]
    df = xl.parse(sheet)

    meter_col  = pick_col(df, METER_CANDS)
    funcloc_col= pick_col(df, FUNCLOC_CANDS)
    acc_col    = pick_col(df, ACC_CANDS)
    reg_col    = pick_col(df, REG_CANDS)
    close_col  = pick_col(df, CLOSE_CANDS)
    status_col = pick_col(df, STATUS_CANDS)
    result_col = pick_col(df, RESULT_CANDS)
    type_col   = pick_col(df, TYPE_CANDS)
    vip_col    = pick_col(df, VIP_CANDS)

    if reg_col:   df[reg_col] = pd.to_datetime(df[reg_col], errors="coerce")
    if close_col: df[close_col] = pd.to_datetime(df[close_col], errors="coerce")

    if "Bucket" not in df.columns:
        df["Bucket"] = "فحص" if is_inspection else infer_bucket_from_name(getattr(file, "name", ""))

    # open/closed
    is_closed = df[close_col].notna() if close_col else False
    df["__is_open__"] = ~is_closed if close_col else True

    # keys
    df["_KEY_METER"] = df[meter_col].astype(str).str.strip() if meter_col else ""
    if funcloc_col:
        df["_KEY_SITE"]  = df[funcloc_col].astype(str).str.strip()

    meta = dict(
        sheet=sheet, meter_col=meter_col, funcloc_col=funcloc_col, acc_col=acc_col,
        reg_col=reg_col, close_col=close_col, status_col=status_col, result_col=result_col,
        type_col=type_col, vip_col=vip_col, columns=list(df.columns)
    )
    return df, meta

# =========================
# UI
# =========================
st.title("🔍 MMS — Disconnected × Inspections × Maintenance (Deep Tracker)")
st.caption("تحليل دقيق وربط ثلاثي: العدادات غير المتصلة ↔ مهام الفحص ↔ مهام الصيانة، مع مقارنة تواريخ (قبل/بعد) آخر اتصال.")

with st.sidebar:
    st.header("📁 ملف العدادات غير المتصلة (إلزامي)")
    dis_file = st.file_uploader("اختر ملف غير المتصلين", type=["xlsx","xls"], accept_multiple_files=False)

    st.header("📁 ملفات مهام الفحص (اختياري 0..N)")
    insp_files = st.file_uploader("اختر ملف/ملفات الفحص", type=["xlsx","xls"], accept_multiple_files=True)

    st.header("📁 ملفات مهام الصيانة (اختياري 0..N)")
    maint_files = st.file_uploader("اختر ملف/ملفات الصيانة", type=["xlsx","xls"], accept_multiple_files=True)

    st.markdown("---")
    st.header("⚙️ إعدادات")
    sla_days = st.number_input("حد التأخير (SLA) بالأيام", 1, 60, 3)
    join_on_site = st.checkbox("اسمح بالربط بالموقع الوظيفي إذا رقم العداد مفقود/غير متطابق", value=True)

if not dis_file:
    st.info("✨ ارفع ملف غير المتصلين أولاً.")
    st.stop()

# Read disconnected
dis_df, dis_meta = read_disconnected(dis_file)
d_meter, d_site, d_last = dis_meta["d_meter"], dis_meta["d_site"], dis_meta["d_last"]
st.success("تم تحميل ملف غير المتصلين ✅")
st.write("**أعمدة مفاتيح غير المتصلين:**", {"Meter": d_meter, "Utility Site": d_site, "Last Daily": d_last})

# Read inspections
insp_df = pd.DataFrame(); insp_metas = []
if insp_files:
    tmp = []
    for f in insp_files:
        df, meta = read_tasks(f, is_inspection=True)
        tmp.append(df); insp_metas.append(meta)
    insp_df = pd.concat(tmp, ignore_index=True, sort=False)
    st.success(f"تم تحميل ملفات الفحص: {len(insp_files)} ✅")

# Read maintenance
maint_df = pd.DataFrame(); maint_metas = []
if maint_files:
    tmp = []
    for f in maint_files:
        df, meta = read_tasks(f, is_inspection=False)
        tmp.append(df); maint_metas.append(meta)
    maint_df = pd.concat(tmp, ignore_index=True, sort=False)
    st.success(f"تم تحميل ملفات الصيانة: {len(maint_files)} ✅")

# Resolve columns
getcol = lambda metas, key: next((m[key] for m in metas if m.get(key)), None) if metas else None
i_meter = getcol(insp_metas, "meter_col")
i_reg   = getcol(insp_metas, "reg_col")
i_close = getcol(insp_metas, "close_col")
i_status= getcol(insp_metas, "status_col")
i_result= getcol(insp_metas, "result_col")

m_meter = getcol(maint_metas, "meter_col")
m_reg   = getcol(maint_metas, "reg_col")
m_close = getcol(maint_metas, "close_col")
m_status= getcol(maint_metas, "status_col")
m_result= getcol(maint_metas, "result_col")

# =========================
# Summary builders
# =========================
def summarize_inspections(df):
    if df.empty: 
        return pd.DataFrame(columns=["_KEY_METER","insp_total","insp_open","insp_latest_result","insp_latest_date","insp_any_before_last","insp_any_after_last"])
    # latest dates
    latest_sort = df[i_close] if (i_close and i_close in df.columns) else df[i_reg] if (i_reg and i_reg in df.columns) else None
    if latest_sort is not None:
        df["_latest_sort"] = pd.to_datetime(latest_sort, errors="coerce")
    else:
        df["_latest_sort"] = pd.NaT

    # open
    if i_close and i_close in df.columns:
        open_mask = df[i_close].isna()
    elif i_status and i_status in df.columns:
        open_mask = df[i_status].astype(str).str.lower().ne("closed")
    else:
        open_mask = pd.Series(True, index=df.index)

    grp = df.groupby("_KEY_METER")
    out = grp.size().reset_index(name="insp_total")
    out = out.merge(df[open_mask].groupby("_KEY_METER").size().rename("insp_open"),
                    how="left", left_on="_KEY_METER", right_index=True)
    latest = df.sort_values("_latest_sort").groupby("_KEY_METER").tail(1)
    show_cols = ["_KEY_METER"]
    if i_result and i_result in latest.columns: show_cols.append(i_result)
    show_cols += [c for c in [i_reg, i_close] if c and c in latest.columns]
    latest = latest[show_cols]
    latest = latest.rename(columns={i_result: "insp_latest_result", i_reg: "insp_reg", i_close:"insp_close"})
    latest["insp_latest_date"] = latest["insp_close"].fillna(latest["insp_reg"])
    out = out.merge(latest, how="left", on="_KEY_METER")
    return out

def summarize_maintenance(df):
    if df.empty:
        return pd.DataFrame(columns=["_KEY_METER","mnt_total","mnt_open","mnt_latest_status","mnt_latest_bucket","mnt_latest_date","mnt_any_before_last","mnt_any_after_last"])
    # latest dates
    latest_sort = df[m_close] if (m_close and m_close in df.columns) else df[m_reg] if (m_reg and m_reg in df.columns) else None
    if latest_sort is not None:
        df["_latest_sort"] = pd.to_datetime(latest_sort, errors="coerce")
    else:
        df["_latest_sort"] = pd.NaT

    # open tasks
    if m_close and m_close in df.columns:
        open_mask = df[m_close].isna()
    elif m_status and m_status in df.columns:
        open_mask = df[m_status].astype(str).str.lower().ne("closed")
    else:
        open_mask = pd.Series(True, index=df.index)

    grp = df.groupby("_KEY_METER")
    out = grp.size().reset_index(name="mnt_total")
    out = out.merge(df[open_mask].groupby("_KEY_METER").size().rename("mnt_open"),
                    how="left", left_on="_KEY_METER", right_index=True)

    latest = df.sort_values("_latest_sort").groupby("_KEY_METER").tail(1)
    cols = ["_KEY_METER"]
    if m_status and m_status in latest.columns: cols.append(m_status)
    cols += ["Bucket"] if "Bucket" in latest.columns else []
    cols += [c for c in [m_reg, m_close] if c and c in latest.columns]
    latest = latest[cols]
    latest = latest.rename(columns={m_status:"mnt_latest_status", m_reg:"mnt_reg", m_close:"mnt_close", "Bucket":"mnt_latest_bucket"})
    latest["mnt_latest_date"] = latest["mnt_close"].fillna(latest["mnt_reg"])
    out = out.merge(latest, how="left", on="_KEY_METER")
    return out

# =========================
# Build core keys on tasks if meter miss & join_on_site
# =========================
if join_on_site and d_site:
    if not insp_df.empty and "_KEY_METER" in insp_df.columns and insp_df["_KEY_METER"].eq("").any() and "_KEY_SITE" in insp_df.columns:
        # fallback: map site to meter using disconnected data
        site_to_meter = dis_df[[d_site, "_KEY_METER"]].dropna().drop_duplicates().set_index(d_site)["_KEY_METER"]
        insp_df.loc[insp_df["_KEY_METER"].eq("") & insp_df["_KEY_SITE"].notna(), "_KEY_METER"] = insp_df["_KEY_SITE"].map(site_to_meter).fillna("")
    if not maint_df.empty and "_KEY_METER" in maint_df.columns and maint_df["_KEY_METER"].eq("").any() and "_KEY_SITE" in maint_df.columns:
        site_to_meter = dis_df[[d_site, "_KEY_METER"]].dropna().drop_duplicates().set_index(d_site)["_KEY_METER"]
        maint_df.loc[maint_df["_KEY_METER"].eq("") & maint_df["_KEY_SITE"].notna(), "_KEY_METER"] = maint_df["_KEY_SITE"].map(site_to_meter).fillna("")

# =========================
# Summaries
# =========================
insp_sum  = summarize_inspections(insp_df) if not insp_df.empty else pd.DataFrame(columns=["_KEY_METER"])
maint_sum = summarize_maintenance(maint_df) if not maint_df.empty else pd.DataFrame(columns=["_KEY_METER"])

# Merge onto disconnected
summary = dis_df.copy()
summary = summary.merge(insp_sum,  how="left", on="_KEY_METER")
summary = summary.merge(maint_sum, how="left", on="_KEY_METER")

# =========================
# Before/After Last Daily flags
# =========================
def compare_to_last(df, time_col, last_col):
    if (time_col is None) or (last_col is None) or (time_col not in df.columns) or (last_col not in df.columns):
        return pd.Series([np.nan]*len(df))
    t = pd.to_datetime(df[time_col], errors="coerce")
    l = pd.to_datetime(df[last_col], errors="coerce")
    # -1: before, 0: equal/unknown, +1: after
    return np.where(t.notna() & l.notna(), np.sign((t - l).dt.total_seconds()), np.nan)

# للإحصاء “هل حصل أي حدث قبل/بعد آخر اتصال؟” نحتاج أدنى/أقصى تواريخ على مستوى العداد:
def any_event_relative(tasks_df, key_col, reg_col, close_col, dis_last_series):
    if tasks_df.empty: 
        return pd.DataFrame(columns=[key_col, "any_before_last", "any_after_last"])
    # تواريخ محتملة للأحداث: التسجيل/الإقفال
    tdf = tasks_df.copy()
    tdf["_EVENT_DATE"] = pd.NaT
    if close_col and close_col in tdf.columns:
        tdf["_EVENT_DATE"] = pd.to_datetime(tdf[close_col], errors="coerce")
    if reg_col and reg_col in tdf.columns:
        # خذ الأقدم/الأحدث لاحقًا، لكن الآن سنسقط الاثنين ونستخرج min/max
        tdf["_REG_DATE__"] = pd.to_datetime(tdf[reg_col], errors="coerce")

    agg = tdf.groupby("_KEY_METER").agg(
        min_event=(" _EVENT_DATE", lambda s: pd.to_datetime(s, errors="coerce").min()) if "_EVENT_DATE" in tdf.columns else ("_EVENT_DATE", "min"),
        max_event=(" _EVENT_DATE", lambda s: pd.to_datetime(s, errors="coerce").max()) if "_EVENT_DATE" in tdf.columns else ("_EVENT_DATE", "max"),
        min_reg=(" _REG_DATE__", lambda s: pd.to_datetime(s, errors="coerce").min()) if "_REG_DATE__" in tdf.columns else ("_REG_DATE__", "min"),
        max_reg=(" _REG_DATE__", lambda s: pd.to_datetime(s, errors="coerce").max()) if "_REG_DATE__" in tdf.columns else ("_REG_DATE__", "max"),
    )
    # أعمدة قد تبدو غير دقيقة بالمسافات؛ نصلحها:
    cols_rename = {}
    for c in list(agg.columns):
        cols_rename[c] = c.replace(" _", "_")
    agg = agg.rename(columns=cols_rename)

    # اختر تواريخ مقارنة (نأخذ أي تاريخ متاح)
    agg["earliest"] = agg[["min_event","min_reg"]].min(axis=1)
    agg["latest"]   = agg[["max_event","max_reg"]].max(axis=1)

    out = agg[["earliest","latest"]].copy()
    out = out.join(dis_last_series.rename("LastDaily"), how="left")
    out["any_before_last"] = (out["earliest"].notna()) & (out["LastDaily"].notna()) & (out["earliest"] < out["LastDaily"])
    out["any_after_last"]  = (out["latest"].notna())   & (out["LastDaily"].notna()) & (out["latest"]  > out["LastDaily"])
    out = out.reset_index().rename(columns={"index":"_KEY_METER"})
    return out[["_KEY_METER","any_before_last","any_after_last"]]

# سلاسل “آخر اتصال” على مستوى العداد
last_series = summary.set_index("_KEY_METER")[d_last] if d_last else pd.Series(dtype="datetime64[ns]")

insp_rel  = any_event_relative(insp_df, "_KEY_METER", i_reg, i_close, last_series)  if not insp_df.empty else pd.DataFrame(columns=["_KEY_METER","any_before_last","any_after_last"])
maint_rel = any_event_relative(maint_df, "_KEY_METER", m_reg, m_close, last_series) if not maint_df.empty else pd.DataFrame(columns=["_KEY_METER","any_before_last","any_after_last"])

summary = summary.merge(insp_rel.add_prefix("insp_") if len(insp_rel) else insp_rel,
                        how="left", left_on="_KEY_METER", right_on="insp__KEY_METER").drop(columns=["insp__KEY_METER"], errors="ignore")
summary = summary.merge(maint_rel.add_prefix("mnt_") if len(maint_rel) else maint_rel,
                        how="left", left_on="_KEY_METER", right_on="mnt__KEY_METER").drop(columns=["mnt__KEY_METER"], errors="ignore")

# =========================
# Next Action classification
# =========================
def next_action(row):
    # أولوية: صيانة مفتوحة > فحص مفتوح > افتح فحص جديد
    mnt_open = (row.get("mnt_open", 0) or 0) > 0
    insp_open= (row.get("insp_open",0) or 0) > 0
    if mnt_open: return "تسريع صيانة مفتوحة"
    if insp_open:return "متابعة فحص مفتوح"
    return "يفتح فحص جديد"

summary["Next Action"] = summary.apply(next_action, axis=1)

# =========================
# KPIs
# =========================
st.markdown("## 📊 مؤشرات عامة")
k1,k2,k3,k4 = st.columns(4)
k1.metric("عدادات غير متصلة", f"{summary['_KEY_METER'].nunique():,}")
k2.metric("لها فحص مفتوح", f"{int(summary['insp_open'].fillna(0).gt(0).sum()) if 'insp_open' in summary.columns else 0:,}")
k3.metric("لها صيانة مفتوحة", f"{int(summary['mnt_open'].fillna(0).gt(0).sum()) if 'mnt_open' in summary.columns else 0:,}")
k4.metric("تجاوز SLA (إرشادي)", f"{sla_days} يوم")

# =========================
# Unified table
# =========================
st.markdown("## 📋 الجدول الموحد لكل عداد")
display_cols = []

# من disconnected
for c in [d_meter, d_site, d_last, "Office", "States", "Logistic State", "Gateway Id", "Latitude", "Longitude"]:
    if c and c in summary.columns: display_cols.append(c)

# مؤشرات الفحص
display_cols += ["insp_total","insp_open","insp_latest_result","insp_latest_date","insp_any_before_last","insp_any_after_last"]
# مؤشرات الصيانة
display_cols += ["mnt_total","mnt_open","mnt_latest_status","mnt_latest_bucket","mnt_latest_date","mnt_any_before_last","mnt_any_after_last"]
# الإجراء القادم
display_cols += ["Next Action"]

display_cols = [c for c in display_cols if c in summary.columns]
st.dataframe(summary[display_cols].sort_values(["Next Action","mnt_open","insp_open"], ascending=[True, False, False]), use_container_width=True)

st.markdown("---")
st.markdown("### تفريغ القوائم حسب الإجراء القادم")
col_a, col_b, col_c = st.columns(3)
with col_a:
    st.markdown("**تسريع صيانة مفتوحة**")
    accel = summary[summary["Next Action"]=="تسريع صيانة مفتوحة"][display_cols]
    st.dataframe(accel, use_container_width=True)
with col_b:
    st.markdown("**متابعة فحص مفتوح**")
    follow = summary[summary["Next Action"]=="متابعة فحص مفتوح"][display_cols]
    st.dataframe(follow, use_container_width=True)
with col_c:
    st.markdown("**يفتح فحص جديد**")
    create = summary[summary["Next Action"]=="يفتح فحص جديد"][display_cols]
    st.dataframe(create, use_container_width=True)

# =========================
# Optional timelines (per meter)
# =========================
st.markdown("---")
st.markdown("## ⏱️ عرض زمني (اختياري) — اختر عدادًا")
sel_meter = st.selectbox("اختر عداد لعرض التسلسل الزمني", options=summary["_KEY_METER"].dropna().unique().tolist())
if sel_meter:
    events = []
    # last daily
    if d_last and d_last in summary.columns:
        ld = summary.loc[summary["_KEY_METER"]==sel_meter, d_last].iloc[0]
        events.append({"when": ld, "type": "LastDaily", "desc": "آخر اتصال"})

    # inspection events
    if not insp_df.empty and "_KEY_METER" in insp_df.columns:
        tmp = insp_df[insp_df["_KEY_METER"]==sel_meter].copy()
        if i_reg:   tmp["__when_reg"]   = pd.to_datetime(tmp[i_reg], errors="coerce")
        if i_close: tmp["__when_close"] = pd.to_datetime(tmp[i_close], errors="coerce")
        for _, r in tmp.iterrows():
            if i_reg and pd.notna(r["__when_reg"]):
                events.append({"when": r["__when_reg"], "type":"Inspection-Reg", "desc":"تسجيل فحص"})
            if i_close and pd.notna(r["__when_close"]):
                events.append({"when": r["__when_close"], "type":"Inspection-Close", "desc":"إقفال فحص"})

    # maintenance events
    if not maint_df.empty and "_KEY_METER" in maint_df.columns:
        tmp = maint_df[maint_df["_KEY_METER"]==sel_meter].copy()
        if m_reg:   tmp["__when_reg"]   = pd.to_datetime(tmp[m_reg], errors="coerce")
        if m_close: tmp["__when_close"] = pd.to_datetime(tmp[m_close], errors="coerce")
        for _, r in tmp.iterrows():
            if m_reg and pd.notna(r["__when_reg"]):
                events.append({"when": r["__when_reg"], "type":"Maint-Reg", "desc":"تسجيل صيانة"})
            if m_close and pd.notna(r["__when_close"]):
                events.append({"when": r["__when_close"], "type":"Maint-Close", "desc":"إقفال صيانة"})

    ev = pd.DataFrame(events)
    if not ev.empty:
        ev = ev.sort_values("when")
        st.dataframe(ev, use_container_width=True)
    else:
        st.info("لا توجد أحداث زمنيّة لهذا العداد.")

# =========================
# Downloads
# =========================
st.markdown("---")
dl1, dl2, dl3 = st.columns(3)
with dl1:
    st.download_button("⬇️ تنزيل الجدول الموحد (CSV)", data=summary.to_csv(index=False).encode("utf-8-sig"),
                       file_name="disconnected_deeptracker_summary.csv", mime="text/csv")
with dl2:
    st.download_button("⬇️ تسريع صيانة مفتوحة (CSV)", data=accel.to_csv(index=False).encode("utf-8-sig"),
                       file_name="accelerate_open_maintenance.csv", mime="text/csv")
with dl3:
    st.download_button("⬇️ يفتح فحص جديد (CSV)", data=create.to_csv(index=False).encode("utf-8-sig"),
                       file_name="create_new_inspection.csv", mime="text/csv")

st.markdown("---")
st.caption("MMS — Disconnected Deep Tracker: ربط وتحليل ثلاثي مع مقارنة قبل/بعد Last Daily، وتتبع شامل للحالة.")
