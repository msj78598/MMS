# -*- coding: utf-8 -*-
# mms_disconnected_deeptracker.py  — with "Start Analysis" button + manual column mapping

import re
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="MMS | Disconnected Deep Tracker", layout="wide")

# ============ helpers ============
def norm_col(c: str) -> str:
    return re.sub(r"\s+", " ", str(c).strip()).lower()

def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    if df is None or df.empty:
        return None
    nm = {norm_col(c): c for c in df.columns}
    for c in candidates:
        if norm_col(c) in nm:
            return nm[norm_col(c)]
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
        "كشف":   ["كشف", "معاينة", "كشف ومعاينة"],
        "فحص":   ["فحص", "inspection", "بور اوف", "power off"]
    }
    for b, kws in hints.items():
        for kw in kws:
            if kw in name:
                return b
    return "غير محدد"

# candidates
DISC_METER = ["HES Device Id", "Meter Number", "Meter No", "رقم العداد"]
DISC_SITE  = ["Utility Site Id", "Functional Location", "الموقع الوظيفي"]
DISC_LAST  = ["Last Daily", "Last Communication", "آخر قراءة", "آخر اتصال"]

METER_C    = ["Meter No", "Meter Number", "HES Device Id", "رقم العداد"]
SITE_C     = ["Functional Location", "Utility Site Id", "الموقع الوظيفي"]
REG_C      = ["Task Registration Date Time", "Request Registration Date Time", "تاريخ تسجيل المهمة", "تاريخ تسجيل الطلب"]
CLOSE_C    = ["Task Closed Time", "Task Completed Time", "وقت إقفال المهمة", "تاريخ إقفال المهمة"]
STATUS_C   = ["Task Status", "Request Status", "حالة المهمة", "حالة الطلب"]
RESULT_C   = ["Technician's Result", "Final Result", "Final Result (Dispatcher's Result)", "نتيجة الفني", "النتيجة النهائية"]

# ============ readers ============
def read_excel_first_sheet(file):
    xl = pd.ExcelFile(file)
    sheet = xl.sheet_names[0]
    df = xl.parse(sheet)
    return df, sheet

def read_disconnected(file):
    df, sheet = read_excel_first_sheet(file)
    d_meter = pick_col(df, DISC_METER)
    d_site  = pick_col(df, DISC_SITE)
    d_last  = pick_col(df, DISC_LAST)
    if d_last: df[d_last] = pd.to_datetime(df[d_last], errors="coerce")
    df["_KEY_METER"] = df[d_meter].astype(str).str.strip() if d_meter else ""
    if d_site: df["_KEY_SITE"] = df[d_site].astype(str).str.strip()
    return df, {"sheet": sheet, "d_meter": d_meter, "d_site": d_site, "d_last": d_last}

# ============ UI: uploads ============
st.title("🔍 MMS — Disconnected × Inspections × Maintenance (Deep Tracker)")
st.caption("لا يبدأ التحليل إلا بعد الضغط على زر **ابدأ التحليل**. يمكنك تعيين الأعمدة يدويًا إذا لم تُكتشف تلقائيًا.")

with st.sidebar:
    st.header("📁 ملف العدادات غير المتصلة (إلزامي)")
    dis_file = st.file_uploader("اختر ملف غير المتصلين", type=["xlsx","xls"], accept_multiple_files=False)

    st.header("📁 ملفات مهام الفحص (اختياري)")
    insp_files = st.file_uploader("اختر ملف/ملفات الفحص", type=["xlsx","xls"], accept_multiple_files=True)

    st.header("📁 ملفات مهام الصيانة (اختياري)")
    maint_files = st.file_uploader("اختر ملف/ملفات الصيانة", type=["xlsx","xls"], accept_multiple_files=True)

    st.markdown("---")
    st.header("⚙️ إعدادات")
    sla_days = st.number_input("حد التأخير (SLA) بالأيام", 1, 60, 3)
    join_on_site = st.checkbox("اسمح بالربط بالموقع الوظيفي إذا رقم العداد مفقود", value=True)

if not dis_file:
    st.info("✨ ارفع ملف غير المتصلين ثم اضغط **ابدأ التحليل**.")
    st.stop()

# disconnected preview
dis_df, dis_meta = read_disconnected(dis_file)
d_meter, d_site, d_last = dis_meta["d_meter"], dis_meta["d_site"], dis_meta["d_last"]
st.success("تم تحميل ملف غير المتصلين ✅")
with st.expander("أعمدة مفاتيح غير المتصلين (اكتشاف تلقائي):", expanded=True):
    st.json({"Meter": d_meter, "Utility Site": d_site, "Last Daily": d_last})

# optional previews for first files to help mapping
def preview_columns(files, title):
    if not files: 
        st.info(f"لم تُرفع {title}.")
        return None, None
    df, sheet = read_excel_first_sheet(files[0])
    with st.expander(f"👀 معاينة أعمدة أول ملف {title} ({getattr(files[0],'name','file')} | ورقة: {sheet})"):
        st.write(list(df.columns))
    return df, sheet

insp_preview_df, _ = preview_columns(insp_files, "للفحص")
maint_preview_df, _ = preview_columns(maint_files, "للصيانة")

# ============ manual mapping form ============
with st.form("mapping_form"):
    st.subheader("🧩 تعيين الأعمدة يدويًا (اختياري)")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**ملفات الفحص**")
        insp_meter_sel = st.selectbox("Meter (فحص)", options=(list(insp_preview_df.columns) if insp_preview_df is not None else []), index=(list(insp_preview_df.columns).index(pick_col(insp_preview_df, METER_C)) if insp_preview_df is not None and pick_col(insp_preview_df, METER_C) in insp_preview_df.columns else 0) if insp_preview_df is not None and len(insp_preview_df.columns)>0 else None)
        insp_reg_sel   = st.selectbox("Reg Time (فحص)", options=(list(insp_preview_df.columns) if insp_preview_df is not None else []), index=(list(insp_preview_df.columns).index(pick_col(insp_preview_df, REG_C)) if insp_preview_df is not None and pick_col(insp_preview_df, REG_C) in insp_preview_df.columns else 0) if insp_preview_df is not None and len(insp_preview_df.columns)>0 else None)
        insp_close_sel = st.selectbox("Close Time (فحص)", options=(list(insp_preview_df.columns) if insp_preview_df is not None else []), index=(list(insp_preview_df.columns).index(pick_col(insp_preview_df, CLOSE_C)) if insp_preview_df is not None and pick_col(insp_preview_df, CLOSE_C) in insp_preview_df.columns else 0) if insp_preview_df is not None and len(insp_preview_df.columns)>0 else None)
        insp_status_sel= st.selectbox("Status (فحص)", options=(list(insp_preview_df.columns) if insp_preview_df is not None else []), index=(list(insp_preview_df.columns).index(pick_col(insp_preview_df, STATUS_C)) if insp_preview_df is not None and pick_col(insp_preview_df, STATUS_C) in insp_preview_df.columns else 0) if insp_preview_df is not None and len(insp_preview_df.columns)>0 else None)
        insp_result_sel= st.selectbox("Result (فحص)", options=(list(insp_preview_df.columns) if insp_preview_df is not None else []), index=(list(insp_preview_df.columns).index(pick_col(insp_preview_df, RESULT_C)) if insp_preview_df is not None and pick_col(insp_preview_df, RESULT_C) in insp_preview_df.columns else 0) if insp_preview_df is not None and len(insp_preview_df.columns)>0 else None)
    with c2:
        st.markdown("**ملفات الصيانة**")
        maint_meter_sel = st.selectbox("Meter (صيانة)", options=(list(maint_preview_df.columns) if maint_preview_df is not None else []), index=(list(maint_preview_df.columns).index(pick_col(maint_preview_df, METER_C)) if maint_preview_df is not None and pick_col(maint_preview_df, METER_C) in maint_preview_df.columns else 0) if maint_preview_df is not None and len(maint_preview_df.columns)>0 else None)
        maint_reg_sel   = st.selectbox("Reg Time (صيانة)", options=(list(maint_preview_df.columns) if maint_preview_df is not None else []), index=(list(maint_preview_df.columns).index(pick_col(maint_preview_df, REG_C)) if maint_preview_df is not None and pick_col(maint_preview_df, REG_C) in maint_preview_df.columns else 0) if maint_preview_df is not None and len(maint_preview_df.columns)>0 else None)
        maint_close_sel = st.selectbox("Close Time (صيانة)", options=(list(maint_preview_df.columns) if maint_preview_df is not None else []), index=(list(maint_preview_df.columns).index(pick_col(maint_preview_df, CLOSE_C)) if maint_preview_df is not None and pick_col(maint_preview_df, CLOSE_C) in maint_preview_df.columns else 0) if maint_preview_df is not None and len(maint_preview_df.columns)>0 else None)
        maint_status_sel= st.selectbox("Status (صيانة)", options=(list(maint_preview_df.columns) if maint_preview_df is not None else []), index=(list(maint_preview_df.columns).index(pick_col(maint_preview_df, STATUS_C)) if maint_preview_df is not None and pick_col(maint_preview_df, STATUS_C) in maint_preview_df.columns else 0) if maint_preview_df is not None and len(maint_preview_df.columns)>0 else None)

    start = st.form_submit_button("🚀 ابدأ التحليل")

if not start:
    st.warning("ارفع الملفات (إن وُجدت) ثم اضغط **ابدأ التحليل**.")
    st.stop()

# ============ load all inspection/maintenance with mapping ============
def load_tasks(files, is_inspection, meter_sel, reg_sel, close_sel, status_sel, result_sel=None):
    if not files:
        return pd.DataFrame()
    dfs = []
    for f in files:
        df, sheet = read_excel_first_sheet(f)
        df = df.copy()
        # cast date columns
        for c in [reg_sel, close_sel]:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors="coerce")
        # add helpers
        df["_KEY_METER"] = df[meter_sel].astype(str).str.strip() if meter_sel in df.columns else ""
        # infer bucket if not exists
        if "Bucket" not in df.columns:
            df["Bucket"] = "فحص" if is_inspection else infer_bucket_from_name(getattr(f, "name", ""))
        dfs.append(df)
    out = pd.concat(dfs, ignore_index=True, sort=False)
    meta = {
        "meter": meter_sel, "reg": reg_sel, "close": close_sel,
        "status": status_sel, "result": result_sel
    }
    return out, meta

insp_df, insp_meta = (pd.DataFrame(), {}) if not insp_files else load_tasks(
    insp_files, True, insp_meter_sel, insp_reg_sel, insp_close_sel, insp_status_sel, insp_result_sel
)
maint_df, maint_meta = (pd.DataFrame(), {}) if not maint_files else load_tasks(
    maint_files, False, maint_meter_sel, maint_reg_sel, maint_close_sel, maint_status_sel
)

# fallback mapping by site if meter missing
if join_on_site and d_site:
    if not insp_df.empty and "_KEY_METER" in insp_df.columns and insp_df["_KEY_METER"].eq("").any():
        site_col = pick_col(insp_df, SITE_C)
        if site_col and d_site in dis_df.columns:
            site_to_meter = dis_df[[d_site, "_KEY_METER"]].dropna().drop_duplicates().set_index(d_site)["_KEY_METER"]
            insp_df["_KEY_METER"] = np.where(insp_df["_KEY_METER"].eq("") & insp_df[site_col].notna(),
                                             insp_df[site_col].astype(str).map(site_to_meter).fillna(insp_df["_KEY_METER"]),
                                             insp_df["_KEY_METER"])
    if not maint_df.empty and "_KEY_METER" in maint_df.columns and maint_df["_KEY_METER"].eq("").any():
        site_col = pick_col(maint_df, SITE_C)
        if site_col and d_site in dis_df.columns:
            site_to_meter = dis_df[[d_site, "_KEY_METER"]].dropna().drop_duplicates().set_index(d_site)["_KEY_METER"]
            maint_df["_KEY_METER"] = np.where(maint_df["_KEY_METER"].eq("") & maint_df[site_col].notna(),
                                              maint_df[site_col].astype(str).map(site_to_meter).fillna(maint_df["_KEY_METER"]),
                                              maint_df["_KEY_METER"])

# ============ summaries ============
i_meter, i_reg, i_close, i_status, i_result = insp_meta.get("meter"), insp_meta.get("reg"), insp_meta.get("close"), insp_meta.get("status"), insp_meta.get("result")
m_meter, m_reg, m_close, m_status         = maint_meta.get("meter"), maint_meta.get("reg"), maint_meta.get("close"), maint_meta.get("status")

def summarize_inspections(df):
    if df.empty: 
        return pd.DataFrame(columns=["_KEY_METER","insp_total","insp_open","insp_latest_result","insp_latest_date"])
    if i_close in df.columns:
        open_mask = df[i_close].isna()
        df["_latest_sort"] = pd.to_datetime(df[i_close], errors="coerce")
    elif i_status in df.columns:
        open_mask = df[i_status].astype(str).str.lower().ne("closed")
        df["_latest_sort"] = pd.to_datetime(df[i_reg], errors="coerce") if i_reg in df.columns else pd.NaT
    else:
        open_mask = pd.Series(True, index=df.index)
        df["_latest_sort"] = pd.to_datetime(df[i_reg], errors="coerce") if i_reg in df.columns else pd.NaT

    out = df.groupby("_KEY_METER").size().reset_index(name="insp_total")
    out = out.merge(df[open_mask].groupby("_KEY_METER").size().rename("insp_open"),
                    how="left", left_on="_KEY_METER", right_index=True)
    latest = df.sort_values("_latest_sort").groupby("_KEY_METER").tail(1)
    cols = ["_KEY_METER"]
    if i_result in latest.columns: cols.append(i_result)
    if i_reg    in latest.columns: cols.append(i_reg)
    if i_close  in latest.columns: cols.append(i_close)
    latest = latest[cols].rename(columns={i_result:"insp_latest_result", i_reg:"insp_reg", i_close:"insp_close"})
    latest["insp_latest_date"] = latest.get("insp_close", pd.Series(dtype="datetime64[ns]")).fillna(latest.get("insp_reg"))
    out = out.merge(latest, how="left", on="_KEY_METER")
    return out

def summarize_maintenance(df):
    if df.empty:
        return pd.DataFrame(columns=["_KEY_METER","mnt_total","mnt_open","mnt_latest_status","mnt_latest_bucket","mnt_latest_date"])
    if m_close in df.columns:
        open_mask = df[m_close].isna()
        df["_latest_sort"] = pd.to_datetime(df[m_close], errors="coerce")
    elif m_status in df.columns:
        open_mask = df[m_status].astype(str).str.lower().ne("closed")
        df["_latest_sort"] = pd.to_datetime(df[m_reg], errors="coerce") if m_reg in df.columns else pd.NaT
    else:
        open_mask = pd.Series(True, index=df.index)
        df["_latest_sort"] = pd.to_datetime(df[m_reg], errors="coerce") if m_reg in df.columns else pd.NaT

    out = df.groupby("_KEY_METER").size().reset_index(name="mnt_total")
    out = out.merge(df[open_mask].groupby("_KEY_METER").size().rename("mnt_open"),
                    how="left", left_on="_KEY_METER", right_index=True)
    latest = df.sort_values("_latest_sort").groupby("_KEY_METER").tail(1)
    cols = ["_KEY_METER"]
    if m_status in latest.columns: cols.append(m_status)
    if "Bucket" in latest.columns: cols.append("Bucket")
    if m_reg    in latest.columns: cols.append(m_reg)
    if m_close  in latest.columns: cols.append(m_close)
    latest = latest[cols].rename(columns={m_status:"mnt_latest_status", "Bucket":"mnt_latest_bucket", m_reg:"mnt_reg", m_close:"mnt_close"})
    latest["mnt_latest_date"] = latest.get("mnt_close", pd.Series(dtype="datetime64[ns]")).fillna(latest.get("mnt_reg"))
    out = out.merge(latest, how="left", on="_KEY_METER")
    return out

insp_sum  = summarize_inspections(insp_df) if not insp_df.empty else pd.DataFrame(columns=["_KEY_METER"])
maint_sum = summarize_maintenance(maint_df) if not maint_df.empty else pd.DataFrame(columns=["_KEY_METER"])

summary = dis_df.copy()
summary = summary.merge(insp_sum,  how="left", on="_KEY_METER")
summary = summary.merge(maint_sum, how="left", on="_KEY_METER")

# before/after last daily flags
def any_event_relative(tasks_df, key_col, reg_col, close_col, last_series):
    if tasks_df is None or tasks_df.empty or key_col not in tasks_df.columns:
        return pd.DataFrame(columns=[key_col, "any_before_last", "any_after_last"])
    t_reg   = pd.to_datetime(tasks_df[reg_col], errors="coerce")  if reg_col  in tasks_df.columns else pd.Series(pd.NaT, index=tasks_df.index)
    t_close = pd.to_datetime(tasks_df[close_col], errors="coerce") if close_col in tasks_df.columns else pd.Series(pd.NaT, index=tasks_df.index)
    event_min = pd.concat([t_close, t_reg], axis=1).min(axis=1)
    event_max = pd.concat([t_close, t_reg], axis=1).max(axis=1)
    agg = pd.DataFrame({key_col: tasks_df[key_col], "_min": event_min, "_max": event_max}) \
            .groupby(key_col).agg(earliest=("_min","min"), latest=("_max","max")).reset_index()
    last_df = last_series.rename("LastDaily").reset_index()
    last_df.columns = [key_col, "LastDaily"]
    out = agg.merge(last_df, how="left", on=key_col)
    out["any_before_last"] = out["earliest"].notna() & out["LastDaily"].notna() & (out["earliest"] < out["LastDaily"])
    out["any_after_last"]  = out["latest"].notna()   & out["LastDaily"].notna() & (out["latest"]  > out["LastDaily"])
    return out[[key_col, "any_before_last", "any_after_last"]]

last_series = summary.set_index("_KEY_METER")[d_last] if d_last else pd.Series(dtype="datetime64[ns]")
insp_rel  = any_event_relative(insp_df, "_KEY_METER", i_reg, m_close, last_series)  if not insp_df.empty else pd.DataFrame(columns=["_KEY_METER","any_before_last","any_after_last"])
maint_rel = any_event_relative(maint_df, "_KEY_METER", m_reg, m_close, last_series) if not maint_df.empty else pd.DataFrame(columns=["_KEY_METER","any_before_last","any_after_last"])

def safe_merge_relative(base: pd.DataFrame, rel_df: pd.DataFrame, prefix: str) -> pd.DataFrame:
    if rel_df is None or rel_df.empty or "_KEY_METER" not in rel_df.columns:
        return base
    rel_pref = rel_df.add_prefix(prefix)  # سيصبح العمود insp__KEY_METER
    key_col = f"{prefix}__KEY_METER"
    if key_col not in rel_pref.columns:
        key_col = f"{prefix}_KEY_METER"
    if key_col not in rel_pref.columns:
        return base
    rel_pref = rel_pref.rename(columns={key_col: "_KEY_METER"})
    return base.merge(rel_pref, how="left", on="_KEY_METER")

summary = safe_merge_relative(summary, insp_rel, "insp_")
summary = safe_merge_relative(summary, maint_rel, "mnt_")

def next_action(r):
    mnt_open = (r.get("mnt_open", 0) or 0) > 0
    insp_open= (r.get("insp_open",0) or 0) > 0
    if mnt_open: return "تسريع صيانة مفتوحة"
    if insp_open:return "متابعة فحص مفتوح"
    return "يفتح فحص جديد"

summary["Next Action"] = summary.apply(next_action, axis=1)

# KPIs
st.markdown("## 📊 مؤشرات عامة")
k1,k2,k3,k4 = st.columns(4)
k1.metric("عدادات غير متصلة", f"{summary['_KEY_METER'].nunique():,}")
k2.metric("لها فحص مفتوح", f"{int(summary.get('insp_open', pd.Series()).fillna(0).gt(0).sum()):,}")
k3.metric("لها صيانة مفتوحة", f"{int(summary.get('mnt_open', pd.Series()).fillna(0).gt(0).sum()):,}")
k4.metric("SLA التوجيهي", f"{sla_days} يوم")

# table
st.markdown("## 📋 الجدول الموحد لكل عداد")
display_cols = []
for c in [dis_meta["d_meter"], dis_meta["d_site"], dis_meta["d_last"], "Office", "States", "Logistic State", "Gateway Id", "Latitude", "Longitude"]:
    if c and c in summary.columns: display_cols.append(c)
display_cols += ["insp_total","insp_open","insp_latest_result","insp_latest_date","insp_any_before_last","insp_any_after_last"]
display_cols += ["mnt_total","mnt_open","mnt_latest_status","mnt_latest_bucket","mnt_latest_date","mnt_any_before_last","mnt_any_after_last"]
display_cols += ["Next Action"]
display_cols = [c for c in display_cols if c in summary.columns]

st.dataframe(summary[display_cols].sort_values(["Next Action",
                                                "mnt_open" if "mnt_open" in summary.columns else display_cols[0],
                                                "insp_open" if "insp_open" in summary.columns else display_cols[0]],
                                               ascending=[True, False, False]),
             use_container_width=True)

# downloads
st.markdown("---")
dl1, dl2, dl3 = st.columns(3)
accel = summary[summary["Next Action"]=="تسريع صيانة مفتوحة"][display_cols]
follow= summary[summary["Next Action"]=="متابعة فحص مفتوح"][display_cols]
create= summary[summary["Next Action"]=="يفتح فحص جديد"][display_cols]
with dl1:
    st.download_button("⬇️ تنزيل الجدول الموحد (CSV)", data=summary.to_csv(index=False).encode("utf-8-sig"),
                       file_name="disconnected_deeptracker_summary.csv", mime="text/csv")
with dl2:
    st.download_button("⬇️ تسريع صيانة مفتوحة (CSV)", data=accel.to_csv(index=False).encode("utf-8-sig"),
                       file_name="accelerate_open_maintenance.csv", mime="text/csv")
with dl3:
    st.download_button("⬇️ يفتح فحص جديد (CSV)", data=create.to_csv(index=False).encode("utf-8-sig"),
                       file_name="create_new_inspection.csv", mime="text/csv")
