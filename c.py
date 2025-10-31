# -*- coding: utf-8 -*-
# mms_premise_tracker.py — Premise as the primary key + robust "Last Daily" parsing
import re
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="MMS | Premise Tracker", layout="wide")

# ================= Helpers =================
def norm_col(c: str) -> str:
    return re.sub(r"\s+", " ", str(c).strip()).lower()

def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    if df is None or df.empty:
        return None
    nm = {norm_col(c): c for c in df.columns}
    for c in candidates:
        if norm_col(c) in nm:
            return nm[norm_col(c)]
    # loose contains
    for c in df.columns:
        for cand in candidates:
            if norm_col(cand) in norm_col(c):
                return c
    return None

def smart_parse_datetime(series: pd.Series) -> pd.Series:
    """Parse mixed text + Excel serials safely."""
    if series is None:
        return pd.Series([], dtype="datetime64[ns]")
    s = series.copy()

    def clean(x):
        if pd.isna(x):
            return np.nan
        x = str(x).strip()
        if x == "" or x.lower() in {"none", "nan", "null", "-", "—", "0"}:
            return np.nan
        return x
    s = s.map(clean)

    parsed = pd.to_datetime(s, errors="coerce", dayfirst=True, infer_datetime_format=True)

    # Excel serial fallback
    need_excel = parsed.isna()
    if need_excel.any():
        as_num = pd.to_numeric(s.where(need_excel), errors="coerce")
        excel_mask = as_num.notna()
        if excel_mask.any():
            parsed.loc[excel_mask] = pd.to_datetime(as_num[excel_mask], unit="d", origin="1899-12-30", errors="coerce")
    return parsed

def to_excel_download(df: pd.DataFrame, filename_prefix="premise_tracker_results") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")
    return output.getvalue()

# ================= UI: Uploads =================
st.title("📊 نظام تتبع العدادات غير المتصلة — Premise Tracker")

with st.sidebar:
    st.header("📁 ملفات الإدخال")
    disconnected_file = st.file_uploader("ملف العدادات غير المتصلة", type=["xlsx","xls"])
    insp_files  = st.file_uploader("ملفات مهام الفحص (0..N)",  type=["xlsx","xls"], accept_multiple_files=True)
    maint_files = st.file_uploader("ملفات مهام الصيانة (0..N)", type=["xlsx","xls"], accept_multiple_files=True)
    st.markdown("---")
    start_btn = st.button("🚀 ابدأ التحليل")

if not start_btn or not disconnected_file:
    st.info("⬆️ ارفع الملفات المطلوبة ثم اضغط **ابدأ التحليل**.")
    st.stop()

# ================= Read disconnected =================
st.subheader("📘 قراءة ملف العدادات غير المتصلة (Premise = المفتاح)")

dis_df = pd.read_excel(disconnected_file)

# Premise في هذا الملف = Utility Site Id
premise_col_dis = pick_col(dis_df, ["Utility Site Id", "Premise", "رقم المكان"])
last_daily_col  = pick_col(dis_df, [
    "Last Daily", "LastDaily", "Last Communication", "Last Comm",
    "آخر قراءة", "آخر اتصال", "اخر اتصال", "اخر قراءة"
])

if not premise_col_dis:
    st.error("لم يتم العثور على عمود رقم المكان في ملف غير المتصلين (توقّع: Utility Site Id).")
    st.stop()

# Create normalized key
dis_df["_KEY_PREMISE"] = dis_df[premise_col_dis].astype(str).str.strip()

# Parse Last Daily if exists
if last_daily_col and last_daily_col in dis_df.columns:
    parsed = smart_parse_datetime(dis_df[last_daily_col])
    dis_df["LastDaily"] = parsed
    success = int(parsed.notna().sum())
    st.success(f"تحويل عمود آخر اتصال '{last_daily_col}': {success}/{len(dis_df)} قيماً صالحة.")
else:
    dis_df["LastDaily"] = pd.NaT
    st.warning("⚠️ لم يُعثر على عمود 'آخر اتصال' — سيتم اعتباره فارغاً (NaT).")

# نحضّر ملخص الأساس
# (لا نفترض وجود عمود LastDaily في الأصل — نستخدم العمود الموحّد الذي أنشأناه)
summary_cols = ["_KEY_PREMISE", "LastDaily"] + [c for c in dis_df.columns if c not in ["_KEY_PREMISE", "LastDaily"]]
summary = dis_df[summary_cols].copy()

# ================= Read tasks (multi files) =================
def load_task_files(files, kind_label: str) -> pd.DataFrame:
    """Normalize multiple files to a unified schema using Premise as key."""
    if not files:
        return pd.DataFrame()

    all_rows = []
    for f in files:
        df = pd.read_excel(f)

        # Premise عمود المفتاح في ملفات الصيانة/الفحص
        premise_col = pick_col(df, ["Premise", "Utility Site Id", "رقم المكان"])
        reg_col     = pick_col(df, ["Task Registration Date Time", "Request Registration Date Time", "تاريخ التسجيل", "تاريخ تسجيل المهمة", "تاريخ تسجيل الطلب"])
        close_col   = pick_col(df, ["Task Closed Time", "Task Completed Time", "تاريخ الإقفال", "وقت إقفال المهمة"])
        status_col  = pick_col(df, ["Task Status", "Request Status", "الحالة", "حالة المهمة", "حالة الطلب"])
        result_col  = pick_col(df, ["Technician's Result", "Final Result", "Final Result (Dispatcher's Result)", "النتيجة النهائية", "نتيجة الفني"])

        if not premise_col:
            # نتجاهل هذا الملف مع تحذير
            st.warning(f"تم تجاهل الملف '{getattr(f, 'name', 'file')}' لعدم وجود عمود Premise/Utility Site Id.")
            continue

        tmp = pd.DataFrame()
        tmp["_KEY_PREMISE"] = df[premise_col].astype(str).str.strip()
        if reg_col   and reg_col   in df.columns: tmp["reg_time"]   = smart_parse_datetime(df[reg_col])
        else:                                      tmp["reg_time"]   = pd.NaT
        if close_col and close_col in df.columns: tmp["close_time"] = smart_parse_datetime(df[close_col])
        else:                                      tmp["close_time"] = pd.NaT
        if status_col and status_col in df.columns: tmp["status"]   = df[status_col].astype(str)
        else:                                        tmp["status"]   = np.nan
        if result_col and result_col in df.columns: tmp["result"]   = df[result_col].astype(str)
        else:                                        tmp["result"]   = np.nan

        tmp["bucket"] = kind_label
        tmp["source"] = getattr(f, "name", kind_label)
        all_rows.append(tmp)

    if not all_rows:
        return pd.DataFrame()
    return pd.concat(all_rows, ignore_index=True, sort=False)

insp_df  = load_task_files(insp_files,  "فحص")
maint_df = load_task_files(maint_files, "صيانة")

# ================= Summaries =================
def summarize_tasks(df: pd.DataFrame, prefix: str) -> pd.DataFrame:
    """Aggregate per Premise to: total, open, latest status/result/date."""
    if df.empty:
        return pd.DataFrame(columns=["_KEY_PREMISE"])

    df = df.copy()
    # تعريف المفتوح = لا يوجد close_time
    df["_is_open"] = df["close_time"].isna()

    # أحدث تاريخ (close ثم reg)
    df["_latest_date"] = df["close_time"].fillna(df["reg_time"])

    # آخر حالة/نتيجة حسب أحدث تاريخ
    idx = df.groupby("_KEY_PREMISE")["_latest_date"].idxmax()
    latest = df.loc[idx, ["_KEY_PREMISE", "status", "result", "_latest_date"]].rename(
        columns={"status": f"{prefix}latest_status", "result": f"{prefix}latest_result", "_latest_date": f"{prefix}latest_date"}
    )

    agg = df.groupby("_KEY_PREMISE").agg(
        **{f"{prefix}total": ("_KEY_PREMISE", "count"),
           f"{prefix}open":  ("_is_open", "sum")}
    ).reset_index()

    out = agg.merge(latest, on="_KEY_PREMISE", how="left")
    return out

insp_sum  = summarize_tasks(insp_df,  "insp_")
maint_sum = summarize_tasks(maint_df, "maint_")

# دمج على الملخص الأساسي
summary = summary.merge(insp_sum,  on="_KEY_PREMISE", how="left")
summary = summary.merge(maint_sum, on="_KEY_PREMISE", how="left")

# ================= Next Action =================
def next_action(row):
    if (row.get("maint_open", 0) or 0) > 0:
        return "تسريع صيانة مفتوحة"
    if (row.get("insp_open", 0) or 0) > 0:
        return "متابعة فحص مفتوح"
    return "يفتح فحص جديد"

summary["Next Action"] = summary.apply(next_action, axis=1)

# ================= KPIs =================
st.subheader("📈 مؤشرات عامة")
c1, c2, c3 = st.columns(3)
c1.metric("عدد العدادات غير المتصلة", f"{summary['_KEY_PREMISE'].nunique():,}")
c2.metric("مهام فحص مفتوحة",        int(summary.get("insp_open",  pd.Series()).fillna(0).sum()))
c3.metric("مهام صيانة مفتوحة",       int(summary.get("maint_open", pd.Series()).fillna(0).sum()))

# ================= Table =================
st.subheader("📋 الجدول الموحد")
display_cols = []

# أعمدة أساسية من ملف غير المتصلين
for c in ["_KEY_PREMISE", "LastDaily"]:
    if c in summary.columns: display_cols.append(c)

# أضف بعض الحقول الأصلية إن وُجدت (مفيدة للتصفية)
for c in ["Office", "States", "Logistic State", "Gateway Id", "Latitude", "Longitude"]:
    if c in summary.columns: display_cols.append(c)

# ملخصات الفحص والصيانة
display_cols += [c for c in ["insp_total","insp_open","insp_latest_status","insp_latest_result","insp_latest_date",
                             "maint_total","maint_open","maint_latest_status","maint_latest_result","maint_latest_date"]
                 if c in summary.columns]
display_cols += ["Next Action"]

st.dataframe(summary[display_cols] if display_cols else summary, use_container_width=True)

# ================= Download (Excel) =================
excel_bytes = to_excel_download(summary)
st.download_button(
    label="⬇️ تنزيل النتائج (Excel)",
    data=excel_bytes,
    file_name=f"premise_tracker_results_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
