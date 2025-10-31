# -*- coding: utf-8 -*-
# mms_premise_tracker.py — تحليل العدادات غير المتصلة باستخدام Premise كمفتاح

import re
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="MMS | Premise Tracker", layout="wide")

# ============ دوال مساعدة ============
def norm_col(c: str) -> str:
    return re.sub(r"\s+", " ", str(c).strip()).lower()

def smart_parse_datetime(series: pd.Series) -> pd.Series:
    """محلل ذكي لتواريخ Excel والنصوص"""
    if series is None:
        return pd.Series([], dtype="datetime64[ns]")
    s = series.copy()

    # تنظيف النصوص
    def clean(x):
        if pd.isna(x):
            return np.nan
        x = str(x).strip()
        if x == "" or x.lower() in {"none", "nan", "null", "-", "—", "0"}:
            return np.nan
        return x
    s = s.map(clean)

    parsed = pd.to_datetime(s, errors="coerce", dayfirst=True, infer_datetime_format=True)
    # تحويل سيريال Excel
    need_excel = parsed.isna()
    if need_excel.any():
        as_num = pd.to_numeric(s.where(need_excel), errors="coerce")
        excel_mask = as_num.notna()
        if excel_mask.any():
            excel_dates = pd.to_datetime(as_num[excel_mask], unit="d", origin="1899-12-30", errors="coerce")
            parsed.loc[excel_mask] = excel_dates
    return parsed

def pick_col(df, candidates):
    nm = {norm_col(c): c for c in df.columns}
    for c in candidates:
        if norm_col(c) in nm:
            return nm[norm_col(c)]
    for c in df.columns:
        for cand in candidates:
            if norm_col(cand) in norm_col(c):
                return c
    return None

def to_excel_download(df: pd.DataFrame) -> bytes:
    """تحويل DataFrame إلى ملف Excel جاهز للتنزيل"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Results")
    return output.getvalue()

# ============ رفع الملفات ============
st.title("📊 نظام تتبع العدادات غير المتصلة - Premise Tracker")

with st.sidebar:
    st.header("📁 ملفات الإدخال")
    disconnected_file = st.file_uploader("ملف العدادات غير المتصلة", type=["xlsx", "xls"])
    insp_file = st.file_uploader("ملف مهام الفحص (اختياري)", type=["xlsx", "xls"])
    maint_file = st.file_uploader("ملف مهام الصيانة (اختياري)", type=["xlsx", "xls"])
    st.markdown("---")
    start_btn = st.button("🚀 ابدأ التحليل")

if not start_btn or not disconnected_file:
    st.info("⬆️ ارفع الملفات المطلوبة ثم اضغط **ابدأ التحليل** لبدء المعالجة.")
    st.stop()

# ============ قراءة ملف العدادات غير المتصلة ============
st.subheader("📘 قراءة ملف العدادات غير المتصلة")
dis_df = pd.read_excel(disconnected_file)
col_premise = pick_col(dis_df, ["Utility Site Id", "Premise", "رقم المكان"])
col_last = pick_col(dis_df, ["Last Daily", "Last Communication", "آخر قراءة", "آخر اتصال"])
if not col_premise:
    st.error("لم يتم العثور على عمود رقم المكان (Utility Site Id) في الملف.")
    st.stop()

dis_df["_KEY_PREMISE"] = dis_df[col_premise].astype(str).str.strip()
if col_last:
    dis_df[col_last] = smart_parse_datetime(dis_df[col_last])
    st.success(f"تم تحويل عمود آخر اتصال ({col_last}) إلى تاريخ بنجاح.")
else:
    st.warning("⚠️ لم يُعثر على عمود 'آخر اتصال' في الملف.")

# ============ قراءة ملفات الفحص والصيانة ============
def read_task_file(f, kind):
    if f is None:
        return pd.DataFrame()
    df = pd.read_excel(f)
    col_pre = pick_col(df, ["Premise", "Utility Site Id", "رقم المكان"])
    col_reg = pick_col(df, ["Task Registration Date Time", "Request Registration Date Time", "تاريخ التسجيل"])
    col_close = pick_col(df, ["Task Closed Time", "Task Completed Time", "تاريخ الإقفال"])
    col_status = pick_col(df, ["Task Status", "Request Status", "الحالة"])
    col_result = pick_col(df, ["Technician's Result", "Final Result", "النتيجة النهائية"])
    df["_KEY_PREMISE"] = df[col_pre].astype(str).str.strip()
    for c in [col_reg, col_close]:
        if c in df.columns:
            df[c] = smart_parse_datetime(df[c])
    df["نوع المهمة"] = kind
    return df

insp_df = read_task_file(insp_file, "فحص")
maint_df = read_task_file(maint_file, "صيانة")

# ============ الدمج والتحليل ============
st.subheader("🔗 تحليل ودمج البيانات")

summary = dis_df.copy()
summary = summary.rename(columns={col_last: "LastDaily"})
summary = summary[["_KEY_PREMISE", "LastDaily"] + [c for c in dis_df.columns if c not in ["_KEY_PREMISE", "LastDaily"]]]

def summarize_tasks(df, prefix):
    if df.empty:
        return pd.DataFrame(columns=["_KEY_PREMISE"])
    reg_col = pick_col(df, ["Task Registration Date Time", "Request Registration Date Time", "تاريخ التسجيل"])
    close_col = pick_col(df, ["Task Closed Time", "Task Completed Time", "تاريخ الإقفال"])
    status_col = pick_col(df, ["Task Status", "Request Status", "الحالة"])
    result_col = pick_col(df, ["Technician's Result", "Final Result", "النتيجة النهائية"])
    df["_open"] = df[close_col].isna() if close_col in df.columns else df[status_col].astype(str).str.lower().ne("closed")
    out = df.groupby("_KEY_PREMISE").agg(
        total_tasks=(status_col, "count"),
        open_tasks=("_open", "sum"),
        last_status=(status_col, "last"),
        last_result=(result_col, "last"),
        last_date=(close_col, "max")
    ).reset_index()
    out = out.add_prefix(prefix)
    out = out.rename(columns={f"{prefix}_KEY_PREMISE": "_KEY_PREMISE"})
    return out

insp_summary = summarize_tasks(insp_df, "insp")
maint_summary = summarize_tasks(maint_df, "maint")

summary = summary.merge(insp_summary, on="_KEY_PREMISE", how="left")
summary = summary.merge(maint_summary, on="_KEY_PREMISE", how="left")

# ============ تحديد الإجراء التالي ============
def next_action(row):
    if row.get("maintopen_tasks", 0) > 0:
        return "تسريع صيانة مفتوحة"
    elif row.get("inspopen_tasks", 0) > 0:
        return "متابعة فحص مفتوح"
    else:
        return "يفتح فحص جديد"

summary["Next Action"] = summary.apply(next_action, axis=1)

# ============ عرض النتائج ============
st.subheader("📈 مؤشرات عامة")
c1, c2, c3 = st.columns(3)
c1.metric("عدد العدادات غير المتصلة", f"{len(summary):,}")
c2.metric("مهام فحص مفتوحة", int(summary["inspopen_tasks"].fillna(0).sum()))
c3.metric("مهام صيانة مفتوحة", int(summary["maintopen_tasks"].fillna(0).sum()))

st.subheader("📋 الجدول الموحد")
st.dataframe(summary, use_container_width=True)

# ============ التنزيل ============
excel_data = to_excel_download(summary)
st.download_button(
    label="⬇️ تنزيل النتائج بصيغة Excel",
    data=excel_data,
    file_name=f"premise_tracker_results_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
