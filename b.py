# -*- coding: utf-8 -*-
# mms_anomaly_crosscheck.py
import re
import math
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="MMS | Cross-Check Anomalies vs Open Tasks", layout="wide")

# -------------------------------
# Helpers
# -------------------------------
def norm_col(c: str) -> str:
    return re.sub(r"\s+", " ", str(c).strip()).lower()

# صيانة: احتمالات أسماء الأعمدة
CLOSED_CANDS = ["Task Closed Time", "Task Completed Time", "وقت إقفال المهمة", "تاريخ إقفال المهمة"]
REG_CANDS    = ["Task Registration Date Time", "Request Registration Date Time", "تاريخ تسجيل المهمة", "تاريخ تسجيل الطلب"]
STATUS_CANDS = ["Task Status", "Request Status", "حالة المهمة", "حالة الطلب"]
METER_CANDS  = ["Meter No", "Meter Number", "رقم العداد"]
VIP_CANDS    = ["Subscription VIP", "VIP", "نوع المشترك", "Account Type", "تصنيف المشترك"]

BUCKET_HINTS = {
    "استبدال": ["استبدال"],
    "تحسين": ["تحسين", "استخراج", "تحسين واستخراج"],
    "صيانة": ["صيانة"],
    "كشف": ["كشف", "معاينة", "كشف ومعاينة"]
}

def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    nm = {norm_col(c): c for c in df.columns}
    for c in candidates:
        if norm_col(c) in nm:
            return nm[norm_col(c)]
    # تطابق جزئي احتياطي
    for c in df.columns:
        for cand in candidates:
            if norm_col(cand) in norm_col(c):
                return c
    return None

def infer_bucket_from_name(name: str) -> str:
    n = (name or "").lower()
    for b, kws in BUCKET_HINTS.items():
        for kw in kws:
            if kw in n:
                return b
    return "غير محدد"

# شذوذ: من ملفك (أهم عمود Meter Number + أعمدة الشذوذ)
ANOM_METER_CANDS = ["Meter Number", "Meter No", "Meter", "رقم العداد"]
ANOM_SEV_CANDS   = ["final_label", "Severity", "خطورة", "أولوية"]
ANOM_LOSS_CANDS  = ["confirmed_loss", "suspected_loss", "bypass_score", "Loss kWh", "Estimated Loss", "فاقد"]

# -------------------------------
# Readers
# -------------------------------
def read_maintenance(file):
    xl = pd.ExcelFile(file)
    sheet = xl.sheet_names[0]
    df = xl.parse(sheet)

    closed_col = pick_col(df, CLOSED_CANDS)
    reg_col    = pick_col(df, REG_CANDS)
    status_col = pick_col(df, STATUS_CANDS)
    meter_col  = pick_col(df, METER_CANDS)
    vip_col    = pick_col(df, VIP_CANDS)

    if closed_col is not None:
        df[closed_col] = pd.to_datetime(df[closed_col], errors="coerce")
        df["Closed Date"] = df[closed_col].dt.date
    else:
        df["Closed Date"] = pd.NaT

    if reg_col is not None:
        df[reg_col] = pd.to_datetime(df[reg_col], errors="coerce")

    if "Bucket" not in df.columns:
        df["Bucket"] = infer_bucket_from_name(getattr(file, "name", ""))

    # تعريف المفتوح/المغلق
    is_closed = df[closed_col].notna() if closed_col else df["Closed Date"].notna()
    df["__is_open__"] = ~is_closed

    # عمر المهمة المفتوحة بالأيام
    if reg_col is not None:
        df["Age Days"] = (pd.to_datetime(datetime.now()) - df[reg_col]).dt.days
    else:
        df["Age Days"] = np.nan

    meta = dict(
        sheet=sheet, closed_col=closed_col, reg_col=reg_col, status_col=status_col,
        meter_col=meter_col, vip_col=vip_col, columns=list(df.columns)
    )
    return df, meta

def read_anomalies(file):
    xl = pd.ExcelFile(file)
    sheet = xl.sheet_names[0]
    df = xl.parse(sheet)

    a_meter = pick_col(df, ANOM_METER_CANDS)
    a_sev   = pick_col(df, ANOM_SEV_CANDS)
    # نجمع مؤشرات الفاقد إن وُجدت
    loss_cols = [c for c in df.columns if norm_col(c) in [norm_col(x) for x in ANOM_LOSS_CANDS]]

    meta = dict(sheet=sheet, a_meter=a_meter, a_sev=a_sev, loss_cols=loss_cols, columns=list(df.columns))
    return df, meta

# -------------------------------
# UI
# -------------------------------
st.title("🔧 MMS — Cross-Check Anomalies vs Open Maintenance")
st.caption("تجنّب تكرار البلاغات: نطابق ملف الحالات الشاذة مع المهام المفتوحة لنفس العدادات ونعطيك قائمتين (تسريع / فتح جديد).")

with st.sidebar:
    st.header("📁 ملفات الصيانة")
    maint_files = st.file_uploader("ارفع ملفات الصيانة (يمكن عدة ملفات)", type=["xlsx","xls"], accept_multiple_files=True)
    st.header("📁 ملف/ملفات الحالات الشاذة")
    anom_files  = st.file_uploader("ارفع ملف/ملفات الشذوذ", type=["xlsx","xls"], accept_multiple_files=True)
    st.markdown("---")
    sla_days = st.number_input("حد التأخير (SLA) بالأيام", min_value=1, max_value=60, value=3)

if not maint_files:
    st.info("✨ ارفع ملفات الصيانة أولاً.")
    st.stop()
if not anom_files:
    st.info("✨ ارفع ملف الحالات الشاذة (detailed_asdct...).")
    st.stop()

# قراءة الصيانة وضمّها
m_dfs, m_metas = [], []
for f in maint_files:
    df, meta = read_maintenance(f)
    m_dfs.append(df)
    m_metas.append(meta)
maintenance = pd.concat(m_dfs, ignore_index=True, sort=False)

# أعمدة معرّفة
closed_col = next((m["closed_col"] for m in m_metas if m["closed_col"]), None)
reg_col    = next((m["reg_col"]    for m in m_metas if m["reg_col"]), None)
meter_col  = next((m["meter_col"]  for m in m_metas if m["meter_col"]), None)
vip_col    = next((m["vip_col"]    for m in m_metas if m["vip_col"]), None)

# فلترة المفتوح
open_tasks = maintenance[maintenance["__is_open__"]].copy()
open_tasks["Is Late"] = open_tasks["Age Days"] > sla_days

# قراءة الشذوذ وضمّها
a_dfs, a_metas = [], []
for f in anom_files:
    df, meta = read_anomalies(f)
    df["Anomaly Source"] = getattr(f, "name", "anomaly.xlsx")
    a_dfs.append(df)
    a_metas.append(meta)
anomalies = pd.concat(a_dfs, ignore_index=True, sort=False)

a_meter = next((m["a_meter"] for m in a_metas if m["a_meter"]), None)
a_sev   = next((m["a_sev"] for m in a_metas if m["a_sev"]), None)
# اجمع أعمدة الفاقد المتاحة
loss_cols = []
for m in a_metas:
    loss_cols.extend([c for c in m["loss_cols"]])
loss_cols = list(dict.fromkeys(loss_cols))

# -------------------------------
# المؤشرات العامة
# -------------------------------
st.markdown("## 1) نظرة عامة")
c1,c2,c3,c4 = st.columns(4)
c1.metric("المهام المفتوحة", f"{len(open_tasks):,}")
c2.metric("المهام المتأخرة", f"{open_tasks['Is Late'].sum():,}")
c3.metric("حالات الشذوذ", f"{len(anomalies):,}")
c4.metric("المفتاح للمطابقة", a_meter or "Meter Number (مفقود)")

# -------------------------------
# المطابقة (Meter Number ↔ Meter Number)
# -------------------------------
st.markdown("## 2) المطابقة")
if not meter_col or not a_meter:
    st.error("لا يمكن المطابقة: لم يتم العثور على عمود رقم العداد في أحد الملفات. تأكد من وجود Meter Number في ملف الشذوذ و Meter No/Meter Number في ملفات الصيانة.")
    st.stop()

left = anomalies.copy()
right = open_tasks.copy()

left[a_meter]  = left[a_meter].astype(str).str.strip()
right[meter_col] = right[meter_col].astype(str).str.strip()

matches = left.merge(
    right,
    how="inner",
    left_on=a_meter,
    right_on=meter_col,
    suffixes=("_ANOM", "_TASK")
)

# -------------------------------
# نتائج المطابقة
# -------------------------------
total_anom = len(anomalies)
matched_meters = matches[[a_meter]].drop_duplicates() if len(matches) else pd.DataFrame(columns=[a_meter])
num_matched = len(matched_meters)

# الحالات التي لا يوجد لها مهام مفتوحة (افتح بلاغ جديد)
unmatched = anomalies.copy()
if num_matched:
    unmatched = anomalies.merge(matched_meters, how="left", on=a_meter, indicator=True)
    unmatched = unmatched[unmatched["_merge"]=="left_only"].drop(columns=["_merge"])

st.markdown("### مؤشرات المطابقة")
d1,d2,d3 = st.columns(3)
d1.metric("وجدنا لها مهام مفتوحة (تسريع)", f"{num_matched:,}")
d2.metric("تحتاج فتح مهمة جديدة", f"{len(unmatched):,}")
late_rate = (open_tasks["Is Late"].mean()*100) if len(open_tasks) else 0
d3.metric("نسبة تأخير المهام المفتوحة", f"{late_rate:,.1f}%")

# أولوية التسريع حسب VIP إن متاح
if vip_col and vip_col in matches.columns:
    st.markdown("#### أولوية التسريع حسب نوع المشترك (VIP)")
    vip_escalate = (matches.groupby(vip_col).size().reset_index(name="OpenTasksForAnomalies")
                    .sort_values("OpenTasksForAnomalies", ascending=False))
    st.dataframe(vip_escalate, use_container_width=True)

# جداول تفصيلية
st.markdown("---")
st.markdown("### ✅ حالات شاذة لها مهام صيانة مفتوحة (سرّع التنفيذ)")
if len(matches):
    cols_show = [a_meter, a_sev, "Anomaly Source"] + loss_cols
    cols_show = [c for c in cols_show if c and c in matches.columns]
    # أعمدة مختارة من الصيانة
    for c in ["Bucket", "Task Code", "Request id", "Task Status", reg_col, closed_col, meter_col, vip_col, "Age Days", "Is Late"]:
        if c and c in matches.columns:
            cols_show.append(c)
    cols_show = list(dict.fromkeys(cols_show))
    st.dataframe(matches[cols_show].sort_values(by=["Is Late","Age Days"] if "Age Days" in cols_show else cols_show[0], ascending=[False, False] if "Age Days" in cols_show else True), use_container_width=True)
else:
    st.info("لا توجد حالات شاذة لها مهام مفتوحة لنفس العداد وفق البيانات الحالية.")

st.markdown("### 🆕 حالات شاذة بلا مهام مطابقة (افتح بلاغ فحص)")
st.dataframe(unmatched, use_container_width=True)

# تنزيل
st.markdown("---")
c_dl1, c_dl2 = st.columns(2)
with c_dl1:
    if len(matches):
        st.download_button(
            "⬇️ تنزيل قائمة التسريع (CSV)",
            data=matches.to_csv(index=False).encode("utf-8-sig"),
            file_name="anomalies_with_open_tasks_to_escalate.csv",
            mime="text/csv"
        )
with c_dl2:
    st.download_button(
        "⬇️ تنزيل الحالات التي تحتاج فتح مهمة (CSV)",
        data=unmatched.to_csv(index=False).encode("utf-8-sig"),
        file_name="anomalies_need_new_tasks.csv",
        mime="text/csv"
    )

st.markdown("---")
st.caption("MMS — Cross-Check: يطابق ملف الشذوذ (Meter Number) مع المهام المفتوحة لتجنب التكرار وتسريع المعالجة عند وجود فاقد.")
