# -*- coding: utf-8 -*-
# meter_maintenance_ops_dashboard.py
import re
import io
import math
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime, date

# -------------------------------
# إعداد الصفحة
# -------------------------------
st.set_page_config(page_title="لوحة صيانة العدادات - إنتاجية ومتابعة", layout="wide")

st.title("📊 لوحة صيانة العدادات")
st.caption("تحليل الإنتاجية اليومية ومتابعة المهام المتبقية (تأخير، تكرار العدادات، نوع المشتركين).")

# -------------------------------
# دوال مساعدة
# -------------------------------
def norm_col(c: str) -> str:
    """تطبيع اسم العمود للمقارنة (لا يُستخدم للعرض)."""
    return re.sub(r"\s+", " ", str(c).strip()).lower()

# احتمالات أسماء الأعمدة الشائعة (عربي/إنجليزي)
CLOSED_CANDIDATES = [
    "Task Closed Time", "Task Completed Time",
    "وقت إقفال المهمة", "تاريخ إقفال المهمة"
]
REG_CANDIDATES = [
    "Task Registration Date Time", "Request Registration Date Time",
    "تاريخ تسجيل المهمة", "تاريخ تسجيل الطلب"
]
STATUS_CANDIDATES = [
    "Task Status", "Request Status", "حالة المهمة", "حالة الطلب"
]
METER_CANDIDATES = [
    "Meter No", "Meter Number", "رقم العداد"
]
TECH_CANDIDATES = [
    "Technician Name", "Field Team Name", "اسم الفني", "اسم فريق الميدان"
]
REQ_TYPE_CANDIDATES = [
    "Request Type", "نوع الطلب"
]
REQ_CHANNEL_CANDIDATES = [
    "Request Channel", "قناة الطلب", "مصدر الطلب"
]
VIP_CANDIDATES = [
    "Subscription VIP", "VIP", "نوع المشترك", "Account Type", "تصنيف المشترك"
]
BUCKET_HINTS = {
    "استبدال": ["استبدال"],
    "تحسين": ["تحسين", "استخراج", "تحسين واستخراج"],
    "صيانة": ["صيانة"],
    "كشف": ["كشف", "معاينة", "كشف ومعاينة"]
}

def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    norm_map = {norm_col(c): c for c in df.columns}
    for c in candidates:
        nc = norm_col(c)
        if nc in norm_map:
            return norm_map[nc]
    # تطابق جزئي إن لزم (احتياط)
    for c in df.columns:
        for cand in candidates:
            if norm_col(cand) in norm_col(c):
                return c
    return None

def infer_bucket_from_filename(name: str) -> str:
    n = (name or "").lower()
    for bucket, kws in BUCKET_HINTS.items():
        for kw in kws:
            if kw in n:
                return bucket
    return "غير محدد"

def read_one_excel(uploaded_file) -> tuple[pd.DataFrame, dict]:
    xl = pd.ExcelFile(uploaded_file)
    sheet = xl.sheet_names[0]
    df = xl.parse(sheet)

    # أعمدة أساسية
    closed_col = pick_col(df, CLOSED_CANDIDATES)
    reg_col    = pick_col(df, REG_CANDIDATES)
    status_col = pick_col(df, STATUS_CANDIDATES)
    meter_col  = pick_col(df, METER_CANDIDATES)
    tech_col   = pick_col(df, TECH_CANDIDATES)
    type_col   = pick_col(df, REQ_TYPE_CANDIDATES)
    chan_col   = pick_col(df, REQ_CHANNEL_CANDIDATES)
    vip_col    = pick_col(df, VIP_CANDIDATES)

    # تواريخ
    if closed_col is not None:
        df[closed_col] = pd.to_datetime(df[closed_col], errors="coerce")
        df["Closed Date"] = df[closed_col].dt.date
    else:
        df["Closed Date"] = pd.NaT

    if reg_col is not None:
        df[reg_col] = pd.to_datetime(df[reg_col], errors="coerce")

    # السلة
    if "Bucket" not in df.columns:
        df["Bucket"] = infer_bucket_from_filename(getattr(uploaded_file, "name", ""))

    # مدة الإنجاز (للمنجزة)
    if (closed_col is not None) and (reg_col is not None):
        df["Duration Hours"] = (df[closed_col] - df[reg_col]).dt.total_seconds() / 3600.0

    meta = dict(
        sheet=sheet, closed_col=closed_col, reg_col=reg_col, status_col=status_col,
        meter_col=meter_col, tech_col=tech_col, type_col=type_col,
        chan_col=chan_col, vip_col=vip_col, columns=list(df.columns)
    )
    return df, meta

def compare_columns(files_meta: dict[str, dict]) -> pd.DataFrame:
    all_cols = []
    for name, info in files_meta.items():
        all_cols.extend(info["columns"])
    ordered = list(dict.fromkeys(all_cols))  # الحفاظ على الترتيب وإزالة التكرار
    out = pd.DataFrame({"Column": ordered})
    for name, info in files_meta.items():
        cols = set(info["columns"])
        out[name] = out["Column"].apply(lambda c: c in cols)
    return out

def today_date() -> date:
    return datetime.now().date()

# -------------------------------
# الشريط الجانبي: رفع الملفات + إعدادات
# -------------------------------
with st.sidebar:
    st.header("📁 رفع الملفات")
    files = st.file_uploader(
        "ارفع ملفات Excel (يمكن عدة ملفات)",
        type=["xlsx", "xls"], accept_multiple_files=True
    )
    st.markdown("---")
    st.header("⚙️ إعدادات")
    sla_days = st.number_input("حد التأخير (SLA) بالأيام", min_value=1, max_value=60, value=3)
    overdue_buckets = st.multiselect(
        "تقسيم فترات التأخير (للمهام المفتوحة)",
        ["≤1 يوم", "2-3 أيام", "4-7 أيام", "8-30 يومًا", ">30 يوم"],
        default=["2-3 أيام", "4-7 أيام", "8-30 يومًا", ">30 يوم"]
    )
    st.markdown("---")
    st.caption("تُستنتج السلة من اسم الملف إن لم يوجد عمود Bucket.")

# -------------------------------
# تحميل ومعالجة البيانات
# -------------------------------
dfs, metas = [], {}
if files:
    for f in files:
        df, meta = read_one_excel(f)
        dfs.append(df)
        metas[f.name] = meta

if not files:
    st.info("✨ ارفع ملفات السلال (استبدال/صيانة/تحسين/كشف) للبدء.")
    st.stop()

data = pd.concat(dfs, ignore_index=True, sort=False)

# عرض مقارنة الأعمدة
with st.expander("🔎 مقارنة الأعمدة بين الملفات"):
    comp = compare_columns(metas)
    st.dataframe(comp, use_container_width=True)

# التعرف على أسماء الأعمدة المختارة
any_meta = next(iter(metas.values()))
closed_col = next((m["closed_col"] for m in metas.values() if m["closed_col"]), None)
reg_col    = next((m["reg_col"]    for m in metas.values() if m["reg_col"]), None)
status_col = next((m["status_col"] for m in metas.values() if m["status_col"]), None)
meter_col  = next((m["meter_col"]  for m in metas.values() if m["meter_col"]), None)
tech_col   = next((m["tech_col"]   for m in metas.values() if m["tech_col"]), None)
type_col   = next((m["type_col"]   for m in metas.values() if m["type_col"]), None)
chan_col   = next((m["chan_col"]   for m in metas.values() if m["chan_col"]), None)
vip_col    = next((m["vip_col"]    for m in metas.values() if m["vip_col"]), None)

# -------------------------------
# فلاتر عامة
# -------------------------------
st.markdown("## 🎯 الفلاتر العامة")
colf1, colf2, colf3 = st.columns([1,1,2])

# فلتر التاريخ (للإنتاجية): حسب Closed Date
if "Closed Date" in data.columns and data["Closed Date"].notna().any():
    min_d = pd.to_datetime(data["Closed Date"]).min().date()
    max_d = pd.to_datetime(data["Closed Date"]).max().date()
    with colf1:
        date_mode = st.radio("نطاق التاريخ (الإنتاجية)", ["يوم واحد", "نطاق أيام"], horizontal=True)
    with colf2:
        if date_mode == "يوم واحد":
            sel_day = st.date_input("اختر اليوم", value=max_d, min_value=min_d, max_value=max_d)
            prod_mask = (pd.to_datetime(data["Closed Date"]).dt.date == sel_day)
        else:
            sel_range = st.date_input("اختر النطاق", value=(min_d, max_d), min_value=min_d, max_value=max_d)
            d1, d2 = sel_range
            prod_mask = (pd.to_datetime(data["Closed Date"]).dt.date >= d1) & (pd.to_datetime(data["Closed Date"]).dt.date <= d2)
else:
    st.warning("لم يتم العثور على Closed Date صالح. سيتم عرض الإنتاجية بدون فلتر زمني.")
    prod_mask = np.array([True] * len(data))

with colf3:
    buckets = sorted(data["Bucket"].dropna().astype(str).unique().tolist()) if "Bucket" in data.columns else []
    sel_buckets = st.multiselect("تصفية حسب السلة", buckets, default=buckets)

mask_bucket = data["Bucket"].astype(str).isin(sel_buckets) if "Bucket" in data.columns and sel_buckets else np.array([True]*len(data))
data_filt_prod = data[prod_mask & mask_bucket].copy()

# -------------------------------
# 1) الإنتاجية اليومية (حسب السلال)
# -------------------------------
st.markdown("## 📈 أولًا: الإنتاجية اليومية لكل السلال")

# KPIs
total_tasks = len(data_filt_prod)
closed_tasks = data_filt_prod["Closed Date"].notna().sum() if "Closed Date" in data_filt_prod.columns else 0
closure_rate = (closed_tasks / total_tasks * 100) if total_tasks else 0
avg_duration = data_filt_prod["Duration Hours"].mean() if "Duration Hours" in data_filt_prod.columns else np.nan

k1, k2, k3, k4 = st.columns(4)
k1.metric("إجمالي المهام (ضمن النطاق)", f"{total_tasks:,}")
k2.metric("المهام المقفلة", f"{closed_tasks:,}")
k3.metric("نسبة الإقفال", f"{closure_rate:,.1f}%")
k4.metric("متوسط مدة الإنجاز (ساعة)", "-" if (math.isnan(avg_duration) if isinstance(avg_duration, float) else pd.isna(avg_duration)) else f"{avg_duration:,.1f}")

# إنتاجية حسب اليوم/السلة
if "Closed Date" in data_filt_prod.columns:
    prod_group = data_filt_prod.groupby(["Closed Date", "Bucket"]).size().reset_index(name="Count")
    st.markdown("#### عدد المهام المقفلة يوميًا حسب السلة")
    st.bar_chart(prod_group.pivot(index="Closed Date", columns="Bucket", values="Count").fillna(0))

# توزيع حسب السلة
if "Bucket" in data_filt_prod.columns:
    st.markdown("#### توزيع المهام حسب السلة (ضمن النطاق)")
    st.bar_chart(data_filt_prod.groupby("Bucket").size())

# أفضل فنيين (اختياري)
if tech_col and tech_col in data_filt_prod.columns:
    top_tech = (
        data_filt_prod.groupby(tech_col).size().reset_index(name="Count")
        .sort_values("Count", ascending=False).head(10)
    )
    st.markdown("#### أعلى 10 فنيين (عدد مهام ضمن النطاق)")
    st.bar_chart(top_tech.set_index(tech_col))

st.markdown("---")

# -------------------------------
# 2) متابعة المهام المتبقية وتحليلها
# -------------------------------
st.markdown("## 🧭 ثانيًا: متابعة المهام المتبقية وتحليلها")

# تعريف "مفتوحة" و "متأخرة"
is_closed = data[closed_col].notna() if closed_col else data["Closed Date"].notna()
is_open = ~is_closed

# عمر المهمة المفتوحة بالأيام (اليوم - تسجيل)
if reg_col and reg_col in data.columns:
    data["Age Days"] = (pd.to_datetime(datetime.now()) - data[reg_col]).dt.days
else:
    data["Age Days"] = np.nan

open_df = data[is_open].copy()
open_df["Is Late"] = False
if "Age Days" in open_df.columns:
    open_df["Is Late"] = open_df["Age Days"] > sla_days

# KPIs للمهام المتبقية
o1, o2, o3, o4 = st.columns(4)
o1.metric("المهام المفتوحة", f"{len(open_df):,}")
o2.metric("المهام المتأخرة (>{} يوم)".format(sla_days), f"{open_df['Is Late'].sum():,}")
if "Bucket" in open_df.columns:
    late_rate = (open_df["Is Late"].mean() * 100) if len(open_df) else 0
else:
    late_rate = 0
o3.metric("نسبة التأخير", f"{late_rate:,.1f}%")
if reg_col and reg_col in open_df.columns:
    avg_age = open_df["Age Days"].mean()
else:
    avg_age = np.nan
o4.metric("متوسط عمر المهمة (يوم)", "-" if (math.isnan(avg_age) if isinstance(avg_age, float) else pd.isna(avg_age)) else f"{avg_age:,.1f}")

# تقسيم فترات التأخير للمهام المفتوحة
def overdue_bucket(d):
    if pd.isna(d):
        return "غير محدد"
    if d <= 1: return "≤1 يوم"
    if d <= 3: return "2-3 أيام"
    if d <= 7: return "4-7 أيام"
    if d <= 30: return "8-30 يومًا"
    return ">30 يوم"

if "Age Days" in open_df.columns:
    open_df["Delay Bucket"] = open_df["Age Days"].apply(overdue_bucket)
    st.markdown("#### توزيع فترات التأخير للمهام المفتوحة")
    # تطبيق اختيار المستخدم
    if overdue_buckets:
        filt_over = open_df["Delay Bucket"].isin(overdue_buckets)
        open_for_buckets = open_df[filt_over]
    else:
        open_for_buckets = open_df
    st.bar_chart(open_for_buckets.groupby("Delay Bucket").size())

# تحليل حسب السلة/الفني/القناة/نوع الطلب
cols = st.columns(3)
if "Bucket" in open_df.columns:
    with cols[0]:
        st.markdown("#### المهام المفتوحة حسب السلة")
        st.bar_chart(open_df.groupby("Bucket").size())
if tech_col and tech_col in open_df.columns:
    with cols[1]:
        st.markdown("#### المهام المفتوحة حسب الفني (Top 10)")
        t = (open_df.groupby(tech_col).size().reset_index(name="Count")
             .sort_values("Count", ascending=False).head(10))
        st.bar_chart(t.set_index(tech_col))
if chan_col and chan_col in open_df.columns:
    with cols[2]:
        st.markdown("#### المهام المفتوحة حسب مصدر الطلب")
        st.bar_chart(open_df.groupby(chan_col).size())

# العدادات ذات أكثر من مهمة (أيًا كانت حالتها)
if meter_col and meter_col in data.columns:
    multi_meter = (data.groupby(meter_col).size()
                   .reset_index(name="TasksCount")
                   .sort_values("TasksCount", ascending=False))
    st.markdown("#### العدادات ذات المهام المتعددة")
    st.dataframe(multi_meter[multi_meter["TasksCount"] > 1].head(200), use_container_width=True)

# تعدد السلال لنفس العداد
if meter_col and "Bucket" in data.columns:
    meter_bucket_multi = (data.groupby([meter_col, "Bucket"]).size()
                          .reset_index(name="Count"))
    st.markdown("#### توزيع المهام لكل عداد عبر السلال")
    st.dataframe(meter_bucket_multi.sort_values(["Count"], ascending=False).head(300), use_container_width=True)

# تحليل نوع الطلب والمتبقي
if type_col and type_col in open_df.columns:
    st.markdown("#### أنواع الطلب للمهام المفتوحة")
    st.bar_chart(open_df.groupby(type_col).size())

# VIP / نوع المشترك
if vip_col and vip_col in data.columns:
    st.markdown("#### متابعة VIP / نوع المشترك")
    vip_open = (open_df.groupby(vip_col).size().reset_index(name="OpenCount")
                .sort_values("OpenCount", ascending=False))
    st.dataframe(vip_open, use_container_width=True)
    if "Is Late" in open_df.columns:
        vip_late = (open_df.groupby(vip_col)["Is Late"].mean().reset_index())
        vip_late["Late %"] = (vip_late["Is Late"] * 100).round(1)
        vip_late = vip_late.drop(columns=["Is Late"])
        st.markdown("**نسبة التأخير حسب نوع المشترك**")
        st.dataframe(vip_late, use_container_width=True)

# جدول تفصيلي للمهام المفتوحة (الأهم تشغيليًا)
st.markdown("### 📋 جدول المهام المفتوحة (الأولوية للتعامل)")
show_cols = []
for c in ["Bucket", "Closed Date", "Duration Hours", "Age Days", "Delay Bucket",
          "Task Code", "Request id", "Request Type", "Task Status",
          "Task Registration Date Time", "Request Registration Date Time",
          "Task Closed Time", "Task Completed Time"]:
    if c in open_df.columns:
        show_cols.append(c)
# تأكد من إضافة حقول مهمة عند توفرها
for c in [meter_col, tech_col, chan_col, vip_col]:
    if c and c in open_df.columns:
        show_cols.append(c)
show_cols = list(dict.fromkeys(show_cols))  # إزالة تكرار
st.dataframe(open_df.sort_values("Age Days", ascending=False)[show_cols] if show_cols else open_df, use_container_width=True)

# تنزيل بيانات
st.markdown("---")
col_dl1, col_dl2 = st.columns(2)
with col_dl1:
    csv_prod = data_filt_prod.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ تنزيل بيانات الإنتاجية (CSV)", data=csv_prod, file_name="productivity_filtered.csv", mime="text/csv")
with col_dl2:
    csv_open = open_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ تنزيل قائمة المهام المفتوحة (CSV)", data=csv_open, file_name="open_tasks.csv", mime="text/csv")

st.markdown("---")
st.caption("© لوحة صيانة العدادات — إنتاجية يومية + متابعة المتبقي (تأخير، تكرار العدادات، VIP).")
