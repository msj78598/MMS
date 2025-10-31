# -*- coding: utf-8 -*-
# mms_inspection_impact.py — إبراز جهود الفحص وربطها بالصيانة والاتصال (Premise Key)

import re
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Inspection Impact — Premise Tracker", layout="wide")

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
    for c in df.columns:
        for cand in candidates:
            if norm_col(cand) in norm_col(c):
                return c
    return None

def smart_parse_datetime(series: pd.Series, excel_origin: str = "1899-12-30") -> pd.Series:
    """تحويل مختلط (نصي/سيريال Excel) إلى datetime؛ يدعم dayfirst و1900/1904."""
    if series is None:
        return pd.Series([], dtype="datetime64[ns]")
    s = series.copy()

    def clean(x):
        if pd.isna(x): return np.nan
        x = str(x).strip()
        if x == "" or x.lower() in {"none", "nan", "null", "-", "—", "0"}: return np.nan
        if re.fullmatch(r"0{2,}[-/:]0{2,}[-/:]0{2,}.*", x): return np.nan
        return x

    s = s.map(clean)
    parsed = pd.to_datetime(s, errors="coerce", dayfirst=True, infer_datetime_format=True)

    # محاولة ثانية
    need2 = parsed.isna()
    if need2.any():
        parsed.loc[need2] = pd.to_datetime(s[need2], errors="coerce", dayfirst=False, infer_datetime_format=True)

    # Excel serial
    need_excel = parsed.isna()
    if need_excel.any():
        as_num = pd.to_numeric(s.where(need_excel), errors="coerce")
        mask = as_num.notna()
        if mask.any():
            parsed.loc[mask] = pd.to_datetime(as_num[mask], unit="d", origin=excel_origin, errors="coerce")
    return parsed

def to_excel_download(df: pd.DataFrame) -> bytes:
    """Excel via xlsxwriter; fallback to openpyxl."""
    bio = BytesIO()
    try:
        with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Results")
    except ModuleNotFoundError:
        with pd.ExcelWriter(bio, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Results")
    return bio.getvalue()

# ================= UI (Uploads & Settings) =================
st.title("📊 إبراز جهود الفحص — Inspection Impact (Premise Key)")

with st.sidebar:
    st.header("📁 ملفات الإدخال")
    dis_file   = st.file_uploader("ملف العدادات غير المتصلة", type=["xlsx","xls"])
    insp_files = st.file_uploader("ملفات مهام الفحص (0..N)",  type=["xlsx","xls"], accept_multiple_files=True)
    mnt_files  = st.file_uploader("ملفات مهام الصيانة (0..N)", type=["xlsx","xls"], accept_multiple_files=True)

    st.markdown("---")
    st.header("⚙️ إعدادات")
    excel_origin = st.selectbox("Excel Origin (للتواريخ الرقمية)", ["1899-12-30", "1904-01-01"], index=0)

    default_closed_terms = """
closed, complete, completed, done, resolved,
مغلق, مغلقة, مقفلة, مقفل, منجز, منجزة, منتهية, منتهي, تمت المعالجة
""".strip()
    closed_terms_input = st.text_area("حالات تعتبر (مقفلة) — مفصولة بفواصل", value=default_closed_terms, height=90)
    CLOSED_TERMS = {w.strip().lower() for w in closed_terms_input.split(",") if w.strip()}

    st.markdown("---")
    start_btn = st.button("🚀 ابدأ التحليل")

if not start_btn or not dis_file:
    st.info("⬆️ ارفع الملفات ثم اضغط **ابدأ التحليل**.")
    st.stop()

# ================= Read Disconnected =================
dis_df = pd.read_excel(dis_file)

PREMISE_CANDS_DIS = ["Utility Site Id", "Premise", "رقم المكان"]
LAST_CANDS = [
    "Last Daily", "LastDaily", "Last Communication", "Last Comm",
    "Last Daily Read", "Last Daily Date",
    "آخر قراءة", "آخر اتصال", "اخر اتصال", "اخر قراءة"
]

premise_dis = pick_col(dis_df, PREMISE_CANDS_DIS)
last_col    = pick_col(dis_df, LAST_CANDS)

if not premise_dis:
    st.error("لم يتم العثور على عمود Premise/Utility Site Id في ملف غير المتصلين.")
    st.stop()

dis_df["_KEY_PREMISE"] = dis_df[premise_dis].astype(str).str.strip()

with st.expander("🔧 اختيار عمود 'آخر اتصال' يدويًا (اختياري)"):
    last_choice = st.selectbox("اختر عمود 'آخر اتصال':", options=["(اكتشاف تلقائي)"] + list(dis_df.columns), index=0)
    if last_choice != "(اكتشاف تلقائي)":
        last_col = last_choice

# LastDaily
if last_col and last_col in dis_df.columns:
    dis_df["LastDaily"] = smart_parse_datetime(dis_df[last_col], excel_origin=excel_origin)
    ok = int(dis_df["LastDaily"].notna().sum())
    st.success(f"تحويل '{last_col}': {ok}/{len(dis_df)} قيماً صالحة.")
else:
    dis_df["LastDaily"] = pd.NaT
    st.warning("⚠️ لم يُعثر على عمود 'آخر اتصال' — سيُترك فارغًا.")

# ================= Read Tasks =================
def load_task_files(files, kind_label: str) -> pd.DataFrame:
    if not files:
        return pd.DataFrame()
    frames = []
    for f in files:
        df = pd.read_excel(f)
        premise_col = pick_col(df, ["Premise", "Utility Site Id", "رقم المكان"])
        reg_col     = pick_col(df, ["Task Registration Date Time", "Request Registration Date Time",
                                    "تاريخ التسجيل", "تاريخ تسجيل المهمة", "تاريخ تسجيل الطلب"])
        close_col   = pick_col(df, ["Task Closed Time", "Task Completed Time",
                                    "تاريخ الإقفال", "وقت إقفال المهمة"])
        status_col  = pick_col(df, ["Task Status", "Request Status", "الحالة",
                                    "حالة المهمة", "حالة الطلب"])
        result_col  = pick_col(df, ["Technician's Result", "Final Result",
                                    "Final Result (Dispatcher's Result)", "النتيجة النهائية", "نتيجة الفني"])

        if not premise_col:
            st.warning(f"تجاهل '{getattr(f,'name','file')}' — لا يحتوي Premise/Utility Site Id.")
            continue

        tmp = pd.DataFrame()
        tmp["_KEY_PREMISE"] = df[premise_col].astype(str).str.strip()
        tmp["reg_time"]   = smart_parse_datetime(df[reg_col],   excel_origin=excel_origin) if (reg_col   and reg_col   in df.columns) else pd.NaT
        tmp["close_time"] = smart_parse_datetime(df[close_col], excel_origin=excel_origin) if (close_col and close_col in df.columns) else pd.NaT
        tmp["status"]     = df[status_col].astype(str) if (status_col and status_col in df.columns) else np.nan
        tmp["result"]     = df[result_col].astype(str) if (result_col and result_col in df.columns) else np.nan
        tmp["bucket"]     = kind_label
        tmp["source"]     = getattr(f, "name", kind_label)
        frames.append(tmp)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)

insp_df  = load_task_files(insp_files,  "فحص")
mnt_df   = load_task_files(mnt_files,   "صيانة")

# ================= Summaries (latest/first) =================
def latest_by_key(df: pd.DataFrame, date_col: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["_KEY_PREMISE"])
    d = df.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    # أحدث سجل بناءً على التاريخ المطلوب ثم reg_time كتعزيز
    d["_sort"] = d[date_col].fillna(d.get("reg_time"))
    d = d.sort_values(["_KEY_PREMISE", "_sort", "reg_time"], na_position="last")
    idx = d.groupby("_KEY_PREMISE").tail(1).index
    return d.loc[idx].drop(columns=["_sort"], errors="ignore")

def first_by_key(df: pd.DataFrame, date_col: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["_KEY_PREMISE"])
    d = df.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d = d.sort_values(["_KEY_PREMISE", date_col])
    idx = d.groupby("_KEY_PREMISE").head(1).index
    return d.loc[idx]

# أحدث فحص/أول صيانة/آخر صيانة
insp_latest = latest_by_key(insp_df, "close_time" if "close_time" in insp_df.columns else "reg_time")
mnt_first   = first_by_key(mnt_df, "reg_time") if not mnt_df.empty else pd.DataFrame(columns=["_KEY_PREMISE"])
mnt_latest  = latest_by_key(mnt_df, "close_time" if "close_time" in mnt_df.columns else "reg_time")

# أعلام الإغلاق للصيانة
def is_closed_status(s: pd.Series, closed_terms: set[str]) -> pd.Series:
    return s.astype(str).str.strip().str.lower().isin(closed_terms)

mnt_latest = mnt_latest.copy()
if not mnt_latest.empty:
    mnt_latest["_closed"] = mnt_latest["close_time"].notna() | is_closed_status(mnt_latest["status"], CLOSED_TERMS)

# ================= Join with disconnected =================
base = dis_df[["_KEY_PREMISE", "LastDaily"]].copy()
base = base.merge(insp_latest[["_KEY_PREMISE","reg_time","close_time","status","result"]]
                  .rename(columns={"reg_time":"insp_reg","close_time":"insp_close","status":"insp_status","result":"insp_result"}),
                  on="_KEY_PREMISE", how="left")
base = base.merge(mnt_first[["_KEY_PREMISE","reg_time"]].rename(columns={"reg_time":"mnt_first_reg"}),
                  on="_KEY_PREMISE", how="left")
base = base.merge(mnt_latest[["_KEY_PREMISE","reg_time","close_time","status","result","_closed"]]
                  .rename(columns={"reg_time":"mnt_last_reg","close_time":"mnt_last_close","status":"mnt_last_status","result":"mnt_last_result","_closed":"mnt_closed"}),
                  on="_KEY_PREMISE", how="left")

# مؤقتات
base["days_from_insp_to_mnt"] = (base["mnt_first_reg"] - base["insp_close"]).dt.days
base["insp_done"]   = base["insp_close"].notna()
base["has_mnt"]     = base["mnt_first_reg"].notna()
base["mnt_open"]    = base["has_mnt"] & ~base["mnt_closed"].fillna(False)
base["mnt_closed"]  = base["mnt_closed"].fillna(False)

# ================= Inspection-focused KPIs =================
st.markdown("## 🧰 مؤشرات جهود الفحص")
insp_count = int(base["insp_done"].sum())
insp_rate  = 100.0 * (base["insp_done"].sum() / len(base)) if len(base) else 0.0
insp_dur   = (insp_df["close_time"] - insp_df["reg_time"]).dt.days.dropna()
avg_insp_days = float(insp_dur.mean()) if not insp_dur.empty else 0.0

k1, k2, k3 = st.columns(3)
k1.metric("عدادات غير متصلة تم فحصها", f"{insp_count:,}")
k2.metric("نسبة الفحص من غير المتصلة", f"{insp_rate:.1f}%")
k3.metric("متوسط مدة الفحص (أيام)", f"{avg_insp_days:.1f}")

# ================= Reports =================
st.markdown("## 📋 تقارير تشغيلية تُبرز مسؤولية الصيانة بعد الفحص")

# 1) فُحصت ولا يوجد صيانة
r1 = base[(base["insp_done"]) & (~base["has_mnt"])].copy()
r1 = r1.sort_values(["insp_close"], ascending=[False])

# 2) صيانة مُقفلة وما زال غير متصل (العميل ما زال ضمن قائمة غير المتصلين = دليل عدم المعالجة)
r2 = base[(base["insp_done"]) & (base["mnt_closed"])].copy()
# (وجوده هنا في ملف غير المتصلين يكفي للدلالة أنه ما عاد يتصل)

# 3) صيانة مفتوحة بعد الفحص (للضغط على الصيانة)
r3 = base[(base["insp_done"]) & (base["mnt_open"])].copy()

# أعمدة عرض مفضلة
cols = ["_KEY_PREMISE", "LastDaily",
        "insp_reg","insp_close","insp_status","insp_result",
        "mnt_first_reg","mnt_last_reg","mnt_last_close","mnt_last_status","mnt_last_result",
        "days_from_insp_to_mnt"]

def show_table_download(title, df, fname):
    st.markdown(f"### {title}  \n({len(df):,}) سجل")
    st.dataframe(df, use_container_width=True)
    st.download_button(
        f"⬇️ تنزيل ({fname}.xlsx)",
        data=to_excel_download(df),
        file_name=f"{fname}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown("---")

show_table_download("1) مفحوصة ولا يوجد لها صيانة", r1[cols], "inspected_no_maintenance")
show_table_download("2) صيانة مُقفلة وما زال غير متصل", r2[cols], "maintenance_closed_still_disconnected")
show_table_download("3) صيانة مفتوحة بعد الفحص", r3[cols], "maintenance_open_post_inspection")

# ================= Optional quick counts =================
st.markdown("## 🔎 ملخص سريع")
c1, c2, c3 = st.columns(3)
c1.metric("مفحوصة بدون صيانة", len(r1))
c2.metric("صيانة مُقفلة — ما زال غير متصل", len(r2))
c3.metric("صيانة مفتوحة بعد الفحص", len(r3))

st.caption("وجود Premise في ملف غير المتصلين يعني أنه ما زال غير متصل حتى لحظة إنشاء التقرير.")
